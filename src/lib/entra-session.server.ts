import {
  createRemoteJWKSet,
  jwtVerify,
} from "jose";

export type EntraExchangeResult = {
  email: string;
  accessToken: string;
  refreshToken: string;
};

const jwksCache = new Map<
  string,
  ReturnType<typeof createRemoteJWKSet>
>();

function getJwks(tenantId: string) {
  let jwks = jwksCache.get(tenantId);

  if (!jwks) {
    jwks = createRemoteJWKSet(
      new URL(
        `https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`,
      ),
    );

    jwksCache.set(tenantId, jwks);
  }

  return jwks;
}

export async function exchangeEntraIdToken(
  idToken: string,
): Promise<EntraExchangeResult> {
  const { supabaseAdmin } = await import(
    "@/integrations/supabase/client.server"
  );

  /* ---------------------------------------------------------------------- */
  /* CONFIGURAÇÃO                                                           */
  /* ---------------------------------------------------------------------- */

  const { data: settings, error: settingsError } =
    await supabaseAdmin
      .from("app_settings")
      .select("key, value")
      .in("key", ["azure_client_id", "azure_tenant_id"]);

  if (settingsError) {
    console.error(
      "[ENTRA SERVER] Erro ao buscar configuração:",
      settingsError,
    );

    throw new Error(
      "Configuração do Microsoft Entra indisponível.",
    );
  }

  const map = Object.fromEntries(
    (settings ?? []).map((setting) => [
      setting.key,
      setting.value,
    ]),
  );

  const clientId = String(
    map["azure_client_id"] ?? "",
  ).trim();

  const tenantId = String(
    map["azure_tenant_id"] ?? "",
  ).trim();

  if (!clientId || !tenantId) {
    throw new Error(
      "Microsoft Entra ID não está configurado no portal.",
    );
  }

  /* ---------------------------------------------------------------------- */
  /* VALIDAR TOKEN                                                          */
  /* ---------------------------------------------------------------------- */

  let payload;

  try {
    const result = await jwtVerify(
      idToken,
      getJwks(tenantId),
      {
        audience: clientId,
        issuer: [
          `https://login.microsoftonline.com/${tenantId}/v2.0`,
          `https://sts.windows.net/${tenantId}/`,
        ],
      },
    );

    payload = result.payload;
  } catch (error) {
    console.error(
      "[ENTRA SERVER] Token Microsoft inválido:",
      error,
    );

    throw new Error(
      "O token do Microsoft Entra ID é inválido ou expirou.",
    );
  }

  /* ---------------------------------------------------------------------- */
  /* VALIDAR TENANT                                                         */
  /* ---------------------------------------------------------------------- */

  const tokenTenantId =
    typeof payload["tid"] === "string"
      ? payload["tid"]
      : "";

  if (!tokenTenantId || tokenTenantId !== tenantId) {
    throw new Error(
      "A conta Microsoft pertence a um tenant não autorizado.",
    );
  }

  /* ---------------------------------------------------------------------- */
  /* OBTER IDENTIDADE                                                       */
  /* ---------------------------------------------------------------------- */

  const email = String(
    payload["email"] ??
      payload["preferred_username"] ??
      payload["upn"] ??
      "",
  )
    .trim()
    .toLowerCase();

  if (!email || !email.includes("@")) {
    throw new Error(
      "A conta Microsoft não possui um e-mail válido.",
    );
  }

  const fullName =
    typeof payload["name"] === "string" &&
    payload["name"].trim()
      ? payload["name"].trim()
      : email.split("@")[0];

  console.log(
    "[ENTRA SERVER] Identidade validada:",
    { email, tenantId },
  );

  /* ---------------------------------------------------------------------- */
  /* PROCURAR / CRIAR USUÁRIO SUPABASE                                     */
  /* ---------------------------------------------------------------------- */

  const { data: usersData, error: usersError } =
    await supabaseAdmin.auth.admin.listUsers({
      page: 1,
      perPage: 1000,
    });

  if (usersError) {
    console.error(
      "[ENTRA SERVER] Erro ao consultar usuários:",
      usersError,
    );

    throw new Error(
      "Não foi possível verificar o usuário do portal.",
    );
  }

  const existingUser = usersData.users.find(
    (user) =>
      user.email?.toLowerCase() === email,
  );

  if (existingUser) {
    console.log(
      "[ENTRA SERVER] Usuário Supabase existente:",
      existingUser.id,
    );
  } else {
    const { data: createdUser, error: createError } =
      await supabaseAdmin.auth.admin.createUser({
        email,
        email_confirm: true,
        user_metadata: {
          full_name: fullName,
          provider: "entra",
        },
      });

    if (createError || !createdUser.user) {
      console.error(
        "[ENTRA SERVER] Erro ao criar usuário:",
        createError,
      );

      throw (
        createError ??
        new Error(
          "Não foi possível criar o usuário do portal.",
        )
      );
    }

    console.log(
      "[ENTRA SERVER] Usuário Supabase criado:",
      createdUser.user.id,
    );
  }

  /* ---------------------------------------------------------------------- */
  /* CRIAR A SESSÃO NO SERVIDOR                                             */
  /* ---------------------------------------------------------------------- */

  /*
   * O fluxo anterior devolvia o hashed_token para o browser e fazia
   * verifyOtp() no client. Isso deixou a criação da sessão dependente de
   * uma segunda chamada ao endpoint /auth/v1/verify e foi justamente onde
   * o portal estava recebendo 400 validation_failed.
   *
   * Agora o token é gerado e consumido no servidor. O servidor devolve
   * somente os tokens de sessão resultantes e o browser usa setSession().
   * Assim, o token_hash nunca precisa trafegar até o client nem ser
   * verificado duas vezes.
   */

  const { data: link, error: linkError } =
    await supabaseAdmin.auth.admin.generateLink({
      type: "magiclink",
      email,
    });

  if (
    linkError ||
    !link?.properties?.hashed_token
  ) {
    console.error(
      "[ENTRA SERVER] Erro ao gerar token de sessão:",
      linkError,
    );

    throw (
      linkError ??
      new Error(
        "Não foi possível gerar a sessão do portal.",
      )
    );
  }

  const { data: verifyData, error: verifyError } =
    await supabaseAdmin.auth.verifyOtp({
      token_hash: link.properties.hashed_token,
      type: "magiclink",
    });

  if (
    verifyError ||
    !verifyData.session?.access_token ||
    !verifyData.session?.refresh_token
  ) {
    console.error(
      "[ENTRA SERVER] Erro ao consumir token e criar sessão:",
      verifyError,
    );

    throw (
      verifyError ??
      new Error(
        "O Supabase não retornou uma sessão válida para o usuário.",
      )
    );
  }

  console.log(
    "[ENTRA SERVER] Sessão Supabase criada no servidor.",
  );

  return {
    email,
    accessToken:
      verifyData.session.access_token,
    refreshToken:
      verifyData.session.refresh_token,
  };
}
