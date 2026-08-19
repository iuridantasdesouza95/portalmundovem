import { createRemoteJWKSet, jwtVerify } from "jose";

/**
 * Troca um id_token do Microsoft Entra ID (obtido no browser pelo MSAL, usando o
 * App Registration do portal) por uma sessão do portal.
 *
 * O token é validado no servidor (assinatura via JWKS do tenant, issuer, audience,
 * expiração) antes de qualquer operação privilegiada.
 */
export type EntraExchangeResult = { email: string; tokenHash: string };

const jwksCache = new Map<string, ReturnType<typeof createRemoteJWKSet>>();

function getJwks(tenantId: string) {
  let jwks = jwksCache.get(tenantId);
  if (!jwks) {
    jwks = createRemoteJWKSet(
      new URL(`https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`),
    );
    jwksCache.set(tenantId, jwks);
  }
  return jwks;
}

export async function exchangeEntraIdToken(idToken: string): Promise<EntraExchangeResult> {
  const { supabaseAdmin } = await import("@/integrations/supabase/client.server");

  const { data: settings, error: settingsError } = await supabaseAdmin
    .from("app_settings")
    .select("key, value")
    .in("key", ["azure_client_id", "azure_tenant_id"]);
  if (settingsError) throw new Error("Configuração do Entra indisponível");

  const map = Object.fromEntries((settings ?? []).map((s) => [s.key, s.value]));
  const clientId = (map["azure_client_id"] ?? "").trim();
  const tenantId = (map["azure_tenant_id"] ?? "").trim();
  if (!clientId || !tenantId) throw new Error("Entra ID não configurado no portal");

  const { payload } = await jwtVerify(idToken, getJwks(tenantId), {
    audience: clientId,
    issuer: [
      `https://login.microsoftonline.com/${tenantId}/v2.0`,
      `https://sts.windows.net/${tenantId}/`,
    ],
  });

  if (payload["tid"] !== tenantId) throw new Error("Tenant não autorizado");

  const email = String(
    payload["email"] ?? payload["preferred_username"] ?? payload["upn"] ?? "",
  ).toLowerCase();
  if (!email.includes("@")) throw new Error("Conta Microsoft sem e-mail corporativo");

  const fullName = String(payload["name"] ?? email.split("@")[0]);

  // Cria o usuário na primeira entrada (sem senha; identidade é o Entra ID).
  const { data: link, error: linkError } = await supabaseAdmin.auth.admin.generateLink({
    type: "magiclink",
    email,
  });

  if (linkError || !link?.properties?.hashed_token) {
    const { error: createError } = await supabaseAdmin.auth.admin.createUser({
      email,
      email_confirm: true,
      user_metadata: { full_name: fullName, provider: "entra" },
    });
    if (createError && !/already/i.test(createError.message)) throw createError;

    const retry = await supabaseAdmin.auth.admin.generateLink({ type: "magiclink", email });
    if (retry.error || !retry.data?.properties?.hashed_token) {
      throw retry.error ?? new Error("Não foi possível criar a sessão do portal");
    }
    return { email, tokenHash: retry.data.properties.hashed_token };
  }

  return { email, tokenHash: link.properties.hashed_token };
}
