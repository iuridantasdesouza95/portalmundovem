import { createFileRoute } from "@tanstack/react-router";

/**
 * Página técnica usada como redirectUri pelos popups do Microsoft Entra ID.
 *
 * IMPORTANTE:
 * - Não contém lógica de autenticação do portal.
 * - Não navega para /inicio.
 * - Não executa login.
 * - Não usa Supabase.
 *
 * O MSAL usa esta página apenas para concluir a comunicação entre a janela
 * popup e a janela principal. Manter a página mínima evita que o popup tente
 * hidratar a tela completa do portal e elimina o React #418 no retorno.
 */
export const Route = createFileRoute("/auth/popup")({
  ssr: false,
  component: AuthPopupPage,
});

function AuthPopupPage() {
  return <div aria-hidden="true" />;
}
