import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";

import { exchangeEntraIdToken } from "./entra-session.server";

const entraSignInSchema = z.object({
  idToken: z
    .string()
    .min(20, "ID Token Microsoft inválido."),
});

export const entraSignIn = createServerFn({
  method: "POST",
})
  .inputValidator((data) =>
    entraSignInSchema.parse(data),
  )
  .handler(async ({ data }) => {
    return exchangeEntraIdToken(data.idToken);
  });
