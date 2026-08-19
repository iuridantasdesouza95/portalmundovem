import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";
import { exchangeEntraIdToken } from "./entra-session.server";

export const entraSignIn = createServerFn({ method: "POST" })
  .inputValidator((data) => z.object({ idToken: z.string().min(20) }).parse(data))
  .handler(async ({ data }) => exchangeEntraIdToken(data.idToken));
