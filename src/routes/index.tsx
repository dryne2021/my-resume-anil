import { createFileRoute, redirect } from "@tanstack/react-router";
import { getSession } from "@/lib/auth.functions";

export const Route = createFileRoute("/")({
  beforeLoad: async () => {
    const { loggedIn } = await getSession();
    throw redirect({ to: loggedIn ? "/dashboard" : "/login" });
  },
});
