import { createFileRoute, Outlet, redirect } from "@tanstack/react-router";
import { getSession } from "@/lib/auth.functions";

export const Route = createFileRoute("/_authenticated")({
  beforeLoad: async () => {
    const { loggedIn } = await getSession();
    if (!loggedIn) throw redirect({ to: "/login" });
  },
  component: () => <Outlet />,
});
