import { createFileRoute, redirect, useNavigate } from "@tanstack/react-router";
import { useState, type FormEvent } from "react";
import { useServerFn } from "@tanstack/react-start";
import { getSession, login } from "@/lib/auth.functions";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { toast } from "sonner";

export const Route = createFileRoute("/login")({
  head: () => ({
    meta: [
      { title: "Sign in — Dryne Agency" },
      { name: "description", content: "Private portal access." },
      { name: "robots", content: "noindex,nofollow" },
    ],
  }),
  beforeLoad: async () => {
    const { loggedIn } = await getSession();
    if (loggedIn) throw redirect({ to: "/dashboard" });
  },
  component: LoginPage,
});

function LoginPage() {
  const navigate = useNavigate();
  const doLogin = useServerFn(login);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [submitting, setSubmitting] = useState(false);

  const onSubmit = async (e: FormEvent) => {
    e.preventDefault();
    setSubmitting(true);
    try {
      const res = await doLogin({ data: { email, password } });
      if (!res.ok) {
        toast.error("Invalid credentials");
        return;
      }
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Sign in failed");
      return;
    } finally {
      setSubmitting(false);
    }
    navigate({ to: "/dashboard" });
  };

  return (
    <div className="min-h-screen grain flex items-center justify-center px-6">
      <div className="w-full max-w-md">
        <div className="mb-10 text-center">
          <p className="text-xs uppercase tracking-[0.3em] text-muted-foreground">Private Portal</p>
          <h1 className="mt-3 text-5xl font-display text-gold">Dryne Agency</h1>
          <p className="mt-2 text-sm text-muted-foreground italic font-display">
            Resume Atelier
          </p>
        </div>

        <form
          onSubmit={onSubmit}
          className="rounded-xl border border-border bg-card/60 backdrop-blur p-8 space-y-5 shadow-2xl"
        >
          <div className="space-y-2">
            <Label htmlFor="email" className="text-xs uppercase tracking-wider text-muted-foreground">
              Email
            </Label>
            <Input
              id="email"
              type="email"
              required
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              autoComplete="email"
              className="bg-background/60"
            />
          </div>
          <div className="space-y-2">
            <Label htmlFor="password" className="text-xs uppercase tracking-wider text-muted-foreground">
              Password
            </Label>
            <Input
              id="password"
              type="password"
              required
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              autoComplete="current-password"
              className="bg-background/60"
            />
          </div>
          <Button
            type="submit"
            disabled={submitting}
            className="w-full bg-gold-gradient text-primary-foreground hover:opacity-90 font-medium"
          >
            {submitting ? "Signing in…" : "Enter Portal"}
          </Button>
        </form>

        <p className="mt-6 text-center text-xs text-muted-foreground">
          Access is restricted to a single authorized user.
        </p>
      </div>
    </div>
  );
}
