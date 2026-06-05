import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import {
  Card,
  CardContent,
  CardFooter,
  CardHeader,
} from "@/components/ui/card";
import { useAuth } from "@/context/AuthContext";
import { toast } from "sonner";

// Import icons
import { Eye, EyeOff, ShieldCheck } from "lucide-react";

const LoginForm = () => {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const { login } = useAuth();

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();

    if (!username || !password) {
      toast.error("Please enter both username and password");
      return;
    }

    setIsLoading(true);

    try {
      const success = await login(username, password);

      if (!success) {
        toast.error("Invalid username or password");
      }
    } catch (error) {
      console.error("Login error:", error);
      toast.error("An error occurred during login");
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <Card className="overflow-hidden rounded-3xl border border-white/25 bg-white/18 text-white shadow-2xl shadow-slate-950/35 backdrop-blur-2xl">
      <div className="h-px bg-gradient-to-r from-transparent via-white/70 to-transparent" />
      <CardHeader className="space-y-5 p-6 pb-4 sm:p-8 sm:pb-5">
        <div className="flex items-center gap-2 rounded-full border border-white/25 bg-white/15 px-4 py-2 text-sm font-medium text-white shadow-inner shadow-white/10 backdrop-blur-xl">
          <ShieldCheck className="h-4 w-4 shrink-0" />
          <span>Secure login for authorized users</span>
        </div>
        <div className="space-y-2">
          <p className="text-xs font-semibold uppercase text-blue-100/80">
            Welcome back
          </p>
          <h2 className="text-2xl font-semibold leading-tight text-white sm:text-3xl">
            Login to your workspace
          </h2>
          <p className="text-sm leading-6 text-slate-100/75">
            Enter your credentials to continue managing orders, accounts, and
            business operations.
          </p>
        </div>
      </CardHeader>

      <CardContent className="p-6 pt-0 sm:px-8">
        <form onSubmit={handleSubmit} className="space-y-5">
          <div className="space-y-2">
            <label
              htmlFor="username"
              className="text-sm font-semibold uppercase text-blue-50/85"
            >
              Username
            </label>
            <Input
              id="username"
              placeholder="Enter your username"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              disabled={isLoading}
              className="h-10 rounded-xl border-white/25 bg-white/18 px-4 text-black shadow-inner shadow-white/5 placeholder:text-slate-100/55 focus-visible:border-white/50 focus-visible:ring-2 focus-visible:ring-white/35"
            />
          </div>

          <div className="space-y-2">
            <label
              htmlFor="password"
              className="text-sm font-semibold uppercase text-blue-50/85"
            >
              Password
            </label>
            <div className="relative">
              <Input
                id="password"
                type={showPassword ? "text" : "password"}
                placeholder="Enter your password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                disabled={isLoading}
                className="h-10 rounded-xl border-white/25 bg-white/18 px-4 pr-11 text-black shadow-inner shadow-white/5 placeholder:text-slate-100/55 focus-visible:border-white/50 focus-visible:ring-2 focus-visible:ring-white/35"
              />
              <button
                type="button"
                onClick={() => setShowPassword((prev) => !prev)}
                className="absolute inset-y-0 right-3 flex items-center text-slate-100/65 transition-colors hover:text-white"
                tabIndex={-1}
                aria-label={showPassword ? "Hide password" : "Show password"}
              >
                {showPassword ? <EyeOff size={18} /> : <Eye size={18} />}
              </button>
            </div>
          </div>

          <Button
            type="submit"
            disabled={isLoading}
            className="h-12 w-full rounded-2xl border border-white/20 bg-gradient-to-r from-royal via-blue-500 to-cyan-500 text-sm font-semibold text-white shadow-xl shadow-blue-950/25 transition hover:from-royal-dark hover:via-blue-600 hover:to-cyan-600 focus-visible:ring-2 focus-visible:ring-white/45 disabled:opacity-70"
          >
            {isLoading ? "Signing in..." : "Login to Seerweb OMS"}
          </Button>
        </form>
      </CardContent>

     
    </Card>
  );
};

export default LoginForm;
