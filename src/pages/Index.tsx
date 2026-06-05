import { useEffect } from "react";
import { useNavigate } from "react-router-dom";
import LoginForm from "@/components/LoginForm";
import { useAuth } from "@/context/AuthContext";
import logo from "../assets/images/react-logo.png";

const Index = () => {
  const { user, isAuthenticated } = useAuth();
  const navigate = useNavigate();

  // Redirect if already logged in
  useEffect(() => {
    if (isAuthenticated) {
      if (user?.role === "admin") {
        navigate("/admin");
      } else if (user?.role === "dealer") {
        navigate("/dealer");
      } else {
        navigate("/retailer/dashboard");
      }
    }
  }, [isAuthenticated, user?.role, navigate]);

  return (
    <main className="relative min-h-screen overflow-hidden bg-slate-950 text-white">
      <img
        src="/images/login-business-owner.jpg"
        alt=""
        aria-hidden="true"
        className="absolute inset-0 h-full w-full scale-105 object-cover object-center blur-sm"
      />
      <div className="absolute inset-0 bg-slate-950/45" />
      <div className="absolute inset-0 bg-[linear-gradient(105deg,rgba(255, 255, 255, 0.82)_0%,rgba(15,23,42,0.54)_42%,rgba(15,23,42,0.2)_100%)]" />
      <div className="absolute inset-x-0 bottom-0 h-1/2 bg-gradient-to-t from-slate-950/75 to-transparent" />

      <div className="relative mx-auto flex min-h-screen w-full max-w-7xl items-center px-4 py-8 sm:px-6 lg:px-8">
        <div className="grid w-full items-center gap-8 lg:grid-cols-[1.05fr_0.95fr] lg:gap-14">
          <section className="max-w-2xl rounded-3xl border border-white/15 bg-white/10 p-5 shadow-2xl shadow-black/20 backdrop-blur-md sm:p-7 lg:bg-white/[0.08]">
            <div className="flex flex-col gap-5 sm:flex-row sm:items-center">
              <div className="flex h-20 w-20 shrink-0 items-center justify-center rounded-2xl border border-white/30 bg-white/90 p-0.5 shadow-xl shadow-black/20">
                <img
                  src={logo}
                  alt="Seerweb Logo"
                  className="h-full w-full object-contain"
                />
              </div>
              <div>
                <p className="text-xs font-semibold uppercase text-blue-100/80">
                  Seerweb OMS
                </p>
                <h1 className="mt-2 max-w-xl text-3xl font-semibold leading-tight text-white sm:text-4xl lg:text-5xl">
                  Seerweb ERP Solutions Pvt Ltd
                </h1>
                <p className="mt-3 max-w-xl text-sm leading-6 text-slate-100/85 sm:text-base">
                  Secure access to business operations, Tally services, ERP
                  workflows, and order management built for growing teams.
                </p>
              </div>
            </div>
          </section>

          <aside className="flex justify-center lg:justify-end">
            <div className="w-full max-w-md">
              <LoginForm />
            </div>
          </aside>
        </div>
      </div>
    </main>
  );
};

export default Index;
