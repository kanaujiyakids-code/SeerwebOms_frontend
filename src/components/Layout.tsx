import React from "react";
import { useAuth } from "@/context/AuthContext";
import { Navigate } from "react-router-dom";
import "../assets/css/style.css";

interface LayoutProps {
  children: React.ReactNode;
  requireAuth?: boolean;
  requiredRole?: "dealer" | "retailer";
}

const Layout: React.FC<LayoutProps> = ({
  children,
  requireAuth = false,
  requiredRole,
}) => {
  const { isAuthenticated, user } = useAuth();

  if (requireAuth && !isAuthenticated) {
    return <Navigate to="/login" replace />;
  }

  if (
    requireAuth &&
    isAuthenticated &&
    requiredRole &&
    user?.role !== requiredRole
  ) {
    return <Navigate to="/" replace />;
  }

  return (
    <div className="relative isolate min-h-screen w-full overflow-x-hidden">
      <div
        aria-hidden="true"
        className="fixed inset-0 -z-20 scale-105 bg-[url('/images/background.png')] bg-cover bg-center bg-no-repeat blur-[2px]"
      />
      <div
        aria-hidden="true"
        className="fixed inset-0 -z-10 bg-gradient-to-br from-slate-50/90 via-white/80 to-blue-50/85 backdrop-blur-[1px]"
      />

      <div className="relative min-h-screen">{children}</div>
    </div>
  );
};

export default Layout;
