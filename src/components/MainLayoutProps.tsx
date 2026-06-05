import { useEffect, useState } from "react";
import Sidebar from "@/components/layout/Sidebar";
import Navbar from "@/components/layout/Navbar";

interface MainLayoutProps {
  children: React.ReactNode;
}

const MainLayout = ({ children }: MainLayoutProps) => {

  // Load sidebar state from localStorage
  const [collapsed, setCollapsed] = useState<boolean>(() => {
    return localStorage.getItem("sidebar-collapsed") === "true";
  });

  // Save sidebar state whenever it changes
  useEffect(() => {
    localStorage.setItem(
      "sidebar-collapsed",
      String(collapsed)
    );
  }, [collapsed]);

  return (
    <div className="relative isolate min-h-screen overflow-x-hidden">
      <div
        aria-hidden="true"
        className="fixed inset-0 -z-20 scale-105 bg-[url('/images/background.png')] bg-cover bg-center bg-no-repeat blur-[2px]"
      />
      <div
        aria-hidden="true"
        className="fixed inset-0 -z-10 bg-gradient-to-br from-slate-50/90 via-white/80 to-blue-50/85 backdrop-blur-[1px]"
      />

      {/* Sidebar */}
      <Sidebar
        collapsed={collapsed}
        setCollapsed={setCollapsed}
      />

      {/* Navbar */}
      <Navbar collapsed={collapsed} />

      {/* Main Content */}
      <main
        className={`relative z-10 pt-20 transition-all duration-300 ${
          collapsed
            ? "lg:ml-20"
            : "lg:ml-64"
        }`}
      >
        <div className="p-6">
          {children}
        </div>
      </main>

    </div>
  );
};

export default MainLayout;
