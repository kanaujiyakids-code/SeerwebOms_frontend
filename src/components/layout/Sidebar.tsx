import { useState } from "react";
import { Link, useNavigate, useLocation } from "react-router-dom";
import { useAuth } from "@/context/AuthContext";
import logo from "@/assets/images/react-logo.png";
import {
  LayoutDashboard,
  ShoppingCart,
  Package,
  Users,
  UserCog,
  ClipboardList,
  BarChart3,
  LogOut,
  Menu,
  X,
  User,
  Box,
  PanelLeftClose,
  PanelLeftOpen,
} from "lucide-react";

import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from "@/components/ui/alert-dialog";

interface SidebarProps {
  collapsed: boolean;
  setCollapsed: React.Dispatch<React.SetStateAction<boolean>>;
}

const Sidebar = ({
  collapsed,
  setCollapsed,
}: SidebarProps) => {
  const { user, logout } = useAuth();

  const [mobileOpen, setMobileOpen] = useState(false);
  const [logoutDialogOpen, setLogoutDialogOpen] = useState(false);

  const navigate = useNavigate();
  const location = useLocation();

  const handleLogout = () => {
    setLogoutDialogOpen(false);
    logout();
    navigate("/");
  };

  const isActive = (path: string) =>
    location.pathname === path;

  const linkClasses = (path: string) =>
    `flex items-center ${
      collapsed ? "justify-center" : "gap-3"
    } rounded-lg px-4 py-3 text-sm font-medium transition-all duration-200 ${
      isActive(path)
        ? "bg-white/10 text-white"
        : "text-gray-300 hover:bg-white/5 hover:text-white"
    }`;

  const renderLinks = () => {
    if (user?.role === "admin") {
      return (
        <>
          <Link to="/admin" className={linkClasses("/admin")}>
            <LayoutDashboard className="h-5 w-5 shrink-0" />
            {!collapsed && <span>Dashboard</span>}
          </Link>

          <Link
            to="/admin/dealers"
            className={linkClasses("/admin/dealers")}
          >
            <Users className="h-5 w-5 shrink-0" />
            {!collapsed && <span>Manage Dealers</span>}
          </Link>
        </>
      );
    }

    if (user?.role === "dealer") {
      return (
        <>
          <Link to="/dealer" className={linkClasses("/dealer")}>
            <LayoutDashboard className="h-5 w-5 shrink-0" />
            {!collapsed && <span>Dashboard</span>}
          </Link>

          <Link
            to="/dealer/orders"
            className={linkClasses("/dealer/orders")}
          >
            <ShoppingCart className="h-5 w-5 shrink-0" />
            {!collapsed && <span>Orders</span>}
          </Link>

          <Link
            to="/dealer/retailers"
            className={linkClasses("/dealer/retailers")}
          >
            <Users className="h-5 w-5 shrink-0" />
            {!collapsed && <span>Customers</span>}
          </Link>

          <Link
            to="/dealer/products"
            className={linkClasses("/dealer/products")}
          >
            <Package className="h-5 w-5 shrink-0" />
            {!collapsed && <span>Products</span>}
          </Link>

          <Link
            to="/dealer/staff"
            className={linkClasses("/dealer/staff")}
          >
            <UserCog className="h-5 w-5 shrink-0" />
            {!collapsed && <span>Staff</span>}
          </Link>

          <Link
            to="/dealer/takeorder"
            className={linkClasses("/dealer/takeorder")}
          >
            <ClipboardList className="h-5 w-5 shrink-0" />
            {!collapsed && <span>Create Order</span>}
          </Link>
        </>
      );
    }

    if (user?.role === "retailer") {
      return (
        <>
          <Link
            to="/retailer/dashboard"
            className={linkClasses("/retailer/dashboard")}
          >
            <LayoutDashboard className="h-5 w-5 shrink-0" />
            {!collapsed && <span>Dashboard</span>}
          </Link>

          <Link
            to="/retailer/products"
            className={linkClasses("/retailer/products")}
          >
            <Box className="h-5 w-5 shrink-0" />
            {!collapsed && <span>Products</span>}
          </Link>

          <Link
            to="/retailer/orders"
            className={linkClasses("/retailer/orders")}
          >
            <ShoppingCart className="h-5 w-5 shrink-0" />
            {!collapsed && <span>My Orders</span>}
          </Link>

          <Link
            to="/retailer/profile"
            className={linkClasses("/retailer/profile")}
          >
            <User className="h-5 w-5 shrink-0" />
            {!collapsed && <span>Profile</span>}
          </Link>
        </>
      );
    }

    if (user?.role === "staff") {
      return (
        <>
          <Link
            to="/staff"
            className={linkClasses("/staff")}
          >
            <LayoutDashboard className="h-5 w-5 shrink-0" />
            {!collapsed && <span>Dashboard</span>}
          </Link>

          <Link
            to="/staff/take_order"
            className={linkClasses("/staff/take_order")}
          >
            <ClipboardList className="h-5 w-5 shrink-0" />
            {!collapsed && <span>Create Order</span>}
          </Link>

          <Link
            to="/staff/sales_report"
            className={linkClasses("/staff/sales_report")}
          >
            <BarChart3 className="h-5 w-5 shrink-0" />
            {!collapsed && <span>Sales Report</span>}
          </Link>
        </>
      );
    }

    return null;
  };

  return (
    <>
      {/* MOBILE TOGGLE */}
      <button
        onClick={() => setMobileOpen(!mobileOpen)}
        className="fixed left-4 top-4 z-50 rounded-md bg-gray-800 p-2 text-white lg:hidden"
      >
        {mobileOpen ? (
          <X className="h-5 w-5" />
        ) : (
          <Menu className="h-5 w-5" />
        )}
      </button>

      {/* MOBILE OVERLAY */}
      {mobileOpen && (
        <div
          className="fixed inset-0 z-30 bg-black/50 lg:hidden"
          onClick={() => setMobileOpen(false)}
        />
      )}

      {/* SIDEBAR */}
      <aside
        className={`fixed left-0 top-0 z-40 flex h-screen flex-col
        bg-gradient-to-b from-[#1f2937] to-[#111827]
        border-r border-white/10
        transition-all duration-300
        ${collapsed ? "w-20" : "w-64"}
        ${
          mobileOpen
            ? "translate-x-0"
            : "-translate-x-full lg:translate-x-0"
        }`}
      >
        {/* HEADER */}
        <div className="flex h-16 items-center justify-between border-b border-white/10 px-4">
          <div
            className={`flex items-center ${
              collapsed
                ? "justify-center w-full"
                : "gap-3"
            } overflow-hidden`}
          >
            <div className="flex h-12 w-12 shrink-0 items-center justify-center rounded-xl bg-white">
              <img src={logo} alt="" />
            </div>

            {!collapsed && (
              <div>
                <h1 className="text-sm font-semibold text-white">
                  SeerWeb OMS
                </h1>

                <p className="text-xs capitalize text-gray-400">
                  {user?.role} Panel
                </p>
              </div>
            )}
          </div>

          {!collapsed && (
            <button
              onClick={() => setCollapsed(true)}
              className="hidden rounded-md p-1 text-white transition hover:bg-white/10 lg:flex"
            >
              <PanelLeftClose className="h-5 w-5" />
            </button>
          )}
        </div>

        {/* COLLAPSED BUTTON */}
        {collapsed && (
          <div className="mt-2 flex justify-center">
            <button
              onClick={() => setCollapsed(false)}
              className="hidden rounded-md p-2 text-white transition hover:bg-white/10 lg:flex"
            >
              <PanelLeftOpen className="h-5 w-5" />
            </button>
          </div>
        )}

        {/* NAVIGATION */}
        <nav className="flex-1 space-y-2 overflow-y-auto px-3 py-4">
          {renderLinks()}
        </nav>

        {/* USER SECTION */}
        <div className="border-t border-white/10 p-4">
          <div
            className={`flex items-center ${
              collapsed
                ? "justify-center"
                : "gap-3"
            }`}
          >
            <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-full bg-white/10 font-semibold text-white">
              {user?.name?.charAt(0)?.toUpperCase() || "U"}
            </div>

            {!collapsed && (
              <>
                <div className="min-w-0 flex-1">
                  <p className="truncate text-sm font-medium text-white">
                    {user?.name}
                  </p>

                  <p className="truncate text-xs text-gray-400">
                    {user?.email}
                  </p>
                </div>

                <button
                  onClick={() => setLogoutDialogOpen(true)}
                  className="text-gray-400 transition hover:text-white"
                >
                  <LogOut className="h-5 w-5" />
                </button>
              </>
            )}
          </div>
        </div>
      </aside>

      {/* LOGOUT DIALOG */}
      <AlertDialog
        open={logoutDialogOpen}
        onOpenChange={setLogoutDialogOpen}
      >
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>
              Confirm Logout
            </AlertDialogTitle>

            <AlertDialogDescription>
              Are you sure you want to logout?
            </AlertDialogDescription>
          </AlertDialogHeader>

          <AlertDialogFooter>
            <AlertDialogCancel>
              Cancel
            </AlertDialogCancel>

            <AlertDialogAction onClick={handleLogout}>
              Yes, Logout
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </>
  );
};

export default Sidebar;