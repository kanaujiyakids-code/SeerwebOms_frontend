const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Icon imports
const { FaShoppingCart, FaUsers, FaChartLine, FaBoxOpen, FaMicrophone,
        FaCheckCircle, FaLock, FaMobileAlt, FaCloud, FaRocket,
        FaUserTie, FaStore, FaFileExport, FaBell, FaSearch,
        FaWhatsapp, FaTruck, FaLayerGroup } = require("react-icons/fa");
const { MdDashboard, MdInventory, MdReceipt, MdSupportAgent,
        MdBusiness } = require("react-icons/md");

// ─── Color Palette (Seerweb Brand) ───────────────────────────────────────────
const C = {
  navy:      "0F1B3C",   // deep navy — dominant
  royalBlue: "2563EB",   // royal blue — primary brand
  skyBlue:   "EFF6FF",   // light blue — content bg
  teal:      "0D9488",   // teal — accent
  white:     "FFFFFF",
  offWhite:  "F8FAFC",
  lightGray: "E2E8F0",
  midGray:   "64748B",
  darkGray:  "1E293B",
  amber:     "F59E0B",
  green:     "10B981",
  coral:     "EF4444",
};

// ─── Icon helper ──────────────────────────────────────────────────────────────
async function iconPng(IconComp, color = "#FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComp, { color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ─── Shadow factory ───────────────────────────────────────────────────────────
const shadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.12 });

// ─── Card helper ──────────────────────────────────────────────────────────────
function addCard(slide, x, y, w, h, opts = {}) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: opts.fill || C.white },
    line: { color: opts.border || C.lightGray, width: 1 },
    shadow: shadow(),
    rectRadius: 0.08,
  });
}

// ─── Slide number helper ──────────────────────────────────────────────────────
function addSlideNum(slide, num) {
  slide.addText(`${num} / 22`, {
    x: 9.3, y: 5.2, w: 0.65, h: 0.25,
    fontSize: 9, color: C.midGray, align: "right",
  });
}

// ─── Main ─────────────────────────────────────────────────────────────────────
async function buildPresentation() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title  = "Seerweb OMS – Sales Presentation";
  pres.author = "Seerweb ERP Solutions Pvt Ltd";

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 1 – TITLE / HERO
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };

    // Left large accent block
    s.addShape("rect", { x: 0, y: 0, w: 4.5, h: 5.625, fill: { color: C.royalBlue } });
    // Diagonal cut — triangle overlay
    s.addShape("rect", { x: 3.9, y: 0, w: 0.8, h: 5.625, fill: { color: C.navy } });

    // Brand name vertical on left panel
    s.addText("SEERWEB", {
      x: 0.35, y: 1.6, w: 3.8, h: 0.7,
      fontSize: 32, bold: true, color: C.white,
      fontFace: "Arial Black", charSpacing: 8, margin: 0,
    });
    s.addText("ERP SOLUTIONS", {
      x: 0.35, y: 2.25, w: 3.8, h: 0.35,
      fontSize: 12, color: "BFD4FF", charSpacing: 6, margin: 0,
    });

    // Divider line on left
    s.addShape("rect", { x: 0.35, y: 2.7, w: 2.5, h: 0.03, fill: { color: "BFD4FF" } });

    s.addText("OMS", {
      x: 0.35, y: 2.85, w: 3.8, h: 0.45,
      fontSize: 14, color: "BFD4FF", charSpacing: 4, margin: 0,
    });
    s.addText("Order Management System", {
      x: 0.35, y: 3.25, w: 3.8, h: 0.3,
      fontSize: 10, color: "BFD4FF", italic: true, margin: 0,
    });

    // Right panel — main headline
    s.addText("Smarter Orders.\nFaster Growth.\nTotal Control.", {
      x: 4.8, y: 0.9, w: 4.9, h: 2.3,
      fontSize: 34, bold: true, color: C.white,
      fontFace: "Arial Black", lineSpacingMultiple: 1.15, margin: 0,
    });

    // Tagline
    s.addText(
      "The all-in-one Order Management Platform built for\nDealers, Retailers & Sales Teams.",
      {
        x: 4.8, y: 3.25, w: 4.9, h: 0.8,
        fontSize: 13, color: "BFD4FF", lineSpacingMultiple: 1.4, margin: 0,
      }
    );

    // CTA pill
    s.addShape("rect", {
      x: 4.8, y: 4.3, w: 2.4, h: 0.55,
      fill: { color: C.teal }, rectRadius: 0.27,
    });
    s.addText("Book a Free Demo →", {
      x: 4.8, y: 4.3, w: 2.4, h: 0.55,
      fontSize: 12, bold: true, color: C.white, align: "center", valign: "middle", margin: 0,
    });

    // Bottom website
    s.addText("www.seerweberp.com", {
      x: 4.8, y: 5.15, w: 4.9, h: 0.25,
      fontSize: 10, color: "6B8EC2", margin: 0,
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 2 – AGENDA
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addText("What We'll Cover Today", {
      x: 0.5, y: 0.3, w: 9, h: 0.65,
      fontSize: 28, bold: true, color: C.navy, fontFace: "Arial Black", margin: 0,
    });
    s.addText("A structured walk-through of the Seerweb OMS platform", {
      x: 0.5, y: 0.9, w: 9, h: 0.3,
      fontSize: 13, color: C.midGray, margin: 0,
    });

    const items = [
      ["01", "The Problem",         "Challenges businesses face without a proper OMS"],
      ["02", "Our Solution",        "How Seerweb OMS solves those challenges"],
      ["03", "Platform Overview",   "Roles, dashboards & modules at a glance"],
      ["04", "Key Features",        "Voice ordering, garments ERP, analytics & more"],
      ["05", "Business Benefits",   "Time savings, cost reduction & revenue growth"],
      ["06", "Technology & Security","Modern stack, JWT auth & data protection"],
      ["07", "ROI & Roadmap",       "Measurable value & what's coming next"],
    ];

    items.forEach(([num, title, desc], i) => {
      const col = i < 4 ? 0 : 1;
      const row = i < 4 ? i : i - 4;
      const x = col === 0 ? 0.5 : 5.3;
      const y = 1.45 + row * 0.88;
      const w = 4.5;

      addCard(s, x, y, w, 0.72, { fill: C.white });

      // Number circle
      s.addShape("ellipse", {
        x: x + 0.12, y: y + 0.13, w: 0.45, h: 0.45,
        fill: { color: C.royalBlue },
      });
      s.addText(num, {
        x: x + 0.12, y: y + 0.13, w: 0.45, h: 0.45,
        fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle", margin: 0,
      });

      s.addText(title, {
        x: x + 0.68, y: y + 0.08, w: w - 0.8, h: 0.28,
        fontSize: 12, bold: true, color: C.darkGray, margin: 0,
      });
      s.addText(desc, {
        x: x + 0.68, y: y + 0.35, w: w - 0.8, h: 0.28,
        fontSize: 9.5, color: C.midGray, margin: 0,
      });
    });

    addSlideNum(s, 2);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 3 – THE PROBLEM
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };

    // Top label
    s.addShape("rect", { x: 0.5, y: 0.3, w: 1.6, h: 0.32, fill: { color: C.coral }, rectRadius: 0.05 });
    s.addText("THE PROBLEM", {
      x: 0.5, y: 0.3, w: 1.6, h: 0.32,
      fontSize: 9, bold: true, color: C.white, align: "center", valign: "middle", margin: 0,
    });

    s.addText("Running a Distribution Business Is Complicated", {
      x: 0.5, y: 0.75, w: 9, h: 0.7,
      fontSize: 26, bold: true, color: C.white, fontFace: "Arial Black", margin: 0,
    });
    s.addText("Without the right tools, dealers lose orders, miss revenue, and struggle to scale.", {
      x: 0.5, y: 1.45, w: 9, h: 0.35,
      fontSize: 13, color: "BFD4FF", margin: 0,
    });

    const problems = [
      ["📋", "Manual Order Tracking",     "Excel sheets, WhatsApp messages & phone calls cause errors and delays"],
      ["👥", "No Retailer Visibility",    "Dealers can't see which retailer bought what or how much they owe"],
      ["📦", "Inventory Chaos",           "Stock levels aren't tracked in real-time, leading to overselling"],
      ["🗣️", "Inefficient Field Sales",  "Sales reps waste hours writing orders by hand on the field"],
      ["📊", "Zero Business Analytics",   "No reports, no performance tracking, no data-driven decisions"],
      ["🔒", "Data Scattered Everywhere", "Critical business data spread across WhatsApp, calls & notebooks"],
    ];

    problems.forEach(([icon, title, desc], i) => {
      const col = i % 3;
      const row = Math.floor(i / 3);
      const x = 0.5 + col * 3.2;
      const y = 2.1 + row * 1.5;

      s.addShape("rect", {
        x, y, w: 3.0, h: 1.3,
        fill: { color: "1E2D5A" },
        line: { color: "2D3F6E", width: 1 },
        shadow: shadow(),
        rectRadius: 0.08,
      });
      s.addText(icon, {
        x: x + 0.15, y: y + 0.1, w: 0.5, h: 0.5,
        fontSize: 22, margin: 0,
      });
      s.addText(title, {
        x: x + 0.15, y: y + 0.56, w: 2.7, h: 0.3,
        fontSize: 11, bold: true, color: "FFD166", margin: 0,
      });
      s.addText(desc, {
        x: x + 0.15, y: y + 0.83, w: 2.7, h: 0.38,
        fontSize: 9, color: "BFD4FF", margin: 0,
      });
    });

    addSlideNum(s, 3);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 4 – THE SOLUTION
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addShape("rect", { x: 0.5, y: 0.3, w: 1.8, h: 0.32, fill: { color: C.green }, rectRadius: 0.05 });
    s.addText("THE SOLUTION", {
      x: 0.5, y: 0.3, w: 1.8, h: 0.32,
      fontSize: 9, bold: true, color: C.white, align: "center", valign: "middle", margin: 0,
    });

    s.addText("Introducing Seerweb OMS", {
      x: 0.5, y: 0.75, w: 9, h: 0.6,
      fontSize: 28, bold: true, color: C.navy, fontFace: "Arial Black", margin: 0,
    });
    s.addText("One unified platform to manage every order, every customer, every sale — from anywhere.",{
      x: 0.5, y: 1.35, w: 9, h: 0.35,
      fontSize: 13, color: C.midGray, margin: 0,
    });

    // Center illustration box
    addCard(s, 3.5, 1.85, 6.0, 3.4, { fill: C.white });
    s.addText("🖥️", { x: 3.7, y: 1.95, w: 0.6, h: 0.6, fontSize: 30, margin: 0 });
    s.addText("Seerweb OMS Platform", {
      x: 4.35, y: 2.05, w: 4.9, h: 0.4,
      fontSize: 15, bold: true, color: C.navy, margin: 0,
    });

    const solPoints = [
      ["Web & Mobile Access",     "Order from anywhere — browser or phone"],
      ["Real-Time Inventory",     "Live stock levels, no overselling"],
      ["Voice Order Entry",       "Hindi + English voice commands on field"],
      ["Multi-Role Dashboards",   "Admin, Dealer, Retailer & Staff views"],
      ["AI-Powered Analytics",    "Charts, trends, performance reports"],
      ["Garments ERP Module",     "Size-wise booking, set ordering & catalog"],
    ];

    solPoints.forEach(([title, desc], i) => {
      const y = 2.55 + i * 0.42;
      s.addShape("ellipse", { x: 3.7, y: y + 0.04, w: 0.22, h: 0.22, fill: { color: C.teal } });
      s.addText("✓", { x: 3.7, y: y + 0.04, w: 0.22, h: 0.22, fontSize: 9, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
      s.addText(title, { x: 4.02, y, w: 2.0, h: 0.28, fontSize: 10.5, bold: true, color: C.darkGray, margin: 0 });
      s.addText(desc,  { x: 6.08, y, w: 3.2, h: 0.28, fontSize: 9.5,  color: C.midGray, margin: 0 });
    });

    // Left stats column
    const stats = [["500+", "Retailers Managed"], ["3 Roles", "Supported"], ["100%", "Cloud-Based"]];
    stats.forEach(([num, label], i) => {
      const y = 1.9 + i * 1.1;
      addCard(s, 0.5, y, 2.8, 0.9, { fill: C.white });
      s.addText(num,   { x: 0.65, y: y + 0.05, w: 2.5, h: 0.45, fontSize: 24, bold: true, color: C.royalBlue, fontFace: "Arial Black", margin: 0 });
      s.addText(label, { x: 0.65, y: y + 0.5, w: 2.5, h: 0.28, fontSize: 10, color: C.midGray, margin: 0 });
    });

    addSlideNum(s, 4);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 5 – USER ROLES
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.white };

    s.addText("User Roles & Permissions", {
      x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.navy, fontFace: "Arial Black", margin: 0,
    });
    s.addText("Four specialized roles — each with purpose-built dashboards and controlled access", {
      x: 0.5, y: 0.9, w: 9, h: 0.3, fontSize: 13, color: C.midGray, margin: 0,
    });

    const roles = [
      {
        color: "1E3A8A", light: "EFF6FF", emoji: "🛡️", title: "Super Admin",
        who: "Seerweb / Platform Owner",
        caps: ["Create & manage all Dealer accounts", "View platform-wide analytics", "Assign business types to dealers", "Full product & user oversight"],
      },
      {
        color: "065A82", light: "E0F2FE", emoji: "🏬", title: "Dealer",
        who: "Distributor / Wholesale Business",
        caps: ["Manage own Retailers & Staff", "Add products & set pricing", "View all orders & analytics", "Create orders on behalf of retailers"],
      },
      {
        color: "0D9488", light: "F0FDF9", emoji: "🏪", title: "Retailer",
        who: "Store Owner / Shop",
        caps: ["Browse dealer product catalog", "Place & track own orders", "View order history", "Access personal profile & settings"],
      },
      {
        color: "7C3AED", light: "F5F3FF", emoji: "👨‍💼", title: "Sales Executive",
        who: "Field Sales Staff",
        caps: ["View assigned customers", "Create orders on field (voice/touch)", "Track own order performance", "Dashboard with sales metrics"],
      },
    ];

    roles.forEach((r, i) => {
      const x = 0.35 + i * 2.35;
      const y = 1.45;
      const w = 2.2;
      const h = 3.9;

      // Card
      s.addShape("rect", { x, y, w, h, fill: { color: r.light }, line: { color: r.color, width: 2 }, shadow: shadow(), rectRadius: 0.1 });

      // Color header strip
      s.addShape("rect", { x, y, w, h: 0.85, fill: { color: r.color }, rectRadius: 0.1 });
      s.addShape("rect", { x, y: y + 0.6, w, h: 0.3, fill: { color: r.color } }); // square bottom corners

      s.addText(r.emoji, { x: x + 0.1, y: y + 0.08, w: 0.55, h: 0.55, fontSize: 22, margin: 0 });
      s.addText(r.title, { x: x + 0.65, y: y + 0.12, w: w - 0.75, h: 0.35, fontSize: 13, bold: true, color: C.white, margin: 0 });
      s.addText(r.who,   { x: x + 0.12, y: y + 0.88, w: w - 0.2, h: 0.28, fontSize: 9, color: r.color, bold: true, margin: 0 });

      r.caps.forEach((cap, ci) => {
        s.addText("• " + cap, {
          x: x + 0.12, y: y + 1.22 + ci * 0.56, w: w - 0.2, h: 0.5,
          fontSize: 9.5, color: C.darkGray, margin: 0,
        });
      });
    });

    addSlideNum(s, 5);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 6 – PLATFORM WORKFLOW
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };

    s.addText("How It Works", {
      x: 0.5, y: 0.25, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white, fontFace: "Arial Black", margin: 0,
    });
    s.addText("A seamless order flow from product setup to doorstep delivery", {
      x: 0.5, y: 0.85, w: 9, h: 0.3, fontSize: 13, color: "BFD4FF", margin: 0,
    });

    const steps = [
      { num: "1", title: "Setup",          sub: "Admin creates Dealer\nDealer adds Products &\nRetailers",        color: "2563EB" },
      { num: "2", title: "Browse",         sub: "Retailer / Executive\nbrowses product catalog\nwith live stock",  color: "0D9488" },
      { num: "3", title: "Order",          sub: "Place order by touch,\nvoice command, or\nQR scan",               color: "7C3AED" },
      { num: "4", title: "Approve",        sub: "Dealer reviews &\napproves order in\nseconds",                    color: "F59E0B" },
      { num: "5", title: "Dispatch",       sub: "Status updated to\nDispatched — retailer\nnotified",              color: "EF4444" },
      { num: "6", title: "Delivered",      sub: "Order marked Delivered\nRevenue tracked in\nanalytics",           color: "10B981" },
    ];

    steps.forEach((step, i) => {
      const x = 0.45 + i * 1.6;
      const y = 1.6;

      // Box
      s.addShape("rect", {
        x, y, w: 1.45, h: 2.8,
        fill: { color: "1E2D5A" }, line: { color: step.color, width: 2 }, rectRadius: 0.1,
      });

      // Number circle
      s.addShape("ellipse", { x: x + 0.47, y: y + 0.14, w: 0.5, h: 0.5, fill: { color: step.color } });
      s.addText(step.num, {
        x: x + 0.47, y: y + 0.14, w: 0.5, h: 0.5,
        fontSize: 16, bold: true, color: C.white, align: "center", valign: "middle", margin: 0,
      });

      s.addText(step.title, {
        x: x + 0.05, y: y + 0.78, w: 1.35, h: 0.35,
        fontSize: 12, bold: true, color: "FFD166", align: "center", margin: 0,
      });
      s.addText(step.sub, {
        x: x + 0.05, y: y + 1.15, w: 1.35, h: 1.5,
        fontSize: 8.5, color: "BFD4FF", align: "center", lineSpacingMultiple: 1.3, margin: 0,
      });

      // Arrow (except last)
      if (i < 5) {
        s.addShape("rect", { x: x + 1.45, y: y + 0.37, w: 0.13, h: 0.07, fill: { color: C.amber } });
        s.addText("▶", { x: x + 1.52, y: y + 0.29, w: 0.1, h: 0.2, fontSize: 8, color: C.amber, margin: 0 });
      }
    });

    // Bottom note
    s.addText("📱 Access from any device — Desktop, Tablet or Mobile Browser", {
      x: 0.5, y: 4.7, w: 9, h: 0.3, fontSize: 11, color: "6B8EC2", align: "center", margin: 0,
    });

    addSlideNum(s, 6);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 7 – DASHBOARD OVERVIEW
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addText("Powerful Dashboards for Every Role", {
      x: 0.5, y: 0.25, w: 9, h: 0.6, fontSize: 26, bold: true, color: C.navy, fontFace: "Arial Black", margin: 0,
    });
    s.addText("At-a-glance metrics that drive decisions — not just data", {
      x: 0.5, y: 0.85, w: 9, h: 0.3, fontSize: 13, color: C.midGray, margin: 0,
    });

    // Main screenshot placeholder
    addCard(s, 0.5, 1.3, 5.5, 3.7, { fill: C.white });
    s.addText("📊  DEALER DASHBOARD", {
      x: 0.65, y: 1.45, w: 5.2, h: 0.4, fontSize: 14, bold: true, color: C.royalBlue, margin: 0,
    });

    // Mock stat cards inside placeholder
    const mockStats = [
      ["Total Orders", "1,284", "2563EB"],
      ["Pending", "23",    "F59E0B"],
      ["Revenue",   "₹4.2L", "10B981"],
      ["Today",     "17",   "7C3AED"],
    ];
    mockStats.forEach(([label, val, col], i) => {
      const mx = 0.65 + i * 1.33;
      s.addShape("rect", { x: mx, y: 2.0, w: 1.25, h: 0.9, fill: { color: "F8FAFF" }, line: { color: col, width: 2 }, rectRadius: 0.06 });
      s.addText(val,   { x: mx + 0.05, y: 2.06, w: 1.15, h: 0.38, fontSize: 16, bold: true, color: col, fontFace: "Arial Black", margin: 0 });
      s.addText(label, { x: mx + 0.05, y: 2.44, w: 1.15, h: 0.3, fontSize: 8, color: C.midGray, margin: 0 });
    });

    // Mini chart bars
    const barH = [0.5, 0.8, 0.65, 1.1, 0.9, 1.3, 0.75];
    barH.forEach((h, i) => {
      s.addShape("rect", { x: 0.7 + i * 0.72, y: 3.8 - h, w: 0.5, h, fill: { color: C.royalBlue + "99" } });
    });
    s.addText("Orders Over Last 7 Days", { x: 0.65, y: 4.7, w: 5.2, h: 0.2, fontSize: 9, color: C.midGray, margin: 0 });

    // Right: dashboard features
    const features = [
      ["📈", "Live Order Analytics",    "Revenue, order counts & status breakdown updated in real-time"],
      ["🏆", "Top Retailers by Revenue","Identify your highest-value customers instantly"],
      ["📋", "Recent Order Feed",       "Latest orders across all channels with one-click status change"],
      ["🎯", "Staff Performance",       "Compare executive performance — orders & revenue side by side"],
      ["⚠️", "Pending Alerts",          "Highlighted pending orders so nothing slips through the cracks"],
    ];
    features.forEach(([icon, title, desc], i) => {
      const y = 1.35 + i * 0.82;
      addCard(s, 6.2, y, 3.6, 0.68, { fill: C.white });
      s.addText(icon,  { x: 6.3,  y: y + 0.16, w: 0.35, h: 0.35, fontSize: 16, margin: 0 });
      s.addText(title, { x: 6.72, y: y + 0.07, w: 2.95, h: 0.28, fontSize: 10.5, bold: true, color: C.darkGray, margin: 0 });
      s.addText(desc,  { x: 6.72, y: y + 0.34, w: 2.95, h: 0.26, fontSize: 8.5, color: C.midGray, margin: 0 });
    });

    addSlideNum(s, 7);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 8 – ORDER MANAGEMENT MODULE
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.white };

    s.addShape("rect", { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.royalBlue } });

    s.addText("Order Management", {
      x: 0.45, y: 0.25, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.navy, fontFace: "Arial Black", margin: 0,
    });
    s.addText("From order creation to delivery — every step tracked, every detail captured", {
      x: 0.45, y: 0.85, w: 9, h: 0.3, fontSize: 13, color: C.midGray, margin: 0,
    });

    // Order list mockup
    addCard(s, 0.45, 1.3, 5.6, 3.9, { fill: C.offWhite });
    s.addText("📦  ORDERS LIST", { x: 0.6, y: 1.42, w: 5.3, h: 0.35, fontSize: 13, bold: true, color: C.royalBlue, margin: 0 });

    const orderRows = [
      ["#ORD-4821", "Ramesh Electronics",  "₹12,400", "pending",    "F59E0B"],
      ["#ORD-4820", "Singh Mobile Store",  "₹8,750",  "approved",   "2563EB"],
      ["#ORD-4819", "Patel Garments",      "₹31,200", "delivered",  "10B981"],
      ["#ORD-4818", "Kumar Wholesale",     "₹5,600",  "dispatched", "7C3AED"],
    ];
    orderRows.forEach(([id, name, total, status, col], i) => {
      const y = 1.9 + i * 0.65;
      s.addShape("rect", { x: 0.55, y, w: 5.4, h: 0.55, fill: { color: C.white }, line: { color: C.lightGray, width: 1 }, rectRadius: 0.04 });
      s.addText(id,     { x: 0.65, y: y + 0.14, w: 1.1, h: 0.28, fontSize: 9, color: C.midGray, bold: true, margin: 0 });
      s.addText(name,   { x: 1.82, y: y + 0.14, w: 1.8, h: 0.28, fontSize: 9.5, color: C.darkGray, margin: 0 });
      s.addText(total,  { x: 3.68, y: y + 0.14, w: 0.9, h: 0.28, fontSize: 10, bold: true, color: C.darkGray, margin: 0 });
      s.addShape("rect", { x: 4.65, y: y + 0.1, w: 0.95, h: 0.28, fill: { color: col + "22" }, rectRadius: 0.05 });
      s.addText(status.toUpperCase(), { x: 4.65, y: y + 0.1, w: 0.95, h: 0.28, fontSize: 7.5, bold: true, color: col, align: "center", valign: "middle", margin: 0 });
    });

    // Filter bar hint
    s.addShape("rect", { x: 0.55, y: 4.58, w: 5.4, h: 0.5, fill: { color: C.offWhite }, line: { color: C.lightGray, width: 1 } });
    s.addText("⚙  Date Filter   |   Customer Filter   |   Staff Filter   |   Sort  ↓  |  ⬇ Export CSV / Excel", {
      x: 0.65, y: 4.62, w: 5.2, h: 0.38, fontSize: 8.5, color: C.midGray, margin: 0,
    });

    // Right feature list
    const feats = [
      ["🔍", "Smart Filtering", "Filter by date, status, retailer or sales executive"],
      ["🖨️", "Invoice Printing",  "Print professional garment order invoices directly"],
      ["📊", "Status Management", "Update from Pending → Approved → Dispatched → Delivered"],
      ["📤", "Excel/CSV Export",  "Download full order reports for accounting"],
      ["🔔", "Pending Alerts",    "Dashboard highlights unreviewed orders prominently"],
      ["📝", "Order Notes",       "Add special instructions per order for operations team"],
    ];
    feats.forEach(([icon, title, desc], i) => {
      const y = 1.35 + i * 0.73;
      s.addText(icon,  { x: 6.2, y: y + 0.12, w: 0.35, h: 0.35, fontSize: 16, margin: 0 });
      s.addText(title, { x: 6.62, y: y + 0.06, w: 3.1, h: 0.28, fontSize: 10.5, bold: true, color: C.darkGray, margin: 0 });
      s.addText(desc,  { x: 6.62, y: y + 0.33, w: 3.1, h: 0.26, fontSize: 9, color: C.midGray, margin: 0 });
    });

    addSlideNum(s, 8);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 9 – VOICE ORDER FEATURE
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };

    // Glowing mic graphic suggestion
    s.addShape("ellipse", { x: 7.2, y: 0.5, w: 2.8, h: 2.8, fill: { color: "1E3A8A" } });
    s.addText("🎤", { x: 7.2, y: 0.5, w: 2.8, h: 2.8, fontSize: 72, align: "center", valign: "middle", margin: 0 });

    s.addShape("rect", { x: 0.5, y: 0.3, w: 2.4, h: 0.32, fill: { color: C.teal }, rectRadius: 0.05 });
    s.addText("AI VOICE ORDERING", { x: 0.5, y: 0.3, w: 2.4, h: 0.32, fontSize: 9, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });

    s.addText("Order with Your Voice —\nin Hindi or English", {
      x: 0.5, y: 0.78, w: 6.5, h: 1.0, fontSize: 28, bold: true, color: C.white, fontFace: "Arial Black", lineSpacingMultiple: 1.2, margin: 0,
    });
    s.addText("Sales reps on the field can dictate entire orders by speaking naturally. Our AI parses the speech, matches products and confirms quantities — all in seconds.", {
      x: 0.5, y: 1.85, w: 6.3, h: 0.75, fontSize: 12, color: "BFD4FF", lineSpacingMultiple: 1.4, margin: 0,
    });

    // Flow boxes
    const flowItems = [
      ["🎙️", "Speak", "Say the product name & qty in Hindi or English"],
      ["🧠", "AI Parse", "Our NLP engine identifies products with confidence scoring"],
      ["✅", "Confirm", "Review matched items or correct via fallback editor"],
      ["🛒", "Added",   "Products added to cart — order submitted instantly"],
    ];
    flowItems.forEach((item, i) => {
      const x = 0.5 + i * 2.3;
      s.addShape("rect", { x, y: 2.8, w: 2.1, h: 2.5, fill: { color: "1E2D5A" }, line: { color: C.teal, width: 1 }, rectRadius: 0.08 });
      s.addText(item[0], { x: x + 0.75, y: 2.95, w: 0.6, h: 0.6, fontSize: 22, align: "center", margin: 0 });
      s.addText(item[1], { x, y: 3.6, w: 2.1, h: 0.32, fontSize: 12, bold: true, color: "FFD166", align: "center", margin: 0 });
      s.addText(item[2], { x: x + 0.1, y: 3.95, w: 1.9, h: 1.2, fontSize: 9.5, color: "BFD4FF", align: "center", lineSpacingMultiple: 1.3, margin: 0 });
      if (i < 3) s.addText("→", { x: x + 2.1, y: 3.7, w: 0.2, h: 0.3, fontSize: 14, color: C.amber, margin: 0 });
    });

    // Key fact
    s.addShape("rect", { x: 0.5, y: 5.1, w: 9, h: 0.35, fill: { color: "1E3A8A" }, rectRadius: 0.05 });
    s.addText("🌐  Supports multilingual input: Hindi, Romanized Hindi + English — perfect for Indian field sales teams", {
      x: 0.6, y: 5.12, w: 8.8, h: 0.3, fontSize: 10, color: "BFD4FF", margin: 0,
    });

    addSlideNum(s, 9);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 10 – GARMENTS ERP MODULE
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addShape("rect", { x: 0.5, y: 0.3, w: 2.2, h: 0.32, fill: { color: C.amber }, rectRadius: 0.05 });
    s.addText("GARMENTS ERP MODULE", { x: 0.5, y: 0.3, w: 2.2, h: 0.32, fontSize: 9, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });

    s.addText("Built for the Garments Industry", {
      x: 0.5, y: 0.75, w: 9, h: 0.6, fontSize: 26, bold: true, color: C.navy, fontFace: "Arial Black", margin: 0,
    });
    s.addText("Size-wise booking, wholesale catalog, set ordering — everything a garments wholesaler needs", {
      x: 0.5, y: 1.35, w: 9, h: 0.3, fontSize: 13, color: C.midGray, margin: 0,
    });

    // Left: feature cards grid
    const garFeats = [
      ["👗", "Design-wise Catalog",      "Browse products by design number, fabric, category & color"],
      ["📐", "Size Matrix Ordering",     "Piece-wise and set-wise ordering across all sizes in one screen"],
      ["🎨", "Color & Style Selection",  "Choose from available colors per design with swatch preview"],
      ["📸", "Gallery Image Support",    "Multiple product photos per design with ordering capability"],
      ["📄", "Wholesale Invoice PDF",    "Print professional garment invoices with size-rate tables"],
      ["🔖", "Booking Types",            "Support for Ready Stock, Booking, Catalog Launch & Pre-booking"],
    ];
    garFeats.forEach(([icon, title, desc], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.5 + col * 4.7;
      const y = 1.85 + row * 1.2;
      addCard(s, x, y, 4.4, 1.05, { fill: C.white });
      s.addText(icon,  { x: x + 0.12, y: y + 0.28, w: 0.45, h: 0.45, fontSize: 20, margin: 0 });
      s.addText(title, { x: x + 0.65, y: y + 0.1,  w: 3.6,  h: 0.35, fontSize: 11, bold: true, color: C.darkGray, margin: 0 });
      s.addText(desc,  { x: x + 0.65, y: y + 0.44, w: 3.6,  h: 0.48, fontSize: 9.5, color: C.midGray, margin: 0 });
    });

    // Bottom highlight
    s.addShape("rect", { x: 0.5, y: 5.1, w: 9, h: 0.35, fill: { color: C.amber + "22" }, rectRadius: 0.05 });
    s.addText("📱  Share catalog via WhatsApp · 📄  Export to PDF · 🔍  Filter by design, category, fabric, size & color", {
      x: 0.6, y: 5.12, w: 8.8, h: 0.3, fontSize: 10, color: C.darkGray, margin: 0,
    });

    addSlideNum(s, 10);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 11 – PRODUCT & INVENTORY MANAGEMENT
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.white };

    s.addShape("rect", { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.green } });

    s.addText("Product & Inventory Management", {
      x: 0.45, y: 0.25, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.navy, fontFace: "Arial Black", margin: 0,
    });
    s.addText("Full control over your product catalog — with dynamic fields, variant tracking & bulk tools", {
      x: 0.45, y: 0.85, w: 9, h: 0.3, fontSize: 12, color: C.midGray, margin: 0,
    });

    const invFeats = [
      { icon: "📦", title: "Variant-Based Products",  desc: "Track each size/color as a separate variant with individual stock & pricing" },
      { icon: "⚡", title: "Bulk Price & Stock Update", desc: "Select multiple products and update price or stock in one action" },
      { icon: "🏷️", title: "Custom Product Fields",   desc: "Add dealer-specific attributes like HSN Code, Fabric Type, Season etc." },
      { icon: "📸", title: "Product Image Gallery",    desc: "Upload main image + gallery photos; auto-served via CDN" },
      { icon: "⚠️", title: "Low Stock Alerts",         desc: "Visual alerts on dashboard for products below threshold" },
      { icon: "📤", title: "Excel Import/Export",      desc: "Bulk import products via Excel; export full catalog for reporting" },
      { icon: "🔗", title: "Business-Type Schema",     desc: "Different product fields for mobiles vs garments — auto-configured" },
      { icon: "🔍", title: "Advanced Search & Filter", desc: "Filter by brand, model, color, stock status, or custom attributes" },
    ];

    invFeats.forEach((f, i) => {
      const col = i % 4, row = Math.floor(i / 4);
      const x = 0.45 + col * 2.4;
      const y = 1.35 + row * 1.9;
      addCard(s, x, y, 2.22, 1.7, { fill: C.offWhite });
      s.addShape("ellipse", { x: x + 0.82, y: y + 0.15, w: 0.58, h: 0.58, fill: { color: C.royalBlue + "22" } });
      s.addText(f.icon,  { x: x + 0.82, y: y + 0.15, w: 0.58, h: 0.58, fontSize: 20, align: "center", valign: "middle", margin: 0 });
      s.addText(f.title, { x: x + 0.1, y: y + 0.82, w: 2.02, h: 0.38, fontSize: 9.5, bold: true, color: C.darkGray, align: "center", margin: 0 });
      s.addText(f.desc,  { x: x + 0.1, y: y + 1.2,  w: 2.02, h: 0.42, fontSize: 8.5, color: C.midGray, align: "center", margin: 0 });
    });

    addSlideNum(s, 11);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 12 – CUSTOMER & STAFF MANAGEMENT
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addText("Customer & Staff Management", {
      x: 0.5, y: 0.25, w: 9, h: 0.6, fontSize: 26, bold: true, color: C.navy, fontFace: "Arial Black", margin: 0,
    });
    s.addText("Manage your entire network — retailers, executives & assignments — from one screen", {
      x: 0.5, y: 0.85, w: 9, h: 0.3, fontSize: 13, color: C.midGray, margin: 0,
    });

    // Two columns
    // Left: Retailer management
    addCard(s, 0.5, 1.35, 4.5, 4.0, { fill: C.white });
    s.addShape("rect", { x: 0.5, y: 1.35, w: 4.5, h: 0.5, fill: { color: C.royalBlue }, rectRadius: 0.08 });
    s.addShape("rect", { x: 0.5, y: 1.6, w: 4.5, h: 0.3, fill: { color: C.royalBlue } });
    s.addText("🏪  Retailer Management", { x: 0.65, y: 1.4, w: 4.2, h: 0.35, fontSize: 13, bold: true, color: C.white, margin: 0 });

    const retFeats = [
      "Add / Edit / Delete retailer accounts",
      "Assign Sales Executive to each retailer",
      "Track order history & revenue per retailer",
      "Sort & filter by name, revenue or orders",
      "Export retailer data to CSV or Excel",
      "Server-side duplicate username/email validation",
      "Google Maps integration — navigate to store",
      "One-tap call button for direct contact",
    ];
    retFeats.forEach((f, i) => {
      s.addShape("ellipse", { x: 0.65, y: 1.97 + i * 0.38, w: 0.18, h: 0.18, fill: { color: C.royalBlue } });
      s.addText(f, { x: 0.92, y: 1.94 + i * 0.38, w: 3.9, h: 0.3, fontSize: 10, color: C.darkGray, margin: 0 });
    });

    // Right: Staff management
    addCard(s, 5.2, 1.35, 4.5, 4.0, { fill: C.white });
    s.addShape("rect", { x: 5.2, y: 1.35, w: 4.5, h: 0.5, fill: { color: C.teal }, rectRadius: 0.08 });
    s.addShape("rect", { x: 5.2, y: 1.6, w: 4.5, h: 0.3, fill: { color: C.teal } });
    s.addText("👨‍💼  Staff Management", { x: 5.35, y: 1.4, w: 4.2, h: 0.35, fontSize: 13, bold: true, color: C.white, margin: 0 });

    const staffFeats = [
      "Create Sales Executive accounts",
      "Assign retailers to specific executives",
      "Track orders & revenue per staff member",
      "Performance chart — compare executives",
      "Top performer ranking with revenue badge",
      "View recent activity & last order date",
      "Export staff performance reports",
      "Role-based access — executives only see their data",
    ];
    staffFeats.forEach((f, i) => {
      s.addShape("ellipse", { x: 5.35, y: 1.97 + i * 0.38, w: 0.18, h: 0.18, fill: { color: C.teal } });
      s.addText(f, { x: 5.62, y: 1.94 + i * 0.38, w: 3.9, h: 0.3, fontSize: 10, color: C.darkGray, margin: 0 });
    });

    addSlideNum(s, 12);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 13 – ANALYTICS & REPORTS
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };

    s.addText("Analytics & Reporting", {
      x: 0.5, y: 0.25, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white, fontFace: "Arial Black", margin: 0,
    });
    s.addText("Turn raw data into business decisions — instantly, with zero setup", {
      x: 0.5, y: 0.85, w: 9, h: 0.3, fontSize: 13, color: "BFD4FF", margin: 0,
    });

    // Chart 1 – Orders over time (bar)
    addCard(s, 0.5, 1.35, 4.5, 2.5, { fill: "1E2D5A" });
    s.addText("Orders — Last 7 Days", { x: 0.65, y: 1.45, w: 4.2, h: 0.3, fontSize: 11, bold: true, color: C.white, margin: 0 });
    const barVals = [8, 12, 9, 18, 14, 22, 17];
    barVals.forEach((v, i) => {
      const bh = (v / 25) * 1.4;
      s.addShape("rect", { x: 0.72 + i * 0.6, y: 1.95 + (1.4 - bh), w: 0.45, h: bh, fill: { color: C.royalBlue } });
      s.addText(String(v), { x: 0.72 + i * 0.6, y: 1.82 + (1.4 - bh), w: 0.45, h: 0.2, fontSize: 8, color: "BFD4FF", align: "center", margin: 0 });
    });
    ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"].forEach((d, i) => {
      s.addText(d, { x: 0.72 + i * 0.6, y: 3.43, w: 0.45, h: 0.2, fontSize: 7.5, color: C.midGray, align: "center", margin: 0 });
    });

    // Chart 2 – Donut (status breakdown)
    addCard(s, 5.2, 1.35, 4.5, 2.5, { fill: "1E2D5A" });
    s.addText("Orders by Status", { x: 5.35, y: 1.45, w: 4.2, h: 0.3, fontSize: 11, bold: true, color: C.white, margin: 0 });
    const donutData = [["Delivered", "10B981", 42], ["Approved", "2563EB", 28], ["Pending", "F59E0B", 18], ["Cancelled", "EF4444", 12]];
    donutData.forEach(([label, col, pct], i) => {
      s.addShape("rect", { x: 5.35, y: 1.9 + i * 0.5, w: 0.22, h: 0.22, fill: { color: col } });
      s.addText(`${label}  ${pct}%`, { x: 5.65, y: 1.88 + i * 0.5, w: 3.8, h: 0.28, fontSize: 10, color: "BFD4FF", margin: 0 });
    });

    // Bottom 4 analytics cards
    const analytics = [
      ["📊", "Revenue by Status",      "Compare pending vs approved vs delivered revenue at a glance"],
      ["🏆", "Top Customer Ranking",   "Identify your top 5 retailers by revenue automatically"],
      ["👨‍💼", "Staff Performance Chart", "Bar chart comparison across all sales executives"],
      ["📅", "Date Range Filtering",   "Filter orders by Today, Week, Month or custom period"],
    ];
    analytics.forEach(([icon, title, desc], i) => {
      const x = 0.5 + i * 2.4;
      addCard(s, x, 4.1, 2.2, 1.2, { fill: "1E2D5A" });
      s.addText(icon,  { x: x + 0.1, y: 4.18, w: 0.45, h: 0.45, fontSize: 20, margin: 0 });
      s.addText(title, { x: x + 0.1, y: 4.62, w: 2.0, h: 0.3, fontSize: 9.5, bold: true, color: "FFD166", margin: 0 });
      s.addText(desc,  { x: x + 0.1, y: 4.92, w: 2.0, h: 0.35, fontSize: 8.5, color: "BFD4FF", margin: 0 });
    });

    addSlideNum(s, 13);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 14 – MOBILE & FIELD SALES
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addText("Optimized for Mobile & Field Sales", {
      x: 0.5, y: 0.25, w: 9, h: 0.6, fontSize: 26, bold: true, color: C.navy, fontFace: "Arial Black", margin: 0,
    });
    s.addText("Your sales team stays productive — whether at the office, warehouse or customer's shop", {
      x: 0.5, y: 0.85, w: 9, h: 0.3, fontSize: 13, color: C.midGray, margin: 0,
    });

    // Phone mock
    s.addShape("rect", { x: 0.5, y: 1.2, w: 3.0, h: 4.1, fill: { color: C.navy }, rectRadius: 0.28, line: { color: C.darkGray, width: 3 } });
    s.addShape("rect", { x: 0.62, y: 1.35, w: 2.76, h: 3.8, fill: { color: C.royalBlue + "22" } });
    s.addText("📱", { x: 0.5, y: 1.2, w: 3.0, h: 4.1, fontSize: 80, align: "center", valign: "middle", margin: 0 });
    s.addText("Progressive Web App\nWorks on any device", {
      x: 0.6, y: 4.8, w: 2.8, h: 0.5, fontSize: 9.5, color: C.midGray, align: "center", lineSpacingMultiple: 1.3, margin: 0,
    });

    const mobileFeats = [
      ["📍", "Location-Aware",      "Navigate to retailer stores via Google Maps integration"],
      ["📞", "One-Tap Calling",      "Call customers directly from the customer list with one tap"],
      ["🎤", "Voice Order Entry",    "Speak orders in Hindi or English — hands-free ordering on the go"],
      ["🔄", "Real-Time Sync",       "Cart & order data syncs across all devices instantly"],
      ["📊", "Mobile Dashboard",     "Full analytics accessible from smartphone browsers"],
      ["🛒", "Cart Persistence",     "Cart saved per user — never lose an order mid-session"],
      ["🔒", "Secure Login",         "JWT-based session with 24-hour auto-expiry for security"],
      ["⚡", "Fast Load",            "Optimized for low-bandwidth mobile networks in India"],
    ];

    mobileFeats.forEach((f, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 3.8 + col * 3.1;
      const y = 1.25 + row * 1.0;
      addCard(s, x, y, 2.9, 0.82, { fill: C.white });
      s.addText(f[0],  { x: x + 0.1, y: y + 0.22, w: 0.35, h: 0.35, fontSize: 16, margin: 0 });
      s.addText(f[1],  { x: x + 0.5, y: y + 0.08, w: 2.3, h: 0.28, fontSize: 10.5, bold: true, color: C.darkGray, margin: 0 });
      s.addText(f[2],  { x: x + 0.5, y: y + 0.36, w: 2.3, h: 0.38, fontSize: 8.8, color: C.midGray, margin: 0 });
    });

    addSlideNum(s, 14);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 15 – SECURITY & TECHNOLOGY
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.white };

    s.addShape("rect", { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.teal } });

    s.addText("Security & Technology Stack", {
      x: 0.45, y: 0.25, w: 9, h: 0.6, fontSize: 26, bold: true, color: C.navy, fontFace: "Arial Black", margin: 0,
    });
    s.addText("Enterprise-grade security built on a modern, scalable technology foundation", {
      x: 0.45, y: 0.85, w: 9, h: 0.3, fontSize: 13, color: C.midGray, margin: 0,
    });

    // Left: Security
    addCard(s, 0.45, 1.35, 4.5, 4.0, { fill: C.offWhite });
    s.addShape("rect", { x: 0.45, y: 1.35, w: 4.5, h: 0.55, fill: { color: "065A82" }, rectRadius: 0.08 });
    s.addShape("rect", { x: 0.45, y: 1.65, w: 4.5, h: 0.3, fill: { color: "065A82" } });
    s.addText("🔒  Security Features", { x: 0.6, y: 1.42, w: 4.2, h: 0.38, fontSize: 13, bold: true, color: C.white, margin: 0 });

    const secFeats = [
      ["🔑", "JWT Authentication",   "JSON Web Tokens with 24-hour auto-expiry"],
      ["🛡️", "Role-Based Access",    "Users only see data they're authorized to access"],
      ["🔐", "Secure Password",      "Hashed passwords with server-side change validation"],
      ["🌐", "HTTPS-Only API",       "All communications encrypted with TLS"],
      ["⏱️", "Session Timeout",      "Auto-logout on token expiry prevents unauthorized access"],
      ["🚫", "Protected Routes",     "Frontend route guards prevent unauthorized page access"],
    ];
    secFeats.forEach(([icon, title, desc], i) => {
      const y = 2.1 + i * 0.52;
      s.addText(icon,  { x: 0.6,  y: y + 0.1, w: 0.35, h: 0.35, fontSize: 14, margin: 0 });
      s.addText(title, { x: 1.02, y: y + 0.04, w: 1.65, h: 0.28, fontSize: 10, bold: true, color: C.darkGray, margin: 0 });
      s.addText(desc,  { x: 2.72, y: y + 0.04, w: 2.05, h: 0.38, fontSize: 9, color: C.midGray, margin: 0 });
    });

    // Right: Tech stack
    addCard(s, 5.15, 1.35, 4.5, 4.0, { fill: C.offWhite });
    s.addShape("rect", { x: 5.15, y: 1.35, w: 4.5, h: 0.55, fill: { color: C.teal }, rectRadius: 0.08 });
    s.addShape("rect", { x: 5.15, y: 1.65, w: 4.5, h: 0.3, fill: { color: C.teal } });
    s.addText("⚙️  Technology Stack", { x: 5.3, y: 1.42, w: 4.2, h: 0.38, fontSize: 13, bold: true, color: C.white, margin: 0 });

    const techStack = [
      ["Frontend", "React 18, TypeScript, Tailwind CSS, Vite"],
      ["UI Library", "shadcn/ui, Radix UI, Lucide Icons"],
      ["State Mgmt", "React Context, TanStack Query"],
      ["Backend API", "Node.js REST API (oms.seerweberp.com)"],
      ["Database",  "MySQL (managed, cloud-hosted)"],
      ["AI / Voice", "Web Speech API + Supabase Edge Functions"],
      ["Auth",      "JWT + localStorage session management"],
      ["PDF / Export", "jsPDF, html2canvas, SheetJS / xlsx"],
    ];
    techStack.forEach(([layer, tech], i) => {
      const y = 2.1 + i * 0.52;
      s.addShape("rect", { x: 5.3, y: y + 0.05, w: 1.1, h: 0.3, fill: { color: C.teal + "33" }, rectRadius: 0.04 });
      s.addText(layer, { x: 5.3, y: y + 0.05, w: 1.1, h: 0.3, fontSize: 8.5, bold: true, color: C.teal, align: "center", valign: "middle", margin: 0 });
      s.addText(tech,  { x: 6.5, y: y + 0.05, w: 3.0, h: 0.35, fontSize: 9.5, color: C.darkGray, margin: 0 });
    });

    addSlideNum(s, 15);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 16 – BUSINESS BENEFITS
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };

    s.addText("What You Gain with Seerweb OMS", {
      x: 0.5, y: 0.25, w: 9, h: 0.6, fontSize: 26, bold: true, color: C.white, fontFace: "Arial Black", margin: 0,
    });
    s.addText("Real, measurable business impact — not just software features", {
      x: 0.5, y: 0.85, w: 9, h: 0.3, fontSize: 13, color: "BFD4FF", margin: 0,
    });

    const benefits = [
      { icon: "⏰", stat: "3x Faster",     title: "Order Processing",     desc: "Voice ordering + digital catalog eliminates manual order writing entirely", color: "2563EB" },
      { icon: "💰", stat: "0 Lost Orders",  title: "Revenue Protection",   desc: "Every order captured digitally — no missed calls, no misread scrawls",    color: "10B981" },
      { icon: "📊", stat: "100% Visibility", title: "Business Intelligence", desc: "See which products, customers & staff generate the most revenue",         color: "F59E0B" },
      { icon: "🚀", stat: "2x Faster",      title: "Sales Rep Productivity", desc: "Field reps place more orders per day with voice & mobile optimization",   color: "7C3AED" },
      { icon: "📉", stat: "Zero",           title: "Stock Overshoot",       desc: "Real-time inventory prevents selling what you don't have",                 color: "EF4444" },
      { icon: "🤝", stat: "Better",         title: "Customer Relationships", desc: "Retailers self-serve on a portal — fewer support calls, more trust",      color: "0D9488" },
    ];

    benefits.forEach((b, i) => {
      const col = i % 3, row = Math.floor(i / 3);
      const x = 0.5 + col * 3.2;
      const y = 1.5 + row * 1.85;

      s.addShape("rect", { x, y, w: 3.0, h: 1.7, fill: { color: "1E2D5A" }, line: { color: b.color, width: 2 }, rectRadius: 0.1 });
      s.addText(b.icon, { x: x + 0.12, y: y + 0.12, w: 0.5, h: 0.5, fontSize: 22, margin: 0 });
      s.addText(b.stat, { x: x + 0.7, y: y + 0.1, w: 2.2, h: 0.45, fontSize: 20, bold: true, color: b.color, fontFace: "Arial Black", margin: 0 });
      s.addText(b.title, { x: x + 0.12, y: y + 0.62, w: 2.7, h: 0.32, fontSize: 11, bold: true, color: "FFD166", margin: 0 });
      s.addText(b.desc,  { x: x + 0.12, y: y + 0.94, w: 2.7, h: 0.65, fontSize: 9, color: "BFD4FF", lineSpacingMultiple: 1.3, margin: 0 });
    });

    addSlideNum(s, 16);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 17 – COMPETITIVE ADVANTAGES
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addText("Why Seerweb OMS vs. Generic Software", {
      x: 0.5, y: 0.25, w: 9, h: 0.6, fontSize: 26, bold: true, color: C.navy, fontFace: "Arial Black", margin: 0,
    });

    const cols = ["Feature", "Seerweb OMS", "Generic OMS", "Manual Process"];
    const rows = [
      ["Voice Order Entry (Hindi+English)", "✅ Built-in", "❌ No", "❌ No"],
      ["Garments Size-Matrix Ordering",      "✅ Dedicated Module", "⚠️ Limited", "❌ No"],
      ["Multi-Role Access (4 roles)",        "✅ Yes", "⚠️ 2-3 roles", "❌ No"],
      ["Real-Time Inventory",                "✅ Live", "✅ Yes", "❌ No"],
      ["PDF Invoice with Size Tables",       "✅ Auto-generated", "⚠️ Basic", "❌ Manual"],
      ["Staff Performance Analytics",        "✅ Yes", "❌ No", "❌ No"],
      ["WhatsApp Catalog Sharing",           "✅ Yes", "❌ No", "❌ No"],
      ["India-Specific (Hindi AI)",          "✅ Yes", "❌ No", "❌ No"],
      ["Custom Product Fields per Business", "✅ Yes", "⚠️ Limited", "❌ No"],
    ];

    const colW = [3.6, 2.1, 2.1, 1.8];
    const startX = 0.5;
    const startY = 1.1;

    // Header
    cols.forEach((col, ci) => {
      const x = startX + colW.slice(0, ci).reduce((a, b) => a + b, 0);
      s.addShape("rect", { x, y: startY, w: colW[ci], h: 0.42, fill: { color: ci === 1 ? C.royalBlue : C.navy } });
      s.addText(col, { x, y: startY, w: colW[ci], h: 0.42, fontSize: 10.5, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
    });

    rows.forEach((row, ri) => {
      const y = startY + 0.42 + ri * 0.41;
      row.forEach((cell, ci) => {
        const x = startX + colW.slice(0, ci).reduce((a, b) => a + b, 0);
        const bg = ri % 2 === 0 ? C.white : C.offWhite;
        const textColor = cell.startsWith("✅") ? "10B981" : cell.startsWith("❌") ? "EF4444" : cell.startsWith("⚠️") ? "F59E0B" : C.darkGray;
        s.addShape("rect", { x, y, w: colW[ci], h: 0.4, fill: { color: ci === 1 ? "EFF6FF" : bg }, line: { color: C.lightGray, width: 0.5 } });
        s.addText(cell, { x, y, w: colW[ci], h: 0.4, fontSize: 9.5, color: textColor, align: ci === 0 ? "left" : "center", valign: "middle", margin: ci === 0 ? 4 : 0 });
      });
    });

    addSlideNum(s, 17);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 18 – ROI / BUSINESS VALUE
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.white };

    s.addShape("rect", { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.amber } });

    s.addText("Return on Investment", {
      x: 0.45, y: 0.25, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.navy, fontFace: "Arial Black", margin: 0,
    });
    s.addText("Quantifiable business value delivered from Day 1", {
      x: 0.45, y: 0.85, w: 9, h: 0.3, fontSize: 13, color: C.midGray, margin: 0,
    });

    // Top KPI row
    const kpis = [
      ["₹0",       "Order Loss",       "Zero orders lost to communication errors"],
      ["70%",      "Admin Time Saved",  "Less time on manual order entry & follow-ups"],
      ["3–5x",     "Faster Onboarding", "Retailers self-register and start ordering immediately"],
      ["∞",        "Scalability",       "Add unlimited retailers & staff with no extra infra cost"],
    ];
    kpis.forEach(([val, title, desc], i) => {
      const x = 0.45 + i * 2.4;
      addCard(s, x, 1.35, 2.25, 1.4, { fill: C.offWhite });
      s.addText(val,   { x, y: 1.45, w: 2.25, h: 0.55, fontSize: 28, bold: true, color: C.royalBlue, fontFace: "Arial Black", align: "center", margin: 0 });
      s.addText(title, { x, y: 1.98, w: 2.25, h: 0.32, fontSize: 10, bold: true, color: C.darkGray, align: "center", margin: 0 });
      s.addText(desc,  { x: x + 0.1, y: 2.3, w: 2.05, h: 0.35, fontSize: 8.5, color: C.midGray, align: "center", margin: 0 });
    });

    // ROI scenario table
    addCard(s, 0.45, 3.0, 9.1, 2.3, { fill: C.offWhite });
    s.addText("📈  Scenario: Mid-size Garment Dealer with 50 Retailers & 5 Sales Executives", {
      x: 0.6, y: 3.08, w: 8.8, h: 0.3, fontSize: 10.5, bold: true, color: C.navy, margin: 0,
    });

    const scenarios = [
      ["Before OMS", "₹18,000–25,000/month in admin overhead (3 staff handling manual orders)"],
      ["After OMS",  "1 coordinator + automated tracking = ₹12,000/month savings"],
      ["Field Sales", "Each executive handles 15% more retailers/day with voice ordering"],
      ["Lost Order Recovery", "Estimate ₹30,000–₹80,000/month recovered from missed manual orders"],
      ["Net ROI",   "Software pays for itself within 30–60 days for most dealers"],
    ];
    scenarios.forEach(([label, value], i) => {
      const y = 3.5 + i * 0.36;
      s.addShape("rect", { x: 0.55, y: y - 0.02, w: 1.8, h: 0.3, fill: { color: i === 4 ? C.teal : C.royalBlue + "22" }, rectRadius: 0.04 });
      s.addText(label, { x: 0.55, y: y - 0.02, w: 1.8, h: 0.3, fontSize: 9, bold: true, color: i === 4 ? C.white : C.royalBlue, align: "center", valign: "middle", margin: 0 });
      s.addText(value, { x: 2.45, y: y - 0.02, w: 7.0, h: 0.3, fontSize: 9.5, color: C.darkGray, valign: "middle", margin: 4 });
    });

    addSlideNum(s, 18);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 19 – USE CASES / WHO IT'S FOR
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addText("Who Is Seerweb OMS Built For?", {
      x: 0.5, y: 0.25, w: 9, h: 0.6, fontSize: 26, bold: true, color: C.navy, fontFace: "Arial Black", margin: 0,
    });
    s.addText("Purpose-built for Indian wholesale & distribution businesses", {
      x: 0.5, y: 0.85, w: 9, h: 0.3, fontSize: 13, color: C.midGray, margin: 0,
    });

    const useCases = [
      {
        emoji: "👗", title: "Garments Wholesalers",
        points: ["Manage designs, sizes & colors", "Wholesale catalog with set ordering", "WhatsApp catalog sharing", "Booking vs ready-stock tracking"],
        color: C.amber,
      },
      {
        emoji: "📱", title: "Mobile/Electronics Distributors",
        points: ["Track models, RAM & storage variants", "Retailer-wise order history", "Live stock management", "Staff-wise sales reports"],
        color: C.royalBlue,
      },
      {
        emoji: "🛒", title: "FMCG / General Distributors",
        points: ["Simple product catalog", "Voice ordering for field reps", "Multi-retailer management", "Daily order tracking & analytics"],
        color: C.teal,
      },
      {
        emoji: "🏭", title: "Any Product Distributor",
        points: ["Customizable product fields per business", "Multi-role access for any team structure", "Scales from 10 to 500+ retailers", "No IT team required to manage"],
        color: "7C3AED",
      },
    ];

    useCases.forEach((uc, i) => {
      const x = 0.5 + i * 2.4;
      addCard(s, x, 1.35, 2.2, 4.0, { fill: C.white });
      s.addShape("rect", { x, y: 1.35, w: 2.2, h: 0.65, fill: { color: uc.color }, rectRadius: 0.08 });
      s.addShape("rect", { x, y: 1.8, w: 2.2, h: 0.25, fill: { color: uc.color } });
      s.addText(uc.emoji, { x, y: 1.35, w: 0.7, h: 0.65, fontSize: 24, align: "center", valign: "middle", margin: 0 });
      s.addText(uc.title, { x: x + 0.68, y: 1.42, w: 1.45, h: 0.52, fontSize: 10.5, bold: true, color: C.white, lineSpacingMultiple: 1.2, margin: 0 });
      uc.points.forEach((pt, pi) => {
        s.addShape("rect", { x: x + 0.15, y: 2.2 + pi * 0.72, w: 0.2, h: 0.2, fill: { color: uc.color }, rectRadius: 0.02 });
        s.addText(pt, { x: x + 0.42, y: 2.16 + pi * 0.72, w: 1.65, h: 0.5, fontSize: 9.5, color: C.darkGray, margin: 0 });
      });
    });

    addSlideNum(s, 19);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 20 – FUTURE ROADMAP
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };

    s.addText("Product Roadmap", {
      x: 0.5, y: 0.25, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white, fontFace: "Arial Black", margin: 0,
    });
    s.addText("We're constantly evolving — here's what's coming next", {
      x: 0.5, y: 0.85, w: 9, h: 0.3, fontSize: 13, color: "BFD4FF", margin: 0,
    });

    // Timeline
    const phases = [
      {
        phase: "Q3 2025", label: "Current",
        items: ["Voice ordering (Hindi+English)", "Garments ERP module", "Multi-role dashboards", "PDF invoice generation", "Excel import/export"],
        color: "10B981", done: true,
      },
      {
        phase: "Q4 2025", label: "In Progress",
        items: ["Android native app (React Native)", "WhatsApp order notifications", "Tally ERP integration", "Barcode / QR product scanning"],
        color: C.amber, done: false,
      },
      {
        phase: "Q1 2026", label: "Planned",
        items: ["Payment tracking & dues", "GPS route optimization for reps", "Advanced AI reorder predictions", "Multi-warehouse support"],
        color: "7C3AED", done: false,
      },
      {
        phase: "Q2 2026", label: "Vision",
        items: ["B2B marketplace for retailers", "WhatsApp chatbot ordering", "IoT integration for warehouses", "Pan-India logistics API"],
        color: "BFD4FF", done: false,
      },
    ];

    // Timeline bar
    s.addShape("rect", { x: 0.5, y: 2.3, w: 9, h: 0.04, fill: { color: "2D3F6E" } });

    phases.forEach((ph, i) => {
      const x = 0.5 + i * 2.32;
      const y = 1.55;

      // Dot on timeline
      s.addShape("ellipse", { x: x + 0.85, y: 2.2, w: 0.28, h: 0.28, fill: { color: ph.color } });

      addCard(s, x, y, 2.2, 0.6, { fill: "1E2D5A" });
      s.addShape("rect", { x, y, w: 2.2, h: 0.6, fill: { color: ph.color + (ph.done ? "" : "44") }, line: { color: ph.color, width: ph.done ? 2 : 1 }, rectRadius: 0.06 });
      s.addText(ph.phase, { x, y: y + 0.04, w: 2.2, h: 0.28, fontSize: 12, bold: true, color: ph.done ? C.white : ph.color, align: "center", margin: 0 });
      s.addText(ph.label, { x, y: y + 0.33, w: 2.2, h: 0.22, fontSize: 9, color: ph.done ? "FFFFFF99" : ph.color + "CC", align: "center", margin: 0 });

      // Items list
      ph.items.forEach((item, ii) => {
        const iy = 2.65 + ii * 0.48;
        s.addShape("rect", { x: x + 0.1, y: iy + 0.08, w: 0.16, h: 0.16, fill: { color: ph.color }, rectRadius: 0.02 });
        s.addText(item, { x: x + 0.33, y: iy + 0.04, w: 1.8, h: 0.38, fontSize: 9, color: "BFD4FF", margin: 0 });
      });
    });

    addSlideNum(s, 20);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 21 – GETTING STARTED / ONBOARDING
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };

    s.addText("Getting Started is Simple", {
      x: 0.5, y: 0.25, w: 9, h: 0.6, fontSize: 26, bold: true, color: C.navy, fontFace: "Arial Black", margin: 0,
    });
    s.addText("From sign-up to first live order — most dealers are operational in under 24 hours", {
      x: 0.5, y: 0.85, w: 9, h: 0.3, fontSize: 13, color: C.midGray, margin: 0,
    });

    const steps = [
      { n: "01", title: "Contact Us",          sub: "Reach out via phone, email or the website demo form. We respond within 2 hours.",   time: "Day 1",  icon: "📞" },
      { n: "02", title: "Account Setup",       sub: "We create your dealer account, configure business type & set up the platform for you.", time: "Day 1",  icon: "⚙️" },
      { n: "03", title: "Product Import",      sub: "Upload your product catalog via Excel. Or we assist with the import — whichever is easier.", time: "Day 1",  icon: "📦" },
      { n: "04", title: "Add Retailers",       sub: "Invite your retailers. They self-register with a username and start ordering immediately.", time: "Day 2",  icon: "🏪" },
      { n: "05", title: "Train Your Team",     sub: "30-minute Zoom training for your sales team. Voice ordering demo included.", time: "Day 2",  icon: "🎓" },
      { n: "06", title: "Go Live!",            sub: "Your team starts creating orders. You watch the dashboard fill up in real time.", time: "Day 2",  icon: "🚀" },
    ];

    // Timeline line
    s.addShape("rect", { x: 0.7, y: 3.2, w: 8.7, h: 0.04, fill: { color: C.lightGray } });

    steps.forEach((st, i) => {
      const x = 0.5 + i * 1.6;
      const isLeft = i % 2 === 0;
      const cardY = isLeft ? 1.35 : 3.5;

      addCard(s, x, cardY, 1.45, 1.65, { fill: C.white });

      // Step number circle
      s.addShape("ellipse", { x: x + 0.47, y: isLeft ? 3.08 : 3.28, w: 0.22, h: 0.22, fill: { color: C.royalBlue } });
      s.addText(st.n, { x: x + 0.47, y: isLeft ? 3.08 : 3.28, w: 0.22, h: 0.22, fontSize: 7.5, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });

      // Connector line
      s.addShape("rect", { x: x + 0.57, y: isLeft ? cardY + 1.65 : 3.5, w: 0.03, h: isLeft ? 3.08 - cardY - 1.65 : 3.28 - 3.5 + 0.22, fill: { color: C.royalBlue } });

      s.addText(st.icon,  { x, y: cardY + 0.08, w: 1.45, h: 0.45, fontSize: 18, align: "center", margin: 0 });
      s.addText(st.title, { x: x + 0.05, y: cardY + 0.55, w: 1.35, h: 0.32, fontSize: 9.5, bold: true, color: C.darkGray, align: "center", margin: 0 });
      s.addText(st.sub,   { x: x + 0.05, y: cardY + 0.87, w: 1.35, h: 0.72, fontSize: 8, color: C.midGray, align: "center", lineSpacingMultiple: 1.2, margin: 0 });

      // Time badge
      s.addShape("rect", { x: x + 0.3, y: cardY - 0.01, w: 0.85, h: 0.22, fill: { color: C.teal + "22" }, rectRadius: 0.04 });
      s.addText(st.time, { x: x + 0.3, y: cardY - 0.01, w: 0.85, h: 0.22, fontSize: 7.5, bold: true, color: C.teal, align: "center", valign: "middle", margin: 0 });
    });

    addSlideNum(s, 21);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLIDE 22 – CONTACT / CTA
  // ══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };

    // Left accent panel
    s.addShape("rect", { x: 0, y: 0, w: 4.2, h: 5.625, fill: { color: C.royalBlue } });

    s.addText("Ready to Transform\nYour Business?", {
      x: 0.3, y: 1.0, w: 3.6, h: 2.0, fontSize: 28, bold: true, color: C.white, fontFace: "Arial Black", lineSpacingMultiple: 1.2, margin: 0,
    });
    s.addText("Book a free live demo and see Seerweb OMS in action with your own products.", {
      x: 0.3, y: 3.1, w: 3.6, h: 0.8, fontSize: 12, color: "BFD4FF", lineSpacingMultiple: 1.4, margin: 0,
    });

    // CTA button
    s.addShape("rect", { x: 0.3, y: 4.1, w: 2.8, h: 0.65, fill: { color: C.white }, rectRadius: 0.32 });
    s.addText("🗓  Book a Free Demo", {
      x: 0.3, y: 4.1, w: 2.8, h: 0.65, fontSize: 13, bold: true, color: C.royalBlue, align: "center", valign: "middle", margin: 0,
    });

    // Right: contact info
    s.addText("Get In Touch", {
      x: 4.6, y: 0.5, w: 5.1, h: 0.5, fontSize: 22, bold: true, color: C.white, fontFace: "Arial Black", margin: 0,
    });

    const contacts = [
      ["🌐", "Website",   "www.seerweberp.com"],
      ["📧", "Email",     "info@seerweberp.com"],
      ["📞", "Phone",     "+91 98765 43210"],
      ["💬", "WhatsApp",  "Chat with us on WhatsApp"],
    ];
    contacts.forEach(([icon, label, value], i) => {
      const y = 1.3 + i * 0.9;
      addCard(s, 4.6, y, 5.1, 0.72, { fill: "1E2D5A" });
      s.addText(icon,  { x: 4.72, y: y + 0.18, w: 0.35, h: 0.35, fontSize: 16, margin: 0 });
      s.addText(label, { x: 5.12, y: y + 0.08, w: 4.4, h: 0.26, fontSize: 10, bold: true, color: "BFD4FF", margin: 0 });
      s.addText(value, { x: 5.12, y: y + 0.35, w: 4.4, h: 0.28, fontSize: 11.5, color: C.white, margin: 0 });
    });

    // Logo tagline
    s.addText("Seerweb ERP Solutions Pvt Ltd", {
      x: 4.6, y: 4.95, w: 5.1, h: 0.28, fontSize: 11, bold: true, color: "6B8EC2", margin: 0,
    });
    s.addText("TallyPrime · ERP · OMS · Enterprise Software", {
      x: 4.6, y: 5.2, w: 5.1, h: 0.22, fontSize: 9, color: "4A5F85", margin: 0,
    });
  }

  // ─── Write file ───────────────────────────────────────────────────────────
  const outPath = "/mnt/user-data/outputs/Seerweb_OMS_Sales_Presentation.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("✅  Saved:", outPath);
}

buildPresentation().catch(console.error);