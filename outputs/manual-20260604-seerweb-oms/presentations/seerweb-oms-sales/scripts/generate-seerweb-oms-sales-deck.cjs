const pptxgen = require("pptxgenjs");
const path = require("path");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Seerweb";
pptx.subject = "Seerweb OMS Sales Presentation";
pptx.title = "Seerweb OMS Sales Presentation";
pptx.company = "Seerweb";
pptx.lang = "en-US";
pptx.theme = {
  headFontFace: "Aptos Display",
  bodyFontFace: "Aptos",
  lang: "en-US",
};
pptx.defineLayout({ name: "CUSTOM_WIDE", width: 13.333, height: 7.5 });
pptx.layout = "CUSTOM_WIDE";

const C = {
  navy: "0F172A",
  blue: "2563EB",
  blue2: "1D4ED8",
  cyan: "06B6D4",
  green: "16A34A",
  amber: "F59E0B",
  red: "EF4444",
  slate: "475569",
  muted: "64748B",
  light: "F8FAFC",
  line: "E2E8F0",
  white: "FFFFFF",
  purple: "7C3AED",
};

const W = 13.333;
const H = 7.5;
const M = 0.45;

function addBg(slide, dark = false) {
  slide.background = { color: dark ? C.navy : C.light };
  if (dark) {
    slide.addShape(pptx.ShapeType.arc, { x: 9.2, y: -0.5, w: 5, h: 5, line: { color: C.blue, transparency: 100 }, fill: { color: "1E3A8A", transparency: 25 }, adjustPoint: 0.2 });
    slide.addShape(pptx.ShapeType.arc, { x: -1.5, y: 5.2, w: 4.5, h: 4.5, line: { color: C.cyan, transparency: 100 }, fill: { color: "155E75", transparency: 45 }, adjustPoint: 0.2 });
  }
}

function footer(slide, n, dark = false) {
  slide.addText("Seerweb OMS", { x: M, y: 7.13, w: 2.4, h: 0.2, fontSize: 8, color: dark ? "CBD5E1" : C.muted, margin: 0 });
  slide.addText(String(n).padStart(2, "0"), { x: 12.42, y: 7.1, w: 0.45, h: 0.24, fontSize: 8, bold: true, color: dark ? C.white : C.slate, align: "right", margin: 0 });
}

function title(slide, t, sub, n, dark = false) {
  addBg(slide, dark);
  slide.addText(t, { x: M, y: 0.42, w: 8.2, h: 0.45, fontSize: 23, bold: true, color: dark ? C.white : C.navy, margin: 0 });
  if (sub) slide.addText(sub, { x: M, y: 0.92, w: 7.7, h: 0.28, fontSize: 9.5, color: dark ? "CBD5E1" : C.muted, margin: 0 });
  footer(slide, n, dark);
}

function pill(slide, text, x, y, w, color = C.blue) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h: 0.36, rectRadius: 0.08, line: { color, transparency: 75 }, fill: { color, transparency: 88 } });
  slide.addText(text, { x: x + 0.12, y: y + 0.095, w: w - 0.24, h: 0.14, fontSize: 8.5, bold: true, color, margin: 0, align: "center" });
}

function box(slide, x, y, w, h, opts = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: 0.06,
    line: { color: opts.line || C.line, transparency: opts.lineT ?? 0 },
    fill: { color: opts.fill || C.white, transparency: opts.transparency ?? 0 },
    shadow: opts.shadow ? { type: "outer", color: "94A3B8", opacity: 0.18, blur: 1, angle: 45, distance: 1 } : undefined,
  });
}

function screenshot(slide, x, y, w, h, label, caption, callouts = []) {
  box(slide, x, y, w, h, { fill: "F1F5F9", line: "CBD5E1", shadow: true });
  slide.addShape(pptx.ShapeType.rect, { x: x + 0.18, y: y + 0.18, w: w - 0.36, h: 0.32, line: { color: "CBD5E1" }, fill: { color: C.white } });
  slide.addText(label, { x: x + 0.25, y: y + 0.27, w: w - 0.5, h: 0.12, fontSize: 7.5, bold: true, color: C.slate, margin: 0 });
  for (let i = 0; i < 4; i++) {
    const yy = y + 0.75 + i * ((h - 1.25) / 4);
    slide.addShape(pptx.ShapeType.roundRect, { x: x + 0.28, y: yy, w: w - 0.56, h: 0.34, rectRadius: 0.04, line: { color: "E2E8F0" }, fill: { color: i % 2 ? C.white : "EAF2FF" } });
  }
  slide.addText(caption, { x, y: y + h + 0.08, w, h: 0.22, fontSize: 8.3, color: C.muted, italic: true, margin: 0, align: "center" });
  callouts.forEach((c, idx) => {
    const cx = x + c[0] * w;
    const cy = y + c[1] * h;
    slide.addShape(pptx.ShapeType.ellipse, { x: cx - 0.13, y: cy - 0.13, w: 0.26, h: 0.26, line: { color: C.white, width: 1 }, fill: { color: c[2] || C.blue } });
    slide.addText(String(idx + 1), { x: cx - 0.05, y: cy - 0.06, w: 0.1, h: 0.1, fontSize: 7, bold: true, color: C.white, margin: 0, align: "center" });
  });
}

function bullets(slide, items, x, y, w, color = C.navy) {
  items.forEach((b, i) => {
    slide.addShape(pptx.ShapeType.ellipse, { x, y: y + i * 0.48 + 0.06, w: 0.13, h: 0.13, line: { color: C.blue }, fill: { color: C.blue } });
    slide.addText(b, { x: x + 0.25, y: y + i * 0.48, w, h: 0.28, fontSize: 11, color, breakLine: false, margin: 0, fit: "shrink" });
  });
}

function metric(slide, x, y, w, label, value, color) {
  box(slide, x, y, w, 0.95, { fill: C.white, line: C.line, shadow: true });
  slide.addText(value, { x: x + 0.18, y: y + 0.18, w: w - 0.36, h: 0.28, fontSize: 18, bold: true, color, margin: 0 });
  slide.addText(label, { x: x + 0.18, y: y + 0.55, w: w - 0.36, h: 0.18, fontSize: 8.5, color: C.muted, margin: 0 });
}

const slides = [];
function s() { const slide = pptx.addSlide(); slides.push(slide); return slide; }

// 1
{
  const slide = s(); addBg(slide, true);
  slide.addText("SEERWEB OMS", { x: 0.62, y: 0.55, w: 3.2, h: 0.28, fontSize: 12, bold: true, color: "93C5FD", margin: 0 });
  slide.addText("Digitize Wholesale Ordering. Accelerate Sales. Control Every Order.", { x: 0.62, y: 1.25, w: 6.25, h: 1.45, fontSize: 31, bold: true, color: C.white, fit: "shrink", margin: 0 });
  slide.addText("A complete order management platform for dealers, retailers, sales executives, and admins.", { x: 0.62, y: 3.0, w: 5.8, h: 0.52, fontSize: 14, color: "CBD5E1", margin: 0 });
  pill(slide, "Sales Presentation", 0.62, 3.85, 1.65, C.cyan);
  pill(slide, "Client Demo Deck", 2.42, 3.85, 1.7, C.green);
  screenshot(slide, 7.25, 1.05, 5.05, 4.75, "Dashboard / Login Screenshot", "Hero product screenshot placement", [[0.2, 0.25, C.green], [0.77, 0.55, C.blue]]);
  footer(slide, 1, true);
}

// 2
{
  const slide = s(); title(slide, "Company Branding", "Position Seerweb as a reliable digital transformation partner.", 2);
  slide.addText("Built for distribution businesses that need speed, visibility, and control across every order touchpoint.", { x: 0.7, y: 1.55, w: 7.1, h: 0.75, fontSize: 22, bold: true, color: C.navy, margin: 0 });
  [["Industry-focused OMS", C.blue], ["Dealer and retailer workflows", C.green], ["Configurable for wholesale operations", C.purple]].forEach((v, i) => {
    metric(slide, 0.75 + i * 4.05, 3.05, 3.4, v[0], i === 0 ? "B2B" : i === 1 ? "Multi-role" : "Scalable", v[1]);
  });
  slide.addText("Branding section: add company logo, brand colors, implementation credentials, support promise, and client-specific vertical messaging.", { x: 1.0, y: 5.0, w: 11.3, h: 0.42, fontSize: 12, color: C.slate, align: "center", margin: 0 });
}

// 3
{
  const slide = s(); title(slide, "Project Overview", "A single operating layer for catalog, orders, sales team, retailers, and reports.", 3);
  screenshot(slide, 6.85, 1.25, 5.75, 4.65, "Dealer Dashboard", "Real-time command center for dealer operations", [[0.2, 0.2, C.blue], [0.55, 0.55, C.green], [0.8, 0.72, C.amber]]);
  slide.addText("Seerweb OMS centralizes product catalogs, customer ordering, sales executive activity, dealer operations, and order analytics.", { x: 0.75, y: 1.55, w: 5.25, h: 0.75, fontSize: 18, bold: true, color: C.navy, margin: 0, fit: "shrink" });
  bullets(slide, ["Dealer dashboard with revenue and order KPIs", "Retailer self-service portal for repeat orders", "Sales executive workflow for field order capture", "Admin panel for platform-level dealer management"], 0.8, 2.75, 5.4);
}

// 4
{
  const slide = s(); title(slide, "The Business Problem", "Wholesale order operations often break down before revenue can be captured.", 4);
  const problems = [
    ["Manual order taking", "Phone, WhatsApp, and paper entries slow teams down."],
    ["Order mistakes", "Wrong customer, quantity, size, or product data creates rework."],
    ["Low visibility", "Managers lack real-time order, revenue, and staff performance data."],
    ["Delayed follow-up", "Retailers wait for updates and sales teams lose productive time."],
  ];
  problems.forEach((p, i) => {
    const x = 0.8 + (i % 2) * 6.1, y = 1.45 + Math.floor(i / 2) * 2.15;
    box(slide, x, y, 5.4, 1.45, { fill: C.white, line: C.line, shadow: true });
    slide.addText(p[0], { x: x + 0.28, y: y + 0.25, w: 4.9, h: 0.22, fontSize: 15, bold: true, color: C.navy, margin: 0 });
    slide.addText(p[1], { x: x + 0.28, y: y + 0.72, w: 4.8, h: 0.36, fontSize: 10.5, color: C.slate, margin: 0, fit: "shrink" });
  });
}

// 5
{
  const slide = s(); title(slide, "The Solution", "One connected platform that turns scattered order activity into a controlled sales engine.", 5);
  screenshot(slide, 0.75, 1.45, 4.55, 3.65, "Product Experience", "Dashboard + order + report collage", [[0.22, 0.2, C.blue], [0.68, 0.62, C.green]]);
  const nodes = ["Catalog setup", "Retailer onboarding", "Field order capture", "Order tracking", "Analytics & export"];
  nodes.forEach((n, i) => {
    box(slide, 6.0, 1.25 + i * 0.82, 5.7, 0.52, { fill: i % 2 ? "F8FAFC" : "EFF6FF", line: "BFDBFE" });
    slide.addText(`${i + 1}. ${n}`, { x: 6.25, y: 1.4 + i * 0.82, w: 4.8, h: 0.14, fontSize: 12, bold: true, color: i % 2 ? C.navy : C.blue, margin: 0 });
  });
  slide.addText("Outcome: faster order capture, fewer mistakes, higher team productivity, and better management visibility.", { x: 6.0, y: 5.75, w: 5.8, h: 0.35, fontSize: 13, bold: true, color: C.green, margin: 0 });
}

// 6
{
  const slide = s(); title(slide, "Users & Roles", "Every user gets the right interface, permissions, and workflow.", 6);
  const roles = [
    ["Super Admin", "Manage dealers and platform access", C.blue],
    ["Dealer", "Manage products, retailers, staff, orders, settings", C.green],
    ["Sales Staff", "Create field orders and track customers", C.amber],
    ["Retailer", "Browse products, place orders, track status", C.purple],
  ];
  roles.forEach((r, i) => {
    box(slide, 0.7 + i * 3.15, 1.55, 2.65, 3.85, { fill: C.white, line: r[2], lineT: 35, shadow: true });
    slide.addShape(pptx.ShapeType.ellipse, { x: 1.55 + i * 3.15, y: 2.0, w: 0.95, h: 0.95, line: { color: r[2], transparency: 30 }, fill: { color: r[2], transparency: 82 } });
    slide.addText(r[0], { x: 0.95 + i * 3.15, y: 3.18, w: 2.15, h: 0.24, fontSize: 15, bold: true, color: C.navy, margin: 0, align: "center" });
    slide.addText(r[1], { x: 0.95 + i * 3.15, y: 3.75, w: 2.15, h: 0.76, fontSize: 10, color: C.slate, margin: 0, align: "center", fit: "shrink" });
  });
}

// 7
{
  const slide = s(); title(slide, "Role-Based Access & Permissions", "Control what each stakeholder can see and do.", 7);
  const headers = ["Module", "Admin", "Dealer", "Staff", "Retailer"];
  const rows = [
    ["Dealer Management", "Full", "-", "-", "-"],
    ["Products & Catalog", "-", "Full", "View", "Browse"],
    ["Retailers", "-", "Full", "Assigned", "Own account"],
    ["Orders", "Monitor", "Full", "Create/View", "Create/View"],
    ["Reports", "Platform", "Business", "Own sales", "Own orders"],
  ];
  headers.forEach((h, i) => slide.addText(h, { x: 0.75 + i * 2.45, y: 1.35, w: 2.2, h: 0.28, fontSize: 11, bold: true, color: C.white, margin: 0, align: "center", fill: { color: C.blue } }));
  rows.forEach((r, ri) => r.forEach((cell, ci) => {
    const x = 0.75 + ci * 2.45, y = 1.78 + ri * 0.62;
    slide.addShape(pptx.ShapeType.rect, { x, y, w: 2.2, h: 0.48, line: { color: C.line }, fill: { color: ci === 0 ? "F1F5F9" : C.white } });
    slide.addText(cell, { x: x + 0.08, y: y + 0.14, w: 2.04, h: 0.12, fontSize: 9.5, color: ci === 0 ? C.navy : C.slate, bold: ci === 0, margin: 0, align: ci === 0 ? "left" : "center" });
  }));
  screenshot(slide, 9.8, 5.2, 2.4, 1.15, "Sidebar", "Role-specific navigation");
}

// 8
{
  const slide = s(); title(slide, "End-to-End Workflow", "From catalog setup to revenue visibility.", 8);
  const steps = ["Catalog", "Retailers", "Order Capture", "Checkout", "Status Tracking", "Analytics"];
  steps.forEach((st, i) => {
    const x = 0.65 + i * 2.05;
    slide.addShape(pptx.ShapeType.ellipse, { x, y: 2.45, w: 1.05, h: 1.05, line: { color: C.blue }, fill: { color: i % 2 ? "DCFCE7" : "DBEAFE" } });
    slide.addText(String(i + 1), { x: x + 0.37, y: 2.75, w: 0.3, h: 0.18, fontSize: 15, bold: true, color: i % 2 ? C.green : C.blue, margin: 0, align: "center" });
    slide.addText(st, { x: x - 0.25, y: 3.72, w: 1.55, h: 0.26, fontSize: 10.5, bold: true, color: C.navy, margin: 0, align: "center" });
    if (i < steps.length - 1) slide.addShape(pptx.ShapeType.chevron, { x: x + 1.22, y: 2.72, w: 0.5, h: 0.45, line: { color: C.line }, fill: { color: C.line } });
  });
  slide.addText("Recommended demo flow: Login as Dealer → open Dashboard → Take Order → Checkout → review Orders/Reports → show Admin and Retailer portals.", { x: 1.15, y: 5.45, w: 11, h: 0.36, fontSize: 12, color: C.slate, align: "center", margin: 0 });
}

// 9
{
  const slide = s(); title(slide, "Dealer Dashboard Overview", "Management visibility for orders, revenue, products, retailers, and staff.", 9);
  screenshot(slide, 0.65, 1.25, 7.0, 4.85, "Dealer Dashboard Screenshot", "Real-time command center for dealer operations", [[0.18, 0.22, C.blue], [0.42, 0.22, C.amber], [0.66, 0.22, C.green], [0.72, 0.64, C.purple]]);
  bullets(slide, ["Total Orders, Pending Orders, Total Revenue, Today's Orders", "Orders and revenue over the last 7 days", "Order status and revenue-by-status charts", "Quick links to products, retailers, staff, and order creation"], 8.25, 1.55, 4.1);
}

// 10
{
  const slide = s(); title(slide, "Product & Catalog Management", "Keep every product, price, image, and garment attribute under control.", 10);
  screenshot(slide, 0.65, 1.25, 5.5, 4.05, "Manage Products", "Centralized catalog and inventory management", [[0.25, 0.28, C.blue], [0.8, 0.32, C.green]]);
  screenshot(slide, 6.65, 1.25, 2.75, 1.8, "Add Product", "Fast product setup");
  screenshot(slide, 9.65, 1.25, 2.75, 1.8, "Size Settings", "Garment size-wise configuration");
  bullets(slide, ["Product images, pricing, stock and brand/model details", "Garment catalog support with size and attribute settings", "Custom fields for business-specific product data"], 6.8, 3.65, 5.1);
}

// 11
{
  const slide = s(); title(slide, "Smart Order Capture", "Help field teams book accurate orders faster.", 11);
  screenshot(slide, 5.95, 1.15, 6.55, 5.05, "Take Order Screen", "Fast order booking for sales teams and dealers", [[0.18, 0.22, C.green], [0.5, 0.42, C.blue], [0.78, 0.72, C.amber]]);
  slide.addText("Built around the real sales flow", { x: 0.75, y: 1.55, w: 4.5, h: 0.32, fontSize: 18, bold: true, color: C.navy, margin: 0 });
  bullets(slide, ["Select retailer before order creation", "Search product catalog quickly", "Add size-wise quantities into cart", "Place order with customer-linked totals"], 0.85, 2.35, 4.7);
}

// 12
{
  const slide = s(); title(slide, "Voice-Assisted Ordering", "Reduce friction when sales executives are on the move.", 12);
  screenshot(slide, 7.2, 1.35, 4.75, 3.65, "Voice Mic / Fallback Modal", "Voice-assisted order entry placement", [[0.5, 0.45, C.red]]);
  slide.addText("Field sales teams can spend less time typing and more time selling.", { x: 0.85, y: 1.55, w: 5.6, h: 0.52, fontSize: 20, bold: true, color: C.navy, margin: 0, fit: "shrink" });
  bullets(slide, ["Voice order hooks for product/quantity capture", "Fallback modal when speech needs confirmation", "Designed for busy wholesale counters and field visits"], 0.9, 2.65, 5.5);
}

// 13
{
  const slide = s(); title(slide, "Cart & Checkout Experience", "Prevent order leakage with customer-linked checkout and clear totals.", 13);
  screenshot(slide, 0.75, 1.3, 6.1, 4.5, "Cart / Checkout Screen", "Accurate customer-linked order confirmation", [[0.2, 0.2, C.green], [0.42, 0.52, C.blue], [0.8, 0.82, C.amber]]);
  bullets(slide, ["Retailer selection required for dealer/staff orders", "Grouped products with variant and quantity controls", "Summary panel with item count, total and checkout validation"], 7.55, 1.8, 4.65);
}

// 14
{
  const slide = s(); title(slide, "Retailer Self-Service Portal", "Let customers browse, order, and track status without manual follow-up.", 14);
  screenshot(slide, 0.8, 1.35, 3.2, 4.55, "Retailer Dashboard", "Retailer order status and quick actions", [[0.5, 0.22, C.blue]]);
  screenshot(slide, 4.55, 1.35, 3.2, 4.55, "Product Browse", "Self-service product ordering");
  screenshot(slide, 8.3, 1.35, 3.2, 4.55, "Retailer Orders", "Order history and tracking");
  slide.addText("Business value: retailers order anytime, while sales teams focus on growth accounts.", { x: 1.35, y: 6.38, w: 10.7, h: 0.28, fontSize: 12.5, bold: true, color: C.green, align: "center", margin: 0 });
}

// 15
{
  const slide = s(); title(slide, "Staff Performance Dashboard", "Give sales executives focus and give managers accountability.", 15);
  screenshot(slide, 5.75, 1.15, 6.65, 5.0, "Staff Dashboard", "Sales executive performance and customer activity", [[0.18, 0.2, C.blue], [0.62, 0.55, C.green], [0.78, 0.76, C.amber]]);
  bullets(slide, ["Assigned customers and direct call/map actions", "Today, monthly and all-time sales indicators", "Top customers and recent orders", "Sales report access"], 0.8, 1.65, 4.75);
}

// 16
{
  const slide = s(); title(slide, "Admin Panel", "Centralized platform control for dealer-led rollout.", 16);
  screenshot(slide, 0.75, 1.35, 5.4, 3.85, "Super Admin Dashboard", "Platform-level dealer monitoring", [[0.28, 0.28, C.blue]]);
  screenshot(slide, 7.05, 1.35, 4.6, 3.85, "Manage Dealers", "Dealer account administration", [[0.75, 0.24, C.green]]);
  slide.addText("Ideal for multi-dealer distribution networks that need controlled onboarding and centralized access management.", { x: 1.05, y: 5.9, w: 11, h: 0.34, fontSize: 13, bold: true, color: C.navy, align: "center", margin: 0 });
}

// 17
{
  const slide = s(); title(slide, "Reports & Analytics", "Turn order data into faster decisions.", 17);
  screenshot(slide, 6.75, 1.2, 5.65, 4.65, "Dashboard Charts / Sales Report", "Sales visibility across orders, revenue, and status", [[0.25, 0.45, C.blue], [0.73, 0.65, C.green]]);
  slide.addText("Decision visibility", { x: 0.8, y: 1.5, w: 4.6, h: 0.28, fontSize: 18, bold: true, color: C.navy, margin: 0 });
  bullets(slide, ["Order and revenue trends", "Pending order alerts", "Order status mix", "Top retailers by revenue", "Executive-level sales reports"], 0.9, 2.25, 4.9);
}

// 18
{
  const slide = s(); title(slide, "Integrations & Export", "Fit naturally into existing finance, catalog, and reporting workflows.", 18);
  const xs = [1.1, 4.7, 8.3];
  [["Seerweb OMS", C.blue], ["Tally / Accounting", C.green], ["PDF / Excel / Reports", C.purple]].forEach((n, i) => {
    box(slide, xs[i], 2.15, 2.8, 1.25, { fill: C.white, line: n[1], lineT: 25, shadow: true });
    slide.addText(n[0], { x: xs[i] + 0.25, y: 2.6, w: 2.3, h: 0.2, fontSize: 13, bold: true, color: n[1], margin: 0, align: "center" });
    if (i < 2) slide.addShape(pptx.ShapeType.chevron, { x: xs[i] + 3.05, y: 2.52, w: 0.55, h: 0.45, line: { color: C.line }, fill: { color: C.line } });
  });
  screenshot(slide, 4.15, 4.3, 4.9, 1.35, "Tally / Export Button or Order Export", "Business-ready export workflow");
  slide.addText("Current hooks include Tally export support, API-driven order/product data, and PDF/catalog export capability.", { x: 1.1, y: 5.95, w: 11, h: 0.28, fontSize: 12, color: C.slate, align: "center", margin: 0 });
}

// 19
{
  const slide = s(); title(slide, "Security, Reliability & Technology Stack", "Modern architecture with controlled access and scalable UI foundations.", 19);
  [["Security Controls", ["JWT/token-based API access", "Role-based routing and protected dashboards", "Separate experiences for admin, dealer, staff, retailer"]], ["Technology Stack", ["React + TypeScript + Vite", "Tailwind CSS + shadcn/ui", "Recharts, React Query, Supabase integration", "REST API integration with export utilities"]]].forEach((col, i) => {
    box(slide, 0.85 + i * 6.05, 1.45, 5.3, 4.5, { fill: C.white, line: C.line, shadow: true });
    slide.addText(col[0], { x: 1.15 + i * 6.05, y: 1.85, w: 4.6, h: 0.25, fontSize: 16, bold: true, color: i ? C.green : C.blue, margin: 0 });
    bullets(slide, col[1], 1.18 + i * 6.05, 2.55, 4.7);
  });
}

// 20
{
  const slide = s(); addBg(slide, true);
  slide.addText("Ready to modernize your order operations?", { x: 0.75, y: 0.95, w: 7.1, h: 0.78, fontSize: 28, bold: true, color: C.white, margin: 0, fit: "shrink" });
  slide.addText("Seerweb OMS helps businesses save time, reduce order errors, improve staff productivity, and unlock real-time revenue visibility.", { x: 0.78, y: 2.05, w: 6.3, h: 0.65, fontSize: 14, color: "CBD5E1", margin: 0 });
  [["Save time", "Faster order capture"], ["Reduce cost", "Less rework and manual follow-up"], ["Grow revenue", "Better retailer service and sales focus"]].forEach((m, i) => {
    box(slide, 0.78 + i * 2.25, 3.25, 1.95, 1.05, { fill: "1E293B", line: "334155" });
    slide.addText(m[0], { x: 0.95 + i * 2.25, y: 3.48, w: 1.6, h: 0.2, fontSize: 12, bold: true, color: C.white, margin: 0, align: "center" });
    slide.addText(m[1], { x: 0.95 + i * 2.25, y: 3.8, w: 1.6, h: 0.26, fontSize: 7.8, color: "CBD5E1", margin: 0, align: "center", fit: "shrink" });
  });
  screenshot(slide, 7.55, 1.05, 4.95, 4.35, "Best Product Screenshot / Dashboard Collage", "Closing product proof", [[0.22, 0.25, C.green], [0.72, 0.62, C.blue]]);
  slide.addShape(pptx.ShapeType.roundRect, { x: 0.78, y: 5.45, w: 3.35, h: 0.58, rectRadius: 0.08, line: { color: C.green }, fill: { color: C.green } });
  slide.addText("Request a Live Demo", { x: 1.05, y: 5.63, w: 2.8, h: 0.16, fontSize: 12, bold: true, color: C.white, align: "center", margin: 0 });
  slide.addText("Contact: [Your Phone]  |  [Your Email]  |  [Website]", { x: 0.78, y: 6.38, w: 5.7, h: 0.2, fontSize: 10.5, color: "CBD5E1", margin: 0 });
  footer(slide, 20, true);
}

pptx.writeFile({
  fileName: path.resolve(__dirname, "../output/seerweb-oms-sales-presentation.pptx"),
});
