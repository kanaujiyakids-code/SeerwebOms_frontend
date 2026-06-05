import { CSSProperties, useEffect, useMemo, useRef, useState } from "react";
import html2canvas from "html2canvas";
import jsPDF from "jspdf";
import { Download, Package2, Search, Share2, SlidersHorizontal } from "lucide-react";
import { Product } from "@/types";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { useAuth } from "@/context/AuthContext";
import { getProxiedImageUrl } from "@/lib/imageUrl";
import { GarmentProductCard } from "./GarmentProductCard";
import { GARMENT_CART_SETTINGS_UPDATED_EVENT, getGarmentCartSettings } from "./cartSettings";
import {
  getGarmentBookingType,
  getGarmentCategory,
  getGarmentColors,
  getGarmentDesignNumber,
  getGarmentFabric,
  getGarmentGallery,
  getGarmentSubCategory,
  isNewArrivalGarment,
  isTrendingGarment,
} from "./productUtils";

interface GarmentCatalogViewProps {
  products: Product[];
  title: string;
  subtitle: string;
}

export function GarmentCatalogView({ products, title, subtitle }: GarmentCatalogViewProps) {
  const { user } = useAuth();
  const [searchQuery, setSearchQuery] = useState("");
  const [cartSettings, setCartSettings] = useState(() => getGarmentCartSettings(user?.id));
  const [exportMenuOpen, setExportMenuOpen] = useState(false);
  const [exporting, setExporting] = useState<null | "download" | "whatsapp">(null);
  const pdfContainerRef = useRef<HTMLDivElement>(null);
  const [filters, setFilters] = useState({
    bookingType: "all",
    brand: "all",
    category: "all",
    subCategory: "all",
    fabricType: "all",
    size: "all",
    color: "all",
    trend: "all",
    design: "",
  });

  useEffect(() => {
    const loadSettings = () => {
      setCartSettings(getGarmentCartSettings(user?.id));
    };

    loadSettings();
    window.addEventListener("storage", loadSettings);
    window.addEventListener(GARMENT_CART_SETTINGS_UPDATED_EVENT, loadSettings);

    return () => {
      window.removeEventListener("storage", loadSettings);
      window.removeEventListener(GARMENT_CART_SETTINGS_UPDATED_EVENT, loadSettings);
    };
  }, [user?.id]);

  useEffect(() => {
    const closeMenu = () => setExportMenuOpen(false);
    window.addEventListener("click", closeMenu);

    return () => {
      window.removeEventListener("click", closeMenu);
    };
  }, []);

  const optionSets = useMemo(() => ({
    bookingTypes: Array.from(new Set(products.map(getGarmentBookingType).filter(Boolean))).sort(),
    brands: Array.from(new Set(products.map((product) => product.brand || product.attributes?.brand).filter(Boolean))).sort(),
    categories: Array.from(new Set(products.map(getGarmentCategory).filter(Boolean))).sort(),
    subCategories: Array.from(new Set(products.map(getGarmentSubCategory).filter(Boolean))).sort(),
    fabrics: Array.from(new Set(products.map(getGarmentFabric).filter(Boolean))).sort(),
    sizes: Array.from(new Set(products.flatMap((product) => (product.variants ?? []).map((variant) => variant.size)).filter(Boolean))).sort()
  }), [products]);

  const filteredProducts = useMemo(() => {
    return products.filter((product) => {
      const query = searchQuery.trim().toLowerCase();
      if (query) {
        const haystack = [
          product.name,
          product.brand,
          getGarmentDesignNumber(product),
          getGarmentCategory(product),
          getGarmentSubCategory(product),
          getGarmentFabric(product),
          ...getGarmentColors(product),
        ]
          .join(" ")
          .toLowerCase();
        if (!haystack.includes(query)) return false;
      }

      if (filters.bookingType !== "all" && getGarmentBookingType(product) !== filters.bookingType) return false;
      if (filters.brand !== "all" && (product.brand || product.attributes?.brand) !== filters.brand) return false;
      if (filters.category !== "all" && getGarmentCategory(product) !== filters.category) return false;
      if (filters.subCategory !== "all" && getGarmentSubCategory(product) !== filters.subCategory) return false;
      if (filters.fabricType !== "all" && getGarmentFabric(product) !== filters.fabricType) return false;
      if (filters.size !== "all" && !(product.variants ?? []).some((variant) => variant.size === filters.size)) return false;
      if (filters.color !== "all" && !getGarmentColors(product).includes(filters.color)) return false;
      if (filters.design && !getGarmentDesignNumber(product).toLowerCase().includes(filters.design.toLowerCase())) return false;
      if (filters.trend === "trending" && !isTrendingGarment(product)) return false;
      if (filters.trend === "new" && !isNewArrivalGarment(product)) return false;
      return true;
    });
  }, [filters, products, searchQuery]);

  const setFilter = (key: keyof typeof filters, value: string) => {
    setFilters((current) => ({ ...current, [key]: value }));
  };

  const productGridStyle = useMemo(() => {
    return {
      "--garment-cards-per-row": String(cartSettings.cardsPerRow),
    } as CSSProperties;
  }, [cartSettings.cardsPerRow]);

  const loadImageAsBase64 = async (url: string, retries = 3): Promise<string> => {
    if (!url) return "";

    const attemptLoad = (delay = 0): Promise<string> =>
      new Promise((resolve) => {
        setTimeout(() => {
          const img = new Image();
          img.crossOrigin = "anonymous";

          const timeout = setTimeout(() => resolve(""), 10000);

          img.onload = () => {
            clearTimeout(timeout);
            try {
              const canvas = document.createElement("canvas");
              const maxDim = 900;
              const ratio = Math.min(maxDim / img.naturalWidth, maxDim / img.naturalHeight, 1.8);
              canvas.width = Math.max(img.naturalWidth * ratio, 1);
              canvas.height = Math.max(img.naturalHeight * ratio, 1);
              const ctx = canvas.getContext("2d");
              if (!ctx) {
                resolve("");
                return;
              }
              ctx.imageSmoothingEnabled = true;
              ctx.imageSmoothingQuality = "high";
              ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
              resolve(canvas.toDataURL("image/jpeg", 0.95));
            } catch {
              resolve("");
            }
          };

          img.onerror = () => {
            clearTimeout(timeout);
            resolve("");
          };

          img.src = `${url}${url.includes("?") ? "&" : "?"}_t=${Date.now()}`;
        }, delay);
      });

    for (let index = 0; index < retries; index += 1) {
      const result = await attemptLoad(index * 1200);
      if (result) return result;
    }

    return "";
  };

  const buildCatalogPdf = async (productsToExport: Product[]) => {
    if (!pdfContainerRef.current) {
      throw new Error("Catalog container is not ready");
    }

    const pdf = new jsPDF("p", "mm", "a4", false);
    const pageWidth = 210;
    const pageHeight = 297;
    const margin = 15;
    const contentWidth = pageWidth - 2 * margin;
    const contentHeight = pageHeight - 2 * margin;
    // ── KEY CHANGE: image gets 80% of height, details get 20% ──
    const imageSectionHeight = contentHeight * 0.80;
    const detailsSectionHeight = contentHeight * 0.20;
    const fileName = `${(user?.company_name || "Product").replace(/\s+/g, "-")}-Catalog.pdf`;

    pdfContainerRef.current.innerHTML = "";

    const productData = await Promise.all(
      productsToExport.map(async (product) => {
        let imgBase64 = "";
        const gallery = getGarmentGallery(product);
        const candidatePath = gallery[0] || product.image || "";

        if (candidatePath) {
          const proxiedUrl = getProxiedImageUrl(candidatePath);
          if (proxiedUrl) {
            imgBase64 = await loadImageAsBase64(proxiedUrl);
          }
        }

        return { product, imgBase64 };
      })
    );

    for (let pageIndex = 0; pageIndex < productData.length; pageIndex += 1) {
      const { product, imgBase64 } = productData[pageIndex];
      const designNo = getGarmentDesignNumber(product) || product.name;
      const variants = product.variants ?? [];
      const sizes = Array.from(new Set(variants.map((variant) => variant.size).filter(Boolean)));
      const rates = sizes.map((size) => {
        const matched = variants.find((variant) => variant.size === size);
        return matched?.rate || matched?.mrp || product.price;
      });

      const tableHTML =
        sizes.length > 0
          ? `
            <div style="width:100%;background:#ffffff;">
              <div style="font-size:11pt;font-weight:700;margin-bottom:2mm;">Design No : ${designNo}</div>
              <table style="width:100%;border-collapse:collapse;font-size:10pt;text-align:center;">
                <tr>
                  <td style="border:1px solid #111827;padding:2mm;font-weight:700;">Size</td>
                  ${sizes
                    .map(
                      (size) =>
                        `<td style="border:1px solid #111827;padding:2mm;font-weight:700;">${size}</td>`
                    )
                    .join("")}
                </tr>
                <tr>
                  <td style="border:1px solid #111827;padding:2mm;font-weight:700;">MRP</td>
                  ${rates
                    .map(
                      (rate) =>
                        `<td style="border:1px solid #111827;padding:2mm;">${Number(rate).toLocaleString("en-IN")}</td>`
                    )
                    .join("")}
                </tr>
              </table>
            </div>
          `
          : `
            <div style="width:100%;background:#ffffff;">
              <div style="font-size:11pt;font-weight:700;margin-bottom:2mm;">Model: ${product.name}</div>
              <table style="width:100%;border-collapse:collapse;font-size:10pt;text-align:center;">
                <tr>
                  <td style="border:1px solid #111827;padding:2mm;font-weight:700;">Config</td>
                  <td style="border:1px solid #111827;padding:2mm;font-weight:700;">N/A/N/A</td>
                </tr>
                <tr>
                  <td style="border:1px solid #111827;padding:2mm;font-weight:700;">MRP</td>
                  <td style="border:1px solid #111827;padding:2mm;">${product.price.toLocaleString("en-IN")}</td>
                </tr>
              </table>
            </div>
          `;

      const productDiv = document.createElement("div");
      productDiv.style.width = `${contentWidth}mm`;
      productDiv.style.height = `${contentHeight}mm`;
      productDiv.style.background = "#ffffff";
      productDiv.style.fontFamily = "Arial, sans-serif";
      productDiv.style.color = "#1f2937";
      productDiv.style.display = "flex";
      productDiv.style.flexDirection = "column";
      // ── Remove any gap between children ──
      productDiv.style.gap = "0";

      productDiv.innerHTML = `
        <div style="
          width:100%;
          height:${imageSectionHeight}mm;
          background:#ffffff;
          position:relative;
          display:flex;
          align-items:center;
          justify-content:center;
          padding:6mm 8mm 0 8mm;
          box-sizing:border-box;
          overflow:hidden;
          flex-shrink:0;
        ">
          ${
            imgBase64
              ? `<img src="${imgBase64}" style="
                  width:100%;
                  height:100%;
                  object-fit:contain;
                " />`
              : `<div style="width:80mm;height:80mm;background:#e5e7eb;border-radius:8mm;display:flex;align-items:center;justify-content:center;color:#94a3b8;font-size:12pt;border:2px dashed #cbd5e1;">No Image</div>`
          }
          <div style="position:absolute;top:4mm;right:4mm;background:rgba(255,255,255,0.88);padding:2mm 4mm;border-radius:2mm;font-size:10pt;font-weight:700;color:#2563eb;max-width:60%;text-align:right;">
            ${user?.company_name || ""}
          </div>
        </div>
        <div style="
          width:100%;
          height:${detailsSectionHeight}mm;
          padding:4mm 8mm;
          box-sizing:border-box;
          display:flex;
          flex-direction:column;
          justify-content:center;
          background:#ffffff;
          flex-shrink:0;
          border-top:1px solid #e5e7eb;
        ">
          <div style="text-align:center;margin-bottom:3mm;">
            <h2 style="font-size:14pt;font-weight:700;color:#1f2937;margin:0;">${product.name}</h2>
          </div>
          ${tableHTML}
        </div>
      `;

      pdfContainerRef.current.appendChild(productDiv);

      const imgElement = productDiv.querySelector("img") as HTMLImageElement | null;
      if (imgElement && imgBase64) {
        await new Promise((resolve) => {
          const done = () => resolve(true);
          if (imgElement.complete && imgElement.naturalWidth > 0) {
            done();
            return;
          }
          imgElement.onload = done;
          imgElement.onerror = done;
          setTimeout(done, 2000);
        });
      }

      const canvas = await html2canvas(productDiv, {
        scale: 2,
        useCORS: true,
        allowTaint: false,
        logging: false,
        backgroundColor: "#ffffff",
        imageTimeout: 3000,
      });

      if (pageIndex > 0) pdf.addPage();
      pdf.addImage(canvas.toDataURL("image/jpeg", 0.92), "JPEG", margin, margin, contentWidth, contentHeight * 0.95);
      pdfContainerRef.current.removeChild(productDiv);
    }

    return {
      blob: pdf.output("blob"),
      fileName,
    };
  };

  const exportCatalog = async (mode: "download" | "whatsapp") => {
    if (filteredProducts.length === 0) {
      toast.error("No products available to export");
      return;
    }

    setExportMenuOpen(false);
    setExporting(mode);
    const toastId = toast.loading("Generating catalog...");

    try {
      const { blob, fileName } = await buildCatalogPdf(filteredProducts);
      const file = new File([blob], fileName, { type: "application/pdf" });
      const objectUrl = URL.createObjectURL(blob);

      if (mode === "download") {
        const link = document.createElement("a");
        link.href = objectUrl;
        link.download = fileName;
        link.click();
        toast.success(`Catalog downloaded for ${filteredProducts.length} products`, { id: toastId });
      } else {
        const shareText = `Catalog for ${user?.company_name || "our products"} (${filteredProducts.length} products)`;
        const navigatorWithShare = navigator as Navigator & {
          canShare?: (data: ShareData) => boolean;
        };

        if (
          typeof navigator.share === "function" &&
          (!navigatorWithShare.canShare || navigatorWithShare.canShare({ files: [file] }))
        ) {
          await navigator.share({
            title: fileName,
            text: shareText,
            files: [file],
          });
          toast.success("Catalog shared", { id: toastId });
        } else {
          const link = document.createElement("a");
          link.href = objectUrl;
          link.download = fileName;
          link.click();
          window.open(`https://wa.me/?text=${encodeURIComponent(`${shareText}. Attach the downloaded PDF.`)}`, "_blank");
          toast.success("Catalog downloaded. Attach the PDF in WhatsApp.", { id: toastId });
        }
      }

      setTimeout(() => URL.revokeObjectURL(objectUrl), 5000);
    } catch (error) {
      console.error(error);
      toast.error("Failed to export catalog", { id: toastId });
    } finally {
      setExporting(null);
    }
  };

  const renderSelect = (key: keyof typeof filters, values: string[], placeholder: string) => (
    <select
      value={filters[key]}
      onChange={(event) => setFilter(key, event.target.value)}
      className="h-11 rounded-2xl border border-slate-200 bg-white px-3 text-sm text-slate-700 shadow-sm focus:border-amber-400 focus:outline-none"
    >
      <option value="all">{placeholder}</option>
      {values.map((value) => (
        <option key={value} value={value}>
          {value}
        </option>
      ))}
    </select>
  );

  return (
    <div className="space-y-6">
      <div className="rounded-[32px] border border-slate-200 bg-[linear-gradient(135deg,_#fff7ed,_#ffffff_45%,_#f8fafc)] p-6 shadow-sm">
        <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
          <div>
            <div className="inline-flex items-center gap-2 rounded-full border border-amber-200 bg-white px-3 py-1 text-xs font-semibold uppercase tracking-[0.16em] text-amber-700">
              <SlidersHorizontal size={12} />
              Booking Catalog
            </div>
            <h1 className="mt-3 text-3xl font-black tracking-tight text-slate-900">{title}</h1>
            <p className="mt-1 text-sm text-slate-500">{subtitle}</p>
          </div>
          <div className="relative w-full max-w-sm">
            <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
            <Input
              value={searchQuery}
              onChange={(event) => setSearchQuery(event.target.value)}
              placeholder="Search design, category, fabric, color..."
              className="h-12 rounded-2xl border-slate-200 pl-9"
            />
          </div>
        </div>
      </div>

      <div className="rounded-[28px] border border-slate-200 bg-white p-4 shadow-sm">
        <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
          <div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-[0.16em] text-slate-500">
            <SlidersHorizontal size={13} />
            Filters
          </div>
          <div className="relative">
            <Button
              type="button"
              variant="outline"
              className="rounded-2xl"
              disabled={exporting !== null}
              onClick={(event) => {
                event.stopPropagation();
                setExportMenuOpen((current) => !current);
              }}
            >
              {exporting ? (
                <>
                  <Download size={15} />
                  Exporting...
                </>
              ) : (
                <>
                  <Download size={15} />
                  Export Catalog
                </>
              )}
            </Button>
            {exportMenuOpen ? (
              <div
                className="absolute right-0 top-[calc(100%+8px)] z-20 min-w-[220px] rounded-2xl border border-slate-200 bg-white p-2 shadow-xl"
                onClick={(event) => event.stopPropagation()}
              >
                <button
                  type="button"
                  className="flex w-full items-center gap-3 rounded-xl px-3 py-2 text-left text-sm font-medium text-slate-700 transition hover:bg-slate-50"
                  onClick={() => exportCatalog("download")}
                >
                  <Download size={16} />
                  Download
                </button>
                <button
                  type="button"
                  className="flex w-full items-center gap-3 rounded-xl px-3 py-2 text-left text-sm font-medium text-slate-700 transition hover:bg-slate-50"
                  onClick={() => exportCatalog("whatsapp")}
                >
                  <Share2 size={16} />
                  Share via WhatsApp
                </button>
              </div>
            ) : null}
          </div>
        </div>
        <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-5">
          {renderSelect("brand", optionSets.brands, "All Brands")}
          {renderSelect("category", optionSets.categories, "All Categories")}
          {renderSelect("subCategory", optionSets.subCategories, "All Sub Categories")}
          {renderSelect("fabricType", optionSets.fabrics, "All Fabrics")}
          <select
            value={filters.trend}
            onChange={(event) => setFilter("trend", event.target.value)}
            className="h-11 rounded-2xl border border-slate-200 bg-white px-3 text-sm text-slate-700 shadow-sm focus:border-amber-400 focus:outline-none"
          >
            <option value="all">Trending / New</option>
            <option value="trending">Trending</option>
            <option value="new">New Arrival</option>
          </select>
          <Input
            value={filters.design}
            onChange={(event) => setFilter("design", event.target.value)}
            placeholder="Design Number"
            className="h-11 rounded-2xl border-slate-200"
          />
        </div>
      </div>

      <div ref={pdfContainerRef} className="pointer-events-none absolute -left-[9999px] top-0 opacity-0" />

      {filteredProducts.length === 0 ? (
        <div className="rounded-[28px] border border-dashed border-slate-300 bg-white px-6 py-16 text-center shadow-sm">
          <Package2 className="mx-auto mb-3 h-12 w-12 text-slate-300" />
          <h3 className="text-lg font-semibold text-slate-900">No garments matched these filters</h3>
          <p className="mt-1 text-sm text-slate-500">Try widening the booking filters or searching with fewer keywords.</p>
        </div>
      ) : (
        <div
          className="grid grid-cols-1 gap-5 md:grid-cols-2 lg:[grid-template-columns:repeat(var(--garment-cards-per-row),minmax(0,1fr))]"
          style={productGridStyle}
        >
          {filteredProducts.map((product) => (
            <GarmentProductCard key={product.id} product={product} />
          ))}
        </div>
      )}
    </div>
  );
}