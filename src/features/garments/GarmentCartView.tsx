import { useEffect, useMemo, useState } from "react";
import { MessageCircle, Minus, Package, Plus, Printer, ShoppingBag, Trash2 } from "lucide-react";
import { useNavigate } from "react-router-dom";
import { toast } from "sonner";
import { useAuth } from "@/context/AuthContext";
import { useCart } from "@/context/CartContext";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { apiUrl } from "@/url";
import { getImageUrl } from "@/lib/imageUrl";

const FALLBACK_IMAGE =
  "data:image/svg+xml;charset=UTF-8," +
  encodeURIComponent(`
    <svg xmlns="http://www.w3.org/2000/svg" width="240" height="240" viewBox="0 0 240 240">
      <rect width="240" height="240" rx="24" fill="#f8fafc" />
      <rect x="20" y="20" width="200" height="200" rx="24" fill="#ffffff" stroke="#e2e8f0" />
      <path d="M78 154h84" stroke="#94a3b8" stroke-width="10" stroke-linecap="round" />
      <path d="M92 106c0-16 13-29 28-29s28 13 28 29" stroke="#cbd5e1" stroke-width="12" stroke-linecap="round" fill="none" />
      <circle cx="120" cy="112" r="16" fill="#e2e8f0" />
    </svg>
  `);

function formatCurrency(value: number) {
  return `Rs.${value.toLocaleString("en-IN")}`;
}

function getCompleteSetCount(
  variants: Array<{ quantity: number }>
) {
  if (variants.length === 0) return 0;

  return Math.min(...variants.map((variant) => variant.quantity));
}

export function GarmentCartView() {
  const navigate = useNavigate();
  const { user } = useAuth();
  const { cart, updateVariantQty, removeVariant, removeFromCart, clearCart, cartTotal } = useCart();
  const [notes, setNotes] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [retailers, setRetailers] = useState<Array<{ id: number; name: string; store_name?: string }>>([]);
  const [selectedRetailerId, setSelectedRetailerId] = useState("");

  const needsRetailerSelection = user?.role === "dealer" || user?.role === "staff";

  useEffect(() => {
    if (!needsRetailerSelection || !user?.id) return;

    (async () => {
      try {
        const token = localStorage.getItem("token");
        const dealerId = user.role === "dealer" ? user.id : user.dealer_id;
        const endpoint =
          user.role === "staff"
            ? `${apiUrl}/staff/get_retailers_by_executive?executiveid=${user.id}`
            : `${apiUrl}/retailers?dealerid=${dealerId}`;
        const response = await fetch(endpoint, {
          headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        });
        if (!response.ok) throw new Error();
        const data = await response.json();
        setRetailers(Array.isArray(data) ? data : []);
      } catch {
        toast.error("Could not load retailers for garments booking");
      }
    })();
  }, [needsRetailerSelection, user]);

  const summary = useMemo(() => {
    const totalPieces = cart.items.reduce(
      (sum, item) => sum + item.variants.reduce((variantSum, variant) => variantSum + variant.quantity, 0),
      0
    );

    const totalSets = cart.items.reduce((sum, item) => sum + getCompleteSetCount(item.variants), 0);

    const gst = cartTotal * 0.05;

    return {
      totalPieces,
      totalSets,
      gst,
      finalAmount: cartTotal + gst,
      productCount: cart.items.length,
    };
  }, [cart.items, cartTotal]);

  const submitOrder = async () => {
    if (!user || cart.items.length === 0) return;
    if (needsRetailerSelection && !selectedRetailerId) {
      toast.error("Select a retailer before submitting this booking");
      return;
    }

    setSubmitting(true);
    try {
      const token = localStorage.getItem("token");
      const selectedRetailer = retailers.find((retailer) => String(retailer.id) === selectedRetailerId);
      const orderItems = cart.items.flatMap((item) =>
        item.variants.map((variant) => ({
          productId: item.productId,
          variantId: variant.variantId,
          size: variant.size,
          color: variant.color,
          quantity: variant.quantity,
          price: variant.price,
          subtotal: variant.price * variant.quantity,
          rack: variant.rack || "",
          attributes_snapshot: {
            ...item.attributes,
            brand: item.brand,
            model: item.model || "",
            business_type_id: item.businessTypeId,
            garment_meta: item.garmentMeta,
            set_quantity: variant.setQuantity ?? 0,
          },
        }))
      );

      const response = await fetch(`${apiUrl}/orders`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          retailerId: user?.role === "retailer" ? user.id : selectedRetailer?.id,
          retailerName: user?.role === "retailer" ? user?.name : selectedRetailer?.name,
          dealerId: user?.dealer_id ?? user?.id,
          total: summary.finalAmount,
          notes,
          order_by: user?.role,
          order_by_id: user?.id,
          items: orderItems,
        }),
      });

      if (!response.ok) {
        throw new Error("Failed to submit garments order");
      }

      toast.success("Garments order submitted");
      clearCart();
      navigate(user?.role === "retailer" ? "/retailer/orders" : "/dealer/orders");
    } catch (error: any) {
      toast.error(error?.message || "Could not submit garments order");
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className="grid gap-6 xl:grid-cols-[minmax(0,1fr)_360px]">
      {/* ── Left: Cart Items ── */}
      <section className="overflow-hidden rounded-[28px] border border-slate-200 bg-white shadow-sm min-w-0">
        <div className="border-b border-slate-200 px-4 sm:px-6 py-5">
          <div className="flex flex-col gap-2 sm:flex-row sm:items-end sm:justify-between">
            <div>
              <h1 className="text-2xl sm:text-3xl font-black tracking-tight text-slate-900">Order Summary</h1>
              <p className="mt-1 text-sm text-slate-500">
                Review product details, update size quantities, and confirm the garment booking summary.
              </p>
            </div>
            <div className="text-sm text-slate-500 shrink-0">Price</div>
          </div>
        </div>

        {cart.items.length === 0 ? (
          <div className="px-6 py-16 text-center">
            <Package className="mx-auto h-10 w-10 text-slate-300" />
            <h2 className="mt-4 text-lg font-semibold text-slate-900">Your garments cart is empty</h2>
            <p className="mt-1 text-sm text-slate-500">Add products and size lines to start building a booking.</p>
          </div>
        ) : (
          <div className="divide-y divide-slate-200">
            {cart.items.map((item) => {
              const imageSrc = item.garmentMeta?.galleryImages?.[0] || getImageUrl(item.image) || FALLBACK_IMAGE;
              const itemPieces = item.variants.reduce((sum, variant) => sum + variant.quantity, 0);
              const itemTotal = item.variants.reduce((sum, variant) => sum + variant.price * variant.quantity, 0);

              return (
                <article key={item.productId} className="px-4 sm:px-6 py-6">
                  {/* Product header: image + info + price */}
                  <div className="flex gap-4 mb-5">
                    <div className="rounded-[20px] border border-slate-200 bg-slate-50 shrink-0 w-24 sm:w-32 flex items-center justify-center p-1">
                      <img
                        src={imageSrc}
                        alt={item.productName}
                        onError={(e) => { e.currentTarget.src = FALLBACK_IMAGE; }}
                        className="w-full h-auto object-contain rounded-[16px]"
                      />
                    </div>

                    <div className="flex-1 min-w-0 flex flex-col justify-between">
                      <div>
                        <div className="text-[11px] font-semibold uppercase tracking-[0.18em] text-amber-700">
                          {item.garmentMeta?.designNumber || "Design"}
                        </div>
                        <h2 className="mt-1 text-base sm:text-lg font-bold leading-tight text-slate-900 truncate">
                          {item.productName}
                        </h2>
                        <div className="mt-2 flex flex-wrap gap-1.5">
                          {item.brand && (
                            <span className="rounded-full bg-slate-100 px-2.5 py-0.5 text-xs font-medium text-slate-700">
                              {item.brand}
                            </span>
                          )}
                          {item.garmentMeta?.fabricType && (
                            <span className="rounded-full bg-slate-100 px-2.5 py-0.5 text-xs font-medium text-slate-700">
                              {item.garmentMeta.fabricType}
                            </span>
                          )}
                          {item.garmentMeta?.selectedColor && (
                            <span className="rounded-full bg-slate-100 px-2.5 py-0.5 text-xs font-medium text-slate-700">
                              {item.garmentMeta.selectedColor}
                            </span>
                          )}
                        </div>
                        <p className="mt-1 text-xs text-slate-500">
                          {itemPieces} piece{itemPieces !== 1 ? "s" : ""} across {item.variants.length} size
                          {item.variants.length !== 1 ? "s" : ""}.
                        </p>
                      </div>

                      <div className="mt-2">
                        <div className="text-xs text-slate-400 uppercase tracking-wider">Rate</div>
                        <div className="text-lg font-black text-slate-900">{formatCurrency(itemTotal)}</div>
                      </div>
                    </div>
                  </div>

                  {/* ── Variant rows — responsive, no overflow ── */}
                  <div className="space-y-2">
                    {item.variants.map((variant) => {
                      const lineTotal = variant.price * variant.quantity;

                      return (
                        <div
                          key={`${item.productId}-${variant.variantId}-${variant.size}-${variant.color}`}
                          className="rounded-2xl border border-slate-200 bg-slate-50 px-3 py-3"
                        >
                          {/* Row 1: Size chip + Rate */}
                          <div className="flex items-center justify-between gap-2 mb-3">
                            <div className="inline-flex items-center gap-1.5 rounded-full bg-slate-900 px-3 py-1 shrink-0">
                              <span className="text-xs font-semibold text-slate-300">Size:</span>
                              <span className="text-xs font-bold text-white">{variant.size || "Single"}</span>
                            </div>
                            <div className="flex items-center gap-1.5 shrink-0">
                              <span className="text-xs text-slate-500">Rate:</span>
                              <span className="text-sm font-bold text-slate-900">{formatCurrency(variant.price)}</span>
                            </div>
                          </div>

                          {/* Row 2: Stepper + Line total + Delete */}
                          <div className="flex items-center gap-2 flex-wrap">
                            {/* Qty stepper */}
                            <div className="flex items-center overflow-hidden rounded-full border-2 border-[#566de2] bg-white shrink-0">
                              <button
                                type="button"
                                onClick={() =>
                                  variant.quantity <= 1
                                    ? removeVariant(item.productId, variant.variantId)
                                    : updateVariantQty(item.productId, variant.variantId, variant.quantity - 1)
                                }
                                className="flex h-8 w-8 items-center justify-center text-slate-700 transition hover:bg-amber-50"
                              >
                                <Minus size={16} />
                              </button>
                              <span className="min-w-[36px] text-center text-base font-semibold text-slate-900 px-1">
                                {variant.quantity}
                              </span>
                              <button
                                type="button"
                                onClick={() =>
                                  updateVariantQty(item.productId, variant.variantId, variant.quantity + 1)
                                }
                                className="flex h-8 w-8 items-center justify-center text-slate-700 transition hover:bg-amber-50"
                              >
                                <Plus size={16} />
                              </button>
                            </div>

                            {/* Line total */}
                            <span className="text-sm font-bold text-slate-700 shrink-0">
                              = {formatCurrency(lineTotal)}
                            </span>

                            {/* Spacer */}
                            <div className="flex-1" />

                            {/* Delete */}
                            <button
                              type="button"
                              onClick={() => removeVariant(item.productId, variant.variantId)}
                              className="shrink-0 text-sm font-medium text-slate-400 hover:text-red-500 transition whitespace-nowrap flex items-center gap-1"
                            >
                              <Trash2 size={13} />
                              <span className="hidden sm:inline">Delete size</span>
                            </button>
                          </div>
                        </div>
                      );
                    })}
                  </div>

                  {/* Delete product link */}
                  <div className="mt-4 flex flex-wrap items-center gap-x-4 gap-y-2 text-sm">
                    <button
                      type="button"
                      onClick={() => removeFromCart(item.productId)}
                      className="inline-flex items-center gap-1.5 text-red-500 hover:text-red-600 transition"
                    >
                      <Trash2 size={14} />
                      Delete product
                    </button>
                    <span className="text-slate-300">|</span>
                    <span className="text-slate-500">
                      {item.garmentMeta?.bookingType || "Piece order"} booking
                    </span>
                  </div>
                </article>
              );
            })}
          </div>
        )}
      </section>

      {/* ── Right: Summary ── */}
      <aside className="space-y-5 xl:sticky xl:top-6 xl:self-start">
        <div className="rounded-[28px] border border-slate-200 bg-white p-5 shadow-sm space-y-4">
          <div>
            <h2 className="text-lg sm:text-xl font-black text-slate-900">
              Subtotal ({summary.totalPieces} items): {formatCurrency(cartTotal)}
            </h2>
            <p className="mt-1 text-sm text-slate-500">
              GST added at checkout summary. Final payable amount:{" "}
              <span className="font-semibold text-slate-900">{formatCurrency(summary.finalAmount)}</span>
            </p>
          </div>

          {needsRetailerSelection && (
            <select
              value={selectedRetailerId}
              onChange={(e) => setSelectedRetailerId(e.target.value)}
              className="h-12 w-full rounded-2xl border border-slate-200 bg-white px-4 text-sm text-slate-700 focus:border-amber-400 focus:outline-none"
            >
              <option value="">Select Retailer</option>
              {retailers.map((retailer) => (
                <option key={retailer.id} value={retailer.id}>
                  {retailer.store_name || retailer.name}
                </option>
              ))}
            </select>
          )}

          {/* Summary breakdown */}
          <div className="space-y-2 rounded-2xl border border-slate-200 bg-slate-50 p-4 text-sm">
            {[
              ["Products", summary.productCount],
              ["Total Pieces", summary.totalPieces],
              ["Total Sets", summary.totalSets],
              ["Dealer Rate Total", formatCurrency(cartTotal)],
              ["GST", formatCurrency(summary.gst)],
            ].map(([label, value]) => (
              <div key={String(label)} className="flex items-center justify-between">
                <span className="text-slate-500">{label}</span>
                <span className="font-semibold text-slate-900">{value}</span>
              </div>
            ))}
          </div>

          <Button
            type="button"
            onClick={submitOrder}
            disabled={submitting || cart.items.length === 0}
            className="h-12 w-full rounded-full bg-[#5d78ff] text-base font-semibold text-white hover:bg-[#5d78ff]/90 disabled:pointer-events-none disabled:bg-slate-200 disabled:text-slate-500"
          >
            <ShoppingBag size={16} className="mr-2" />
            {submitting ? "Submitting..." : "Proceed to Booking"}
          </Button>

          {/* Booking notes */}
          <div className="rounded-2xl border border-slate-200 p-4">
            <div className="mb-3 text-sm font-semibold text-slate-900">Booking Notes</div>
            <Textarea
              rows={4}
              value={notes}
              onChange={(e) => setNotes(e.target.value)}
              placeholder="Notes for packing, dispatch, assortments, or booking instructions"
              className="rounded-2xl border-slate-200 bg-slate-50 text-sm resize-none"
            />
          </div>

          
        </div>
      </aside>
    </div>
  );
}
