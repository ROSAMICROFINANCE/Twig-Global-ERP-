import React, { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, Legend, ResponsiveContainer } from "recharts";
import {
  LayoutGrid, ShoppingCart, Truck, Boxes, Calculator,
  Plus, Trash2, Printer, Upload, X, Check, AlertTriangle, Search,
  ArrowRight, Loader2, ChevronRight, Package,
  Settings as SettingsIcon, LogOut, Lock, User as UserIcon, Shield, MapPin
} from "lucide-react";
import {
  onAuthStateChanged, signInWithEmailAndPassword, createUserWithEmailAndPassword,
  signOut
} from "firebase/auth";
import {
  doc, getDoc, setDoc, onSnapshot, collection, deleteDoc
} from "firebase/firestore";
import { auth, db, getSecondaryAuth, toAuthEmail } from "./firebase";

/* ---------------------------------------------------------
   Helpers
--------------------------------------------------------- */
const money = (n) =>
  "$" + (Number(n) || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const todayISO = () => new Date().toISOString().slice(0, 10);

const daysSince = (iso) => {
  if (!iso) return null;
  const d = Math.floor((Date.now() - new Date(iso).getTime()) / 86400000);
  return d < 0 ? 0 : d;
};

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
const nextNumber = (prefix, counter) => `${prefix}-${String(counter).padStart(4, "0")}`;
const slugify = (s) => s.toLowerCase().trim().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") || uid();

const startOfWeek = (d) => {
  const date = new Date(d);
  const day = date.getDay();
  const diff = date.getDate() - day + (day === 0 ? -6 : 1);
  const monday = new Date(date.setDate(diff));
  monday.setHours(0, 0, 0, 0);
  return monday;
};

const groupByDay = (list, dateKey, valueFn) => {
  const map = {};
  list.forEach((item) => {
    const d = item[dateKey];
    if (!map[d]) map[d] = { date: d, total: 0, count: 0 };
    map[d].total += valueFn(item);
    map[d].count += 1;
  });
  return Object.values(map).sort((a, b) => (a.date < b.date ? 1 : -1));
};

const MODULES = [
  { id: "dashboard", label: "Dashboard", icon: LayoutGrid },
  { id: "sales", label: "Sales", icon: ShoppingCart },
  { id: "purchase", label: "Purchase", icon: Truck },
  { id: "stock", label: "Stock", icon: Boxes },
  { id: "accounting", label: "Accounting", icon: Calculator },
];

const emptyData = () => ({
  products: [], quotes: [], salesOrders: [], purchaseOrders: [], transfers: [], expenses: [], returns: [],
  counters: { quote: 1, sales: 1, purchase: 1, transfer: 1, exp: 1, ret: 1 },
  invoiceSettings: { shopName: "", tagline: "", phone: "", footerNote: "" },
});

const canAccess = (user, moduleId) => !!user && (user.isAdmin || (user.permissions && user.permissions[moduleId]));

const accessibleLocations = (user, locations) => {
  if (!user) return [];
  if (user.isAdmin) return locations;
  return locations.filter((l) => user.locationIds && user.locationIds.includes(l.id));
};

/* ---------------------------------------------------------
   Firestore data layer
   - /business/locations         -> { list: [{id,name}] }
   - /business/{locationId}      -> that shop's products, orders, etc.
   - /users/{uid}                -> profile: name, username, rights, locationIds
--------------------------------------------------------- */
const LOCATIONS_DOC = doc(db, "meta", "locations");
const USERS_COL = collection(db, "users");
const locationDocRef = (locationId) => doc(db, "business", locationId);

async function ensureLocationsList() {
  const snap = await getDoc(LOCATIONS_DOC);
  if (!snap.exists()) {
    await setDoc(LOCATIONS_DOC, { list: [{ id: "main", name: "Main Shop" }] });
  }
}
async function ensureLocationDoc(locationId) {
  const snap = await getDoc(locationDocRef(locationId));
  if (!snap.exists()) await setDoc(locationDocRef(locationId), emptyData());
}
async function saveLocationData(locationId, next) {
  try { await setDoc(locationDocRef(locationId), next); } catch (e) { console.error("Save failed", e); }
}
async function saveLocationsList(list) {
  try { await setDoc(LOCATIONS_DOC, { list }); } catch (e) { console.error("Save failed", e); }
}

/* ---------------------------------------------------------
   Small UI atoms
--------------------------------------------------------- */
const Tag = ({ children, tone = "neutral" }) => {
  const tones = {
    neutral: "bg-slate-100 text-slate-600 border-slate-200",
    low: "bg-amber-50 text-amber-800 border-amber-300",
    danger: "bg-red-50 text-red-700 border-red-300",
    good: "bg-emerald-50 text-emerald-700 border-emerald-300",
    pending: "bg-blue-50 text-blue-700 border-blue-300",
  };
  return (
    <span className={`inline-flex items-center gap-1 text-[11px] font-mono uppercase tracking-wide border px-1.5 py-0.5 rounded ${tones[tone]}`}>
      {children}
    </span>
  );
};

const Btn = ({ children, onClick, variant = "primary", className = "", type = "button", disabled }) => {
  const base = "inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm font-medium transition-colors disabled:opacity-40 disabled:cursor-not-allowed";
  const variants = {
    primary: "bg-[#2B4C7E] text-white hover:bg-[#1F3A63]",
    amber: "bg-[#D9A441] text-[#1F2428] hover:bg-[#c7912e]",
    ghost: "bg-transparent text-[#2B4C7E] hover:bg-[#eef2f7]",
    danger: "bg-transparent text-red-600 hover:bg-red-50",
    outline: "bg-white text-[#2B4C7E] border border-[#c7d2e0] hover:bg-[#eef2f7]",
  };
  return (
    <button type={type} disabled={disabled} onClick={onClick} className={`${base} ${variants[variant]} ${className}`}>
      {children}
    </button>
  );
};

const Field = ({ label, children }) => (
  <label className="block">
    <span className="block text-[11px] font-mono uppercase tracking-wide text-slate-500 mb-1">{label}</span>
    {children}
  </label>
);

const inputCls = "w-full border border-slate-300 rounded-md px-2.5 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-[#2B4C7E]/30 focus:border-[#2B4C7E]";

const TabBar = ({ tabs, active, onChange }) => (
  <div className="flex flex-wrap gap-2 mb-5 border-b border-slate-200 pb-3">
    {tabs.map((t) => (
      <button key={t.id} onClick={() => onChange(t.id)}
        className={`px-3 py-1.5 rounded-full text-sm font-medium transition-colors ${active === t.id ? "bg-[#2B4C7E] text-white" : "bg-slate-100 text-slate-600 hover:bg-slate-200"}`}>
        {t.label}
      </button>
    ))}
  </div>
);

/* ---------------------------------------------------------
   Login / first-run admin setup / self sign-up
--------------------------------------------------------- */
function LoginScreen({ firstRun, onLogin, onCreateAdmin, onSignUp, busy }) {
  const [mode, setMode] = useState("signin");
  const [name, setName] = useState("");
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  const showNameField = firstRun || mode === "signup";

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    try {
      if (firstRun) {
        if (!name || !username || !password) return setError("Fill in every field.");
        if (password.length < 6) return setError("Password needs at least 6 characters.");
        await onCreateAdmin({ name, username, password });
      } else if (mode === "signup") {
        if (!name || !username || !password) return setError("Fill in every field.");
        if (password.length < 6) return setError("Password needs at least 6 characters.");
        await onSignUp({ name, username, password });
      } else {
        await onLogin({ username, password });
      }
    } catch (err) {
      setError(err.message.includes("invalid-credential") || err.message.includes("wrong-password") || err.message.includes("user-not-found")
        ? "Wrong username or password." : err.message.includes("email-already-in-use") ? "That username is already taken."
        : "Something went wrong. Try again.");
    }
  };

  return (
    <div className="min-h-screen bg-[#1F2428] flex items-center justify-center p-4">
      <div className="w-full max-w-sm bg-white rounded-lg p-6">
        <div className="flex items-center gap-2 mb-1">
          <Package size={22} className="text-[#D9A441]" />
          <span className="font-semibold text-lg">HardwareERP</span>
        </div>
        <p className="text-sm text-slate-500 mb-5">
          {firstRun ? "First time setup — create the admin account." : mode === "signup" ? "Create your account — an admin will grant you access after." : "Sign in to your account."}
        </p>
        <form onSubmit={submit} className="space-y-3">
          {showNameField && (
            <Field label="Full name"><input className={inputCls} value={name} onChange={(e) => setName(e.target.value)} /></Field>
          )}
          <Field label="Username"><input className={inputCls} value={username} onChange={(e) => setUsername(e.target.value)} autoCapitalize="none" /></Field>
          <Field label="Password"><input type="password" className={inputCls} value={password} onChange={(e) => setPassword(e.target.value)} /></Field>
          {error && <div className="text-xs text-red-600">{error}</div>}
          <Btn type="submit" className="w-full justify-center" disabled={busy}>
            {busy ? <Loader2 size={14} className="animate-spin" /> : firstRun ? <><Shield size={14} /> Create admin account</> : mode === "signup" ? <><Check size={14} /> Create account</> : <><Lock size={14} /> Sign in</>}
          </Btn>
        </form>
        {!firstRun && (
          <button onClick={() => { setMode(mode === "signup" ? "signin" : "signup"); setError(""); }} className="text-xs text-[#2B4C7E] mt-4 block mx-auto">
            {mode === "signup" ? "Already have an account? Sign in" : "New here? Create an account"}
          </button>
        )}
      </div>
    </div>
  );
}

function AccessPendingScreen({ user, missingModules, missingLocation, onLogout }) {
  return (
    <div className="min-h-screen bg-[#F4F4F2] flex items-center justify-center p-4">
      <div className="w-full max-w-sm bg-white rounded-lg p-6 text-center border border-slate-200">
        <Shield size={28} className="text-[#D9A441] mx-auto mb-3" />
        <h2 className="font-semibold text-lg mb-1">Account created, {user.name.split(" ")[0]}</h2>
        <p className="text-sm text-slate-500 mb-2">
          {missingLocation && "You haven't been assigned to a shop location yet. "}
          {missingModules && "No modules have been assigned to your account yet. "}
        </p>
        <p className="text-sm text-slate-500 mb-5">Ask your admin to grant you access in Settings.</p>
        <Btn variant="outline" onClick={onLogout} className="w-full justify-center"><LogOut size={14} /> Sign out</Btn>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Line item editor
--------------------------------------------------------- */
function LineItemsEditor({ items, setItems, products, priceField, showCost }) {
  const addLine = () => setItems([...items, { id: uid(), productId: "", qty: 1, price: 0 }]);
  const removeLine = (id) => setItems(items.filter((l) => l.id !== id));
  const updateLine = (id, patch) => setItems(items.map((l) => (l.id === id ? { ...l, ...patch } : l)));
  const onProductChange = (id, productId) => {
    const p = products.find((p) => p.id === productId);
    updateLine(id, { productId, price: p ? p[priceField] : 0 });
  };
  return (
    <div className="border border-slate-200 rounded-md overflow-hidden">
      <table className="w-full text-sm">
        <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
          <tr>
            <th className="text-left px-2 py-2">Product</th>
            <th className="text-right px-2 py-2 w-20">Qty</th>
            <th className="text-right px-2 py-2 w-28">{showCost ? "Unit Cost" : "Unit Price"}</th>
            <th className="text-right px-2 py-2 w-28">Line Total</th>
            <th className="w-8"></th>
          </tr>
        </thead>
        <tbody>
          {items.map((line) => {
            const product = products.find((p) => p.id === line.productId);
            return (
              <tr key={line.id} className="border-t border-slate-100">
                <td className="px-2 py-1.5">
                  <select className={inputCls} value={line.productId} onChange={(e) => onProductChange(line.id, e.target.value)}>
                    <option value="">Select product…</option>
                    {products.map((p) => (
                      <option key={p.id} value={p.id}>{p.name} {p.sku ? `(${p.sku})` : ""}</option>
                    ))}
                  </select>
                  {product && <div className="text-[11px] text-slate-400 mt-0.5 font-mono">In stock: {product.stock}</div>}
                </td>
                <td className="px-2 py-1.5">
                  <input type="number" min="0" className={inputCls + " text-right"} value={line.qty}
                    onChange={(e) => updateLine(line.id, { qty: Number(e.target.value) })} />
                </td>
                <td className="px-2 py-1.5">
                  <input type="number" min="0" step="0.01" className={inputCls + " text-right"} value={line.price}
                    onChange={(e) => updateLine(line.id, { price: Number(e.target.value) })} />
                </td>
                <td className="px-2 py-1.5 text-right font-mono">{money(line.qty * line.price)}</td>
                <td className="px-1"><button onClick={() => removeLine(line.id)} className="text-slate-400 hover:text-red-600"><Trash2 size={15} /></button></td>
              </tr>
            );
          })}
          {items.length === 0 && (
            <tr><td colSpan={5} className="text-center text-slate-400 text-sm py-4">No items yet. Add a line below.</td></tr>
          )}
        </tbody>
      </table>
      <div className="px-2 py-2 bg-slate-50 border-t border-slate-200">
        <Btn variant="ghost" onClick={addLine}><Plus size={14} /> Add line</Btn>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Printable documents — pull shop letterhead from invoiceSettings
--------------------------------------------------------- */
function PrintableDoc({ title, number, date, partyLabel, partyName, items, total, footer, settings, extraLine }) {
  const s = settings || {};
  return (
    <div id="print-doc" className="hidden print:block p-8 text-black">
      <div className="flex justify-between items-start border-b-2 border-black pb-4 mb-4">
        <div>
          <div className="text-xl font-bold">{s.shopName || "YOUR SHOP NAME"}</div>
          <div className="text-xs">{s.tagline || "Set your shop name and details in Settings → Invoice Layout"}</div>
          {s.phone && <div className="text-xs">{s.phone}</div>}
        </div>
        <div className="text-right">
          <div className="text-lg font-bold">{title}</div>
          <div className="text-sm font-mono">{number}</div>
          <div className="text-sm">{date}</div>
        </div>
      </div>
      <div className="mb-2 text-sm"><span className="text-slate-600">{partyLabel}: </span><span className="font-semibold">{partyName || "—"}</span></div>
      {extraLine && <div className="mb-2 text-sm text-slate-600">{extraLine}</div>}
      <table className="w-full text-sm border-collapse mt-2">
        <thead><tr className="border-b-2 border-black"><th className="text-left py-1.5">Item</th><th className="text-right py-1.5">Qty</th><th className="text-right py-1.5">Unit Price</th><th className="text-right py-1.5">Total</th></tr></thead>
        <tbody>{items.map((l, i) => (
          <tr key={i} className="border-b border-slate-300"><td className="py-1.5">{l.name}</td><td className="text-right py-1.5">{l.qty}</td><td className="text-right py-1.5">{money(l.price)}</td><td className="text-right py-1.5">{money(l.qty * l.price)}</td></tr>
        ))}</tbody>
      </table>
      <div className="flex justify-end mt-3"><div className="text-right"><div className="text-sm text-slate-600">Total</div><div className="text-2xl font-bold">{money(total)}</div></div></div>
      {s.footerNote && <div className="mt-8 text-xs text-slate-500">{s.footerNote}</div>}
      {footer && <div className="mt-2 text-xs text-slate-400">{footer}</div>}
    </div>
  );
}

function ReceiptDoc({ number, date, partyName, items, total, settings }) {
  const s = settings || {};
  return (
    <div id="print-receipt" className="p-1" style={{ width: "58mm" }}>
      <div className="text-center font-bold text-[13px]">{s.shopName || "YOUR SHOP NAME"}</div>
      <div className="text-center text-[10px] mb-1">Sales Receipt</div>
      <div className="text-[10px] font-mono border-t border-dashed border-black pt-1">{number}<br />{date}<br />{partyName ? `Customer: ${partyName}` : ""}</div>
      <div className="border-t border-dashed border-black mt-1 pt-1">
        {items.map((l, i) => (
          <div key={i} className="text-[10px] mb-0.5">
            <div>{l.name}</div>
            <div className="flex justify-between font-mono"><span>{l.qty} x {Number(l.price).toFixed(2)}</span><span>{(l.qty * l.price).toFixed(2)}</span></div>
          </div>
        ))}
      </div>
      <div className="border-t border-dashed border-black mt-1 pt-1 flex justify-between text-[12px] font-bold"><span>TOTAL</span><span>{money(total)}</span></div>
      <div className="text-center text-[9px] mt-2">{s.footerNote || "Thank you!"}</div>
    </div>
  );
}

function usePrint() {
  const [printJob, setPrintJob] = useState(null);
  useEffect(() => {
    if (!printJob) return;
    const t = setTimeout(() => { window.print(); setPrintJob(null); }, 80);
    return () => clearTimeout(t);
  }, [printJob]);
  return [printJob, setPrintJob];
}

/* ---------------------------------------------------------
   Dashboard
--------------------------------------------------------- */
function Dashboard({ data, setView, user }) {
  const lowStock = data.products.filter((p) => p.stock <= p.minStock);
  const revenue = data.salesOrders.reduce((s, o) => s + o.total, 0);
  const profit = data.salesOrders.reduce((s, o) => s + o.items.reduce((a, l) => a + (l.price - (l.cost || 0)) * l.qty, 0), 0);
  const pendingPOs = data.purchaseOrders.filter((p) => p.status === "pending");
  const cards = [
    { label: "Products tracked", value: data.products.length },
    { label: "Total revenue", value: money(revenue) },
    { label: "Gross profit", value: money(profit) },
    { label: "Low stock items", value: lowStock.length, tone: lowStock.length ? "low" : "neutral" },
  ];

  const weekComparison = useMemo(() => {
    const thisMonday = startOfWeek(new Date());
    const lastMonday = new Date(thisMonday);
    lastMonday.setDate(lastMonday.getDate() - 7);
    const labels = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
    return labels.map((label, i) => {
      const td = new Date(thisMonday); td.setDate(td.getDate() + i);
      const ld = new Date(lastMonday); ld.setDate(ld.getDate() + i);
      const tIso = td.toISOString().slice(0, 10);
      const lIso = ld.toISOString().slice(0, 10);
      const thisTotal = data.salesOrders.filter((o) => o.date === tIso).reduce((s, o) => s + o.total, 0);
      const lastTotal = data.salesOrders.filter((o) => o.date === lIso).reduce((s, o) => s + o.total, 0);
      return { day: label, "This week": thisTotal, "Last week": lastTotal };
    });
  }, [data.salesOrders]);

  const expenseDaily = useMemo(() => {
    const days = [];
    for (let i = 13; i >= 0; i--) {
      const d = new Date(); d.setDate(d.getDate() - i);
      const iso = d.toISOString().slice(0, 10);
      const total = data.expenses.filter((e) => e.date === iso).reduce((s, e) => s + e.amount, 0);
      days.push({ day: iso.slice(5), total });
    }
    return days;
  }, [data.expenses]);

  const thisWeekTotal = weekComparison.reduce((s, d) => s + d["This week"], 0);
  const lastWeekTotal = weekComparison.reduce((s, d) => s + d["Last week"], 0);
  const weekChange = lastWeekTotal ? (((thisWeekTotal - lastWeekTotal) / lastWeekTotal) * 100) : null;

  return (
    <div>
      <h1 className="text-xl font-semibold text-[#1F2428] mb-1">Welcome, {user.name.split(" ")[0]}</h1>
      <p className="text-sm text-slate-500 mb-5">Shop overview at a glance.</p>
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
        {cards.map((c) => (
          <div key={c.label} className="border border-slate-200 rounded-lg p-4 bg-white">
            <div className="text-[11px] font-mono uppercase tracking-wide text-slate-400 mb-1">{c.label}</div>
            <div className="text-2xl font-semibold text-[#1F2428]">{c.value}</div>
          </div>
        ))}
      </div>

      <div className="border border-slate-200 rounded-lg bg-white p-4 mb-4">
        <div className="flex items-center justify-between mb-3">
          <div className="text-sm font-medium">Weekly performance — this week vs last week</div>
          {weekChange !== null && (
            <Tag tone={weekChange >= 0 ? "good" : "danger"}>{weekChange >= 0 ? "+" : ""}{weekChange.toFixed(1)}% vs last week</Tag>
          )}
        </div>
        <div style={{ width: "100%", height: 220 }}>
          <ResponsiveContainer>
            <BarChart data={weekComparison}>
              <XAxis dataKey="day" tick={{ fontSize: 11 }} />
              <YAxis tick={{ fontSize: 11 }} />
              <Tooltip formatter={(v) => money(v)} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="Last week" fill="#c7d2e0" radius={[3, 3, 0, 0]} />
              <Bar dataKey="This week" fill="#2B4C7E" radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="border border-slate-200 rounded-lg bg-white p-4 mb-6">
        <div className="text-sm font-medium mb-3">Daily expenses — last 14 days</div>
        <div style={{ width: "100%", height: 200 }}>
          <ResponsiveContainer>
            <BarChart data={expenseDaily}>
              <XAxis dataKey="day" tick={{ fontSize: 10 }} />
              <YAxis tick={{ fontSize: 10 }} />
              <Tooltip formatter={(v) => money(v)} />
              <Bar dataKey="total" fill="#C0392B" radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="grid md:grid-cols-2 gap-4">
        <div className="border border-slate-200 rounded-lg bg-white">
          <div className="px-4 py-3 border-b border-slate-100 flex items-center justify-between">
            <span className="font-medium text-sm">Items needing restock</span>
            {canAccess(user, "stock") && <button className="text-xs text-[#2B4C7E] flex items-center gap-0.5" onClick={() => setView("stock")}>View stock <ChevronRight size={13} /></button>}
          </div>
          <div className="p-4">
            {lowStock.length === 0 && <div className="text-sm text-slate-400">All stocked up. Nothing below minimum.</div>}
            {lowStock.slice(0, 6).map((p) => (
              <div key={p.id} className="flex justify-between text-sm py-1"><span>{p.name}</span><Tag tone="low">{p.stock} left / min {p.minStock}</Tag></div>
            ))}
          </div>
        </div>
        <div className="border border-slate-200 rounded-lg bg-white">
          <div className="px-4 py-3 border-b border-slate-100 flex items-center justify-between">
            <span className="font-medium text-sm">Purchase orders pending receipt</span>
            {canAccess(user, "purchase") && <button className="text-xs text-[#2B4C7E] flex items-center gap-0.5" onClick={() => setView("purchase")}>View purchases <ChevronRight size={13} /></button>}
          </div>
          <div className="p-4">
            {pendingPOs.length === 0 && <div className="text-sm text-slate-400">No purchase orders waiting.</div>}
            {pendingPOs.slice(0, 6).map((po) => (
              <div key={po.id} className="flex justify-between text-sm py-1"><span>{po.number} — {po.supplier}</span><Tag tone="pending">{money(po.total)}</Tag></div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Stock module: Stock Levels + Goods Returns (folded in)
--------------------------------------------------------- */
function StockLevels({ data, setData, save }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [search, setSearch] = useState("");
  const [importing, setImporting] = useState(false);
  const fileRef = useRef(null);
  const blank = { name: "", sku: "", category: "", unit: "pc", costPrice: 0, sellPrice: 0, stock: 0, minStock: 0 };
  const [form, setForm] = useState(blank);

  const openNew = () => { setForm(blank); setEditing(null); setShowForm(true); };
  const openEdit = (p) => { setForm(p); setEditing(p.id); setShowForm(true); };

  const submit = (e) => {
    e.preventDefault();
    let products;
    if (editing) products = data.products.map((p) => (p.id === editing ? { ...form } : p));
    else products = [...data.products, { ...form, id: uid(), createdAt: todayISO(), lastRestocked: todayISO() }];
    const next = { ...data, products };
    setData(next); save(next);
    setShowForm(false);
  };

  const removeProduct = (id) => {
    const next = { ...data, products: data.products.filter((p) => p.id !== id) };
    setData(next); save(next);
  };

  const handleImport = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setImporting(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      const norm = (obj, keys) => {
        for (const k of Object.keys(obj)) {
          const lk = k.toLowerCase().trim();
          for (const target of keys) {
            if (lk === target || lk.replace(/[\s_]/g, "") === target.replace(/[\s_]/g, "")) return obj[k];
          }
        }
        return undefined;
      };
      const imported = rows.map((r) => ({
        id: uid(),
        name: norm(r, ["name", "product", "product name"]) || "Unnamed",
        sku: norm(r, ["sku", "code"]) || "",
        category: norm(r, ["category"]) || "",
        unit: norm(r, ["unit"]) || "pc",
        costPrice: Number(norm(r, ["costprice", "cost", "cost price"])) || 0,
        sellPrice: Number(norm(r, ["sellprice", "price", "sell price", "selling price"])) || 0,
        stock: Number(norm(r, ["stock", "qty", "quantity"])) || 0,
        minStock: Number(norm(r, ["minstock", "min stock", "minimum stock", "reorder level"])) || 0,
        createdAt: todayISO(), lastRestocked: todayISO(),
      }));
      const next = { ...data, products: [...data.products, ...imported] };
      setData(next); save(next);
    } catch (err) {
      alert("Could not read that file. Use an .xlsx or .csv with columns like Name, SKU, Cost Price, Sell Price, Stock, Min Stock.");
    }
    setImporting(false);
    if (fileRef.current) fileRef.current.value = "";
  };

  const filtered = data.products.filter((p) => (p.name + p.sku + p.category).toLowerCase().includes(search.toLowerCase()));

  return (
    <div>
      <h2 className="text-lg font-semibold text-[#1F2428] mb-1">Stock Levels</h2>
      <p className="text-sm text-slate-500 mb-4">Stock levels, stock age, and reorder tags.</p>
      <div className="flex flex-wrap items-center gap-2 mb-4">
        <div className="relative flex-1 min-w-[180px]">
          <Search size={14} className="absolute left-2.5 top-2.5 text-slate-400" />
          <input className={inputCls + " pl-8"} placeholder="Search products…" value={search} onChange={(e) => setSearch(e.target.value)} />
        </div>
        <input type="file" accept=".xlsx,.xls,.csv" ref={fileRef} className="hidden" onChange={handleImport} />
        <Btn variant="outline" onClick={() => fileRef.current.click()} disabled={importing}>
          {importing ? <Loader2 size={14} className="animate-spin" /> : <Upload size={14} />} Import Excel
        </Btn>
        <Btn onClick={openNew}><Plus size={14} /> Add product</Btn>
      </div>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr>
              <th className="text-left px-3 py-2">Product</th><th className="text-left px-3 py-2">SKU</th>
              <th className="text-right px-3 py-2">Stock</th><th className="text-right px-3 py-2">Min</th>
              <th className="text-right px-3 py-2">Age</th><th className="text-right px-3 py-2">Cost</th>
              <th className="text-right px-3 py-2">Price</th><th className="px-3 py-2"></th><th className="w-16"></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((p) => {
              const low = p.stock <= p.minStock;
              const age = daysSince(p.lastRestocked);
              return (
                <tr key={p.id} className="border-t border-slate-100 hover:bg-slate-50">
                  <td className="px-3 py-2"><div className="font-medium">{p.name}</div><div className="text-[11px] text-slate-400">{p.category}</div></td>
                  <td className="px-3 py-2 font-mono text-xs text-slate-500">{p.sku}</td>
                  <td className="px-3 py-2 text-right font-mono">{p.stock} {p.unit}</td>
                  <td className="px-3 py-2 text-right font-mono text-slate-400">{p.minStock}</td>
                  <td className="px-3 py-2 text-right font-mono text-slate-400">{age !== null ? `${age}d` : "—"}</td>
                  <td className="px-3 py-2 text-right font-mono">{money(p.costPrice)}</td>
                  <td className="px-3 py-2 text-right font-mono">{money(p.sellPrice)}</td>
                  <td className="px-3 py-2">{low && <Tag tone="low"><AlertTriangle size={11} /> Reorder</Tag>}</td>
                  <td className="px-3 py-2 text-right whitespace-nowrap">
                    <button onClick={() => openEdit(p)} className="text-slate-400 hover:text-[#2B4C7E] mr-2 text-xs">Edit</button>
                    <button onClick={() => removeProduct(p.id)} className="text-slate-400 hover:text-red-600"><Trash2 size={14} /></button>
                  </td>
                </tr>
              );
            })}
            {filtered.length === 0 && <tr><td colSpan={9} className="text-center text-slate-400 py-8 text-sm">No products yet. Add one or import from Excel.</td></tr>}
          </tbody>
        </table>
      </div>
      {showForm && (
        <div className="fixed inset-0 bg-black/30 flex items-center justify-center z-50 p-4" onClick={() => setShowForm(false)}>
          <div className="bg-white rounded-lg w-full max-w-md p-5" onClick={(e) => e.stopPropagation()}>
            <div className="flex justify-between items-center mb-4">
              <h2 className="font-semibold">{editing ? "Edit product" : "Add product"}</h2>
              <button onClick={() => setShowForm(false)}><X size={18} className="text-slate-400" /></button>
            </div>
            <form onSubmit={submit} className="space-y-3">
              <Field label="Name"><input required className={inputCls} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
              <div className="grid grid-cols-2 gap-3">
                <Field label="SKU / Code"><input className={inputCls} value={form.sku} onChange={(e) => setForm({ ...form, sku: e.target.value })} /></Field>
                <Field label="Category"><input className={inputCls} value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} /></Field>
              </div>
              <div className="grid grid-cols-3 gap-3">
                <Field label="Unit"><input className={inputCls} value={form.unit} onChange={(e) => setForm({ ...form, unit: e.target.value })} /></Field>
                <Field label="Cost price"><input type="number" step="0.01" className={inputCls} value={form.costPrice} onChange={(e) => setForm({ ...form, costPrice: Number(e.target.value) })} /></Field>
                <Field label="Sell price"><input type="number" step="0.01" className={inputCls} value={form.sellPrice} onChange={(e) => setForm({ ...form, sellPrice: Number(e.target.value) })} /></Field>
              </div>
              <div className="grid grid-cols-2 gap-3">
                <Field label="Current stock"><input type="number" className={inputCls} value={form.stock} onChange={(e) => setForm({ ...form, stock: Number(e.target.value) })} /></Field>
                <Field label="Minimum stock"><input type="number" className={inputCls} value={form.minStock} onChange={(e) => setForm({ ...form, minStock: Number(e.target.value) })} /></Field>
              </div>
              <div className="flex justify-end gap-2 pt-2">
                <Btn variant="ghost" onClick={() => setShowForm(false)}>Cancel</Btn>
                <Btn type="submit"><Check size={14} /> Save product</Btn>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

function GoodsReturnsRegister({ data, setData, save, user }) {
  const [showForm, setShowForm] = useState(false);
  const blank = { type: "good", productId: "", qty: 1, reason: "" };
  const [form, setForm] = useState(blank);

  const submit = (e) => {
    e.preventDefault();
    const product = data.products.find((p) => p.id === form.productId);
    if (!product) return;
    const qty = Number(form.qty) || 0;
    const stockChange = form.type === "good" ? qty : -qty;
    const products = data.products.map((p) => (p.id === product.id ? { ...p, stock: p.stock + stockChange } : p));
    const entry = {
      id: uid(), number: nextNumber("RET", data.counters.ret), date: todayISO(),
      type: form.type, productId: form.productId, productName: product.name, qty,
      reason: form.reason, processedBy: user.username,
      value: form.type === "damaged" ? qty * product.costPrice : 0,
    };
    const next = { ...data, products, returns: [entry, ...data.returns], counters: { ...data.counters, ret: data.counters.ret + 1 } };
    setData(next); save(next);
    setForm(blank); setShowForm(false);
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-1">
        <h2 className="text-lg font-semibold text-[#1F2428]">Goods Returns</h2>
        <Btn onClick={() => setShowForm(true)}><Plus size={14} /> Record return</Btn>
      </div>
      <p className="text-sm text-slate-500 mb-4">Good returns go back into stock. Damaged stock is written off and logged as a loss.</p>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr><th className="text-left px-3 py-2">Number</th><th className="text-left px-3 py-2">Date</th><th className="text-left px-3 py-2">Product</th>
              <th className="text-right px-3 py-2">Qty</th><th className="text-left px-3 py-2">Type</th><th className="text-left px-3 py-2">Reason</th><th className="text-left px-3 py-2">By</th></tr>
          </thead>
          <tbody>
            {data.returns.map((r) => (
              <tr key={r.id} className="border-t border-slate-100">
                <td className="px-3 py-2 font-mono text-xs">{r.number}</td>
                <td className="px-3 py-2">{r.date}</td>
                <td className="px-3 py-2">{r.productName}</td>
                <td className="px-3 py-2 text-right font-mono">{r.qty}</td>
                <td className="px-3 py-2"><Tag tone={r.type === "good" ? "good" : "danger"}>{r.type === "good" ? "Good return" : "Damaged"}</Tag></td>
                <td className="px-3 py-2 text-slate-500">{r.reason}</td>
                <td className="px-3 py-2 text-slate-500 font-mono text-xs">{r.processedBy}</td>
              </tr>
            ))}
            {data.returns.length === 0 && <tr><td colSpan={7} className="text-center text-slate-400 py-8 text-sm">No returns recorded yet.</td></tr>}
          </tbody>
        </table>
      </div>
      {showForm && (
        <div className="fixed inset-0 bg-black/30 flex items-center justify-center z-50 p-4" onClick={() => setShowForm(false)}>
          <div className="bg-white rounded-lg w-full max-w-md p-5" onClick={(e) => e.stopPropagation()}>
            <div className="flex justify-between items-center mb-4">
              <h2 className="font-semibold">Record a return</h2>
              <button onClick={() => setShowForm(false)}><X size={18} className="text-slate-400" /></button>
            </div>
            <form onSubmit={submit} className="space-y-3">
              <Field label="Type">
                <div className="flex gap-2">
                  <button type="button" onClick={() => setForm({ ...form, type: "good" })}
                    className={`flex-1 py-1.5 rounded-md text-sm border ${form.type === "good" ? "bg-emerald-50 border-emerald-400 text-emerald-700" : "border-slate-300 text-slate-500"}`}>Good return</button>
                  <button type="button" onClick={() => setForm({ ...form, type: "damaged" })}
                    className={`flex-1 py-1.5 rounded-md text-sm border ${form.type === "damaged" ? "bg-red-50 border-red-400 text-red-700" : "border-slate-300 text-slate-500"}`}>Damaged / write-off</button>
                </div>
              </Field>
              <Field label="Product">
                <select required className={inputCls} value={form.productId} onChange={(e) => setForm({ ...form, productId: e.target.value })}>
                  <option value="">Select product…</option>
                  {data.products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                </select>
              </Field>
              <Field label="Quantity"><input type="number" min="1" className={inputCls} value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} /></Field>
              <Field label="Reason"><input className={inputCls} value={form.reason} onChange={(e) => setForm({ ...form, reason: e.target.value })} placeholder={form.type === "good" ? "e.g. wrong size ordered" : "e.g. cracked in transit"} /></Field>
              <div className="flex justify-end gap-2 pt-2">
                <Btn variant="ghost" onClick={() => setShowForm(false)}>Cancel</Btn>
                <Btn type="submit"><Check size={14} /> Save</Btn>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

function StockHub({ data, setData, save, user }) {
  const [tab, setTab] = useState("levels");
  const tabs = [
    { id: "levels", label: "Stock Levels" },
    { id: "returns", label: "Goods Returns" },
  ];
  return (
    <div>
      <h1 className="text-xl font-semibold text-[#1F2428] mb-4">Stock</h1>
      <TabBar tabs={tabs} active={tab} onChange={setTab} />
      {tab === "levels" && <StockLevels data={data} setData={setData} save={save} />}
      {tab === "returns" && <GoodsReturnsRegister data={data} setData={setData} save={save} user={user} />}
    </div>
  );
}

/* ---------------------------------------------------------
   Generic doc list
--------------------------------------------------------- */
function DocList({ title, subtitle, docs, columns, onNew, onOpen, renderTag }) {
  return (
    <div>
      <div className="flex items-center justify-between mb-1">
        {title && <h2 className="text-lg font-semibold text-[#1F2428]">{title}</h2>}
        {onNew && <Btn onClick={onNew}><Plus size={14} /> New</Btn>}
      </div>
      {subtitle && <p className="text-sm text-slate-500 mb-4">{subtitle}</p>}
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr>{columns.map((c) => <th key={c} className="text-left px-3 py-2">{c}</th>)}<th className="px-3 py-2"></th></tr>
          </thead>
          <tbody>
            {docs.map((d) => (
              <tr key={d.id} className={`border-t border-slate-100 hover:bg-slate-50 ${onOpen ? "cursor-pointer" : ""}`} onClick={() => onOpen && onOpen(d)}>
                <td className="px-3 py-2 font-mono text-xs">{d.number}</td>
                <td className="px-3 py-2">{d.date}</td>
                <td className="px-3 py-2">{d.customer || d.supplier}</td>
                <td className="px-3 py-2 text-right font-mono">{money(d.total)}</td>
                <td className="px-3 py-2">{renderTag ? renderTag(d) : null}</td>
                <td className="px-3 py-2 text-right">{onOpen && <ChevronRight size={14} className="text-slate-300" />}</td>
              </tr>
            ))}
            {docs.length === 0 && <tr><td colSpan={columns.length + 1} className="text-center text-slate-400 py-8 text-sm">Nothing here yet.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Quotation register
--------------------------------------------------------- */
function QuotationModule({ data, setData, save, setPrintJob, user }) {
  const [mode, setMode] = useState("list");
  const [current, setCurrent] = useState(null);
  const [customer, setCustomer] = useState("");
  const [items, setItems] = useState([]);
  const openNew = () => { setCustomer(""); setItems([]); setCurrent(null); setMode("form"); };
  const openView = (q) => { setCurrent(q); setMode("view"); };
  const total = items.reduce((s, l) => s + l.qty * l.price, 0);

  const submit = (e) => {
    e.preventDefault();
    const number = nextNumber("QUO", data.counters.quote);
    const quote = { id: uid(), number, customer, date: todayISO(), items, total, status: "draft", createdBy: user.username };
    const next = { ...data, quotes: [quote, ...data.quotes], counters: { ...data.counters, quote: data.counters.quote + 1 } };
    setData(next); save(next);
    setMode("list");
  };

  const convertToSalesOrder = (quote) => {
    for (const l of quote.items) {
      const p = data.products.find((p) => p.id === l.productId);
      if (p && p.stock < l.qty) {
        if (!confirm(`${p.name} only has ${p.stock} in stock but order needs ${l.qty}. Convert anyway?`)) return;
        break;
      }
    }
    const number = nextNumber("SO", data.counters.sales);
    const soItems = quote.items.map((l) => {
      const p = data.products.find((p) => p.id === l.productId);
      return { ...l, cost: p ? p.costPrice : 0 };
    });
    const salesOrder = { id: uid(), number, customer: quote.customer, date: todayISO(), items: soItems, total: quote.total, quoteId: quote.id, createdBy: user.username, paymentType: "cash" };
    const products = data.products.map((p) => {
      const l = quote.items.find((l) => l.productId === p.id);
      return l ? { ...p, stock: p.stock - l.qty } : p;
    });
    const quotes = data.quotes.map((q) => (q.id === quote.id ? { ...q, status: "converted" } : q));
    const next = { ...data, products, quotes, salesOrders: [salesOrder, ...data.salesOrders], counters: { ...data.counters, sales: data.counters.sales + 1 } };
    setData(next); save(next);
    setMode("list");
  };

  if (mode === "form") {
    return (
      <div>
        <div className="flex items-center gap-2 mb-4 text-sm text-slate-500">
          <button onClick={() => setMode("list")} className="hover:text-[#2B4C7E]">Quotations</button><ChevronRight size={13} /> <span className="text-slate-800">New quotation</span>
        </div>
        <form onSubmit={submit} className="space-y-4 max-w-3xl">
          <Field label="Customer name"><input required className={inputCls} value={customer} onChange={(e) => setCustomer(e.target.value)} /></Field>
          <LineItemsEditor items={items} setItems={setItems} products={data.products} priceField="sellPrice" />
          <div className="flex justify-between items-center">
            <div className="text-lg font-semibold">Total: {money(total)}</div>
            <div className="flex gap-2"><Btn variant="ghost" onClick={() => setMode("list")}>Cancel</Btn><Btn type="submit"><Check size={14} /> Save quotation</Btn></div>
          </div>
        </form>
      </div>
    );
  }

  if (mode === "view" && current) {
    return (
      <div>
        <div className="flex items-center gap-2 mb-4 text-sm text-slate-500">
          <button onClick={() => setMode("list")} className="hover:text-[#2B4C7E]">Quotations</button><ChevronRight size={13} /> <span className="text-slate-800">{current.number}</span>
        </div>
        <div className="max-w-3xl border border-slate-200 rounded-lg bg-white p-5">
          <div className="flex justify-between mb-4">
            <div><div className="text-lg font-semibold">{current.customer}</div><div className="text-xs text-slate-400 font-mono">{current.number} · {current.date}</div></div>
            <Tag tone={current.status === "converted" ? "good" : "pending"}>{current.status}</Tag>
          </div>
          <table className="w-full text-sm mb-4">
            <thead className="text-[11px] font-mono uppercase text-slate-400 border-b"><tr><th className="text-left py-1">Item</th><th className="text-right py-1">Qty</th><th className="text-right py-1">Price</th><th className="text-right py-1">Total</th></tr></thead>
            <tbody>
              {current.items.map((l, i) => {
                const p = data.products.find((p) => p.id === l.productId);
                return <tr key={i} className="border-b border-slate-100"><td className="py-1.5">{p ? p.name : "—"}</td><td className="text-right py-1.5 font-mono">{l.qty}</td><td className="text-right py-1.5 font-mono">{money(l.price)}</td><td className="text-right py-1.5 font-mono">{money(l.qty * l.price)}</td></tr>;
              })}
            </tbody>
          </table>
          <div className="flex justify-between items-center">
            <div className="text-lg font-semibold">Total: {money(current.total)}</div>
            <div className="flex gap-2">
              <Btn variant="outline" onClick={() => setPrintJob({ mode: "doc", payload: {
                title: "QUOTATION", number: current.number, date: current.date, partyLabel: "Customer",
                partyName: current.customer, items: current.items.map(l => ({ ...l, name: (data.products.find(p=>p.id===l.productId)||{}).name || "—" })), total: current.total
              }})}><Printer size={14} /> Print</Btn>
              {current.status !== "converted" && <Btn onClick={() => convertToSalesOrder(current)}><ArrowRight size={14} /> Convert to sales order</Btn>}
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <DocList title="Quotations" subtitle="Draft quotes for customers, then convert to a sales order."
      docs={data.quotes} columns={["Number", "Date", "Customer", "Total", "Status"]}
      onNew={openNew} onOpen={openView}
      renderTag={(d) => <Tag tone={d.status === "converted" ? "good" : "pending"}>{d.status}</Tag>} />
  );
}

/* ---------------------------------------------------------
   Sales Order register — now with Cash / Credit
--------------------------------------------------------- */
function SalesOrderModule({ data, setData, save, setPrintJob, user }) {
  const [mode, setMode] = useState("list");
  const [current, setCurrent] = useState(null);
  const [customer, setCustomer] = useState("");
  const [items, setItems] = useState([]);
  const [paymentType, setPaymentType] = useState("cash");
  const openNew = () => { setCustomer(""); setItems([]); setPaymentType("cash"); setMode("form"); };
  const openView = (d) => { setCurrent(d); setMode("view"); };
  const total = items.reduce((s, l) => s + l.qty * l.price, 0);

  const submit = (e) => {
    e.preventDefault();
    for (const l of items) {
      const p = data.products.find((p) => p.id === l.productId);
      if (p && p.stock < l.qty) { if (!confirm(`${p.name} only has ${p.stock} in stock but order needs ${l.qty}. Save anyway?`)) return; }
    }
    const number = nextNumber("SO", data.counters.sales);
    const soItems = items.map((l) => { const p = data.products.find((p) => p.id === l.productId); return { ...l, cost: p ? p.costPrice : 0 }; });
    const salesOrder = {
      id: uid(), number, customer, date: todayISO(), items: soItems, total, createdBy: user.username,
      paymentType, paidStatus: paymentType === "credit" ? "unpaid" : "paid",
    };
    const products = data.products.map((p) => { const l = items.find((l) => l.productId === p.id); return l ? { ...p, stock: p.stock - l.qty } : p; });
    const next = { ...data, products, salesOrders: [salesOrder, ...data.salesOrders], counters: { ...data.counters, sales: data.counters.sales + 1 } };
    setData(next); save(next);
    setMode("list");
  };

  const markPaid = (order) => {
    const salesOrders = data.salesOrders.map((o) => (o.id === order.id ? { ...o, paidStatus: "paid" } : o));
    const next = { ...data, salesOrders };
    setData(next); save(next);
    setCurrent({ ...order, paidStatus: "paid" });
  };

  if (mode === "form") {
    return (
      <div>
        <div className="flex items-center gap-2 mb-4 text-sm text-slate-500">
          <button onClick={() => setMode("list")} className="hover:text-[#2B4C7E]">Sales orders</button><ChevronRight size={13} /> <span className="text-slate-800">New sales order</span>
        </div>
        <form onSubmit={submit} className="space-y-4 max-w-3xl">
          <Field label="Customer name"><input required className={inputCls} value={customer} onChange={(e) => setCustomer(e.target.value)} /></Field>
          <Field label="Payment">
            <div className="flex gap-2">
              <button type="button" onClick={() => setPaymentType("cash")}
                className={`flex-1 py-1.5 rounded-md text-sm border ${paymentType === "cash" ? "bg-emerald-50 border-emerald-400 text-emerald-700" : "border-slate-300 text-slate-500"}`}>Cash</button>
              <button type="button" onClick={() => setPaymentType("credit")}
                className={`flex-1 py-1.5 rounded-md text-sm border ${paymentType === "credit" ? "bg-amber-50 border-amber-400 text-amber-800" : "border-slate-300 text-slate-500"}`}>Credit</button>
            </div>
          </Field>
          <LineItemsEditor items={items} setItems={setItems} products={data.products} priceField="sellPrice" />
          <div className="flex justify-between items-center">
            <div className="text-lg font-semibold">Total: {money(total)}</div>
            <div className="flex gap-2"><Btn variant="ghost" onClick={() => setMode("list")}>Cancel</Btn><Btn type="submit"><Check size={14} /> Save — deducts stock</Btn></div>
          </div>
          <p className="text-xs text-slate-400">Saving this order automatically removes the quantities above from stock. Credit sales appear in the Creditors Report until marked paid.</p>
        </form>
      </div>
    );
  }

  if (mode === "view" && current) {
    return (
      <div>
        <div className="flex items-center gap-2 mb-4 text-sm text-slate-500">
          <button onClick={() => setMode("list")} className="hover:text-[#2B4C7E]">Sales orders</button><ChevronRight size={13} /> <span className="text-slate-800">{current.number}</span>
        </div>
        <div className="max-w-3xl border border-slate-200 rounded-lg bg-white p-5">
          <div className="flex justify-between mb-4">
            <div><div className="text-lg font-semibold">{current.customer}</div><div className="text-xs text-slate-400 font-mono">{current.number} · {current.date}{current.createdBy ? ` · by ${current.createdBy}` : ""}</div></div>
            <div className="flex gap-1.5">
              <Tag tone={current.paymentType === "credit" ? "pending" : "good"}>{current.paymentType === "credit" ? "Credit" : "Cash"}</Tag>
              {current.paymentType === "credit" && <Tag tone={current.paidStatus === "paid" ? "good" : "danger"}>{current.paidStatus}</Tag>}
            </div>
          </div>
          <table className="w-full text-sm mb-4">
            <thead className="text-[11px] font-mono uppercase text-slate-400 border-b"><tr><th className="text-left py-1">Item</th><th className="text-right py-1">Qty</th><th className="text-right py-1">Price</th><th className="text-right py-1">Total</th></tr></thead>
            <tbody>
              {current.items.map((l, i) => {
                const p = data.products.find((p) => p.id === l.productId);
                return <tr key={i} className="border-b border-slate-100"><td className="py-1.5">{p ? p.name : "—"}</td><td className="text-right py-1.5 font-mono">{l.qty}</td><td className="text-right py-1.5 font-mono">{money(l.price)}</td><td className="text-right py-1.5 font-mono">{money(l.qty * l.price)}</td></tr>;
              })}
            </tbody>
          </table>
          <div className="flex justify-between items-center">
            <div className="text-lg font-semibold">Total: {money(current.total)}</div>
            <div className="flex gap-2 flex-wrap justify-end">
              <Btn variant="outline" onClick={() => setPrintJob({ mode: "doc", payload: {
                title: "SALES ORDER", number: current.number, date: current.date, partyLabel: "Customer",
                partyName: current.customer, items: current.items.map(l => ({ ...l, name: (data.products.find(p=>p.id===l.productId)||{}).name || "—" })), total: current.total,
                extraLine: `Payment: ${current.paymentType === "credit" ? "Credit" : "Cash"}`
              }})}><Printer size={14} /> Print A4 / PDF</Btn>
              <Btn onClick={() => setPrintJob({ mode: "receipt", payload: {
                number: current.number, date: current.date, partyName: current.customer,
                items: current.items.map(l => ({ ...l, name: (data.products.find(p=>p.id===l.productId)||{}).name || "—" })), total: current.total
              }})}><Printer size={14} /> Print receipt (58mm)</Btn>
              {current.paymentType === "credit" && current.paidStatus !== "paid" && (
                <Btn variant="amber" onClick={() => markPaid(current)}><Check size={14} /> Mark as paid</Btn>
              )}
            </div>
          </div>
        </div>
      </div>
    );
  }

  return <DocList title="Sales Orders" subtitle="Confirmed sales — saving one removes stock automatically."
    docs={data.salesOrders} columns={["Number", "Date", "Customer", "Total"]} onNew={openNew} onOpen={openView}
    renderTag={(d) => d.paymentType === "credit" ? <Tag tone={d.paidStatus === "paid" ? "good" : "danger"}>Credit · {d.paidStatus}</Tag> : <Tag tone="good">Cash</Tag>} />;
}

/* ---------------------------------------------------------
   Daily Creditors Report (Sales)
--------------------------------------------------------- */
function DailyCreditorsReport({ data }) {
  const creditOrders = data.salesOrders.filter((o) => o.paymentType === "credit");
  const rows = useMemo(() => {
    const map = {};
    creditOrders.forEach((o) => {
      if (!map[o.date]) map[o.date] = { date: o.date, total: 0, outstanding: 0, count: 0 };
      map[o.date].total += o.total;
      map[o.date].outstanding += o.paidStatus === "paid" ? 0 : o.total;
      map[o.date].count += 1;
    });
    return Object.values(map).sort((a, b) => (a.date < b.date ? 1 : -1));
  }, [creditOrders]);

  return (
    <div>
      <h2 className="text-lg font-semibold text-[#1F2428] mb-1">Daily Creditors Report</h2>
      <p className="text-sm text-slate-500 mb-4">Credit sales by day, and how much of each day's credit is still outstanding.</p>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr><th className="text-left px-3 py-2">Date</th><th className="text-right px-3 py-2">Credit Sales</th><th className="text-right px-3 py-2">Total</th><th className="text-right px-3 py-2">Outstanding</th></tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.date} className="border-t border-slate-100">
                <td className="px-3 py-2">{r.date}</td>
                <td className="px-3 py-2 text-right font-mono">{r.count}</td>
                <td className="px-3 py-2 text-right font-mono">{money(r.total)}</td>
                <td className="px-3 py-2 text-right font-mono text-red-600">{money(r.outstanding)}</td>
              </tr>
            ))}
            {rows.length === 0 && <tr><td colSpan={4} className="text-center text-slate-400 py-8 text-sm">No credit sales recorded yet.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Purchase Order register
--------------------------------------------------------- */
function PurchaseOrderModule({ data, setData, save, setPrintJob, user }) {
  const [mode, setMode] = useState("list");
  const [current, setCurrent] = useState(null);
  const [supplier, setSupplier] = useState("");
  const [items, setItems] = useState([]);
  const openNew = () => { setSupplier(""); setItems([]); setMode("form"); };
  const openView = (d) => { setCurrent(d); setMode("view"); };
  const total = items.reduce((s, l) => s + l.qty * l.price, 0);

  const submit = (e) => {
    e.preventDefault();
    const number = nextNumber("PO", data.counters.purchase);
    const po = { id: uid(), number, supplier, date: todayISO(), items, total, status: "pending", createdBy: user.username };
    const next = { ...data, purchaseOrders: [po, ...data.purchaseOrders], counters: { ...data.counters, purchase: data.counters.purchase + 1 } };
    setData(next); save(next);
    setMode("list");
  };

  const markReceived = (po) => {
    const products = data.products.map((p) => {
      const l = po.items.find((l) => l.productId === p.id);
      return l ? { ...p, stock: p.stock + l.qty, lastRestocked: todayISO(), costPrice: l.price || p.costPrice } : p;
    });
    const purchaseOrders = data.purchaseOrders.map((p) => (p.id === po.id ? { ...p, status: "received", receivedDate: todayISO() } : p));
    const next = { ...data, products, purchaseOrders };
    setData(next); save(next);
    setCurrent({ ...po, status: "received" });
  };

  if (mode === "form") {
    return (
      <div>
        <div className="flex items-center gap-2 mb-4 text-sm text-slate-500">
          <button onClick={() => setMode("list")} className="hover:text-[#2B4C7E]">Purchase orders</button><ChevronRight size={13} /> <span className="text-slate-800">New purchase order</span>
        </div>
        <form onSubmit={submit} className="space-y-4 max-w-3xl">
          <Field label="Supplier name"><input required className={inputCls} value={supplier} onChange={(e) => setSupplier(e.target.value)} /></Field>
          <LineItemsEditor items={items} setItems={setItems} products={data.products} priceField="costPrice" showCost />
          <div className="flex justify-between items-center">
            <div className="text-lg font-semibold">Total: {money(total)}</div>
            <div className="flex gap-2"><Btn variant="ghost" onClick={() => setMode("list")}>Cancel</Btn><Btn type="submit"><Check size={14} /> Save purchase order</Btn></div>
          </div>
        </form>
      </div>
    );
  }

  if (mode === "view" && current) {
    return (
      <div>
        <div className="flex items-center gap-2 mb-4 text-sm text-slate-500">
          <button onClick={() => setMode("list")} className="hover:text-[#2B4C7E]">Purchase orders</button><ChevronRight size={13} /> <span className="text-slate-800">{current.number}</span>
        </div>
        <div className="max-w-3xl border border-slate-200 rounded-lg bg-white p-5">
          <div className="flex justify-between mb-4">
            <div><div className="text-lg font-semibold">{current.supplier}</div><div className="text-xs text-slate-400 font-mono">{current.number} · {current.date}</div></div>
            <Tag tone={current.status === "received" ? "good" : "pending"}>{current.status}</Tag>
          </div>
          <table className="w-full text-sm mb-4">
            <thead className="text-[11px] font-mono uppercase text-slate-400 border-b"><tr><th className="text-left py-1">Item</th><th className="text-right py-1">Qty</th><th className="text-right py-1">Unit cost</th><th className="text-right py-1">Total</th></tr></thead>
            <tbody>
              {current.items.map((l, i) => {
                const p = data.products.find((p) => p.id === l.productId);
                return <tr key={i} className="border-b border-slate-100"><td className="py-1.5">{p ? p.name : "—"}</td><td className="text-right py-1.5 font-mono">{l.qty}</td><td className="text-right py-1.5 font-mono">{money(l.price)}</td><td className="text-right py-1.5 font-mono">{money(l.qty * l.price)}</td></tr>;
              })}
            </tbody>
          </table>
          <div className="flex justify-between items-center">
            <div className="text-lg font-semibold">Total: {money(current.total)}</div>
            <div className="flex gap-2">
              <Btn variant="outline" onClick={() => setPrintJob({ mode: "doc", payload: {
                title: "PURCHASE ORDER", number: current.number, date: current.date, partyLabel: "Supplier",
                partyName: current.supplier, items: current.items.map(l => ({ ...l, name: (data.products.find(p=>p.id===l.productId)||{}).name || "—" })), total: current.total
              }})}><Printer size={14} /> Print</Btn>
              {current.status !== "received" && <Btn onClick={() => markReceived(current)}><Check size={14} /> Mark received — adds stock</Btn>}
            </div>
          </div>
        </div>
      </div>
    );
  }

  return <DocList title="Purchase Orders" subtitle="Order stock from suppliers. Marking received adds stock automatically."
    docs={data.purchaseOrders} columns={["Number", "Date", "Supplier", "Total", "Status"]} onNew={openNew} onOpen={openView}
    renderTag={(d) => <Tag tone={d.status === "received" ? "good" : "pending"}>{d.status}</Tag>} />;
}

/* ---------------------------------------------------------
   Internal Transfer Out (Sales)
--------------------------------------------------------- */
function TransferOutRegister({ data, setData, save, user }) {
  const [showForm, setShowForm] = useState(false);
  const blank = { productId: "", qty: 1, party: "", notes: "" };
  const [form, setForm] = useState(blank);
  const outTransfers = data.transfers.filter((t) => t.type === "out");

  const submit = (e) => {
    e.preventDefault();
    const product = data.products.find((p) => p.id === form.productId);
    if (!product) return;
    const qty = Number(form.qty) || 0;
    if (product.stock < qty) { if (!confirm(`Only ${product.stock} in stock. Send out anyway?`)) return; }
    const entry = {
      id: uid(), number: nextNumber("TRF-OUT", data.counters.transfer), type: "out", date: todayISO(),
      productId: product.id, productName: product.name, qty, party: form.party, notes: form.notes,
      valueAtCost: qty * product.costPrice, createdBy: user.username,
    };
    const products = data.products.map((p) => (p.id === product.id ? { ...p, stock: p.stock - qty } : p));
    const next = { ...data, products, transfers: [entry, ...data.transfers], counters: { ...data.counters, transfer: data.counters.transfer + 1 } };
    setData(next); save(next);
    setForm(blank); setShowForm(false);
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-1">
        <h2 className="text-lg font-semibold text-[#1F2428]">Internal Transfer Out</h2>
        <Btn onClick={() => setShowForm(true)}><Plus size={14} /> Record transfer out</Btn>
      </div>
      <p className="text-sm text-slate-500 mb-4">Stock lent out (e.g. to a friend) — removed from stock at cost. No profit is recorded.</p>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr><th className="text-left px-3 py-2">Number</th><th className="text-left px-3 py-2">Date</th><th className="text-left px-3 py-2">Product</th>
              <th className="text-right px-3 py-2">Qty</th><th className="text-left px-3 py-2">To</th><th className="text-right px-3 py-2">Value (cost)</th><th className="text-left px-3 py-2">Notes</th></tr>
          </thead>
          <tbody>
            {outTransfers.map((t) => (
              <tr key={t.id} className="border-t border-slate-100">
                <td className="px-3 py-2 font-mono text-xs">{t.number}</td>
                <td className="px-3 py-2">{t.date}</td>
                <td className="px-3 py-2">{t.productName}</td>
                <td className="px-3 py-2 text-right font-mono">{t.qty}</td>
                <td className="px-3 py-2">{t.party}</td>
                <td className="px-3 py-2 text-right font-mono">{money(t.valueAtCost)}</td>
                <td className="px-3 py-2 text-slate-500">{t.notes}</td>
              </tr>
            ))}
            {outTransfers.length === 0 && <tr><td colSpan={7} className="text-center text-slate-400 py-8 text-sm">No outgoing transfers yet.</td></tr>}
          </tbody>
        </table>
      </div>
      {showForm && (
        <div className="fixed inset-0 bg-black/30 flex items-center justify-center z-50 p-4" onClick={() => setShowForm(false)}>
          <div className="bg-white rounded-lg w-full max-w-md p-5" onClick={(e) => e.stopPropagation()}>
            <div className="flex justify-between items-center mb-4">
              <h2 className="font-semibold">Record transfer out</h2>
              <button onClick={() => setShowForm(false)}><X size={18} className="text-slate-400" /></button>
            </div>
            <form onSubmit={submit} className="space-y-3">
              <Field label="Product">
                <select required className={inputCls} value={form.productId} onChange={(e) => setForm({ ...form, productId: e.target.value })}>
                  <option value="">Select product…</option>
                  {data.products.map((p) => <option key={p.id} value={p.id}>{p.name} (in stock: {p.stock})</option>)}
                </select>
              </Field>
              <Field label="Quantity"><input type="number" min="1" className={inputCls} value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} /></Field>
              <Field label="Given to (name)"><input required className={inputCls} value={form.party} onChange={(e) => setForm({ ...form, party: e.target.value })} /></Field>
              <Field label="Notes"><input className={inputCls} value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} placeholder="e.g. borrowed, to be returned" /></Field>
              <div className="flex justify-end gap-2 pt-2">
                <Btn variant="ghost" onClick={() => setShowForm(false)}>Cancel</Btn>
                <Btn type="submit"><Check size={14} /> Save</Btn>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Internal Transfer In (Purchase)
--------------------------------------------------------- */
function TransferInRegister({ data, setData, save, user }) {
  const [showForm, setShowForm] = useState(false);
  const blank = { productId: "", qty: 1, party: "", notes: "" };
  const [form, setForm] = useState(blank);
  const inTransfers = data.transfers.filter((t) => t.type === "in");

  const submit = (e) => {
    e.preventDefault();
    const product = data.products.find((p) => p.id === form.productId);
    if (!product) return;
    const qty = Number(form.qty) || 0;
    const entry = {
      id: uid(), number: nextNumber("TRF-IN", data.counters.transfer), type: "in", date: todayISO(),
      productId: product.id, productName: product.name, qty, party: form.party, notes: form.notes,
      valueAtCost: qty * product.costPrice, createdBy: user.username,
    };
    const products = data.products.map((p) => (p.id === product.id ? { ...p, stock: p.stock + qty } : p));
    const next = { ...data, products, transfers: [entry, ...data.transfers], counters: { ...data.counters, transfer: data.counters.transfer + 1 } };
    setData(next); save(next);
    setForm(blank); setShowForm(false);
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-1">
        <h2 className="text-lg font-semibold text-[#1F2428]">Internal Transfer In</h2>
        <Btn onClick={() => setShowForm(true)}><Plus size={14} /> Record transfer in</Btn>
      </div>
      <p className="text-sm text-slate-500 mb-4">Stock returned into the system — added back to stock at cost.</p>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr><th className="text-left px-3 py-2">Number</th><th className="text-left px-3 py-2">Date</th><th className="text-left px-3 py-2">Product</th>
              <th className="text-right px-3 py-2">Qty</th><th className="text-left px-3 py-2">From</th><th className="text-right px-3 py-2">Value (cost)</th><th className="text-left px-3 py-2">Notes</th></tr>
          </thead>
          <tbody>
            {inTransfers.map((t) => (
              <tr key={t.id} className="border-t border-slate-100">
                <td className="px-3 py-2 font-mono text-xs">{t.number}</td>
                <td className="px-3 py-2">{t.date}</td>
                <td className="px-3 py-2">{t.productName}</td>
                <td className="px-3 py-2 text-right font-mono">{t.qty}</td>
                <td className="px-3 py-2">{t.party}</td>
                <td className="px-3 py-2 text-right font-mono">{money(t.valueAtCost)}</td>
                <td className="px-3 py-2 text-slate-500">{t.notes}</td>
              </tr>
            ))}
            {inTransfers.length === 0 && <tr><td colSpan={7} className="text-center text-slate-400 py-8 text-sm">No incoming transfers yet.</td></tr>}
          </tbody>
        </table>
      </div>
      {showForm && (
        <div className="fixed inset-0 bg-black/30 flex items-center justify-center z-50 p-4" onClick={() => setShowForm(false)}>
          <div className="bg-white rounded-lg w-full max-w-md p-5" onClick={(e) => e.stopPropagation()}>
            <div className="flex justify-between items-center mb-4">
              <h2 className="font-semibold">Record transfer in</h2>
              <button onClick={() => setShowForm(false)}><X size={18} className="text-slate-400" /></button>
            </div>
            <form onSubmit={submit} className="space-y-3">
              <Field label="Product">
                <select required className={inputCls} value={form.productId} onChange={(e) => setForm({ ...form, productId: e.target.value })}>
                  <option value="">Select product…</option>
                  {data.products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                </select>
              </Field>
              <Field label="Quantity"><input type="number" min="1" className={inputCls} value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} /></Field>
              <Field label="Returned by (name)"><input required className={inputCls} value={form.party} onChange={(e) => setForm({ ...form, party: e.target.value })} /></Field>
              <Field label="Notes"><input className={inputCls} value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} /></Field>
              <div className="flex justify-end gap-2 pt-2">
                <Btn variant="ghost" onClick={() => setShowForm(false)}>Cancel</Btn>
                <Btn type="submit"><Check size={14} /> Save</Btn>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Daily Sales Report (Sales)
--------------------------------------------------------- */
function DailySalesReport({ data }) {
  const rows = groupByDay(data.salesOrders, "date", (o) => o.total);
  return (
    <div>
      <h2 className="text-lg font-semibold text-[#1F2428] mb-1">Daily Sales Report</h2>
      <p className="text-sm text-slate-500 mb-4">Gross sales per day, before any expenses.</p>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr><th className="text-left px-3 py-2">Date</th><th className="text-right px-3 py-2">Orders</th><th className="text-right px-3 py-2">Total Sales</th></tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.date} className="border-t border-slate-100"><td className="px-3 py-2">{r.date}</td><td className="px-3 py-2 text-right font-mono">{r.count}</td><td className="px-3 py-2 text-right font-mono">{money(r.total)}</td></tr>
            ))}
            {rows.length === 0 && <tr><td colSpan={3} className="text-center text-slate-400 py-8 text-sm">No sales recorded yet.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Daily Expense Report + entry (Sales)
--------------------------------------------------------- */
function DailyExpenseRegister({ data, setData, save, user }) {
  const [showForm, setShowForm] = useState(false);
  const blank = { date: todayISO(), category: "", description: "", amount: 0 };
  const [form, setForm] = useState(blank);

  const submit = (e) => {
    e.preventDefault();
    const entry = { id: uid(), number: nextNumber("EXP", data.counters.exp), ...form, amount: Number(form.amount), recordedBy: user.username };
    const next = { ...data, expenses: [entry, ...data.expenses], counters: { ...data.counters, exp: data.counters.exp + 1 } };
    setData(next); save(next);
    setForm(blank); setShowForm(false);
  };

  const removeExpense = (id) => {
    const next = { ...data, expenses: data.expenses.filter((e) => e.id !== id) };
    setData(next); save(next);
  };

  const daily = groupByDay(data.expenses, "date", (e) => e.amount);

  return (
    <div>
      <div className="flex items-center justify-between mb-1">
        <h2 className="text-lg font-semibold text-[#1F2428]">Daily Expense Report</h2>
        <Btn onClick={() => setShowForm(true)}><Plus size={14} /> Add expense</Btn>
      </div>
      <p className="text-sm text-slate-500 mb-4">All expenses, grouped by day.</p>

      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white mb-5">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr><th className="text-left px-3 py-2">Date</th><th className="text-right px-3 py-2">Entries</th><th className="text-right px-3 py-2">Total</th></tr>
          </thead>
          <tbody>
            {daily.map((r) => (
              <tr key={r.date} className="border-t border-slate-100"><td className="px-3 py-2">{r.date}</td><td className="px-3 py-2 text-right font-mono">{r.count}</td><td className="px-3 py-2 text-right font-mono">{money(r.total)}</td></tr>
            ))}
            {daily.length === 0 && <tr><td colSpan={3} className="text-center text-slate-400 py-6 text-sm">No expenses recorded yet.</td></tr>}
          </tbody>
        </table>
      </div>

      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <div className="px-4 py-2.5 border-b border-slate-100 text-sm font-medium">All entries</div>
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr><th className="text-left px-3 py-2">Date</th><th className="text-left px-3 py-2">Category</th><th className="text-left px-3 py-2">Description</th><th className="text-right px-3 py-2">Amount</th><th className="w-10"></th></tr>
          </thead>
          <tbody>
            {data.expenses.map((e) => (
              <tr key={e.id} className="border-t border-slate-100">
                <td className="px-3 py-2">{e.date}</td>
                <td className="px-3 py-2"><Tag>{e.category || "Other"}</Tag></td>
                <td className="px-3 py-2 text-slate-600">{e.description}</td>
                <td className="px-3 py-2 text-right font-mono">{money(e.amount)}</td>
                <td className="px-3 py-2 text-right"><button onClick={() => removeExpense(e.id)} className="text-slate-400 hover:text-red-600"><Trash2 size={14} /></button></td>
              </tr>
            ))}
            {data.expenses.length === 0 && <tr><td colSpan={5} className="text-center text-slate-400 py-8 text-sm">No expenses recorded yet.</td></tr>}
          </tbody>
        </table>
      </div>

      {showForm && (
        <div className="fixed inset-0 bg-black/30 flex items-center justify-center z-50 p-4" onClick={() => setShowForm(false)}>
          <div className="bg-white rounded-lg w-full max-w-md p-5" onClick={(e) => e.stopPropagation()}>
            <div className="flex justify-between items-center mb-4">
              <h2 className="font-semibold">Add expense</h2>
              <button onClick={() => setShowForm(false)}><X size={18} className="text-slate-400" /></button>
            </div>
            <form onSubmit={submit} className="space-y-3">
              <Field label="Date"><input type="date" className={inputCls} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></Field>
              <Field label="Category"><input className={inputCls} placeholder="e.g. Rent, Wages, Transport" value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} /></Field>
              <Field label="Description"><input className={inputCls} value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} /></Field>
              <Field label="Amount"><input type="number" step="0.01" min="0" className={inputCls} value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} /></Field>
              <div className="flex justify-end gap-2 pt-2">
                <Btn variant="ghost" onClick={() => setShowForm(false)}>Cancel</Btn>
                <Btn type="submit"><Check size={14} /> Save expense</Btn>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Daily Cash Book (Sales)
--------------------------------------------------------- */
function DailyCashBook({ data }) {
  const rows = useMemo(() => {
    const dates = new Set([...data.salesOrders.map((o) => o.date), ...data.expenses.map((e) => e.date)]);
    const sorted = Array.from(dates).sort();
    let running = 0;
    return sorted.map((date) => {
      const debit = data.salesOrders.filter((o) => o.date === date).reduce((s, o) => s + o.total, 0);
      const credit = data.expenses.filter((e) => e.date === date).reduce((s, e) => s + e.amount, 0);
      const net = debit - credit;
      running += net;
      return { date, debit, credit, net, cashInHand: running };
    }).reverse();
  }, [data.salesOrders, data.expenses]);

  return (
    <div>
      <h2 className="text-lg font-semibold text-[#1F2428] mb-1">Daily Cash Book</h2>
      <p className="text-sm text-slate-500 mb-4">Sales (debit) against expenses (credit), with running cash in hand.</p>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr>
              <th className="text-left px-3 py-2">Date</th>
              <th className="text-right px-3 py-2">Debit (Sales)</th>
              <th className="text-right px-3 py-2">Credit (Expenses)</th>
              <th className="text-right px-3 py-2">Net for day</th>
              <th className="text-right px-3 py-2">Cash in hand</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.date} className="border-t border-slate-100">
                <td className="px-3 py-2">{r.date}</td>
                <td className="px-3 py-2 text-right font-mono text-emerald-700">{money(r.debit)}</td>
                <td className="px-3 py-2 text-right font-mono text-red-600">{money(r.credit)}</td>
                <td className={`px-3 py-2 text-right font-mono ${r.net >= 0 ? "text-emerald-700" : "text-red-600"}`}>{money(r.net)}</td>
                <td className="px-3 py-2 text-right font-mono font-semibold">{money(r.cashInHand)}</td>
              </tr>
            ))}
            {rows.length === 0 && <tr><td colSpan={5} className="text-center text-slate-400 py-8 text-sm">No activity recorded yet.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Daily Purchase Report (Purchase)
--------------------------------------------------------- */
function DailyPurchaseReport({ data }) {
  const rows = groupByDay(data.purchaseOrders, "date", (o) => o.total);
  return (
    <div>
      <h2 className="text-lg font-semibold text-[#1F2428] mb-1">Daily Purchase Report</h2>
      <p className="text-sm text-slate-500 mb-4">Purchases placed per day.</p>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr><th className="text-left px-3 py-2">Date</th><th className="text-right px-3 py-2">Orders</th><th className="text-right px-3 py-2">Total Purchases</th></tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.date} className="border-t border-slate-100"><td className="px-3 py-2">{r.date}</td><td className="px-3 py-2 text-right font-mono">{r.count}</td><td className="px-3 py-2 text-right font-mono">{money(r.total)}</td></tr>
            ))}
            {rows.length === 0 && <tr><td colSpan={3} className="text-center text-slate-400 py-8 text-sm">No purchases recorded yet.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Costing register (Accounting)
--------------------------------------------------------- */
function CostingRegister({ data }) {
  const rows = data.products.map((p) => ({ ...p, stockValue: p.stock * p.costPrice }));
  const totalValue = rows.reduce((s, r) => s + r.stockValue, 0);
  return (
    <div>
      <h2 className="text-lg font-semibold text-[#1F2428] mb-1">Costing</h2>
      <p className="text-sm text-slate-500 mb-4">Cost per product, based on the last purchase price paid.</p>
      <div className="border border-slate-200 rounded-lg p-4 bg-white mb-4 inline-block">
        <div className="text-[11px] font-mono uppercase text-slate-400 mb-1">Total stock value (at cost)</div>
        <div className="text-xl font-semibold">{money(totalValue)}</div>
      </div>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr><th className="text-left px-3 py-2">Product</th><th className="text-left px-3 py-2">SKU</th><th className="text-right px-3 py-2">Last Cost</th><th className="text-right px-3 py-2">Stock Qty</th><th className="text-right px-3 py-2">Stock Value</th></tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.id} className="border-t border-slate-100">
                <td className="px-3 py-2 font-medium">{r.name}</td>
                <td className="px-3 py-2 font-mono text-xs text-slate-500">{r.sku}</td>
                <td className="px-3 py-2 text-right font-mono">{money(r.costPrice)}</td>
                <td className="px-3 py-2 text-right font-mono">{r.stock}</td>
                <td className="px-3 py-2 text-right font-mono">{money(r.stockValue)}</td>
              </tr>
            ))}
            {rows.length === 0 && <tr><td colSpan={5} className="text-center text-slate-400 py-8 text-sm">No products yet.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Income Statement (Accounting)
--------------------------------------------------------- */
function IncomeStatement({ data }) {
  const revenue = data.salesOrders.reduce((s, o) => s + o.total, 0);
  const cogs = data.salesOrders.reduce((s, o) => s + o.items.reduce((a, l) => a + (l.cost || 0) * l.qty, 0), 0);
  const grossProfit = revenue - cogs;
  const totalExpenses = data.expenses.reduce((s, e) => s + e.amount, 0);
  const netProfit = grossProfit - totalExpenses;
  const outOnLoan = data.transfers.filter((t) => t.type === "out").reduce((s, t) => s + t.valueAtCost, 0)
    - data.transfers.filter((t) => t.type === "in").reduce((s, t) => s + t.valueAtCost, 0);

  const Row = ({ label, value, bold, indent }) => (
    <div className={`flex justify-between py-2 ${bold ? "font-semibold border-t border-slate-300" : ""} ${indent ? "pl-4 text-slate-600" : ""}`}>
      <span>{label}</span><span className="font-mono">{money(value)}</span>
    </div>
  );

  return (
    <div>
      <h2 className="text-lg font-semibold text-[#1F2428] mb-1">Income Statement</h2>
      <p className="text-sm text-slate-500 mb-4">Company-wide, based on all recorded sales and expenses.</p>
      <div className="max-w-lg border border-slate-200 rounded-lg bg-white p-5">
        <Row label="Sales Revenue" value={revenue} />
        <Row label="Cost of Goods Sold" value={-cogs} indent />
        <Row label="Gross Profit" value={grossProfit} bold />
        <Row label="Operating Expenses" value={-totalExpenses} indent />
        <Row label="Net Profit" value={netProfit} bold />
      </div>
      {outOnLoan !== 0 && (
        <p className="text-xs text-slate-400 mt-3 max-w-lg">
          Informational only — not included above: {money(Math.abs(outOnLoan))} of stock currently {outOnLoan > 0 ? "out on internal loan and not yet returned" : "returned in excess of what's on loan"}.
        </p>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Creditors Report — overall (Accounting)
--------------------------------------------------------- */
function CreditorsReport({ data }) {
  const rows = useMemo(() => {
    const unpaid = data.salesOrders.filter((o) => o.paymentType === "credit" && o.paidStatus !== "paid");
    const map = {};
    unpaid.forEach((o) => {
      if (!map[o.customer]) map[o.customer] = { customer: o.customer, total: 0, count: 0, oldest: o.date };
      map[o.customer].total += o.total;
      map[o.customer].count += 1;
      if (o.date < map[o.customer].oldest) map[o.customer].oldest = o.date;
    });
    return Object.values(map).sort((a, b) => b.total - a.total);
  }, [data.salesOrders]);
  const grandTotal = rows.reduce((s, r) => s + r.total, 0);

  return (
    <div>
      <h2 className="text-lg font-semibold text-[#1F2428] mb-1">Creditors Report</h2>
      <p className="text-sm text-slate-500 mb-4">Everyone who still owes money on credit sales, company-wide.</p>
      <div className="border border-slate-200 rounded-lg p-4 bg-white mb-4 inline-block">
        <div className="text-[11px] font-mono uppercase text-slate-400 mb-1">Total outstanding</div>
        <div className="text-xl font-semibold text-red-600">{money(grandTotal)}</div>
      </div>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr><th className="text-left px-3 py-2">Customer</th><th className="text-right px-3 py-2">Unpaid Orders</th><th className="text-right px-3 py-2">Outstanding</th><th className="text-right px-3 py-2">Oldest since</th></tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.customer} className="border-t border-slate-100">
                <td className="px-3 py-2 font-medium">{r.customer}</td>
                <td className="px-3 py-2 text-right font-mono">{r.count}</td>
                <td className="px-3 py-2 text-right font-mono text-red-600">{money(r.total)}</td>
                <td className="px-3 py-2 text-right font-mono text-slate-500">{r.oldest} ({daysSince(r.oldest)}d)</td>
              </tr>
            ))}
            {rows.length === 0 && <tr><td colSpan={4} className="text-center text-slate-400 py-8 text-sm">No outstanding credit — nothing owed.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Reports hub (Accounting)
--------------------------------------------------------- */
function ReportsHub({ data }) {
  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-lg font-semibold text-[#1F2428] mb-1">Reports</h2>
        <p className="text-sm text-slate-500">All quotations, sales orders, and purchase orders on record.</p>
      </div>
      <DocList title="Quotations" docs={data.quotes} columns={["Number", "Date", "Customer", "Total", "Status"]}
        renderTag={(d) => <Tag tone={d.status === "converted" ? "good" : "pending"}>{d.status}</Tag>} />
      <DocList title="Sales Orders" docs={data.salesOrders} columns={["Number", "Date", "Customer", "Total"]} />
      <DocList title="Purchase Orders" docs={data.purchaseOrders} columns={["Number", "Date", "Supplier", "Total", "Status"]}
        renderTag={(d) => <Tag tone={d.status === "received" ? "good" : "pending"}>{d.status}</Tag>} />
    </div>
  );
}

/* ---------------------------------------------------------
   Sales module (hub with sub-registers)
--------------------------------------------------------- */
function SalesHub({ data, setData, save, setPrintJob, user }) {
  const [tab, setTab] = useState("orders");
  const tabs = [
    { id: "orders", label: "Sales Orders" },
    { id: "quotes", label: "Quotations" },
    { id: "transferOut", label: "Internal Transfer Out" },
    { id: "dailySales", label: "Daily Sales Report" },
    { id: "dailyExpense", label: "Daily Expense Report" },
    { id: "cashBook", label: "Daily Cash Book" },
    { id: "creditors", label: "Creditors Report" },
  ];
  return (
    <div>
      <h1 className="text-xl font-semibold text-[#1F2428] mb-4">Sales</h1>
      <TabBar tabs={tabs} active={tab} onChange={setTab} />
      {tab === "orders" && <SalesOrderModule data={data} setData={setData} save={save} setPrintJob={setPrintJob} user={user} />}
      {tab === "quotes" && <QuotationModule data={data} setData={setData} save={save} setPrintJob={setPrintJob} user={user} />}
      {tab === "transferOut" && <TransferOutRegister data={data} setData={setData} save={save} user={user} />}
      {tab === "dailySales" && <DailySalesReport data={data} />}
      {tab === "dailyExpense" && <DailyExpenseRegister data={data} setData={setData} save={save} user={user} />}
      {tab === "cashBook" && <DailyCashBook data={data} />}
      {tab === "creditors" && <DailyCreditorsReport data={data} />}
    </div>
  );
}

/* ---------------------------------------------------------
   Purchase module (hub with sub-registers)
--------------------------------------------------------- */
function PurchaseHub({ data, setData, save, setPrintJob, user }) {
  const [tab, setTab] = useState("orders");
  const tabs = [
    { id: "orders", label: "Purchase Orders" },
    { id: "transferIn", label: "Internal Transfer In" },
    { id: "dailyPurchase", label: "Daily Purchase Report" },
  ];
  return (
    <div>
      <h1 className="text-xl font-semibold text-[#1F2428] mb-4">Purchase</h1>
      <TabBar tabs={tabs} active={tab} onChange={setTab} />
      {tab === "orders" && <PurchaseOrderModule data={data} setData={setData} save={save} setPrintJob={setPrintJob} user={user} />}
      {tab === "transferIn" && <TransferInRegister data={data} setData={setData} save={save} user={user} />}
      {tab === "dailyPurchase" && <DailyPurchaseReport data={data} />}
    </div>
  );
}

/* ---------------------------------------------------------
   Accounting module (hub with sub-registers)
--------------------------------------------------------- */
function AccountingHub({ data }) {
  const [tab, setTab] = useState("costing");
  const tabs = [
    { id: "costing", label: "Costing" },
    { id: "income", label: "Income Statement" },
    { id: "creditors", label: "Creditors Report" },
    { id: "reports", label: "Reports" },
  ];
  return (
    <div>
      <h1 className="text-xl font-semibold text-[#1F2428] mb-4">Accounting</h1>
      <TabBar tabs={tabs} active={tab} onChange={setTab} />
      {tab === "costing" && <CostingRegister data={data} />}
      {tab === "income" && <IncomeStatement data={data} />}
      {tab === "creditors" && <CreditorsReport data={data} />}
      {tab === "reports" && <ReportsHub data={data} />}
    </div>
  );
}

/* ---------------------------------------------------------
   Settings: Users & Rights
--------------------------------------------------------- */
function UsersSettings({ users, refreshUsers, locations }) {
  const [showForm, setShowForm] = useState(false);
  const [editingUser, setEditingUser] = useState(null);
  const blank = { name: "", username: "", password: "", permissions: {}, locationIds: [] };
  const [form, setForm] = useState(blank);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  const openNew = () => { setForm(blank); setEditingUser(null); setError(""); setShowForm(true); };
  const openEdit = (u) => { setForm({ name: u.name, username: u.username, password: "", permissions: u.permissions || {}, locationIds: u.locationIds || [] }); setEditingUser(u); setError(""); setShowForm(true); };
  const togglePerm = (moduleId) => setForm({ ...form, permissions: { ...form.permissions, [moduleId]: !form.permissions[moduleId] } });
  const toggleLoc = (locId) => setForm({ ...form, locationIds: form.locationIds.includes(locId) ? form.locationIds.filter((x) => x !== locId) : [...form.locationIds, locId] });

  const submit = async (e) => {
    e.preventDefault();
    setError(""); setBusy(true);
    try {
      if (editingUser) {
        await setDoc(doc(db, "users", editingUser.id), {
          name: form.name, username: editingUser.username, isAdmin: false, permissions: form.permissions, locationIds: form.locationIds,
        });
      } else {
        if (form.password.length < 6) { setError("Password needs at least 6 characters."); setBusy(false); return; }
        const { authInstance, appInstance } = getSecondaryAuth();
        const cred = await createUserWithEmailAndPassword(authInstance, toAuthEmail(form.username), form.password);
        await setDoc(doc(db, "users", cred.user.uid), {
          name: form.name, username: form.username.trim().toLowerCase(), isAdmin: false, permissions: form.permissions, locationIds: form.locationIds,
        });
        await signOut(authInstance);
        appInstance.delete && appInstance.delete();
      }
      setShowForm(false);
      refreshUsers();
    } catch (err) {
      setError(err.message.includes("email-already-in-use") ? "That username is already taken." : err.message.includes("weak-password") ? "Password needs at least 6 characters." : "Could not save user.");
    }
    setBusy(false);
  };

  const removeUser = async (u) => {
    if (u.isAdmin) return;
    if (!confirm(`Remove ${u.name}'s access? They will be signed out and won't be able to log in again.`)) return;
    await deleteDoc(doc(db, "users", u.id));
    refreshUsers();
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-1">
        <h2 className="text-lg font-semibold text-[#1F2428]">Users & Rights</h2>
        <Btn onClick={openNew}><Plus size={14} /> Add user</Btn>
      </div>
      <p className="text-sm text-slate-500 mb-4">New team members can create their own login on the sign-in screen — they'll see nothing until you grant modules and a location below.</p>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr><th className="text-left px-3 py-2">Name</th><th className="text-left px-3 py-2">Username</th><th className="text-left px-3 py-2">Modules</th><th className="text-left px-3 py-2">Locations</th><th className="w-24"></th></tr>
          </thead>
          <tbody>
            {users.map((u) => (
              <tr key={u.id} className="border-t border-slate-100">
                <td className="px-3 py-2 font-medium">{u.name}</td>
                <td className="px-3 py-2 font-mono text-xs">{u.username}</td>
                <td className="px-3 py-2">
                  {u.isAdmin ? <Tag tone="good">Full access</Tag> :
                    <div className="flex flex-wrap gap-1">
                      {MODULES.filter((m) => u.permissions && u.permissions[m.id]).map((m) => <Tag key={m.id}>{m.label}</Tag>)}
                      {!(u.permissions && Object.values(u.permissions).some(Boolean)) && <span className="text-slate-400 text-xs">None</span>}
                    </div>}
                </td>
                <td className="px-3 py-2">
                  {u.isAdmin ? <Tag tone="good">All locations</Tag> :
                    <div className="flex flex-wrap gap-1">
                      {(locations || []).filter((l) => u.locationIds && u.locationIds.includes(l.id)).map((l) => <Tag key={l.id}>{l.name}</Tag>)}
                      {!(u.locationIds && u.locationIds.length) && <span className="text-slate-400 text-xs">None</span>}
                    </div>}
                </td>
                <td className="px-3 py-2 text-right whitespace-nowrap">
                  {!u.isAdmin && <>
                    <button onClick={() => openEdit(u)} className="text-slate-400 hover:text-[#2B4C7E] mr-2 text-xs">Edit</button>
                    <button onClick={() => removeUser(u)} className="text-slate-400 hover:text-red-600"><Trash2 size={14} /></button>
                  </>}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {showForm && (
        <div className="fixed inset-0 bg-black/30 flex items-center justify-center z-50 p-4" onClick={() => setShowForm(false)}>
          <div className="bg-white rounded-lg w-full max-w-md p-5 max-h-[90vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
            <div className="flex justify-between items-center mb-4">
              <h2 className="font-semibold">{editingUser ? "Edit user" : "Add user"}</h2>
              <button onClick={() => setShowForm(false)}><X size={18} className="text-slate-400" /></button>
            </div>
            <form onSubmit={submit} className="space-y-3">
              <Field label="Full name"><input required className={inputCls} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
              <div className="grid grid-cols-2 gap-3">
                <Field label="Username">
                  <input required disabled={!!editingUser} className={inputCls} value={form.username} onChange={(e) => setForm({ ...form, username: e.target.value })} />
                </Field>
                {!editingUser && (
                  <Field label="Password"><input required type="password" className={inputCls} value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} /></Field>
                )}
              </div>
              {editingUser && <p className="text-xs text-slate-400">Username and password can't be changed here once created — remove and re-add the user if that's needed.</p>}
              <Field label="Modules this person can access">
                <div className="grid grid-cols-2 gap-1.5 mt-1">
                  {MODULES.map((m) => (
                    <label key={m.id} className="flex items-center gap-2 text-sm border border-slate-200 rounded-md px-2 py-1.5 cursor-pointer">
                      <input type="checkbox" checked={!!form.permissions[m.id]} onChange={() => togglePerm(m.id)} />
                      {m.label}
                    </label>
                  ))}
                </div>
              </Field>
              <Field label="Locations this person can access">
                <div className="grid grid-cols-2 gap-1.5 mt-1">
                  {(locations || []).map((l) => (
                    <label key={l.id} className="flex items-center gap-2 text-sm border border-slate-200 rounded-md px-2 py-1.5 cursor-pointer">
                      <input type="checkbox" checked={form.locationIds.includes(l.id)} onChange={() => toggleLoc(l.id)} />
                      {l.name}
                    </label>
                  ))}
                  {(!locations || locations.length === 0) && <span className="text-slate-400 text-xs col-span-2">No locations set up yet — add one in the Locations tab first.</span>}
                </div>
              </Field>
              {error && <div className="text-xs text-red-600">{error}</div>}
              <div className="flex justify-end gap-2 pt-2">
                <Btn variant="ghost" onClick={() => setShowForm(false)}>Cancel</Btn>
                <Btn type="submit" disabled={busy}>{busy ? <Loader2 size={14} className="animate-spin" /> : <><Check size={14} /> Save user</>}</Btn>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Settings: Locations
--------------------------------------------------------- */
function LocationsSettings({ locations, refreshLocations }) {
  const [name, setName] = useState("");
  const [busy, setBusy] = useState(false);
  const [renaming, setRenaming] = useState(null);
  const [renameValue, setRenameValue] = useState("");

  const addLocation = async (e) => {
    e.preventDefault();
    if (!name.trim()) return;
    setBusy(true);
    let id = slugify(name);
    if (locations.some((l) => l.id === id)) id = id + "-" + uid().slice(0, 4);
    const list = [...locations, { id, name: name.trim() }];
    await saveLocationsList(list);
    await ensureLocationDoc(id);
    setName("");
    refreshLocations();
    setBusy(false);
  };

  const saveRename = async (loc) => {
    const list = locations.map((l) => (l.id === loc.id ? { ...l, name: renameValue } : l));
    await saveLocationsList(list);
    setRenaming(null);
    refreshLocations();
  };

  const removeLocation = async (loc) => {
    if (!confirm(`Remove "${loc.name}" from the location list? Its data stays saved but no one will be able to select it unless you add it back.`)) return;
    const list = locations.filter((l) => l.id !== loc.id);
    await saveLocationsList(list);
    refreshLocations();
  };

  return (
    <div>
      <h2 className="text-lg font-semibold text-[#1F2428] mb-1">Locations</h2>
      <p className="text-sm text-slate-500 mb-4">Each location keeps its own products, stock, and sales — separate from the others.</p>
      <form onSubmit={addLocation} className="flex gap-2 mb-4 max-w-md">
        <input className={inputCls} placeholder="e.g. Downtown Branch" value={name} onChange={(e) => setName(e.target.value)} />
        <Btn type="submit" disabled={busy}>{busy ? <Loader2 size={14} className="animate-spin" /> : <Plus size={14} />} Add</Btn>
      </form>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white max-w-md">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr><th className="text-left px-3 py-2">Location</th><th className="w-28"></th></tr>
          </thead>
          <tbody>
            {locations.map((l) => (
              <tr key={l.id} className="border-t border-slate-100">
                <td className="px-3 py-2">
                  {renaming === l.id ? (
                    <input autoFocus className={inputCls} value={renameValue} onChange={(e) => setRenameValue(e.target.value)} onBlur={() => saveRename(l)} onKeyDown={(e) => e.key === "Enter" && saveRename(l)} />
                  ) : l.name}
                </td>
                <td className="px-3 py-2 text-right whitespace-nowrap">
                  <button onClick={() => { setRenaming(l.id); setRenameValue(l.name); }} className="text-slate-400 hover:text-[#2B4C7E] mr-2 text-xs">Rename</button>
                  <button onClick={() => removeLocation(l)} className="text-slate-400 hover:text-red-600"><Trash2 size={14} /></button>
                </td>
              </tr>
            ))}
            {locations.length === 0 && <tr><td colSpan={2} className="text-center text-slate-400 py-6 text-sm">No locations yet.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Settings: Invoice Layout (for the currently active location)
--------------------------------------------------------- */
function InvoiceLayoutSettings({ data, setData, save, activeLocationName }) {
  const [form, setForm] = useState(data.invoiceSettings || { shopName: "", tagline: "", phone: "", footerNote: "" });

  const submit = (e) => {
    e.preventDefault();
    const next = { ...data, invoiceSettings: form };
    setData(next); save(next);
  };

  return (
    <div>
      <h2 className="text-lg font-semibold text-[#1F2428] mb-1">Invoice Layout — {activeLocationName}</h2>
      <p className="text-sm text-slate-500 mb-4">What appears at the top and bottom of every printed quotation, sales order, and purchase order for this location.</p>
      <form onSubmit={submit} className="space-y-3 max-w-md">
        <Field label="Shop name (shown large at the top)"><input className={inputCls} value={form.shopName} onChange={(e) => setForm({ ...form, shopName: e.target.value })} placeholder="e.g. Twig Global Hardware" /></Field>
        <Field label="Tagline / address"><input className={inputCls} value={form.tagline} onChange={(e) => setForm({ ...form, tagline: e.target.value })} placeholder="e.g. Plot 12, Main Street" /></Field>
        <Field label="Phone"><input className={inputCls} value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
        <Field label="Footer note (shown at the bottom)"><input className={inputCls} value={form.footerNote} onChange={(e) => setForm({ ...form, footerNote: e.target.value })} placeholder="e.g. Goods sold are not returnable after 7 days." /></Field>
        <Btn type="submit"><Check size={14} /> Save layout</Btn>
      </form>
    </div>
  );
}

/* ---------------------------------------------------------
   Settings hub
--------------------------------------------------------- */
function SettingsHub({ users, refreshUsers, locations, refreshLocations, data, setData, save, activeLocationName }) {
  const [tab, setTab] = useState("users");
  const tabs = [
    { id: "users", label: "Users & Rights" },
    { id: "locations", label: "Locations" },
    { id: "invoice", label: "Invoice Layout" },
  ];
  return (
    <div>
      <h1 className="text-xl font-semibold text-[#1F2428] mb-4">Settings</h1>
      <TabBar tabs={tabs} active={tab} onChange={setTab} />
      {tab === "users" && <UsersSettings users={users} refreshUsers={refreshUsers} locations={locations} />}
      {tab === "locations" && <LocationsSettings locations={locations} refreshLocations={refreshLocations} />}
      {tab === "invoice" && <InvoiceLayoutSettings data={data} setData={setData} save={save} activeLocationName={activeLocationName} />}
    </div>
  );
}

/* ---------------------------------------------------------
   App shell
--------------------------------------------------------- */
export default function App() {
  const [locations, setLocations] = useState([]);
  const [locationsLoaded, setLocationsLoaded] = useState(false);
  const [activeLocationId, setActiveLocationId] = useState(null);
  const [data, setData] = useState(emptyData());
  const [users, setUsers] = useState([]);
  const [usersLoaded, setUsersLoaded] = useState(false);
  const [currentUser, setCurrentUser] = useState(null);
  const [authChecked, setAuthChecked] = useState(false);
  const [authBusy, setAuthBusy] = useState(false);
  const [view, setView] = useState("dashboard");
  const [printJob, setPrintJob] = usePrint();

  // Locations list (shared meta)
  useEffect(() => {
    ensureLocationsList();
    const unsub = onSnapshot(LOCATIONS_DOC, (snap) => {
      setLocations(snap.exists() ? (snap.data().list || []) : []);
      setLocationsLoaded(true);
    });
    return unsub;
  }, []);

  // Users directory
  useEffect(() => {
    const unsub = onSnapshot(USERS_COL, (snap) => {
      setUsers(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
      setUsersLoaded(true);
    });
    return unsub;
  }, []);

  // Auth session
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, async (fbUser) => {
      if (!fbUser) { setCurrentUser(null); setAuthChecked(true); return; }
      const snap = await getDoc(doc(db, "users", fbUser.uid));
      if (snap.exists()) {
        setCurrentUser({ id: fbUser.uid, ...snap.data() });
      } else {
        await signOut(auth);
        setCurrentUser(null);
      }
      setAuthChecked(true);
    });
    return unsub;
  }, []);

  // Keep permissions/locations fresh if admin edits them live
  useEffect(() => {
    if (currentUser) {
      const fresh = users.find((u) => u.id === currentUser.id);
      if (fresh && JSON.stringify(fresh) !== JSON.stringify(currentUser)) setCurrentUser(fresh);
    }
  }, [users]); // eslint-disable-line

  // Pick a default active location once we know what the user can access
  const myLocations = useMemo(() => accessibleLocations(currentUser, locations), [currentUser, locations]);
  useEffect(() => {
    if (myLocations.length === 0) { setActiveLocationId(null); return; }
    if (!activeLocationId || !myLocations.some((l) => l.id === activeLocationId)) {
      setActiveLocationId(myLocations[0].id);
    }
  }, [myLocations]); // eslint-disable-line

  // Subscribe to the active location's data
  useEffect(() => {
    if (!activeLocationId) return;
    ensureLocationDoc(activeLocationId);
    const unsub = onSnapshot(locationDocRef(activeLocationId), (snap) => {
      if (snap.exists()) setData({ ...emptyData(), ...snap.data() });
    });
    return unsub;
  }, [activeLocationId]);

  const save = (next) => { if (activeLocationId) saveLocationData(activeLocationId, next); };

  const handleCreateAdmin = async ({ name, username, password }) => {
    setAuthBusy(true);
    try {
      const cred = await createUserWithEmailAndPassword(auth, toAuthEmail(username), password);
      await setDoc(doc(db, "users", cred.user.uid), { name, username: username.trim().toLowerCase(), isAdmin: true, permissions: {}, locationIds: [] });
    } finally { setAuthBusy(false); }
  };
  const handleSignUp = async ({ name, username, password }) => {
    setAuthBusy(true);
    try {
      const cred = await createUserWithEmailAndPassword(auth, toAuthEmail(username), password);
      await setDoc(doc(db, "users", cred.user.uid), { name, username: username.trim().toLowerCase(), isAdmin: false, permissions: {}, locationIds: [] });
    } finally { setAuthBusy(false); }
  };
  const handleLogin = async ({ username, password }) => {
    setAuthBusy(true);
    try { await signInWithEmailAndPassword(auth, toAuthEmail(username), password); }
    finally { setAuthBusy(false); }
  };
  const handleLogout = () => signOut(auth);

  if (!authChecked || !usersLoaded || !locationsLoaded) {
    return <div className="h-screen flex items-center justify-center bg-[#F4F4F2]"><Loader2 className="animate-spin text-[#2B4C7E]" size={28} /></div>;
  }

  if (!currentUser) {
    return <LoginScreen firstRun={users.length === 0} onLogin={handleLogin} onCreateAdmin={handleCreateAdmin} onSignUp={handleSignUp} busy={authBusy} />;
  }

  const visibleNav = MODULES.filter((m) => canAccess(currentUser, m.id));
  const missingModules = visibleNav.length === 0;
  const missingLocation = myLocations.length === 0;
  if (!currentUser.isAdmin && (missingModules || missingLocation)) {
    return <AccessPendingScreen user={currentUser} missingModules={missingModules} missingLocation={missingLocation} onLogout={handleLogout} />;
  }
  if (currentUser.isAdmin) visibleNav.push({ id: "settings", label: "Settings", icon: SettingsIcon });
  const activeView = visibleNav.some((n) => n.id === view) ? view : (visibleNav[0] ? visibleNav[0].id : "dashboard");
  const activeLocation = locations.find((l) => l.id === activeLocationId);

  return (
    <div className="min-h-screen bg-[#F4F4F2] text-[#1F2428] flex" style={{ fontFamily: "Inter, system-ui, sans-serif" }}>
      <style>{`
        @media print {
          body * { visibility: hidden; }
          #print-doc, #print-doc * { visibility: visible; }
          #print-doc { position: absolute; top: 0; left: 0; width: 100%; }
        }
      `}</style>

      <aside className="w-56 bg-[#1F2428] text-white flex-shrink-0 hidden md:flex flex-col print:hidden">
        <div className="px-5 py-5 border-b border-white/10">
          <div className="flex items-center gap-2"><Package size={20} className="text-[#D9A441]" /><span className="font-semibold tracking-tight">HardwareERP</span></div>
          <div className="text-[10px] font-mono uppercase tracking-widest text-white/40 mt-0.5">Shop System</div>
        </div>

        {myLocations.length > 1 && (
          <div className="px-5 py-3 border-b border-white/10">
            <div className="text-[10px] font-mono uppercase text-white/40 mb-1 flex items-center gap-1"><MapPin size={11} /> Location</div>
            <select value={activeLocationId || ""} onChange={(e) => setActiveLocationId(e.target.value)}
              className="w-full bg-white/10 border border-white/20 rounded-md px-2 py-1.5 text-sm text-white">
              {myLocations.map((l) => <option key={l.id} value={l.id} className="text-black">{l.name}</option>)}
            </select>
          </div>
        )}
        {myLocations.length === 1 && (
          <div className="px-5 py-2 border-b border-white/10 text-xs text-white/50 flex items-center gap-1"><MapPin size={11} /> {myLocations[0].name}</div>
        )}

        <nav className="flex-1 py-3">
          {visibleNav.map((n) => {
            const Icon = n.icon;
            const active = activeView === n.id;
            return (
              <button key={n.id} onClick={() => setView(n.id)}
                className={`w-full flex items-center gap-2.5 px-5 py-2.5 text-sm text-left transition-colors ${active ? "bg-white/10 text-[#D9A441] border-r-2 border-[#D9A441]" : "text-white/70 hover:bg-white/5 hover:text-white"}`}>
                <Icon size={16} /> {n.label}
              </button>
            );
          })}
        </nav>
        <div className="px-5 py-3 border-t border-white/10">
          <div className="flex items-center gap-2 mb-2 text-xs text-white/70"><UserIcon size={13} /> {currentUser.name} {currentUser.isAdmin && <Tag tone="good">admin</Tag>}</div>
          <button onClick={handleLogout} className="flex items-center gap-1.5 text-xs text-white/50 hover:text-white"><LogOut size={13} /> Sign out</button>
        </div>
      </aside>

      <div className="md:hidden fixed bottom-0 left-0 right-0 bg-[#1F2428] flex justify-around py-2 z-40 print:hidden overflow-x-auto">
        {visibleNav.map((n) => {
          const Icon = n.icon;
          const active = activeView === n.id;
          return (
            <button key={n.id} onClick={() => setView(n.id)} className={`flex flex-col items-center gap-0.5 px-2 flex-shrink-0 ${active ? "text-[#D9A441]" : "text-white/60"}`}>
              <Icon size={17} /><span className="text-[9px]">{n.label.split(" ")[0]}</span>
            </button>
          );
        })}
      </div>

      <main className="flex-1 p-5 md:p-8 pb-20 md:pb-8 max-w-6xl print:hidden">
        {myLocations.length > 1 && (
          <div className="md:hidden mb-4">
            <select value={activeLocationId || ""} onChange={(e) => setActiveLocationId(e.target.value)} className={inputCls}>
              {myLocations.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
            </select>
          </div>
        )}
        {activeView === "dashboard" && <Dashboard data={data} setView={setView} user={currentUser} />}
        {activeView === "sales" && <SalesHub data={data} setData={setData} save={save} setPrintJob={setPrintJob} user={currentUser} />}
        {activeView === "purchase" && <PurchaseHub data={data} setData={setData} save={save} setPrintJob={setPrintJob} user={currentUser} />}
        {activeView === "stock" && <StockHub data={data} setData={setData} save={save} user={currentUser} />}
        {activeView === "accounting" && <AccountingHub data={data} />}
        {activeView === "settings" && currentUser.isAdmin && (
          <SettingsHub users={users} refreshUsers={() => {}} locations={locations} refreshLocations={() => {}}
            data={data} setData={setData} save={save} activeLocationName={activeLocation ? activeLocation.name : ""} />
        )}
      </main>

      {printJob && printJob.mode === "doc" && <PrintableDoc {...printJob.payload} settings={data.invoiceSettings} footer="Generated by HardwareERP" />}
      {printJob && printJob.mode === "receipt" && <div className="hidden print:block" style={{ width: "58mm" }}><ReceiptDoc {...printJob.payload} settings={data.invoiceSettings} /></div>}
    </div>
  );
}
