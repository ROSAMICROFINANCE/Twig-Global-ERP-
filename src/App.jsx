import React, { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer } from "recharts";
import {
  LayoutGrid, FileText, ShoppingCart, Truck, Boxes, Calculator,
  Plus, Trash2, Printer, Upload, X, Check, AlertTriangle, Search,
  ArrowRight, Loader2, ChevronRight, Package, RotateCcw, Wallet,
  BarChart3, Settings as SettingsIcon, LogOut, Lock, User as UserIcon, Shield
} from "lucide-react";
import {
  onAuthStateChanged, signInWithEmailAndPassword, createUserWithEmailAndPassword,
  signOut, deleteUser
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

const startOfWeek = (d) => {
  const date = new Date(d);
  const day = date.getDay();
  const diff = date.getDate() - day + (day === 0 ? -6 : 1);
  const monday = new Date(date.setDate(diff));
  monday.setHours(0, 0, 0, 0);
  return monday;
};

const MODULES = [
  { id: "dashboard", label: "Dashboard", icon: LayoutGrid },
  { id: "quotes", label: "Quotations", icon: FileText },
  { id: "sales", label: "Sales Orders", icon: ShoppingCart },
  { id: "purchase", label: "Purchase Orders", icon: Truck },
  { id: "stock", label: "Stock", icon: Boxes },
  { id: "returns", label: "Goods Returns", icon: RotateCcw },
  { id: "expenses", label: "Expenses", icon: Wallet },
  { id: "reports", label: "Reports", icon: BarChart3 },
  { id: "accounting", label: "Accounting", icon: Calculator },
];

const emptyData = () => ({
  products: [], quotes: [], salesOrders: [], purchaseOrders: [], returns: [], expenses: [],
  counters: { quote: 1, sales: 1, purchase: 1, ret: 1, exp: 1 },
});

const canAccess = (user, moduleId) => !!user && (user.isAdmin || (user.permissions && user.permissions[moduleId]));

/* ---------------------------------------------------------
   Firestore data layer
   - /business/shopData  -> one shared document: products, orders, etc.
   - /users/{uid}         -> one profile per login: name, username, rights
   Passwords are never stored here — Firebase Authentication owns them,
   hashed and salted server-side. This app never sees or stores a raw
   password after account creation.
--------------------------------------------------------- */
const BUSINESS_DOC = doc(db, "business", "shopData");
const USERS_COL = collection(db, "users");

async function ensureBusinessDoc() {
  const snap = await getDoc(BUSINESS_DOC);
  if (!snap.exists()) await setDoc(BUSINESS_DOC, emptyData());
}
async function saveData(next) {
  try { await setDoc(BUSINESS_DOC, next); } catch (e) { console.error("Save failed", e); }
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

/* ---------------------------------------------------------
   Login / first-run admin setup
--------------------------------------------------------- */
function LoginScreen({ firstRun, onLogin, onCreateAdmin, busy }) {
  const [name, setName] = useState("");
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    try {
      if (firstRun) {
        if (!name || !username || !password) return setError("Fill in every field.");
        if (password.length < 6) return setError("Password needs at least 6 characters.");
        await onCreateAdmin({ name, username, password });
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
          {firstRun ? "First time setup — create the admin account." : "Sign in to your account."}
        </p>
        <form onSubmit={submit} className="space-y-3">
          {firstRun && (
            <Field label="Full name"><input className={inputCls} value={name} onChange={(e) => setName(e.target.value)} /></Field>
          )}
          <Field label="Username"><input className={inputCls} value={username} onChange={(e) => setUsername(e.target.value)} autoCapitalize="none" /></Field>
          <Field label="Password"><input type="password" className={inputCls} value={password} onChange={(e) => setPassword(e.target.value)} /></Field>
          {error && <div className="text-xs text-red-600">{error}</div>}
          <Btn type="submit" className="w-full justify-center" disabled={busy}>
            {busy ? <Loader2 size={14} className="animate-spin" /> : firstRun ? <><Shield size={14} /> Create admin account</> : <><Lock size={14} /> Sign in</>}
          </Btn>
        </form>
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
   Printable documents
--------------------------------------------------------- */
function PrintableDoc({ title, number, date, partyLabel, partyName, items, total, footer }) {
  return (
    <div id="print-doc" className="hidden print:block p-8 text-black">
      <div className="flex justify-between items-start border-b-2 border-black pb-4 mb-4">
        <div><div className="text-xl font-bold">HARDWARE SHOP</div><div className="text-xs">Point of Sale & Inventory System</div></div>
        <div className="text-right"><div className="text-lg font-bold">{title}</div><div className="text-sm font-mono">{number}</div><div className="text-sm">{date}</div></div>
      </div>
      <div className="mb-4 text-sm"><span className="text-slate-600">{partyLabel}: </span><span className="font-semibold">{partyName || "—"}</span></div>
      <table className="w-full text-sm border-collapse">
        <thead><tr className="border-b-2 border-black"><th className="text-left py-1.5">Item</th><th className="text-right py-1.5">Qty</th><th className="text-right py-1.5">Unit Price</th><th className="text-right py-1.5">Total</th></tr></thead>
        <tbody>{items.map((l, i) => (
          <tr key={i} className="border-b border-slate-300"><td className="py-1.5">{l.name}</td><td className="text-right py-1.5">{l.qty}</td><td className="text-right py-1.5">{money(l.price)}</td><td className="text-right py-1.5">{money(l.qty * l.price)}</td></tr>
        ))}</tbody>
      </table>
      <div className="flex justify-end mt-3"><div className="text-right"><div className="text-sm text-slate-600">Total</div><div className="text-2xl font-bold">{money(total)}</div></div></div>
      {footer && <div className="mt-8 text-xs text-slate-500">{footer}</div>}
    </div>
  );
}

function ReceiptDoc({ number, date, partyName, items, total }) {
  return (
    <div id="print-receipt" className="p-1" style={{ width: "58mm" }}>
      <div className="text-center font-bold text-[13px]">HARDWARE SHOP</div>
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
      <div className="text-center text-[9px] mt-2">Thank you!</div>
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
   Stock module
--------------------------------------------------------- */
function StockModule({ data, setData, save }) {
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
      <h1 className="text-xl font-semibold text-[#1F2428] mb-1">Stock</h1>
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

/* ---------------------------------------------------------
   Generic doc list
--------------------------------------------------------- */
function DocList({ title, subtitle, docs, columns, onNew, onOpen, renderTag }) {
  return (
    <div>
      <div className="flex items-center justify-between mb-1">
        <h1 className="text-xl font-semibold text-[#1F2428]">{title}</h1>
        {onNew && <Btn onClick={onNew}><Plus size={14} /> New</Btn>}
      </div>
      <p className="text-sm text-slate-500 mb-4">{subtitle}</p>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr>{columns.map((c) => <th key={c} className="text-left px-3 py-2">{c}</th>)}<th className="px-3 py-2"></th></tr>
          </thead>
          <tbody>
            {docs.map((d) => (
              <tr key={d.id} className="border-t border-slate-100 hover:bg-slate-50 cursor-pointer" onClick={() => onOpen(d)}>
                <td className="px-3 py-2 font-mono text-xs">{d.number}</td>
                <td className="px-3 py-2">{d.date}</td>
                <td className="px-3 py-2">{d.customer || d.supplier}</td>
                <td className="px-3 py-2 text-right font-mono">{money(d.total)}</td>
                <td className="px-3 py-2">{renderTag ? renderTag(d) : null}</td>
                <td className="px-3 py-2 text-right"><ChevronRight size={14} className="text-slate-300" /></td>
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
   Quotation module
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
    const salesOrder = { id: uid(), number, customer: quote.customer, date: todayISO(), items: soItems, total: quote.total, quoteId: quote.id, createdBy: user.username };
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
   Sales Order module
--------------------------------------------------------- */
function SalesOrderModule({ data, setData, save, setPrintJob, user }) {
  const [mode, setMode] = useState("list");
  const [current, setCurrent] = useState(null);
  const [customer, setCustomer] = useState("");
  const [items, setItems] = useState([]);
  const openNew = () => { setCustomer(""); setItems([]); setMode("form"); };
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
    const salesOrder = { id: uid(), number, customer, date: todayISO(), items: soItems, total, createdBy: user.username };
    const products = data.products.map((p) => { const l = items.find((l) => l.productId === p.id); return l ? { ...p, stock: p.stock - l.qty } : p; });
    const next = { ...data, products, salesOrders: [salesOrder, ...data.salesOrders], counters: { ...data.counters, sales: data.counters.sales + 1 } };
    setData(next); save(next);
    setMode("list");
  };

  if (mode === "form") {
    return (
      <div>
        <div className="flex items-center gap-2 mb-4 text-sm text-slate-500">
          <button onClick={() => setMode("list")} className="hover:text-[#2B4C7E]">Sales orders</button><ChevronRight size={13} /> <span className="text-slate-800">New sales order</span>
        </div>
        <form onSubmit={submit} className="space-y-4 max-w-3xl">
          <Field label="Customer name"><input required className={inputCls} value={customer} onChange={(e) => setCustomer(e.target.value)} /></Field>
          <LineItemsEditor items={items} setItems={setItems} products={data.products} priceField="sellPrice" />
          <div className="flex justify-between items-center">
            <div className="text-lg font-semibold">Total: {money(total)}</div>
            <div className="flex gap-2"><Btn variant="ghost" onClick={() => setMode("list")}>Cancel</Btn><Btn type="submit"><Check size={14} /> Save — deducts stock</Btn></div>
          </div>
          <p className="text-xs text-slate-400">Saving this order automatically removes the quantities above from stock.</p>
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
          <div className="mb-4"><div className="text-lg font-semibold">{current.customer}</div><div className="text-xs text-slate-400 font-mono">{current.number} · {current.date}{current.createdBy ? ` · by ${current.createdBy}` : ""}</div></div>
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
                title: "SALES ORDER", number: current.number, date: current.date, partyLabel: "Customer",
                partyName: current.customer, items: current.items.map(l => ({ ...l, name: (data.products.find(p=>p.id===l.productId)||{}).name || "—" })), total: current.total
              }})}><Printer size={14} /> Print A4 / PDF</Btn>
              <Btn onClick={() => setPrintJob({ mode: "receipt", payload: {
                number: current.number, date: current.date, partyName: current.customer,
                items: current.items.map(l => ({ ...l, name: (data.products.find(p=>p.id===l.productId)||{}).name || "—" })), total: current.total
              }})}><Printer size={14} /> Print receipt (58mm)</Btn>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return <DocList title="Sales Orders" subtitle="Confirmed sales — saving one removes stock automatically."
    docs={data.salesOrders} columns={["Number", "Date", "Customer", "Total"]} onNew={openNew} onOpen={openView} />;
}

/* ---------------------------------------------------------
   Purchase Order module
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
   Goods Returns module
--------------------------------------------------------- */
function ReturnsModule({ data, setData, save, user }) {
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
        <h1 className="text-xl font-semibold text-[#1F2428]">Goods Returns</h1>
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

/* ---------------------------------------------------------
   Expenses module
--------------------------------------------------------- */
function ExpensesModule({ data, setData, save, user }) {
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

  const today = todayISO();
  const weekStart = startOfWeek(new Date());
  const monthStart = today.slice(0, 7);
  const todayTotal = data.expenses.filter((e) => e.date === today).reduce((s, e) => s + e.amount, 0);
  const weekTotal = data.expenses.filter((e) => new Date(e.date) >= weekStart).reduce((s, e) => s + e.amount, 0);
  const monthTotal = data.expenses.filter((e) => e.date.startsWith(monthStart)).reduce((s, e) => s + e.amount, 0);

  return (
    <div>
      <div className="flex items-center justify-between mb-1">
        <h1 className="text-xl font-semibold text-[#1F2428]">Expenses</h1>
        <Btn onClick={() => setShowForm(true)}><Plus size={14} /> Add expense</Btn>
      </div>
      <p className="text-sm text-slate-500 mb-4">Daily running costs — rent, wages, transport, and more.</p>
      <div className="grid grid-cols-3 gap-3 mb-5">
        <div className="border border-slate-200 rounded-lg p-4 bg-white"><div className="text-[11px] font-mono uppercase text-slate-400 mb-1">Today</div><div className="text-xl font-semibold">{money(todayTotal)}</div></div>
        <div className="border border-slate-200 rounded-lg p-4 bg-white"><div className="text-[11px] font-mono uppercase text-slate-400 mb-1">This week</div><div className="text-xl font-semibold">{money(weekTotal)}</div></div>
        <div className="border border-slate-200 rounded-lg p-4 bg-white"><div className="text-[11px] font-mono uppercase text-slate-400 mb-1">This month</div><div className="text-xl font-semibold">{money(monthTotal)}</div></div>
      </div>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
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
   Reports module
--------------------------------------------------------- */
function ReportsModule({ data }) {
  const today = todayISO();
  const weekStart = startOfWeek(new Date());
  const monthStart = today.slice(0, 7);
  const todayOrders = data.salesOrders.filter((o) => o.date === today);
  const weekOrders = data.salesOrders.filter((o) => new Date(o.date) >= weekStart);
  const monthOrders = data.salesOrders.filter((o) => o.date.startsWith(monthStart));
  const sum = (arr) => arr.reduce((s, o) => s + o.total, 0);

  const last14 = useMemo(() => {
    const days = [];
    for (let i = 13; i >= 0; i--) {
      const d = new Date();
      d.setDate(d.getDate() - i);
      const iso = d.toISOString().slice(0, 10);
      const total = data.salesOrders.filter((o) => o.date === iso).reduce((s, o) => s + o.total, 0);
      days.push({ day: iso.slice(5), total });
    }
    return days;
  }, [data.salesOrders]);

  const cards = [
    { label: "Today's sales", orders: todayOrders.length, total: sum(todayOrders) },
    { label: "This week's sales", orders: weekOrders.length, total: sum(weekOrders) },
    { label: "This month's sales", orders: monthOrders.length, total: sum(monthOrders) },
  ];

  return (
    <div>
      <h1 className="text-xl font-semibold text-[#1F2428] mb-1">Reports</h1>
      <p className="text-sm text-slate-500 mb-4">Sales performance by day, week, and month.</p>
      <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mb-6">
        {cards.map((c) => (
          <div key={c.label} className="border border-slate-200 rounded-lg p-4 bg-white">
            <div className="text-[11px] font-mono uppercase text-slate-400 mb-1">{c.label}</div>
            <div className="text-2xl font-semibold">{money(c.total)}</div>
            <div className="text-xs text-slate-400 mt-1">{c.orders} order{c.orders === 1 ? "" : "s"}</div>
          </div>
        ))}
      </div>
      <div className="border border-slate-200 rounded-lg bg-white p-4 mb-6">
        <div className="text-sm font-medium mb-3">Last 14 days</div>
        <div style={{ width: "100%", height: 220 }}>
          <ResponsiveContainer>
            <BarChart data={last14}>
              <XAxis dataKey="day" tick={{ fontSize: 10 }} />
              <YAxis tick={{ fontSize: 10 }} />
              <Tooltip formatter={(v) => money(v)} />
              <Bar dataKey="total" fill="#2B4C7E" radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <div className="px-4 py-3 border-b border-slate-100 font-medium text-sm">Today's orders</div>
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr><th className="text-left px-3 py-2">Number</th><th className="text-left px-3 py-2">Customer</th><th className="text-right px-3 py-2">Total</th></tr>
          </thead>
          <tbody>
            {todayOrders.map((o) => (
              <tr key={o.id} className="border-t border-slate-100"><td className="px-3 py-2 font-mono text-xs">{o.number}</td><td className="px-3 py-2">{o.customer}</td><td className="px-3 py-2 text-right font-mono">{money(o.total)}</td></tr>
            ))}
            {todayOrders.length === 0 && <tr><td colSpan={3} className="text-center text-slate-400 py-6 text-sm">No sales recorded today yet.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Accounting module
--------------------------------------------------------- */
function AccountingModule({ data }) {
  const rows = useMemo(() => {
    const map = {};
    data.salesOrders.forEach((o) => {
      o.items.forEach((l) => {
        if (!map[l.productId]) map[l.productId] = { qty: 0, revenue: 0, cost: 0 };
        map[l.productId].qty += l.qty;
        map[l.productId].revenue += l.qty * l.price;
        map[l.productId].cost += l.qty * (l.cost || 0);
      });
    });
    return data.products.map((p) => {
      const m = map[p.id] || { qty: 0, revenue: 0, cost: 0 };
      const profit = m.revenue - m.cost;
      const margin = m.revenue ? (profit / m.revenue) * 100 : 0;
      return { ...p, ...m, profit, margin };
    }).filter((r) => r.qty > 0).sort((a, b) => b.profit - a.profit);
  }, [data]);

  const totalRevenue = rows.reduce((s, r) => s + r.revenue, 0);
  const totalCost = rows.reduce((s, r) => s + r.cost, 0);
  const totalExpenses = data.expenses.reduce((s, e) => s + e.amount, 0);
  const damagedLoss = data.returns.filter((r) => r.type === "damaged").reduce((s, r) => s + (r.value || 0), 0);
  const grossProfit = totalRevenue - totalCost;
  const netProfit = grossProfit - totalExpenses - damagedLoss;
  const totalMargin = totalRevenue ? (grossProfit / totalRevenue) * 100 : 0;

  return (
    <div>
      <h1 className="text-xl font-semibold text-[#1F2428] mb-1">Accounting</h1>
      <p className="text-sm text-slate-500 mb-4">Cost, gross profit and margin — company-wide and per product.</p>
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-3">
        <div className="border border-slate-200 rounded-lg p-4 bg-white"><div className="text-[11px] font-mono uppercase text-slate-400 mb-1">Revenue</div><div className="text-xl font-semibold">{money(totalRevenue)}</div></div>
        <div className="border border-slate-200 rounded-lg p-4 bg-white"><div className="text-[11px] font-mono uppercase text-slate-400 mb-1">Cost of goods</div><div className="text-xl font-semibold">{money(totalCost)}</div></div>
        <div className="border border-slate-200 rounded-lg p-4 bg-white"><div className="text-[11px] font-mono uppercase text-slate-400 mb-1">Gross profit</div><div className="text-xl font-semibold text-emerald-700">{money(grossProfit)}</div></div>
        <div className="border border-slate-200 rounded-lg p-4 bg-white"><div className="text-[11px] font-mono uppercase text-slate-400 mb-1">Margin</div><div className="text-xl font-semibold">{totalMargin.toFixed(1)}%</div></div>
      </div>
      <div className="grid grid-cols-2 md:grid-cols-3 gap-3 mb-6">
        <div className="border border-slate-200 rounded-lg p-4 bg-white"><div className="text-[11px] font-mono uppercase text-slate-400 mb-1">Expenses</div><div className="text-xl font-semibold">{money(totalExpenses)}</div></div>
        <div className="border border-slate-200 rounded-lg p-4 bg-white"><div className="text-[11px] font-mono uppercase text-slate-400 mb-1">Damaged stock loss</div><div className="text-xl font-semibold text-red-600">{money(damagedLoss)}</div></div>
        <div className="border border-slate-200 rounded-lg p-4 bg-white"><div className="text-[11px] font-mono uppercase text-slate-400 mb-1">Net profit</div><div className={`text-xl font-semibold ${netProfit >= 0 ? "text-emerald-700" : "text-red-600"}`}>{money(netProfit)}</div></div>
      </div>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr><th className="text-left px-3 py-2">Product</th><th className="text-right px-3 py-2">Qty sold</th><th className="text-right px-3 py-2">Revenue</th><th className="text-right px-3 py-2">Cost</th><th className="text-right px-3 py-2">Profit</th><th className="text-right px-3 py-2">Margin</th></tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.id} className="border-t border-slate-100">
                <td className="px-3 py-2 font-medium">{r.name}</td>
                <td className="px-3 py-2 text-right font-mono">{r.qty}</td>
                <td className="px-3 py-2 text-right font-mono">{money(r.revenue)}</td>
                <td className="px-3 py-2 text-right font-mono">{money(r.cost)}</td>
                <td className="px-3 py-2 text-right font-mono text-emerald-700">{money(r.profit)}</td>
                <td className="px-3 py-2 text-right font-mono">{r.margin.toFixed(1)}%</td>
              </tr>
            ))}
            {rows.length === 0 && <tr><td colSpan={6} className="text-center text-slate-400 py-8 text-sm">No sales recorded yet.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Settings module (admin only) — manage users & module rights
   Passwords are set once at account creation via Firebase Auth
   and are never stored or shown here again.
--------------------------------------------------------- */
function SettingsModule({ users, refreshUsers }) {
  const [showForm, setShowForm] = useState(false);
  const [editingUser, setEditingUser] = useState(null);
  const blank = { name: "", username: "", password: "", permissions: {} };
  const [form, setForm] = useState(blank);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  const openNew = () => { setForm(blank); setEditingUser(null); setError(""); setShowForm(true); };
  const openEdit = (u) => { setForm({ name: u.name, username: u.username, password: "", permissions: u.permissions || {} }); setEditingUser(u); setError(""); setShowForm(true); };
  const togglePerm = (moduleId) => setForm({ ...form, permissions: { ...form.permissions, [moduleId]: !form.permissions[moduleId] } });

  const submit = async (e) => {
    e.preventDefault();
    setError(""); setBusy(true);
    try {
      if (editingUser) {
        // Editing an existing account: update name + permissions only.
        // Password changes for other staff require the Firebase console
        // or an admin backend function — not available from the browser alone.
        await setDoc(doc(db, "users", editingUser.id), {
          name: form.name, username: editingUser.username, isAdmin: false, permissions: form.permissions,
        });
      } else {
        if (form.password.length < 6) { setError("Password needs at least 6 characters."); setBusy(false); return; }
        // Create the new login using a secondary Firebase app instance so
        // the admin's own session stays signed in.
        const { authInstance, appInstance } = getSecondaryAuth();
        const cred = await createUserWithEmailAndPassword(authInstance, toAuthEmail(form.username), form.password);
        await setDoc(doc(db, "users", cred.user.uid), {
          name: form.name, username: form.username.trim().toLowerCase(), isAdmin: false, permissions: form.permissions,
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
        <h1 className="text-xl font-semibold text-[#1F2428]">Settings — Users & Rights</h1>
        <Btn onClick={openNew}><Plus size={14} /> Add user</Btn>
      </div>
      <p className="text-sm text-slate-500 mb-4">Only the admin account can create logins and choose which modules each person can see and use.</p>
      <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-[11px] font-mono uppercase tracking-wide text-slate-500">
            <tr><th className="text-left px-3 py-2">Name</th><th className="text-left px-3 py-2">Username</th><th className="text-left px-3 py-2">Access</th><th className="w-24"></th></tr>
          </thead>
          <tbody>
            {users.map((u) => (
              <tr key={u.id} className="border-t border-slate-100">
                <td className="px-3 py-2 font-medium">{u.name}</td>
                <td className="px-3 py-2 font-mono text-xs">{u.username}</td>
                <td className="px-3 py-2">
                  {u.isAdmin ? <Tag tone="good">Full admin access</Tag> :
                    <div className="flex flex-wrap gap-1">
                      {MODULES.filter((m) => u.permissions && u.permissions[m.id]).map((m) => <Tag key={m.id}>{m.label}</Tag>)}
                      {!(u.permissions && Object.values(u.permissions).some(Boolean)) && <span className="text-slate-400 text-xs">No modules assigned</span>}
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
   App shell
--------------------------------------------------------- */
export default function App() {
  const [data, setData] = useState(emptyData());
  const [users, setUsers] = useState([]);
  const [usersLoaded, setUsersLoaded] = useState(false);
  const [currentUser, setCurrentUser] = useState(null);
  const [authChecked, setAuthChecked] = useState(false);
  const [authBusy, setAuthBusy] = useState(false);
  const [view, setView] = useState("dashboard");
  const [printJob, setPrintJob] = usePrint();

  // Live business data (shared across the whole team)
  useEffect(() => {
    ensureBusinessDoc();
    const unsub = onSnapshot(BUSINESS_DOC, (snap) => {
      if (snap.exists()) setData({ ...emptyData(), ...snap.data() });
    });
    return unsub;
  }, []);

  // Live user directory (for first-run detection and Settings)
  useEffect(() => {
    const unsub = onSnapshot(USERS_COL, (snap) => {
      setUsers(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
      setUsersLoaded(true);
    });
    return unsub;
  }, []);

  // Track who's signed in
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, async (fbUser) => {
      if (!fbUser) { setCurrentUser(null); setAuthChecked(true); return; }
      const snap = await getDoc(doc(db, "users", fbUser.uid));
      if (snap.exists()) {
        setCurrentUser({ id: fbUser.uid, ...snap.data() });
      } else {
        // Their profile was removed by the admin — access revoked.
        await signOut(auth);
        setCurrentUser(null);
      }
      setAuthChecked(true);
    });
    return unsub;
  }, []);

  // Keep currentUser's permissions fresh if admin edits them live
  useEffect(() => {
    if (currentUser) {
      const fresh = users.find((u) => u.id === currentUser.id);
      if (fresh && JSON.stringify(fresh) !== JSON.stringify(currentUser)) setCurrentUser(fresh);
    }
  }, [users]); // eslint-disable-line

  const save = (next) => saveData(next);

  const handleCreateAdmin = async ({ name, username, password }) => {
    setAuthBusy(true);
    try {
      const cred = await createUserWithEmailAndPassword(auth, toAuthEmail(username), password);
      await setDoc(doc(db, "users", cred.user.uid), { name, username: username.trim().toLowerCase(), isAdmin: true, permissions: {} });
    } finally { setAuthBusy(false); }
  };
  const handleLogin = async ({ username, password }) => {
    setAuthBusy(true);
    try { await signInWithEmailAndPassword(auth, toAuthEmail(username), password); }
    finally { setAuthBusy(false); }
  };
  const handleLogout = () => signOut(auth);

  if (!authChecked || !usersLoaded) {
    return <div className="h-screen flex items-center justify-center bg-[#F4F4F2]"><Loader2 className="animate-spin text-[#2B4C7E]" size={28} /></div>;
  }

  if (!currentUser) {
    return <LoginScreen firstRun={users.length === 0} onLogin={handleLogin} onCreateAdmin={handleCreateAdmin} busy={authBusy} />;
  }

  const visibleNav = MODULES.filter((m) => canAccess(currentUser, m.id));
  if (currentUser.isAdmin) visibleNav.push({ id: "settings", label: "Settings", icon: SettingsIcon });
  const activeView = visibleNav.some((n) => n.id === view) ? view : (visibleNav[0] ? visibleNav[0].id : "dashboard");

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
        {activeView === "dashboard" && <Dashboard data={data} setView={setView} user={currentUser} />}
        {activeView === "quotes" && <QuotationModule data={data} setData={setData} save={save} setPrintJob={setPrintJob} user={currentUser} />}
        {activeView === "sales" && <SalesOrderModule data={data} setData={setData} save={save} setPrintJob={setPrintJob} user={currentUser} />}
        {activeView === "purchase" && <PurchaseOrderModule data={data} setData={setData} save={save} setPrintJob={setPrintJob} user={currentUser} />}
        {activeView === "stock" && <StockModule data={data} setData={setData} save={save} />}
        {activeView === "returns" && <ReturnsModule data={data} setData={setData} save={save} user={currentUser} />}
        {activeView === "expenses" && <ExpensesModule data={data} setData={setData} save={save} user={currentUser} />}
        {activeView === "reports" && <ReportsModule data={data} />}
        {activeView === "accounting" && <AccountingModule data={data} />}
        {activeView === "settings" && currentUser.isAdmin && <SettingsModule users={users} refreshUsers={() => {}} />}
      </main>

      {printJob && printJob.mode === "doc" && <PrintableDoc {...printJob.payload} footer="Generated by HardwareERP" />}
      {printJob && printJob.mode === "receipt" && <div className="hidden print:block" style={{ width: "58mm" }}><ReceiptDoc {...printJob.payload} /></div>}
    </div>
  );
}
