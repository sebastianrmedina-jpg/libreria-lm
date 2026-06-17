/* eslint-disable */
import { createClient } from "@supabase/supabase-js";
import React, { useState, useMemo, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── MOBILE HOOK ─────────────────────────────────────────────────────────────
function useIsMobile() {
  const [mob, setMob] = useState(typeof window !== "undefined" && window.innerWidth < 768);
  useEffect(() => {
    const fn = () => setMob(window.innerWidth < 768);
    window.addEventListener("resize", fn);
    return () => window.removeEventListener("resize", fn);
  }, []);
  return mob;
}

// ─── SUPABASE CONFIG ──────────────────────────────────────────────────────────
const SUPA_URL = "https://pqwcegwadffzqecmbqbe.supabase.co";
const SUPA_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBxd2NlZ3dhZGZmenFlY21icWJlIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODA0OTgzNjgsImV4cCI6MjA5NjA3NDM2OH0.XnmgmzabW4YV4SrP1YNDtRElp7aNoGjbG37XG6VvXak";
const SUPA_SERVICE = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBxd2NlZ3dhZGZmenFlY21icWJlIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDQ5ODM2OCwiZXhwIjoyMDk2MDc0MzY4fQ.dPLccDyKEZ4YPZovD4iXrSJe-DoNf6UFDMNWCmWiK84";
const supabase = createClient(SUPA_URL, SUPA_ANON);
const supaAdmin = supabase;

const mapProduct = r => ({id:r.id,name:r.name,category:r.category,costPrice:r.cost_price,salePrice:r.sale_price,stock:r.stock});
const mapOrder = r => ({id:r.id,client:r.client,vendedor:r.vendedor,notes:r.notes,total:r.total,stage:r.stage,date:r.date,items:r.items||[],docNum:r.doc_num||"",compNum:r.comp_num||"",isTest:r.is_test||false,isSandbox:r.is_sandbox||false,internalNote:r.internal_note||"",editStatus:r.edit_status||"",editReason:r.edit_reason||"",editItems:r.edit_items||null,editRejectReason:r.edit_reject_reason||""});
const mapQuote = r => ({id:r.id,client:r.client,vendedor:r.vendedor,notes:r.notes,total:r.total,date:r.date,items:r.items||[],validity:r.validity||"",docNum:r.doc_num||"",convertida:r.convertida||false,ordenId:r.orden_id||"",extendida:r.extendida||false,extendReason:r.extend_reason||"",extendDate:r.extend_date||"",globalDisc:r.global_disc||null,subtotal:r.subtotal||0});


// ─── CORRELATIVE NUMBER HELPERS ───────────────────────────────────────────────
// Table lm_counters: { id: "reserva"|"comp"|"presu", value: number }
// SQL: CREATE TABLE lm_counters (id TEXT PRIMARY KEY, value INTEGER DEFAULT 0);
// INSERT INTO lm_counters (id,value) VALUES ('reserva',0),('comp',0),('presu',0);
const padNum = (n) => String(n).padStart(6,"0");
const fmtDocNum = (prefix, n) => `${prefix}-${padNum(n)}`;

// ─── TEST MODE ────────────────────────────────────────────────────────────────
// Vendedor "Prueba": no incrementa contadores, numera como TEST-000000
// y agrega marca de agua en el PDF
const TEST_VENDOR = "Prueba";
const isTestOrder = (vendedor) => vendedor === TEST_VENDOR;

const db = {
  getUsers:     async () => { const {data,error} = await supabase.from("lm_users").select("*").order("name"); if(error) throw error; return (data||[]).map(u=>({...u,priceList:u.price_list||"default",vendedor:u.vendedor||"",canSeeAll:u.can_see_all!==false,phone:u.phone||"",cargo:u.cargo||"",avatar:u.avatar||"",barcodeEnabled:u.barcode_enabled||false})); },
  saveUser:     async (u) => {
    const {error} = await supaAdmin.from("lm_users").upsert({
      id: u.id,
      username: u.username,
      password: u.password,
      name: u.name,
      role: u.role||"vendedor",
      email: u.email||"",
      vendedor: u.vendedor||"",
      price_list: u.priceList||"default",
      can_see_all: u.canSeeAll!==false,
      phone: u.phone||"",
      cargo: u.cargo||"",
      avatar: u.avatar||"",
      barcode_enabled: u.barcodeEnabled||false,
    });
    if(error) { console.error("saveUser error:", error); throw error; }
  },
  deleteUser:   async (id) => { const {error} = await supaAdmin.from("lm_users").delete().eq("id",id); if(error) throw error; },

  getVendors:   async () => { const {data,error} = await supabase.from("lm_vendors").select("name").order("name"); if(error) throw error; return (data||[]).map(v=>v.name); },
  addVendor:    async (name) => { const {error} = await supaAdmin.from("lm_vendors").insert({name}); if(error) throw error; },
  deleteVendor: async (name) => { const {error} = await supaAdmin.from("lm_vendors").delete().eq("name",name); if(error) throw error; },
  updateVendor: async (old,nw) => { const {error} = await supaAdmin.from("lm_vendors").update({name:nw}).eq("name",old); if(error) throw error; },

  getProducts:  async () => {
    // Supabase tiene limite de 1000 filas por query — traemos paginas hasta tener todos
    let all = [];
    let from = 0;
    const pageSize = 1000;
    while(true) {
      const {data,error} = await supabase.from("lm_products").select("*").order("name").range(from, from+pageSize-1);
      if(error) throw error;
      if(!data || data.length===0) break;
      all = all.concat(data);
      if(data.length < pageSize) break;
      from += pageSize;
    }
    return all.map(mapProduct);
  },
  upsertProduct: async (p) => { const {error} = await supaAdmin.from("lm_products").upsert({id:p.id,name:p.name,category:p.category||"Importado",cost_price:p.costPrice||0,sale_price:p.salePrice||0,stock:p.stock||0}); if(error) throw error; },
  upsertProducts: async (arr) => { const {error} = await supaAdmin.from("lm_products").upsert(arr.map(p=>({id:p.id,name:p.name,category:p.category||"Importado",cost_price:p.costPrice||0,sale_price:p.salePrice||0,stock:p.stock||0}))); if(error) throw error; },
  deleteProduct: async (id) => { const {error} = await supaAdmin.from("lm_products").delete().eq("id",id); if(error) throw error; },

  getOrders:    async () => { const {data,error} = await supabase.from("lm_orders").select("*").order("date",{ascending:false}); if(error) throw error; return (data||[]).map(mapOrder); },
  upsertOrder:  async (o) => { const {error} = await supaAdmin.from("lm_orders").upsert({id:o.id,client:o.client,vendedor:o.vendedor||"",notes:o.notes||"",total:o.total,stage:o.stage,date:o.date,items:o.items,doc_num:o.docNum||"",comp_num:o.compNum||"",is_test:o.isTest||false,is_sandbox:o.isSandbox||false,internal_note:o.internalNote||"",edit_status:o.editStatus||"",edit_reason:o.editReason||"",edit_items:o.editItems||null,edit_reject_reason:o.editRejectReason||""}); if(error) throw error; },
  deleteOrder:  async (id) => { const {error} = await supaAdmin.from("lm_orders").delete().eq("id",id); if(error) throw error; },

  getQuotes:    async () => { const {data,error} = await supabase.from("lm_quotes").select("*").order("date",{ascending:false}); if(error) throw error; return (data||[]).map(mapQuote); },
  upsertQuote:  async (q) => { const {error} = await supaAdmin.from("lm_quotes").upsert({id:q.id,client:q.client,vendedor:q.vendedor||"",notes:q.notes||"",total:q.total,date:q.date,items:q.items,validity:q.validity||"",doc_num:q.docNum||"",convertida:q.convertida||false,orden_id:q.ordenId||""}); if(error) throw error; },
  deleteQuote:  async (id) => { const {error} = await supaAdmin.from("lm_quotes").delete().eq("id",id); if(error) throw error; },

  getStockLog:  async () => { const {data,error} = await supabase.from("lm_stocklog").select("*").order("fecha",{ascending:false}); if(error) throw error; return (data||[]).map(r=>({...r,productoId:r.producto_id,stockAntes:r.stock_antes,stockDespues:r.stock_despues})); },
  addStockLog:  async (e) => { const {error} = await supaAdmin.from("lm_stocklog").insert({id:e.id,fecha:e.fecha,usuario:e.usuario,rol:e.rol,tipo:e.tipo,producto_id:e.productoId,producto:e.producto,stock_antes:e.stockAntes,stock_despues:e.stockDespues,cambio:e.cambio,motivo:e.motivo}); if(error) throw error; },
  clearStockLog: async () => { const {error} = await supaAdmin.from("lm_stocklog").delete().neq("id","none"); if(error) throw error; },

  getNotifs:    async () => { const {data,error} = await supabase.from("lm_notifs").select("*").order("fecha",{ascending:false}); if(error) throw error; return data||[]; },
  addNotif:     async (n) => { const {error} = await supaAdmin.from("lm_notifs").insert(n); if(error) throw error; },
  updateNotif:  async (id,data) => { const {error} = await supaAdmin.from("lm_notifs").update(data).eq("id",id); if(error) throw error; },
  deleteNotif:  async (id) => { const {error} = await supaAdmin.from("lm_notifs").delete().eq("id",id); if(error) throw error; },
  clearNotifs:  async () => { const {error} = await supaAdmin.from("lm_notifs").delete().neq("id","none"); if(error) throw error; },
  // Clients
  getClients:   async () => { const {data,error} = await supabase.from("lm_clients").select("*").order("name"); if(error) throw error; return (data||[]).map(r=>({id:r.id,name:r.name,phone:r.phone||"",email:r.email||"",cuit:r.cuit||"",address:r.address||"",notes:r.notes||"",deleteRequested:r.delete_requested||false,deleteReason:r.delete_reason||"",createdBy:r.created_by||"",createdAt:r.created_at||""})); },
  saveClient:   async (c) => { const {error} = await supaAdmin.from("lm_clients").upsert({id:c.id,name:c.name,phone:c.phone||"",email:c.email||"",cuit:c.cuit||"",address:c.address||"",notes:c.notes||"",delete_requested:c.deleteRequested||false,delete_reason:c.deleteReason||"",created_by:c.createdBy||"",created_at:c.createdAt||""}); if(error) throw error; },
  deleteClient: async (id) => { const {error} = await supaAdmin.from("lm_clients").delete().eq("id",id); if(error) throw error; },
  // Counters
  nextCounter: async (id) => {
    const {data,error} = await supaAdmin.rpc("increment_counter", {counter_id: id});
    if(error) {
      const {data:row} = await supaAdmin.from("lm_counters").select("value").eq("id",id).single();
      const next = (row?.value||0) + 1;
      await supaAdmin.from("lm_counters").upsert({id, value:next});
      return next;
    }
    return data;
  },
  // Price lists
  getPriceLists: async () => { const {data,error} = await supabase.from("lm_pricelists").select("*").order("name"); if(error) throw error; return data||[]; },
  savePriceList: async (pl) => { const {error} = await supaAdmin.from("lm_pricelists").upsert(pl); if(error) throw error; },
  deletePriceList: async (id) => { const {error} = await supaAdmin.from("lm_pricelists").delete().eq("id",id); if(error) throw error; },
  // Purchase orders (solicitudes de compra)
  // SQL: CREATE TABLE lm_purchase_orders (id TEXT PRIMARY KEY, fecha TEXT, vendedor TEXT, estado TEXT DEFAULT 'abierta', items JSONB DEFAULT '[]', notas TEXT DEFAULT '', fecha_cierre TEXT DEFAULT '');
  getPurchaseOrders: async () => { const {data,error} = await supabase.from("lm_purchase_orders").select("*").order("fecha",{ascending:false}); if(error) throw error; return data||[]; },
  savePurchaseOrder: async (po) => { const {error} = await supaAdmin.from("lm_purchase_orders").upsert({id:po.id,fecha:po.fecha,vendedor:po.vendedor,estado:po.estado,items:po.items,notas:po.notas||"",fecha_cierre:po.fechaCierre||""}); if(error) throw error; },
  deletePurchaseOrder: async (id) => { const {error} = await supaAdmin.from("lm_purchase_orders").delete().eq("id",id); if(error) throw error; },
  // Activity log
  // SQL: CREATE TABLE lm_activity (id TEXT PRIMARY KEY, fecha TEXT, usuario TEXT, rol TEXT, accion TEXT, detalle TEXT, ref_id TEXT, ref_tipo TEXT);
  getActivity:  async () => { const {data,error} = await supabase.from("lm_activity").select("*").order("fecha",{ascending:false}).limit(500); if(error) throw error; return data||[]; },
  addActivity:  async (a) => { try { await supaAdmin.from("lm_activity").insert(a); } catch(e) {} }, // silent fail - non critical
  clearActivity: async () => { const {error} = await supaAdmin.from("lm_activity").delete().neq("id","none"); if(error) throw error; },
};

const RED = "#c0392b", REDD = "#922b21";
const VINO = "#7b1a1a";          // header, máxima jerarquía
const GOLD = "#c9a96a";          // detalles, líneas separadoras
const IVORY = "#fdfbf7";         // fondos cálidos
const IVORY_BORDER = "#e5ddd0";
const SERIF = "Georgia,'Times New Roman',serif";
const STAGE_COLORS = {
  reserva:     "#b7770d",
  confirmado:  "#1e8449",
  "en armado": "#1a5276",
  entregado:   "#6c3483",
};
const STAGES = ["reserva","confirmado","en armado","entregado"];
const SCFG = {
  reserva:     {label:"Reserva",    color:"#b7770d", bg:"#fef9e7", icon:"🕐"},
  confirmado:  {label:"Confirmado", color:"#1e8449", bg:"#eafaf1", icon:"✅"},
  "en armado": {label:"En Armado",  color:"#1a5276", bg:"#eaf4fc", icon:"📦"},
  entregado:   {label:"Entregado",  color:"#6c3483", bg:"#f5eef8", icon:"🎉"},
};

const LOGO = "/logo.png";
const fARS = n => "$" + Number(n).toLocaleString("es-AR",{minimumFractionDigits:2,maximumFractionDigits:2});
const genId = () => Date.now().toString(36)+Math.random().toString(36).slice(2);
const today = () => new Date().toLocaleDateString("es-AR");

const DEFAULT_USERS = [{id:"u1",username:"admin",password:"admin123",role:"admin",name:"Administrador"}];

function useLocalData(key, initial) {
  const [data, setData] = useState(() => {
    try { const s = localStorage.getItem(key); return s ? JSON.parse(s) : initial; } catch { return initial; }
  });
  const setter = (val) => {
    const next = typeof val === "function" ? val(data) : val;
    try { localStorage.setItem(key, JSON.stringify(next)); } catch {}
    setData(next);
  };
  return [data, setter];
}

function SPill({n}) {
  const [c,b] = n===0?["#c0392b","#fdecea"]:n<=5?["#e67e22","#fef9e7"]:["#1e8449","#d5f5e3"];
  return <span style={{background:b,color:c,borderRadius:10,padding:"2px 8px",fontSize:11,fontWeight:700}}>{n===0?"Sin stock":n<=5?`🎉 ${n}`:n}</span>;
}
function Bdg({stage}) {
  const c=SCFG[stage]||{};
  return <span style={{background:c.bg,color:c.color,border:`1px solid ${c.color}44`,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{c.icon} {c.label}</span>;
}
// ─── SAVE SPINNER ────────────────────────────────────────────────────────────
function SaveSpinner({label="Guardando...", color="#c0392b"}) {
  return (
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"60px 20px",gap:16}}>
      <div style={{width:48,height:48,borderRadius:"50%",border:`4px solid ${color}22`,borderTop:`4px solid ${color}`,animation:"lm-spin 0.8s linear infinite"}}/>
      <div style={{fontWeight:700,fontSize:14,color:"#555"}}>{label}</div>
      <style>{`@keyframes lm-spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}

function Field({label,children}) {
  return <div style={{marginBottom:12}}>
    <label style={{fontSize:12,fontWeight:600,color:"#666",display:"block",marginBottom:4}}>{label}</label>
    {children}
  </div>;
}
const inputStyle = {width:"100%",padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",boxSizing:"border-box",background:"#fff"};

// ─── DISCOUNT HELPERS ────────────────────────────────────────────────────────
// Returns discounted price for one item
const applyItemDiscount = (price, qty, disc) => {
  if(!disc || (!disc.value)) return price * qty;
  const v = parseFloat(disc.value) || 0;
  if(disc.type === "%") return Math.max(0, price - price * v / 100) * qty;
  return Math.max(0, price * qty - v);
};
// Returns final total after global discount
const applyGlobalDiscount = (subtotal, disc) => {
  if(!disc || (!disc.value)) return subtotal;
  const v = parseFloat(disc.value) || 0;
  if(disc.type === "%") return Math.max(0, subtotal - subtotal * v / 100);
  return Math.max(0, subtotal - v);
};
// Format discount label for display
// ─── WHATSAPP MESSAGE BUILDER ────────────────────────────────────────────────
function buildWAOrder(o) {
  const num   = o.compNum || o.docNum || "";
  const tipo  = o.compNum ? "Comprobante" : "Reserva";
  const stage = {"reserva":"Reserva","confirmado":"Confirmado","armado":"En Armado","entregado":"Entregado"}[o.stage] || o.stage;
  const sep   = "─────────────────────";
  const items = o.items.map(it => {
    const sub  = it.price * it.qty;
    const disc = it.disc?.value ? ` (-${it.disc.value}${it.disc.type==="%"?"%":"$"})` : "";
    return `\u25b8 ${it.name}\n  ${it.qty} u. \u00d7 ${fARS(it.price)}${disc} = *${fARS(sub)}*`;
  }).join("\n");
  const discLine = o.subtotal && o.subtotal !== o.total
    ? `\n_Descuento: -${fARS(o.subtotal - o.total)}_` : "";
  const notes = o.notes ? `\n\uD83D\uDCAC _${o.notes}_` : "";
  return [
    `\uD83C\uDFEA *LIBRERIA MADRID*`,
    `\uD83D\uDCC4 *${tipo}: ${num}*`,
    sep,
    `\uD83D\uDC64 *Cliente:* ${o.client}`,
    `\uD83D\uDCC5 *Fecha:* ${o.date}`,
    `\uD83D\uDCE6 *Estado:* ${stage}`,
    sep,
    items,
    sep,
    `${discLine ? `Subtotal: ${fARS(o.subtotal)}${discLine}\n` : ""}*TOTAL: ${fARS(o.total)}*`,
    notes,
    sep,
    `_Libreria Madrid \u2014 ${o.vendedor || ""}_`,
  ].filter(Boolean).join("\n");
}

function buildWAQuote(q) {
  const sep   = "─────────────────────";
  const items = q.items.map(it => {
    const sub  = it.price * it.qty;
    const disc = it.disc?.value ? ` (-${it.disc.value}${it.disc.type==="%"?"%":"$"})` : "";
    return `\u25b8 ${it.name}\n  ${it.qty} u. \u00d7 ${fARS(it.price)}${disc} = *${fARS(sub)}*`;
  }).join("\n");
  const discLine = q.subtotal && q.subtotal !== q.total
    ? `\n_Descuento: -${fARS(q.subtotal - q.total)}_` : "";
  const notes = q.notes ? `\n\uD83D\uDCAC _${q.notes}_` : "";
  return [
    `\uD83C\uDFEA *LIBRERIA MADRID*`,
    `\uD83D\uDCCB *Cotizacion: ${q.docNum || ""}*`,
    sep,
    `\uD83D\uDC64 *Cliente:* ${q.client}`,
    `\uD83D\uDCC5 *Fecha:* ${q.date}`,
    `\u23F3 *Valida:* ${q.validity || "48 horas"}`,
    sep,
    items,
    sep,
    `${discLine ? `Subtotal: ${fARS(q.subtotal)}${discLine}\n` : ""}*TOTAL: ${fARS(q.total)}*`,
    notes,
    sep,
    `_Libreria Madrid \u2014 ${q.vendedor || ""}_`,
  ].filter(Boolean).join("\n");
}

const fmtDisc = (disc) => {
  if(!disc || !disc.value) return null;
  const v = parseFloat(disc.value) || 0;
  if(!v) return null;
  return disc.type === "%" ? `-${v}%` : `-${fARS(v)}`;
};

// Normaliza texto para búsqueda: minúsculas + sin acentos
// "Bolígrafo" -> "boligrafo", "Ñoño" -> "nono"
const norm = (s) => String(s||"").toLowerCase()
  .normalize("NFD").replace(/[̀-ͯ]/g, "");
// Normaliza SKU quitando guiones, puntos, espacios para búsqueda flexible
const normSKU = (s) => String(s||"").toLowerCase().replace(/[-.\s_]/g, "");

// ─── QUOTE EXPIRY HELPERS ─────────────────────────────────────────────────────
// Parsea fecha "dd/mm/yyyy" a Date
function parseDate(dateStr) {
  if(!dateStr) return null;
  const parts = dateStr.split("/");
  if(parts.length !== 3) return null;
  return new Date(parseInt(parts[2]), parseInt(parts[1])-1, parseInt(parts[0]));
}

// Estado de vencimiento de una cotización
function quoteStatus(q) {
  if(q.convertida) return "convertida";
  const created = parseDate(q.date);
  if(!created) return "vigente";
  const now = new Date();
  const hoursElapsed = (now - created) / (1000 * 60 * 60);
  const extendedHours = q.extendida ? 96 : 48; // 48hs + 48hs extensión
  if(hoursElapsed <= extendedHours) {
    const hoursLeft = Math.ceil(extendedHours - hoursElapsed);
    return { status: q.extendida ? "extendida" : "vigente", hoursLeft };
  }
  return { status: "vencida", hoursLeft: 0 };
}

// ─── ONESIGNAL PUSH NOTIFICATIONS ───────────────────────────────────────────
const OS_APP_ID = "00dbe0c1-e7bf-4f80-885c-1680381ed121";

// Espera hasta que OneSignal esté listo (máx 10 segundos)
function getOneSignal() {
  return new Promise((resolve) => {
    if(window.OneSignal && window._oneSignalReady) {
      resolve(window.OneSignal);
      return;
    }
    const handler = () => resolve(window.OneSignal);
    window.addEventListener("oneSignalReady", handler, { once: true });
    // Timeout de seguridad
    setTimeout(() => {
      window.removeEventListener("oneSignalReady", handler);
      resolve(window.OneSignal || null);
    }, 10000);
  });
}
const OS_API_KEY = "os_v2_app_adn6bqphx5hybcc4c2adqhwreg4ekar7am5ej5fv4uwxgtau3pqgjfzynkjg226awxtc4737kumrypjz7ju777gjpna26pw6epg2syi";

// Registrar el dispositivo en OneSignal con el username como external_id
async function registerOneSignal(username, role) {
  try {
    const OS = await getOneSignal();
    if(!OS) return;
    await OS.login(username);
    // OneSignal v16: set external_id tags via User.addTag
    try {
      await OS.User.addTag("username", username);
      await OS.User.addTag("role", role||"vendedor");
    } catch(e) { console.warn("addTag:", e); }
  } catch(e) {
    console.warn("OneSignal register error:", e);
  }
}

// Enviar notificación push via OneSignal REST API a un usuario específico
// targetUsername: username del destinatario ("admin" significa enviar a todos los admins)
async function sendPushNotif({title, body, targetUsername}) {
  try {
    // Si el destinatario es "admin", enviar a todos con rol admin
    // Si es un username específico, enviar solo a ese usuario
    const filters = targetUsername === "admin"
      ? [{ field: "tag", key: "role", relation: "=", value: "admin" }]
      : [{ field: "tag", key: "username", relation: "=", value: targetUsername }];

    const payload = {
      app_id: OS_APP_ID,
      headings: { en: title, es: title },
      contents: { en: body, es: body },
      filters,
    };
    // Llamada a la API de OneSignal
    await fetch("https://onesignal.com/api/v1/notifications", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Key ${OS_API_KEY}`,
      },
      body: JSON.stringify(payload),
    });
  } catch(e) {
    console.warn("OneSignal push error:", e);
  }
}

// ─── CROSS-DEVICE NOTIFICATIONS via lm_notifs + OneSignal ────────────────────
async function sendCrossNotif(db, setNotifs, {title, body, tag, para, de}) {
  const notif = {
    id: genId(),
    fecha: new Date().toLocaleString("es-AR"),
    title, body,
    tag: tag||"lm",
    para, de,
    leida: false,
  };
  // Guardar en lm_notifs para Realtime (cuando la app está abierta)
  await db.addNotif(notif);
  setNotifs(n=>[notif,...n]);
  // Enviar push via OneSignal (cuando la app está cerrada)
  await sendPushNotif({ title, body, targetUsername: para });
}

// ─── PUSH NOTIFICATIONS ──────────────────────────────────────────────────────
const VAPID_PUBLIC_KEY = null; // Sin servidor VAPID — usamos Notification API directa

async function requestNotifPermission() {
  if (!("Notification" in window)) return "unsupported";
  if (Notification.permission === "granted") return "granted";
  if (Notification.permission === "denied") return "denied";
  const result = await Notification.requestPermission();
  return result;
}

function sendLocalNotif(title, body, tag = "lm") {
  if (!("Notification" in window) || Notification.permission !== "granted") return;
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.ready.then(reg => {
      reg.showNotification(title, {
        body,
        icon: "/icon-192.png",
        badge: "/icon-192.png",
        tag,
        renotify: true,
        vibrate: [200, 100, 200],
      });
    });
  } else {
    new Notification(title, { body, icon: "/icon-192.png" });
  }
}

// ─── PDF / PRINT ──────────────────────────────────────────────────────────────
// tipo: "reserva" | "confirmado" | "cotizacion"
// doc must have: docNum, compNum (orders), client, vendedor, date, items, total, notes, validity
function printDoc(doc, tipo) {
  const doRender = (logoSrc) => {
    // Determine display number and badge
    let docLabel, badgeColor, docNumDisplay;
    if(tipo === "cotizacion") {
      docLabel = "COTIZACIÓN";
      badgeColor = "#6c3483";
      docNumDisplay = doc.docNum || "Presu-------";
    } else if(tipo === "confirmado") {
      docLabel = "COMPROBANTE";
      badgeColor = "#1a5276";
      docNumDisplay = doc.compNum || doc.docNum || "Comp-------";
    } else {
      // reserva
      docLabel = "RESERVA";
      badgeColor = "#c0392b";
      docNumDisplay = doc.docNum || "Reserva-------";
    }

    const itemRows = doc.items.map(it => {
      const hasDisc = it.disc && parseFloat(it.disc.value) > 0;
      const lineTotal = applyItemDiscount(it.price, it.qty, it.disc);
      const originalTotal = it.price * it.qty;
      const discLabel = hasDisc ? fmtDisc(it.disc) : "";
      return `<tr>
        <td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;font-size:15px;">${it.name||""}</td>
        <td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;text-align:center;font-size:15px;">${it.qty}</td>
        <td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;text-align:right;font-size:15px;">
          ${fARS(it.price)}
          ${hasDisc?`<div style="font-size:11px;color:#1e8449;font-weight:700;">${discLabel}</div>`:""}
        </td>
        <td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;text-align:right;font-weight:700;font-size:15px;">
          ${hasDisc?`<div style="font-size:11px;color:#aaa;text-decoration:line-through;">${fARS(originalTotal)}</div>`:""}
          <span style="color:${hasDisc?"#1e8449":"inherit"}">${fARS(lineTotal)}</span>
        </td>
      </tr>`;
    }).join("");

    // Build discount summary rows for the total box
    const subtotalVal = doc.items.reduce((s,it)=>s+applyItemDiscount(it.price,it.qty,it.disc),0);
    const hasItemDiscs = doc.items.some(it=>it.disc&&parseFloat(it.disc.value)>0);
    const hasGlobalDisc = doc.globalDisc && parseFloat(doc.globalDisc.value)>0;
    const globalDiscAmt = hasGlobalDisc ? subtotalVal - applyGlobalDiscount(subtotalVal, doc.globalDisc) : 0;
    const discountRows = (hasItemDiscs || hasGlobalDisc) ? `
      ${hasItemDiscs ? `<div style="display:flex;justify-content:space-between;font-size:13px;color:#888;margin-bottom:4px;"><span>Subtotal s/dto</span><span>${fARS(doc.items.reduce((s,it)=>s+it.price*it.qty,0))}</span></div>` : ""}
      ${hasItemDiscs ? `<div style="display:flex;justify-content:space-between;font-size:13px;color:#1e8449;margin-bottom:4px;"><span>Dto. por item</span><span>-${fARS(doc.items.reduce((s,it)=>s+it.price*it.qty,0)-subtotalVal)}</span></div>` : ""}
      ${hasGlobalDisc ? `<div style="display:flex;justify-content:space-between;font-size:13px;color:#1e8449;margin-bottom:4px;"><span>Dto. global (${fmtDisc(doc.globalDisc)})</span><span>-${fARS(globalDiscAmt)}</span></div>` : ""}
      <div style="border-top:1px solid #e0e0e0;margin:6px 0;"></div>
    ` : "";

    const validityHtml = (tipo==="cotizacion" && doc.validity)
      ? `<div style="background:#fef9e7;border-left:3px solid #f1c40f;padding:8px 14px;border-radius:0 8px 8px 0;font-size:13px;color:#7d6608;margin-bottom:16px;">
          ⏳ <strong>Válida hasta:</strong> ${doc.validity}
        </div>` : "";

    const html = `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"/>
<title>${docLabel} ${docNumDisplay}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0;font-family:'Segoe UI',Arial,sans-serif;}
  html,body{background:#fdfbf7;color:#1a1a1a;width:210mm;}
  @media screen{body{max-width:210mm;margin:0 auto;box-shadow:0 0 20px #0002;}}
  @media print{.no-print{display:none!important;}@page{margin:3mm;size:A4 portrait;}body{width:100%;}}
  .print-btn{display:block;margin:12px auto;padding:9px 28px;background:linear-gradient(135deg,#7b1a1a,${badgeColor});color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;}
  .pdf-wrap{padding:0 0 16px;}

  /* HEADER — logo circular centrado + texto de marca + línea dorada de acento */
  .pdf-header{
    display:flex;
    align-items:center;
    gap:18px;
    padding:22px 28px;
    background:linear-gradient(135deg,#7b1a1a,#9c2a1f);
  }
  .pdf-logo-circle{
    width:64px;
    height:64px;
    border-radius:50%;
    overflow:hidden;
    border:2px solid #ffffff33;
    flex-shrink:0;
    background:#fff;
    display:flex;
    align-items:center;
    justify-content:center;
  }
  .pdf-logo-circle img{
    width:88%;
    height:88%;
    object-fit:contain;
  }
  .pdf-brand-name{
    color:#fff;
    font-family:Georgia,'Times New Roman',serif;
    font-size:26px;
    font-weight:700;
    letter-spacing:0.3px;
  }
  .pdf-brand-sub{
    color:#e8c9a8;
    font-size:10px;
    letter-spacing:2px;
    text-transform:uppercase;
    margin-top:3px;
  }
  .gold-line{height:3px;background:linear-gradient(90deg,#c9a96a,#e8d4a8,#c9a96a);}

  .content{padding:14px 18px 0;}

  /* DOC META row */
  .doc-meta{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;padding-bottom:12px;border-bottom:2px solid #f0ece2;}
  .doc-left{}
  .doc-type-label{font-size:10px;color:#999;text-transform:uppercase;letter-spacing:1.5px;margin-bottom:5px;font-weight:700;}
  .doc-badge{display:inline-block;padding:5px 16px;border-radius:6px;font-size:13px;font-weight:800;letter-spacing:1px;background:${badgeColor};color:#fff;}
  .doc-right{text-align:right;}
  .doc-num{font-size:23px;font-weight:700;color:#1a1a1a;letter-spacing:0.3px;font-family:Georgia,'Times New Roman',serif;}
  .doc-date{font-size:11px;color:#888;margin-top:3px;}

  /* INFO GRID */
  .info-grid{display:grid;grid-template-columns:1fr 1fr;gap:7px 16px;margin-bottom:14px;}
  .info-box{background:#fdfbf7;border-radius:6px;padding:8px 12px;border-left:3px solid #e5ddd0;}
  .info-box.hl{border-left-color:${badgeColor};}
  .info-label{font-size:8px;color:#999;text-transform:uppercase;letter-spacing:.7px;margin-bottom:3px;font-weight:700;}
  .info-value{font-size:13px;font-weight:700;color:#1a1a1a;}

  /* VALIDITY */
  .validity-bar{background:#fef9e7;border-left:3px solid #c9a96a;padding:6px 12px;border-radius:0 6px 6px 0;font-size:12px;color:#7d6608;margin-bottom:10px;}

  /* TABLE */
  table{width:100%;border-collapse:collapse;margin-bottom:14px;}
  thead tr{background:#fdfbf7;}
  th{padding:9px 10px;text-align:left;font-size:10px;color:#999;text-transform:uppercase;letter-spacing:.6px;font-weight:700;border-bottom:2px solid #f0ece2;}
  th.r{text-align:right;}th.c{text-align:center;}
  td{padding:9px 10px;border-bottom:1px solid #f5f1ea;font-size:15px;color:#1a1a1a;vertical-align:middle;}
  td.r{text-align:right;}td.c{text-align:center;}
  tbody tr:nth-child(even){background:#fdfbf7;}
  tbody tr:last-child td{border-bottom:none;}

  /* TOTAL */
  .total-wrap{display:flex;justify-content:flex-end;margin-bottom:14px;}
  .total-box{background:${tipo==="cotizacion"?"#f5eef8":"#fdf3f1"};border-radius:10px;padding:14px 22px;min-width:280px;border:1px solid ${badgeColor}22;}
  .disc-row{display:flex;justify-content:space-between;font-size:13px;margin-bottom:3px;}
  .disc-row.green{color:#1e8449;}
  .disc-row.grey{color:#888;}
  .disc-divider{border:none;border-top:1px solid #e5ddd0;margin:6px 0;}
  .total-final{display:flex;justify-content:space-between;align-items:center;margin-top:4px;}
  .total-label{font-size:14px;color:#555;font-weight:600;}
  .total-amount{font-size:28px;font-weight:700;color:${badgeColor};font-family:Georgia,'Times New Roman',serif;}

  /* NOTES */
  .notes{background:#fdfbf7;border-left:3px solid ${badgeColor};padding:9px 12px;border-radius:0 6px 6px 0;font-size:13px;color:#555;margin-bottom:14px;}

  /* FOOTER */
  .footer{border-top:1px solid #f0ece2;padding-top:9px;margin:0 0;font-size:10px;color:#bbb;display:flex;justify-content:space-between;align-items:center;}
  .footer-brand{color:#7b1a1a;font-weight:700;font-family:Georgia,serif;}
  /* TEST WATERMARK */
  .test-banner{background:#f1c40f;color:#1a1a1a;text-align:center;padding:8px;font-weight:900;font-size:13px;letter-spacing:2px;border-bottom:2px solid #d4ac0d;}
  .watermark{position:fixed;top:50%;left:50%;transform:translate(-50%,-50%) rotate(-35deg);font-size:80px;font-weight:900;color:rgba(241,196,15,0.15);pointer-events:none;white-space:nowrap;z-index:0;}
  .pdf-wrap{position:relative;z-index:1;}
</style></head><body>
<button class="no-print print-btn" onclick="window.print()">🖨️ Imprimir / Guardar PDF</button>
${doc.isTest ? `<div class="test-banner">⚠️ DOCUMENTO DE PRUEBA — NO VÁLIDO COMO COMPROBANTE</div>` : ""}
${doc.isTest ? `<div class="watermark">PRUEBA</div>` : ""}
<div class="pdf-wrap">
  <div class="pdf-header">
    <div class="pdf-logo-circle"><img src="${logoSrc}" alt="LM" onerror="this.style.display='none'"/></div>
    <div>
      <div class="pdf-brand-name">Libreria Madrid</div>
      <div class="pdf-brand-sub">Sistema de Gestión Comercial</div>
    </div>
  </div>
  <div class="gold-line"></div>
  <div class="content">
    <div class="doc-meta">
      <div class="doc-left">
        <div class="doc-type-label">Comprobante de</div>
        <div class="doc-badge">${docLabel}</div>
      </div>
      <div class="doc-right">
        <div class="doc-num">${docNumDisplay}</div>
        <div class="doc-date">${doc.date}</div>
      </div>
    </div>
    <div class="info-grid">
      <div class="info-box hl"><div class="info-label">Cliente</div><div class="info-value">${doc.client}</div></div>
      <div class="info-box"><div class="info-label">Vendedor</div><div class="info-value">${doc.vendedor||"—"}</div></div>
      <div class="info-box"><div class="info-label">Fecha</div><div class="info-value">${doc.date}</div></div>
      ${doc.validity?`<div class="info-box"><div class="info-label">Válida hasta</div><div class="info-value" style="color:#7d6608">${doc.validity}</div></div>`:""}
    </div>
    ${validityHtml}
    <table>
      <thead><tr>
        <th style="width:48%">Descripción</th>
        <th class="c" style="width:10%">Cant.</th>
        <th class="r" style="width:21%">P. Unit.</th>
        <th class="r" style="width:21%">Subtotal</th>
      </tr></thead>
      <tbody>${itemRows}</tbody>
    </table>
    <div class="total-wrap">
      <div class="total-box">
        ${discountRows}
        <div class="total-final">
          <span class="total-label">TOTAL</span>
          <span class="total-amount">${fARS(doc.total)}</span>
        </div>
      </div>
    </div>
    ${doc.notes?`<div class="notes"><strong>Notas:</strong> ${doc.notes}</div>`:""}
    <div class="footer">
      <span><span class="footer-brand">Libreria Madrid</span> — madrid.libreria · +54 9 11 2502-0640</span>
      <span>Emitido el ${new Date().toLocaleString("es-AR")}</span>
    </div>
  </div>
</div>
</body></html>`;

    const w = window.open("","_blank","width=820,height=750");
    if(w){ w.document.write(html); w.document.close(); }
  };

  // Banner incrustado como base64 — siempre disponible sin depender del servidor
  doRender("data:image/jpeg;base64,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");
}


// ─── NOTIFICATION TYPES CONFIG ───────────────────────────────────────────────
const NOTIF_TYPES = {
  NUEVO_PEDIDO:   {label:"Nuevo pedido",         icon:"🛒", color:"#1a5276", bg:"#d6eaf8"},
  CAMBIO_ESTADO:  {label:"Cambio de estado",      icon:"📋", color:"#6c3483", bg:"#e8daef"},
  ALTA_MERCADERIA:{label:"Alta de mercadería",    icon:"📦", color:"#1e8449", bg:"#d5f5e3"},
  PEDIDOS_PEND:   {label:"Pedidos pendientes",    icon:"⏰", color:"#e67e22", bg:"#fef9e7"},
};

// ─── NOTIF PANEL ─────────────────────────────────────────────────────────────
function NotifPanel({notifs,setNotifs,currentUser,users,onClose,onMarkAllRead,pushNotif,orders}) {
  const myNotifs = notifs.filter(n =>
    n.para === "todos" || n.para === currentUser.role || n.para === currentUser.id
  );
  const unread = myNotifs.filter(n => !n.leida.includes(currentUser.id));
  const markRead = async (id) => {
    const n = notifs.find(x=>x.id===id);
    if(!n||n.leida.includes(currentUser.id)) return;
    const updated = [...n.leida, currentUser.id];
    setNotifs(ns=>ns.map(x=>x.id===id?{...x,leida:updated}:x));
    await db.updateNotif(id, {leida:updated});
  };
  const delNotif = async (id) => { setNotifs(ns=>ns.filter(n=>n.id!==id)); await db.deleteNotif(id); };
  const pendingOrders = orders.filter(o=>o.stage!=="entregado");
  return (
    <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:998}} onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{position:"fixed",top:70,right:16,width:380,maxHeight:"80vh",background:"#fff",borderRadius:16,boxShadow:"0 12px 40px #0003",border:"1px solid #f0f0f0",display:"flex",flexDirection:"column",zIndex:999,overflow:"hidden"}}>
        <div style={{padding:"14px 18px",borderBottom:"1px solid #f5f5f5",display:"flex",alignItems:"center",justifyContent:"space-between",background:`linear-gradient(135deg,${REDD},${RED})`,borderRadius:"16px 16px 0 0"}}>
          <div style={{color:"#fff",fontWeight:800,fontSize:15}}>
            🔔 Notificaciones {unread.length>0&&<span style={{background:"#f1c40f",color:"#1a1a1a",borderRadius:10,fontSize:11,padding:"1px 6px",marginLeft:6,fontWeight:800}}>{unread.length}</span>}
          </div>
          <div style={{display:"flex",gap:6}}>
            {unread.length>0&&<button onClick={onMarkAllRead} style={{padding:"4px 10px",borderRadius:6,border:"none",background:"#ffffff33",color:"#fff",fontSize:11,fontWeight:600,cursor:"pointer"}}>Marcar todas leídas</button>}
            <button onClick={onClose} style={{background:"#ffffff22",border:"none",color:"#fff",borderRadius:6,padding:"4px 8px",cursor:"pointer",fontSize:14}}>x</button>
          </div>
        </div>
        {currentUser.role==="vendedor" && pendingOrders.filter(o=>o.vendedor===currentUser.name).length>0&&(
          <div style={{background:"#fef9e7",borderBottom:"1px solid #f1c40f22",padding:"10px 16px",display:"flex",gap:8,alignItems:"center"}}>
            <span style={{fontSize:18}}>⏰</span>
            <div>
              <div style={{fontWeight:700,fontSize:12,color:"#7d6608"}}>Tenés {pendingOrders.filter(o=>o.vendedor===currentUser.name).length} pedido(s) pendiente(s)</div>
              <div style={{fontSize:11,color:"#9a7d0a"}}>Revisá el estado de tus pedidos en Central</div>
            </div>
          </div>
        )}
        <div style={{overflowY:"auto",flex:1}}>
          {myNotifs.length===0
            ? <div style={{textAlign:"center",padding:40,color:"#aaa"}}><div style={{fontSize:36,marginBottom:8}}>🔕</div><div>No tenés notificaciones</div></div>
            : myNotifs.map(n=>{
                const cfg = NOTIF_TYPES[n.tipo]||{icon:"-",color:"#666",bg:"#f5f5f5"};
                const isRead = n.leida.includes(currentUser.id);
                return (
                  <div key={n.id} onClick={()=>markRead(n.id)} style={{padding:"12px 16px",borderBottom:"1px solid #f9f9f9",background:isRead?"#fff":"#fafbff",cursor:"pointer",display:"flex",gap:10,alignItems:"flex-start"}}>
                    <span style={{fontSize:20,flexShrink:0,marginTop:2}}>{cfg.icon}</span>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:6}}>
                        <div style={{fontWeight:isRead?500:700,fontSize:13,color:"#1a1a1a",lineHeight:1.3}}>{n.titulo}</div>
                        {!isRead&&<span style={{width:8,height:8,borderRadius:"50%",background:RED,flexShrink:0,marginTop:4}}/>}
                      </div>
                      <div style={{fontSize:12,color:"#666",marginTop:2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{n.cuerpo}</div>
                      <div style={{fontSize:10,color:"#aaa",marginTop:4}}>{n.fecha}</div>
                    </div>
                    <button onClick={e=>{e.stopPropagation();delNotif(n.id);}} style={{background:"none",border:"none",color:"#ddd",cursor:"pointer",fontSize:16,flexShrink:0,padding:0,lineHeight:1}}>x</button>
                  </div>
                );
              })
          }
        </div>
      </div>
    </div>
  );
}

function NotifConfig({users,setUsers,notifs,setNotifs}) {
  const DEFAULT_PREFS = {
    admin:    {NUEVO_PEDIDO:true,  CAMBIO_ESTADO:true,  ALTA_MERCADERIA:true,  PEDIDOS_PEND:false},
    vendedor: {NUEVO_PEDIDO:false, CAMBIO_ESTADO:true,  ALTA_MERCADERIA:false, PEDIDOS_PEND:true},
  };
  const getPrefs = (u) => u.notifPrefs || DEFAULT_PREFS[u.role] || DEFAULT_PREFS.vendedor;
  const togglePref = (uid, tipo) => {
    setUsers(us => us.map(u => {
      if(u.id!==uid) return u;
      const prefs = getPrefs(u);
      return {...u, notifPrefs:{...prefs, [tipo]:!prefs[tipo]}};
    }));
  };
  return (
    <div>
      <div style={{background:"#d6eaf8",border:"1px solid #aed6f1",borderRadius:12,padding:"12px 16px",marginBottom:16,fontSize:13,color:"#1a5276"}}>
        <strong>🔔 Centro de notificaciones</strong> &mdash; Configurá qué alertas recibe cada usuario dentro de la app.<br/>
        <span style={{fontSize:11,marginTop:4,display:"block",color:"#1a5276bb"}}>Para envío de emails automáticos conectá <strong>EmailJS</strong> cuando la app esté en producción.</span>
      </div>
      {users.map(u=>{
        const prefs = getPrefs(u);
        return (
          <div key={u.id} style={{background:"#fff",borderRadius:12,padding:18,marginBottom:12,boxShadow:"0 1px 4px #0001"}}>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
              <span style={{fontSize:24}}>{u.role==="admin"?"👑":"👤"}</span>
              <div>
                <div style={{fontWeight:800,fontSize:14}}>{u.name}</div>
                <div style={{fontSize:11,color:"#888"}}>@{u.username} . <span style={{color:u.role==="admin"?RED:"#1a5276",fontWeight:600}}>{u.role==="admin"?"Admin":"Vendedor"}</span>{u.email&&<span style={{marginLeft:6,color:"#aaa"}}>. 📧 {u.email}</span>}{!u.email&&<span style={{marginLeft:6,color:"#e67e22",fontSize:10}}>🎉 Sin email configurado</span>}</div>
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:8}}>
              {Object.entries(NOTIF_TYPES).map(([tipo,cfg])=>{
                const active = prefs[tipo]||false;
                return (
                  <div key={tipo} onClick={()=>togglePref(u.id,tipo)} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",borderRadius:10,border:`1.5px solid ${active?cfg.color:"#e5e5e5"}`,background:active?cfg.bg:"#fafafa",cursor:"pointer"}}>
                    <span style={{fontSize:18}}>{cfg.icon}</span>
                    <div style={{flex:1}}><div style={{fontWeight:600,fontSize:12,color:active?cfg.color:"#666"}}>{cfg.label}</div></div>
                    <div style={{width:18,height:18,borderRadius:4,border:`2px solid ${active?cfg.color:"#ccc"}`,background:active?cfg.color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                      {active&&<span style={{color:"#fff",fontSize:11,fontWeight:800}}>v</span>}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
      {notifs.length>0&&(<div style={{marginTop:8,textAlign:"right"}}><button onClick={async()=>{setNotifs([]);await db.clearNotifs();}} style={{padding:"7px 14px",borderRadius:8,border:"1.5px solid #fcc",background:"#fff",color:RED,cursor:"pointer",fontSize:12,fontWeight:600}}>🗑 Limpiar todas las notificaciones ({notifs.length})</button></div>)}
    </div>
  );
}

// ─── LOGIN ────────────────────────────────────────────────────────────────────
function Login({users, onLogin}) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [showPass, setShowPass] = useState(false);
  const [checking, setChecking] = useState(false);
  const handleLogin = async () => {
    if(!username.trim()||!password) return;
    setChecking(true); setError("");
    try {
      const fresh = await db.getUsers();
      const u = fresh.find(u => u.username === username.trim() && u.password === password);
      if (u) { onLogin(u); return; }
      const u2 = users.find(u => u.username === username.trim() && u.password === password);
      if(u2) { onLogin(u2); return; }
      setError("Usuario o contraseña incorrectos");
    } catch(e) {
      const u = users.find(u => u.username === username.trim() && u.password === password);
      if (u) { onLogin(u); return; }
      setError("Usuario o contraseña incorrectos");
    } finally { setChecking(false); }
  };
  return (
    <div style={{minHeight:"100vh",background:`linear-gradient(135deg,${REDD},${RED})`,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"#fff",borderRadius:20,padding:40,width:"100%",maxWidth:380,boxShadow:"0 20px 60px #0004"}}>
        <div style={{textAlign:"center",marginBottom:28}}>
          <img src={LOGO} alt="LM" style={{width:90,height:90,borderRadius:"50%",objectFit:"cover",marginBottom:12,boxShadow:"0 4px 16px #0002"}}/>
          <div style={{fontWeight:800,fontSize:22,fontFamily:"Georgia,serif",color:"#1a1a1a"}}>Libreria Madrid</div>
          <div style={{fontSize:12,color:"#aaa",letterSpacing:2,textTransform:"uppercase",marginTop:2}}>Sistema de Gestión</div>
        </div>
        <Field label="Usuario"><input value={username} onChange={e=>setUsername(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleLogin()} placeholder="Tu usuario" style={inputStyle}/></Field>
        <Field label="Contraseña">
          <div style={{position:"relative"}}>
            <input type={showPass?"text":"password"} value={password} onChange={e=>setPassword(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleLogin()} placeholder="Tu contraseña" style={{...inputStyle,paddingRight:40}}/>
            <button onClick={()=>setShowPass(s=>!s)} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:16,color:"#aaa"}}>{showPass?"🙈":"👁️"}</button>
          </div>
        </Field>
        {error && <div style={{background:"#fdecea",color:RED,borderRadius:8,padding:"8px 12px",fontSize:13,marginBottom:12,textAlign:"center"}}>{error}</div>}
        <button onClick={handleLogin} disabled={checking} style={{width:"100%",padding:12,borderRadius:10,border:"none",background:checking?"#aaa":`linear-gradient(135deg,${REDD},${RED})`,color:"#fff",fontWeight:800,fontSize:15,cursor:checking?"not-allowed":"pointer",marginTop:4}}>
          {checking?"Verificando...":"Ingresar"}
        </button>
      </div>
    </div>
  );
}


// ─── ROOT APP ─────────────────────────────────────────────────────────────────
export default function App() {
  // ── SANDBOX STOCK — copia para el vendedor Prueba, persiste en localStorage ──
  const [sandboxStock, setSandboxStock] = useState({});
  const isSandboxUser = (user) => user && isTestOrder(user.vendedor || user.name);

  const [currentUser, setCurrentUser] = useState(() => {
    // Recuperar sesión guardada al iniciar
    try {
      const saved = localStorage.getItem("lm_session");
      return saved ? JSON.parse(saved) : null;
    } catch { return null; }
  });
  const [users, setUsers]       = useState([]);
  const [vendors, setVendors]   = useState([]);
  const [products, setProducts] = useState([]);
  const [orders, setOrders]     = useState([]);
  const [quotes, setQuotes]     = useState([]);
  const [stockLog, setStockLog] = useState([]);
  const [activity, setActivity] = useState([]);
  const [priceLists, setPriceLists] = useState([{id:"default",name:"Normal",discount:0}]);
  const [purchaseOrders, setPurchaseOrders] = useState([]);
  const [notifs, setNotifs]     = useState([]);
  const [clients, setClients]   = useState([]);
  const [loading, setLoading]   = useState(true);
  const [error, setError]       = useState(null);

  useEffect(() => {
    async function loadAll() {
      try {
        const [u,v,p,o,q,sl,act,pl,po,n,cl] = await Promise.all([
          db.getUsers(), db.getVendors(), db.getProducts(),
          db.getOrders(), db.getQuotes(), db.getStockLog(), db.getActivity(), db.getPriceLists(), db.getPurchaseOrders(), db.getNotifs(), db.getClients(),
        ]);
        setUsers(u); setVendors(v); setProducts(p);
        setOrders(o); setQuotes(q); setStockLog(sl); setActivity(act); setPriceLists(pl); setPurchaseOrders(po); setNotifs(n); setClients(cl);
        // Inicializar sandbox — recuperar de localStorage si existe, sino copiar stock real
        try {
          const saved = localStorage.getItem("lm_sandbox_stock");
          if(saved) {
            setSandboxStock(JSON.parse(saved));
          } else {
            const sbInit = {};
            p.forEach(prod => { sbInit[prod.id] = prod.stock; });
            setSandboxStock(sbInit);
          }
        } catch {
          const sbInit = {};
          p.forEach(prod => { sbInit[prod.id] = prod.stock; });
          setSandboxStock(sbInit);
        }
        // Refrescar sesión guardada con los datos actuales del usuario
        try {
          const saved = localStorage.getItem("lm_session");
          if(saved) {
            const savedUser = JSON.parse(saved);
            const fresh = u.find(x => x.id === savedUser.id);
            if(fresh) {
              localStorage.setItem("lm_session", JSON.stringify(fresh));
              setCurrentUser(fresh);
            } else {
              // Usuario eliminado — cerrar sesión
              localStorage.removeItem("lm_session");
              setCurrentUser(null);
            }
          }
        } catch {}

        // Realtime se inicia en MainApp una vez que el usuario está logueado
      } catch(e) {
        setError("No se pudo conectar con la base de datos. Verificá tu conexión.");
      } finally { setLoading(false); }
    }
    loadAll();
  }, []);

  if(loading) return (
    <div style={{minHeight:"100vh",background:`linear-gradient(135deg,#922b21,#c0392b)`,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:16}}>
      <img src="/logo.png" alt="LM" style={{width:80,height:80,borderRadius:"50%",objectFit:"cover"}}/>
      <div style={{color:"#fff",fontWeight:700,fontSize:16}}>Cargando Libreria Madrid...</div>
      <div style={{color:"#ffcccc",fontSize:13}}>Conectando con la base de datos</div>
    </div>
  );
  if(error) return (
    <div style={{minHeight:"100vh",background:"#f5f5f5",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:"#fff",borderRadius:16,padding:32,maxWidth:400,textAlign:"center",boxShadow:"0 4px 20px #0002"}}>
        <div style={{fontSize:48,marginBottom:12}}>🎉️</div>
        <div style={{fontWeight:800,fontSize:16,marginBottom:8}}>Error de conexión</div>
        <div style={{color:"#666",fontSize:13,marginBottom:20}}>{error}</div>
        <button onClick={()=>window.location.reload()} style={{padding:"10px 24px",borderRadius:8,border:"none",background:"#c0392b",color:"#fff",fontWeight:700,cursor:"pointer"}}>Reintentar</button>
      </div>
    </div>
  );
  if(!currentUser) return <Login users={users} onLogin={u=>{
    localStorage.setItem("lm_session", JSON.stringify(u));
    setCurrentUser(u);
    // Registrar en OneSignal (espera a que esté listo)
    registerOneSignal(u.username, u.role);
  }}/>;
  return <MainApp
    currentUser={currentUser} onLogout={async()=>{
      try { const OS=await getOneSignal(); if(OS) await OS.logout(); } catch(e) {}
      localStorage.removeItem("lm_session");
      setCurrentUser(null);
    }}
    users={users} setUsers={setUsers}
    vendors={vendors} setVendors={setVendors}
    products={products} setProducts={setProducts}
    orders={orders} setOrders={setOrders}
    quotes={quotes} setQuotes={setQuotes}
    stockLog={stockLog} setStockLog={setStockLog}
    activity={activity} setActivity={setActivity}
    priceLists={priceLists} setPriceLists={setPriceLists}
    purchaseOrders={purchaseOrders} setPurchaseOrders={setPurchaseOrders}
    notifs={notifs} setNotifs={setNotifs}
    sandboxStock={sandboxStock} setSandboxStock={setSandboxStock}
    clients={clients} setClients={setClients}
    onLogout={async()=>{
      try {
        const OS = window.OneSignal;
        if(OS) await OS.logout();
      } catch(e) {}
      localStorage.removeItem("lm_session");
      setCurrentUser(null);
    }}
  />;
}

// ─── MAIN APP (authenticated) ─────────────────────────────────────────────────
function MainApp({currentUser,onLogout,users,setUsers,vendors,setVendors,products,setProducts,orders,setOrders,quotes,setQuotes,stockLog,setStockLog,activity,setActivity,priceLists,setPriceLists,purchaseOrders,setPurchaseOrders,notifs,setNotifs,sandboxStock,setSandboxStock,clients,setClients}) {
  // updateSandboxStock persists to localStorage on every change
  const updateSandboxStock = (updater) => {
    setSandboxStock(prev => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      try { localStorage.setItem("lm_sandbox_stock", JSON.stringify(next)); } catch {}
      return next;
    });
  };
  const isAdmin = currentUser.role === "admin";
  const [tab, setTab] = useState("central");
  const [showNotifs, setShowNotifs] = useState(false);

  const pushNotif = async (notif) => {
    const full = {id:genId(), fecha:new Date().toLocaleString("es-AR"), leida:[], ...notif};
    setNotifs(n=>[full,...n]); await db.addNotif(full);
  };
  const unreadCount = notifs.filter(n =>
    !n.leida.includes(currentUser.id) &&
    (n.para==="todos"||n.para===currentUser.role||n.para===currentUser.id)
  ).length;
  const markAllRead = async () => {
    const toUpdate = notifs.filter(n=>!n.leida.includes(currentUser.id)&&(n.para==="todos"||n.para===currentUser.role||n.para===currentUser.id));
    for(const n of toUpdate) await db.updateNotif(n.id,{leida:[...n.leida,currentUser.id]});
    setNotifs(ns=>ns.map(n=>n.leida.includes(currentUser.id)?n:{...n,leida:[...n.leida,currentUser.id]}));
  };
  // ── PRICE LIST ──
  // Admin can preview any list; other users use their assigned list
  const [previewListId, setPreviewListId] = useState(null);
  const activeListId = currentUser.role==="admin" && previewListId
    ? previewListId
    : (currentUser.priceList||"default");
  const activeList = priceLists.find(pl=>pl.id===activeListId) || {id:"default",name:"Normal",discount:0};
  // Apply list discount to a base price
  const getPrice = (basePrice) => {
    if(!activeList || activeList.discount===0) return basePrice;
    return Math.round(basePrice * (1 - activeList.discount/100) * 100) / 100;
  };
  // Products with prices adjusted for active list
  const isTestUser = isTestOrder(currentUser.vendedor || currentUser.name);
  const pricedProducts = useMemo(()=>products.map(p=>({
    ...p,
    salePrice: getPrice(p.salePrice),
    _basePrice: p.salePrice,
    // Vendedor Prueba ve el stock sandbox, no el real
    stock: isTestUser ? (sandboxStock[p.id] ?? p.stock) : p.stock,
  })),[products, activeList, isTestUser, sandboxStock]);

  const addLog = async (entry) => {
    const full = {id:genId(),fecha:new Date().toLocaleString("es-AR"),usuario:currentUser.name,rol:currentUser.role,...entry};
    setStockLog(l=>[full,...l]); await db.addStockLog(full);
  };
  const logActivity = async (accion, detalle, refId="", refTipo="") => {
    const entry = {id:genId(),fecha:new Date().toLocaleString("es-AR"),usuario:currentUser.name,rol:currentUser.role,accion,detalle,ref_id:refId,ref_tipo:refTipo};
    setActivity(a=>[entry,...a]);
    await db.addActivity(entry);
  };
  const addOrder = async (order) => {
    const test = isTestOrder(order.vendedor);
    const n = test ? 0 : await db.nextCounter("reserva");
    const orderWithNum = {...order, docNum: test ? "TEST-000000" : fmtDocNum("Reserva", n), isTest: test, isSandbox: test};
    const isSandbox = test;

    if(isSandbox) {
      // SANDBOX: solo en memoria + localStorage, no tocar Supabase ni notificaciones
      updateSandboxStock(prev => {
        const next = {...prev};
        orderWithNum.items.forEach(it => {
          next[it.pid] = Math.max(0, (next[it.pid] ?? 0) - it.qty);
        });
        return next;
      });
      setOrders(o=>[{...orderWithNum, isSandbox: true},...o]);
      if(!order.fromQuote) setTimeout(() => printDoc(orderWithNum, "reserva"), 400);
      return; // ← no Supabase, no notifs, no push
    }

    // REAL: descontar stock, guardar en Supabase, notificar
    const updatedProds = products.map(x=>{const it=orderWithNum.items.find(i=>i.pid===x.id);return it?{...x,stock:Math.max(0,x.stock-it.qty)}:x;});
    setProducts(updatedProds);
    for(const p of updatedProds.filter(p=>orderWithNum.items.find(i=>i.pid===p.id))) await db.upsertProduct(p);
    setOrders(o=>[{...orderWithNum, isSandbox: false},...o]);
    await db.upsertOrder({...orderWithNum, isSandbox: false});
    const notif={id:genId(),fecha:new Date().toLocaleString("es-AR"),leida:[],tipo:"NUEVO_PEDIDO",para:"admin",icono:"🛒",titulo:"Nuevo pedido registrado",cuerpo:`${orderWithNum.client} - ${fARS(orderWithNum.total)} - ${orderWithNum.docNum}`,ref:orderWithNum.id};
    await db.addNotif(notif); setNotifs(n=>[notif,...n]);
    await logActivity("Nuevo pedido", `${orderWithNum.docNum} - ${orderWithNum.client} - ${fARS(orderWithNum.total)} - Vendedor: ${orderWithNum.vendedor||"-"}`, orderWithNum.id, "pedido");
    if(!order.fromQuote) setTimeout(() => printDoc(orderWithNum, "reserva"), 400);
  };
  const [compPopup, setCompPopup] = useState(null);
  const [showChangePass, setShowChangePass] = useState(false);

  const setStage = async (id,stage) => {
    const ord = orders.find(o=>o.id===id);
    let updated = {...ord, stage};

    // When moving to "confirmado": assign Comp-XXXXXX (skip counter for test orders)
    if(stage === "confirmado" && !ord.compNum) {
      const test = isTestOrder(ord.vendedor);
      const n = test ? 0 : await db.nextCounter("comp");
      updated = {...updated, compNum: test ? "TEST-000000" : fmtDocNum("Comp", n), isTest: test};
    }

    setOrders(o=>o.map(x=>x.id===id?updated:x));
    await db.upsertOrder(updated);

    // Show popup with Comp number when confirmed
    if(stage === "confirmado" && updated.compNum) {
      setCompPopup(updated);
    }

    if(ord){
      const cfg=SCFG[stage]||{};
      const n1={id:genId(),fecha:new Date().toLocaleString("es-AR"),leida:[],tipo:"CAMBIO_ESTADO",para:"admin",icono:cfg.icon||"📋",titulo:`Pedido paso a ${cfg.label}`,cuerpo:`${ord.client} - ${fARS(ord.total)} - ${updated.compNum||""}`,ref:id};
      await db.addNotif(n1); setNotifs(n=>[n1,...n]);
      await logActivity(`Cambio estado: ${cfg.label}`, `${ord.docNum||""} ${updated.compNum||""} - ${ord.client} - ${fARS(ord.total)}`.trim(), id, "pedido");
      const vendUser=users.find(u=>u.name===ord.vendedor||u.username===ord.vendedor);
      if(vendUser&&vendUser.id!==currentUser.id){
        const n2={id:genId(),fecha:new Date().toLocaleString("es-AR"),leida:[],tipo:"CAMBIO_ESTADO",para:vendUser.id,icono:cfg.icon||"📋",titulo:`Tu pedido paso a ${cfg.label}`,cuerpo:`${ord.client} - ${fARS(ord.total)}`,ref:id};
        await db.addNotif(n2); setNotifs(n=>[n2,...n]);
      }
    }
  };
  const delOrder = async (id) => {
    const ord=orders.find(o=>o.id===id);
    // Detectar sandbox por isSandbox flag O por vendedor de prueba
    const isSandboxOrder = ord && (ord.isSandbox || isTestOrder(ord.vendedor));
    if(ord && ord.stage!=="entregado") {
      if(isSandboxOrder) {
        // SANDBOX: devolver stock al paralelo en memoria, nunca tocar Supabase
        updateSandboxStock(prev => {
          const next = {...prev};
          ord.items.forEach(it => { next[it.pid] = (next[it.pid] ?? 0) + it.qty; });
          return next;
        });
      } else {
        const updatedProds=products.map(x=>{const it=ord.items.find(i=>i.pid===x.id);return it?{...x,stock:x.stock+it.qty}:x;});
        setProducts(updatedProds);
        for(const p of updatedProds.filter(p=>ord.items.find(i=>i.pid===p.id))) await db.upsertProduct(p);
      }
    }
    const ordDel=orders.find(o=>o.id===id);
    setOrders(o=>o.filter(x=>x.id!==id)); await db.deleteOrder(id);
    if(ordDel) await logActivity("Pedido eliminado", `${ordDel.docNum||ordDel.compNum||""} - ${ordDel.client} - ${fARS(ordDel.total)}`, id, "pedido");
  };
  const saveNote = async (id, note) => {
    const updated = orders.map(o=>o.id===id ? {...o, internalNote:note} : o);
    setOrders(updated);
    const ord = updated.find(o=>o.id===id);
    await db.upsertOrder(ord);
  };

  // ── EDIT REQUEST FLOW ─────────────────────────────────────────────────────
  // Fase 1: vendedor solicita edición
  const requestEdit = async (id, reason) => {
    const updated = orders.map(o=>o.id===id ? {...o, editStatus:"solicitada", editReason:reason} : o);
    setOrders(updated);
    await db.upsertOrder(updated.find(o=>o.id===id));
    const ord0 = orders.find(o=>o.id===id);
    await sendCrossNotif(db, setNotifs, {title:"✏️ Solicitud de edición", body:`${ord0?.vendedor} quiere editar un pedido (${ord0?.client})`, tag:`edit-req-${id}`, para:"admin", de:ord0?.vendedor||""});
    sendLocalNotif("✏️ Solicitud enviada", `Esperá la aprobación del admin`, `edit-req-${id}`);
  };

  // Fase 2a: admin aprueba la solicitud → vendedor puede editar
  const approveEditRequest = async (id) => {
    const updated = orders.map(o=>o.id===id ? {...o, editStatus:"aprobada", editRejectReason:""} : o);
    setOrders(updated);
    await db.upsertOrder(updated.find(o=>o.id===id));
    const ordApr = orders.find(o=>o.id===id);
    await sendCrossNotif(db, setNotifs, {title:"✅ Edición aprobada", body:`Tu solicitud para editar el pedido de ${ordApr?.client} fue aprobada`, tag:`edit-apr-${id}`, para:ordApr?.vendedor||"", de:"admin"});
  };

  // Fase 2b: admin rechaza la solicitud
  const rejectEditRequest = async (id, reason) => {
    const updated = orders.map(o=>o.id===id ? {...o, editStatus:"rechazada", editRejectReason:reason} : o);
    setOrders(updated);
    await db.upsertOrder(updated.find(o=>o.id===id));
    const ordRej = orders.find(o=>o.id===id);
    await sendCrossNotif(db, setNotifs, {title:"❌ Edición rechazada", body:`Tu solicitud para editar el pedido de ${ordRej?.client} fue rechazada. Motivo: "${reason}"`, tag:`edit-rej-${id}`, para:ordRej?.vendedor||"", de:"admin"});
  };

  // Fase 3: vendedor guarda los cambios editados
  const submitEdit = async (id, newItems, newTotal) => {
    const updated = orders.map(o=>o.id===id ? {...o, editStatus:"en revisión", editItems:newItems, editTotal:newTotal} : o);
    setOrders(updated);
    await db.upsertOrder(updated.find(o=>o.id===id));
    const ordSub = orders.find(o=>o.id===id);
    await sendCrossNotif(db, setNotifs, {title:"👀 Cambios para revisar", body:`${ordSub?.vendedor} editó el pedido de ${ordSub?.client} — revisá los cambios`, tag:`edit-sub-${id}`, para:"admin", de:ordSub?.vendedor||""});
    sendLocalNotif("📤 Cambios enviados", `El admin revisará tu edición`, `edit-sub-${id}`);
  };

  // Fase 4a: admin aprueba los cambios finales
  const approveEdit = async (id) => {
    const ord = orders.find(o=>o.id===id);
    if(!ord) return;
    // Aplicar cambios: restaurar stock viejo, descontar nuevo
    if(!ord.isSandbox) {
      // Devolver stock de items originales
      let prods = [...products];
      ord.items.forEach(it => {
        const idx = prods.findIndex(p=>p.id===it.pid);
        if(idx>=0) prods[idx] = {...prods[idx], stock: prods[idx].stock + it.qty};
      });
      // Descontar stock de items nuevos
      ord.editItems.forEach(it => {
        const idx = prods.findIndex(p=>p.id===it.pid);
        if(idx>=0) prods[idx] = {...prods[idx], stock: Math.max(0, prods[idx].stock - it.qty)};
      });
      setProducts(prods);
      for(const p of prods.filter(p=>ord.items.find(i=>i.pid===p.id)||ord.editItems.find(i=>i.pid===p.id))) await db.upsertProduct(p);
    }
    const newTotal = ord.editItems.reduce((s,it)=>s+it.price*it.qty,0);
    const updated = orders.map(o=>o.id===id ? {...o, items:ord.editItems, total:newTotal, editStatus:"", editItems:null, editReason:"", editRejectReason:""} : o);
    setOrders(updated);
    await db.upsertOrder(updated.find(o=>o.id===id));
    const ordOk = updated.find(o=>o.id===id);
    await sendCrossNotif(db, setNotifs, {title:"✅ Cambios aprobados", body:`Tu edición del pedido de ${ordOk?.client} fue aprobada`, tag:`edit-ok-${id}`, para:ordOk?.vendedor||"", de:"admin"});
    await logActivity("Edición aprobada", `Pedido ${ord.docNum||ord.compNum||""} editado`, id, "pedido");
  };

  // Fase 4b: admin rechaza los cambios finales
  const rejectEdit = async (id, reason) => {
    const updated = orders.map(o=>o.id===id ? {...o, editStatus:"cambios rechazados", editRejectReason:reason, editItems:null} : o);
    setOrders(updated);
    await db.upsertOrder(updated.find(o=>o.id===id));
    const ordNo = orders.find(o=>o.id===id);
    await sendCrossNotif(db, setNotifs, {title:"❌ Cambios rechazados", body:`El admin rechazó tu edición del pedido de ${ordNo?.client}. Motivo: "${reason}"`, tag:`edit-no-${id}`, para:ordNo?.vendedor||"", de:"admin"});
  };
  const addQuote = async (quote) => {
    const test = isTestOrder(quote.vendedor);
    const n = test ? 0 : await db.nextCounter("presu");
    const quoteWithNum = {...quote, docNum: test ? "TEST-000000" : fmtDocNum("Presu", n), isTest: test};
    setQuotes(q=>[quoteWithNum,...q]);
    if(!test) {
      await db.upsertQuote(quoteWithNum);
      await logActivity("Nueva cotización", `${quoteWithNum.docNum} - ${quoteWithNum.client} - ${fARS(quoteWithNum.total)} - Vendedor: ${quoteWithNum.vendedor||"-"}`, quoteWithNum.id, "cotizacion");
    }
  };
  const delQuote = async (id) => {
    const q = quotes.find(x=>x.id===id);
    setQuotes(qs=>qs.filter(x=>x.id!==id));
    if(q && !isTestOrder(q.vendedor)) await db.deleteQuote(id);
  };
  const extendQuote = async (id, reason) => {
    const updated = quotes.map(q => q.id===id ? {...q, extendida:true, extendReason:reason, extendDate:today()} : q);
    setQuotes(updated);
    const quo = updated.find(x=>x.id===id);
    if(quo && !isTestOrder(quo.vendedor)) {
      await db.upsertQuote({...quo, extend_reason:reason, extend_date:today()});
    }
  };

  // Convierte una cotización en reserva — descuenta stock y arranca el circuito de ventas
  const convertQuoteToOrder = async (quote) => {
    const isSandbox = isTestOrder(quote.vendedor);
    const order = {
      id: genId(),
      client: quote.client,
      clientId: quote.clientId || "",
      vendedor: quote.vendedor,
      notes: quote.notes,
      items: quote.items,
      total: quote.total,
      subtotal: quote.subtotal,
      globalDisc: quote.globalDisc,
      stage: "reserva",
      date: today(),
      fromQuote: true,
    };
    await addOrder(order);
    // Marcar la cotización como convertida
    const updated = {...quote, convertida: true, ordenId: order.id};
    setQuotes(q=>q.map(x=>x.id===quote.id?updated:x));
    // Solo tocar Supabase si NO es sandbox
    if(!isSandbox) {
      await db.upsertQuote({...updated});
      await logActivity("Cotización convertida a reserva", `${quote.docNum||""} - ${quote.client} - ${fARS(quote.total)}`, quote.id, "cotizacion");
    }
  };
  const updProd = async (upd) => {
    const prev = products.find(p=>p.id===upd.id);
    setProducts(p=>p.map(x=>x.id===upd.id?upd:x)); await db.upsertProduct(upd);
    if(prev) await logActivity("Precio/stock editado", `${upd.name} - Venta: ${fARS(upd.salePrice)} (antes ${fARS(prev.salePrice)}) - Stock: ${upd.stock}`, upd.id, "producto");
  };
  const addStock = async (pid,qty,newCost) => {
    const prod=products.find(p=>p.id===pid);
    const updatedProds=products.map(x=>{if(x.id!==pid)return x;const u={...x,stock:x.stock+qty};if(newCost){u.costPrice=newCost;u.salePrice=Math.round(newCost*1.5*100)/100;}return u;});
    setProducts(updatedProds);
    const updProd=updatedProds.find(p=>p.id===pid);
    if(updProd)await db.upsertProduct(updProd);
    if(prod){
      const notif={id:genId(),fecha:new Date().toLocaleString("es-AR"),leida:[],tipo:"ALTA_MERCADERIA",para:"admin",icono:"📦",titulo:"Alta de mercaderia",cuerpo:`${prod.name} - +${qty} unidades${newCost?` - Nuevo costo: ${fARS(newCost)}`:""}`,ref:pid};
      await db.addNotif(notif); setNotifs(n=>[notif,...n]);
    }
  };

  const pending = orders.filter(o=>o.stage!=="entregado"&&!o.isSandbox).length;
  const TABS = [
    {k:"central",   label:"Central",           icon:"📋", roles:["admin","vendedor"]},
    {k:"nuevo",     label:"Nuevo Pedido",       icon:"🛒", roles:["admin","vendedor"]},
    {k:"clientes",  label:"Clientes",           icon:"👥", roles:["admin","vendedor"]},
    {k:"cotizacion",label:"Cotizaciones",       icon:"📄", roles:["admin","vendedor"]},
    {k:"precios",   label:"Precios",            icon:"💲", roles:["admin","vendedor"]},
    {k:"stock",     label:"Stock",              icon:"📦", roles:["admin","vendedor"]},
    {k:"compras",   label:"Alta de Mercancía",icon:"🏪", roles:["admin","vendedor"]},
    {k:"solicitud",  label:"Solicitud Compra", icon:"📋", roles:["admin","vendedor"]},
    {k:"admin",     label:"Administracion",     icon:"★",   roles:["admin"]},
  ].filter(t=>t.roles.includes(currentUser.role));

  // ── CLIENT FUNCTIONS ──────────────────────────────────────────────────────
  const saveClient = async (client) => {
    const isNew = !clients.find(c=>c.id===client.id);
    setClients(prev => isNew ? [client,...prev] : prev.map(c=>c.id===client.id?client:c));
    await db.saveClient(client);
    if(isNew) await logActivity("Cliente creado", `${client.name}`, client.id, "cliente");
  };
  const deleteClient = async (id) => {
    setClients(prev=>prev.filter(c=>c.id!==id));
    await db.deleteClient(id);
    await logActivity("Cliente eliminado", `ID: ${id}`, id, "cliente");
  };
  const requestDeleteClient = async (id, reason) => {
    const client = clients.find(c=>c.id===id);
    if(!client) return;
    const updated = {...client, deleteRequested:true, deleteReason:reason};
    setClients(prev=>prev.map(c=>c.id===id?updated:c));
    await db.saveClient(updated);
    await sendCrossNotif(db, setNotifs, {
      title:"🗑 Solicitud de baja de cliente",
      body:`${currentUser.name} solicita eliminar a "${client.name}". Motivo: "${reason}"`,
      tag:`del-client-${id}`, para:"admin", de:currentUser.name
    });
  };

  const isMobile = useIsMobile();
  const [mobileMenu, setMobileMenu] = useState(false);
  const [notifPermission, setNotifPermission] = useState(
    "Notification" in window ? Notification.permission : "unsupported"
  );

  // Escuchar el evento de logout desde el botón atrás mobile
  useEffect(() => {
    const handler = () => onLogout();
    window.addEventListener('lm-logout', handler);
    return () => window.removeEventListener('lm-logout', handler);
  }, [onLogout]);

  // ── Supabase Realtime — se inicia aquí donde currentUser ya existe ──
  useEffect(() => {
    const channel = supabase.channel("lm-realtime-" + currentUser.id)
      // PEDIDOS
      .on("postgres_changes", {event:"INSERT", schema:"public", table:"lm_orders"}, (payload) => {
        const o = mapOrder(payload.new);
        setOrders(prev => prev.find(x=>x.id===o.id) ? prev : [o,...prev]);
        if(o.vendedor !== currentUser.vendedor && o.vendedor !== currentUser.name) {
          sendLocalNotif("📋 Nuevo pedido", `${o.vendedor} · ${o.client}`, `order-${o.id}`);
        }
      })
      .on("postgres_changes", {event:"UPDATE", schema:"public", table:"lm_orders"}, (payload) => {
        const o = mapOrder(payload.new);
        setOrders(prev => prev.map(x => x.id===o.id ? o : x));
        const stageLabels = {confirmado:"✅ Pedido confirmado","en armado":"📦 Pedido en armado",entregado:"🎉 Pedido entregado"};
        if(stageLabels[o.stage] && o.vendedor !== currentUser.vendedor && o.vendedor !== currentUser.name) {
          sendLocalNotif(stageLabels[o.stage], `${o.client} — ${o.docNum||o.compNum||""}`, `stage-${o.id}`);
        }
      })
      .on("postgres_changes", {event:"DELETE", schema:"public", table:"lm_orders"}, (payload) => {
        setOrders(prev => prev.filter(x => x.id !== payload.old.id));
      })
      // COTIZACIONES
      .on("postgres_changes", {event:"INSERT", schema:"public", table:"lm_quotes"}, (payload) => {
        const q = mapQuote(payload.new);
        setQuotes(prev => prev.find(x=>x.id===q.id) ? prev : [q,...prev]);
        if(q.vendedor !== currentUser.vendedor && q.vendedor !== currentUser.name) {
          sendLocalNotif("📄 Nueva cotización", `${q.vendedor} · ${q.client}`, `quote-${q.id}`);
        }
      })
      .on("postgres_changes", {event:"UPDATE", schema:"public", table:"lm_quotes"}, (payload) => {
        const q = mapQuote(payload.new);
        setQuotes(prev => prev.map(x => x.id===q.id ? q : x));
      })
      .on("postgres_changes", {event:"DELETE", schema:"public", table:"lm_quotes"}, (payload) => {
        setQuotes(prev => prev.filter(x => x.id !== payload.old.id));
      })
      // NOTIFICACIONES CRUZADAS — escuchar lm_notifs dirigidas a este usuario
      .on("postgres_changes", {event:"INSERT", schema:"public", table:"lm_notifs"}, (payload) => {
        const n = payload.new;
        // Mostrar solo si está dirigida a este usuario o al rol admin
        const esAdmin = currentUser.role === "admin";
        const dirigidaAMi = n.para === currentUser.username || n.para === currentUser.name || n.para === currentUser.vendedor;
        const dirigidaAAdmin = n.para === "admin" && esAdmin;
        if((dirigidaAMi || dirigidaAAdmin) && n.de !== currentUser.name && n.de !== currentUser.vendedor) {
          sendLocalNotif(n.title, n.body, n.tag);
          setNotifs(prev => prev.find(x=>x.id===n.id) ? prev : [n,...prev]);
        }
      })
      .on("postgres_changes", {event:"INSERT", schema:"public", table:"lm_purchase_orders"}, (payload) => {
        const po = payload.new;
        setPurchaseOrders(prev => prev.find(x=>x.id===po.id) ? prev : [{...po, fechaCierre:po.fecha_cierre},...prev]);
        if(po.vendedor !== currentUser.vendedor && po.vendedor !== currentUser.name) {
          sendLocalNotif("📋 Nueva solicitud de compra", `${po.vendedor} creó una solicitud`, `po-${po.id}`);
        }
      })
      .on("postgres_changes", {event:"UPDATE", schema:"public", table:"lm_purchase_orders"}, (payload) => {
        const po = payload.new;
        setPurchaseOrders(prev => prev.map(x => x.id===po.id ? {...po, fechaCierre:po.fecha_cierre} : x));
      })
      .on("postgres_changes", {event:"DELETE", schema:"public", table:"lm_purchase_orders"}, (payload) => {
        setPurchaseOrders(prev => prev.filter(x => x.id !== payload.old.id));
      })
      .subscribe((status) => {
        console.log("Realtime status:", status);
      });

    return () => { supabase.removeChannel(channel); };
  }, [currentUser.id]);

  // Chequear estado de notificaciones al iniciar (sin pedir permiso automáticamente)
  useEffect(() => {
    const check = async () => {
      try {
        // Intentar con OneSignal primero (espera hasta 5s)
        const OS = await Promise.race([
          getOneSignal(),
          new Promise(r => setTimeout(()=>r(null), 5000))
        ]);
        if(OS) {
          const granted = OS.Notifications.permission;
          setNotifPermission(granted ? "granted" : "default");
        } else if("Notification" in window) {
          setNotifPermission(Notification.permission);
        }
      } catch(e) {}
    };
    check();
  }, []);

  // Confirmación al cerrar/retroceder (funciona en desktop y mobile)
  useEffect(() => {
    // Desktop: diálogo nativo al cerrar pestaña
    const beforeUnload = (e) => {
      e.preventDefault();
      e.returnValue = '';
      return '';
    };
    window.addEventListener('beforeunload', beforeUnload);

    // Mobile: intercepta el botón "atrás" del navegador
    // Empujamos un estado extra al historial; cuando el usuario retrocede
    // volvemos a este estado y mostramos el confirm
    window.history.pushState({ appActive: true }, '');
    const popState = (e) => {
      const salir = window.confirm('¿Seguro que querés salir de la app?');
      if (salir) {
        window.removeEventListener('popstate', popState);
        // Disparar evento custom para que onLogout lo reciba
        window.dispatchEvent(new CustomEvent('lm-logout'));
      } else {
        // Volvemos a empujar el estado para que el botón atrás vuelva a funcionar
        window.history.pushState({ appActive: true }, '');
      }
    };
    window.addEventListener('popstate', popState);

    return () => {
      window.removeEventListener('beforeunload', beforeUnload);
      window.removeEventListener('popstate', popState);
    };
  }, []);

  return (
    <div style={{minHeight:"100vh",background:IVORY,fontFamily:"'Segoe UI',system-ui,sans-serif"}}>

      {/* ── HEADER ── */}
      <div style={{background:`linear-gradient(135deg,${VINO},#9c2a1f)`,boxShadow:"0 4px 16px #0004",position:"sticky",top:0,zIndex:100,borderBottom:`3px solid ${GOLD}`}}>
        {isMobile ? (
          <div>
            <div style={{display:"flex",alignItems:"center",padding:"12px 14px 8px",gap:10}}>
              <img src={LOGO} alt="LM" style={{width:38,height:38,borderRadius:"50%",objectFit:"cover",border:"2px solid #ffffff44",flexShrink:0}}/>
              <div style={{flex:1,minWidth:0}}>
                <div style={{color:"#fff",fontWeight:800,fontSize:15,fontFamily:"Georgia,serif"}}>Libreria Madrid</div>
                <div style={{color:"#ffcccc",fontSize:11}}>👤 {currentUser.name}{activeList.discount>0&&<span style={{marginLeft:6,background:"#f1c40f",color:"#1a1a1a",borderRadius:4,padding:"1px 5px",fontSize:10,fontWeight:800}}>{activeList.name}</span>}</div>
              </div>
              <div style={{display:"flex",gap:6,flexShrink:0}}>
                <button onClick={()=>setShowNotifs(s=>!s)} style={{width:34,height:34,borderRadius:9,background:"#ffffff1a",border:"1px solid #ffffff2a",color:"#fff",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",position:"relative"}}>
                  🔔{unreadCount>0&&<span style={{position:"absolute",top:-3,right:-3,background:"#f1c40f",color:"#1a1a1a",borderRadius:"50%",width:15,height:15,fontSize:8,fontWeight:900,display:"flex",alignItems:"center",justifyContent:"center"}}>{unreadCount}</span>}
                  {notifPermission==="default"&&<span style={{position:"absolute",top:-3,right:-3,background:"#e67e22",borderRadius:"50%",width:10,height:10}}/>}
                </button>
                <button onClick={()=>setMobileMenu(o=>!o)} style={{width:34,height:34,borderRadius:9,background:"#ffffff1a",border:"1px solid #ffffff2a",color:"#fff",fontSize:18,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer"}}>☰</button>
              </div>
            </div>
            <div style={{display:"flex",overflowX:"auto",gap:1,padding:"0 10px",scrollbarWidth:"none",alignItems:"center"}}>
              {TABS.map(t=>(
                <button key={t.k} onClick={()=>setTab(t.k)}
                  style={{padding:"7px 13px",border:"none",cursor:"pointer",fontSize:11,color:tab===t.k?"#fff":"#ffbbbb",fontWeight:tab===t.k?700:600,borderRadius:"8px 8px 0 0",background:tab===t.k?"#ffffff18":"transparent",borderBottom:tab===t.k?"3px solid #fff":"3px solid transparent",whiteSpace:"nowrap",flexShrink:0}}>
                  {t.icon} {t.label}
                </button>
              ))}
              <button onClick={()=>{if(window.confirm("¿Seguro que querés salir?")) onLogout();}}
                style={{padding:"5px 10px",border:"1px solid #ffffff33",cursor:"pointer",fontSize:11,color:"#ffbbbb",fontWeight:600,borderRadius:7,background:"transparent",whiteSpace:"nowrap",flexShrink:0,marginLeft:6}}>
                🚪 Salir
              </button>
            </div>
          </div>
        ) : (
          <div style={{maxWidth:1200,margin:"0 auto",padding:"0 16px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap"}}>
            <div style={{display:"flex",alignItems:"center",gap:12,padding:"12px 0"}}>
              <img src={LOGO} alt="LM Logo" style={{width:54,height:54,borderRadius:"50%",objectFit:"cover",boxShadow:"0 2px 8px #0003"}}/>
              <div>
                <div style={{color:"#fff",fontWeight:800,fontSize:20,fontFamily:"Georgia,serif"}}>Libreria Madrid</div>
                <div style={{color:"#ffcccc",fontSize:10,letterSpacing:2,textTransform:"uppercase"}}>Sistema de Gestión</div>
              </div>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:4,flexWrap:"wrap"}}>
              <nav style={{display:"flex",gap:2,flexWrap:"wrap"}}>
                {TABS.map(t=>(
                  <button key={t.k} onClick={()=>setTab(t.k)} style={{padding:"11px 14px",border:"none",cursor:"pointer",fontSize:13,background:tab===t.k?"#fff":"transparent",color:tab===t.k?RED:"#ffcccc",fontWeight:tab===t.k?700:500,borderRadius:"8px 8px 0 0",position:"relative"}}>
                    {t.icon} {t.label}
                    {t.k==="central"&&pending>0&&<span style={{position:"absolute",top:5,right:3,background:"#fff",color:RED,borderRadius:10,fontSize:10,padding:"1px 5px",fontWeight:800,border:`1.5px solid ${RED}`}}>{pending}</span>}
                  </button>
                ))}
              </nav>
              <div style={{display:"flex",alignItems:"center",gap:8,padding:"0 8px",borderLeft:"1px solid #ffffff33",marginLeft:4}}>
                {isAdmin && priceLists.length>1 && (
                  <select value={previewListId||"default"} onChange={e=>setPreviewListId(e.target.value==="default"?null:e.target.value)}
                    style={{background:"#ffffff22",border:"1px solid #ffffff44",color:"#fff",borderRadius:6,padding:"5px 8px",cursor:"pointer",fontSize:11,fontWeight:600}}>
                    {priceLists.map(pl=><option key={pl.id} value={pl.id} style={{color:"#1a1a1a"}}>{pl.name}{pl.discount>0?` (-${pl.discount}%)`:""}</option>)}
                  </select>
                )}
                <button onClick={()=>setShowChangePass(true)}
                  style={{background:"#ffffff15",border:"1px solid #ffffff33",color:"#ffeeee",borderRadius:6,padding:"5px 10px",cursor:"pointer",fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:5}}>
                  👤 {currentUser.name}
                  {activeList.discount>0&&<span style={{background:"#f1c40f",color:"#1a1a1a",borderRadius:4,padding:"1px 5px",fontSize:10,fontWeight:800,marginLeft:3}}>{activeList.name}</span>}
                  <span style={{fontSize:10,opacity:.7}}>🔑</span>
                </button>
                <div style={{position:"relative"}}>
                  <button onClick={()=>setShowNotifs(s=>!s)} style={{background:"#ffffff22",border:"none",color:"#fff",borderRadius:6,padding:"5px 8px",cursor:"pointer",fontSize:16,lineHeight:1,position:"relative"}}>
                    🔔
                    {unreadCount>0&&<span style={{position:"absolute",top:-4,right:-4,background:"#f1c40f",color:"#1a1a1a",borderRadius:10,fontSize:9,padding:"1px 4px",fontWeight:800,minWidth:14,textAlign:"center"}}>{unreadCount}</span>}
                  </button>
                </div>
                <button onClick={onLogout} style={{background:"#ffffff22",border:"none",color:"#fff",borderRadius:6,padding:"5px 8px",cursor:"pointer",fontSize:11,fontWeight:600}}>Salir</button>
              </div>
              {showNotifs&&<NotifPanel notifs={notifs} setNotifs={setNotifs} currentUser={currentUser} users={users} onClose={()=>setShowNotifs(false)} onMarkAllRead={markAllRead} pushNotif={pushNotif} orders={orders}/>}
            </div>
          </div>
        )}
      </div>

      {/* MOBILE DRAWER MENU */}
      {isMobile && mobileMenu && (
        <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,width:"100vw",height:"100vh",background:"#0007",zIndex:9999}} onClick={()=>setMobileMenu(false)}>
          <div style={{width:"80%",maxWidth:300,height:"100%",background:"#fff",boxShadow:"4px 0 24px #0005",paddingTop:48,display:"flex",flexDirection:"column",overflowY:"auto"}} onClick={e=>e.stopPropagation()}>
            <div style={{padding:"0 20px 16px",borderBottom:"1px solid #f0f0f0",marginBottom:8}}>
              <div style={{fontWeight:800,fontSize:17,fontFamily:"Georgia,serif"}}>Libreria Madrid</div>
              <div style={{fontSize:11,color:"#aaa",marginTop:2}}>Sistema de Gestión</div>
            </div>
            {TABS.map(t=>(
              <div key={t.k} onClick={()=>{setTab(t.k);setMobileMenu(false);}}
                style={{display:"flex",alignItems:"center",gap:14,padding:"13px 20px",fontSize:14,fontWeight:600,color:tab===t.k?RED:"#333",background:tab===t.k?"#fdecea":"transparent",cursor:"pointer"}}>
                <span style={{fontSize:20,width:28,textAlign:"center"}}>{t.icon}</span>{t.label}
              </div>
            ))}
            <div style={{borderTop:"2px solid #e5e5e5",marginTop:216,paddingTop:8}}>
              {/* Botón de notificaciones OneSignal */}
              <div onClick={async()=>{
                try {
                  const OS = await getOneSignal();
                  if(!OS) {
                    alert("El sistema de notificaciones no está disponible. Verificá tu conexión y recargá la app.");
                    setMobileMenu(false);
                    return;
                  }
                  const granted = await OS.Notifications.requestPermission();
                  setNotifPermission(granted ? "granted" : "denied");
                  if(granted) {
                    await OS.login(currentUser.username);
                    try {
                      await OS.User.addTag("username", currentUser.username);
                      await OS.User.addTag("role", currentUser.role);
                    } catch(e) { console.warn("addTag:", e); }
                  }
                } catch(e) {
                  console.warn("OneSignal:", e);
                  alert("Hubo un problema al activar las notificaciones. Intentá de nuevo.");
                }
                setMobileMenu(false);
              }} style={{display:"flex",alignItems:"center",gap:14,padding:"13px 20px",fontSize:14,fontWeight:600,
                color:notifPermission==="granted"?"#1e8449":notifPermission==="denied"?"#aaa":"#e67e22",
                cursor:"pointer",
                background:notifPermission==="granted"?"#eafaf1":"transparent"}}>
                <span style={{fontSize:20,width:28,textAlign:"center"}}>
                  {notifPermission==="granted"?"🔔":"🔕"}
                </span>
                {notifPermission==="granted"
                  ? "Notificaciones activas ✓"
                  : notifPermission==="denied"
                    ? "Notificaciones bloqueadas"
                    : "Activar notificaciones push"
                }
              </div>
              <div onClick={()=>{if(window.confirm("¿Seguro que querés salir de la app?")) onLogout();}} style={{display:"flex",alignItems:"center",gap:14,padding:"13px 20px",fontSize:14,fontWeight:600,color:RED,cursor:"pointer",background:"#fdecea"}}>
                <span style={{fontSize:20,width:28,textAlign:"center"}}>🚪</span>Salir
              </div>
              <div onClick={()=>{setShowChangePass(true);setMobileMenu(false);}} style={{display:"flex",alignItems:"center",gap:14,padding:"13px 20px",fontSize:14,fontWeight:600,color:"#333",cursor:"pointer"}}>
                <span style={{fontSize:20,width:28,textAlign:"center"}}>🔑</span>Cambiar contraseña
              </div>
              {isAdmin && priceLists.length>1 && (
                <div style={{padding:"8px 20px"}}>
                  <div style={{fontSize:10,fontWeight:700,color:"#aaa",marginBottom:4}}>LISTA DE PRECIOS</div>
                  <select value={previewListId||"default"} onChange={e=>setPreviewListId(e.target.value==="default"?null:e.target.value)}
                    style={{width:"100%",padding:"8px 10px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,background:"#fff",cursor:"pointer"}}>
                    {priceLists.map(pl=><option key={pl.id} value={pl.id}>{pl.name}{pl.discount>0?` (-${pl.discount}%)`:""}</option>)}
                  </select>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {compPopup && <CompPopup order={compPopup} onClose={()=>setCompPopup(null)}/>}
      {showChangePass && <ChangePasswordModal currentUser={currentUser} users={users} setUsers={setUsers} onClose={(updated)=>{setShowChangePass(false);}}/>}
      {showNotifs && isMobile && <NotifPanel notifs={notifs} setNotifs={setNotifs} currentUser={currentUser} users={users} onClose={()=>setShowNotifs(false)} onMarkAllRead={markAllRead} pushNotif={pushNotif} orders={orders}/>}

      <div style={{maxWidth:isMobile?undefined:1200,margin:"0 auto",padding:isMobile?"12px 0":"20px 16px"}}>
        {tab==="central"    && <Central
          orders={isAdmin
            ? orders  // admin ve todos
            : isTestUser
              ? orders.filter(o=>o.isSandbox)  // usuario Prueba ve solo sandbox
              : currentUser.canSeeAll===false
                ? orders.filter(o=>o.vendedor===currentUser.vendedor||o.vendedor===currentUser.name||o.vendedor===currentUser.username)
                : orders}  // vendedor con canSeeAll ve todos
          products={pricedProducts} onStage={setStage} onDel={delOrder} onSaveNote={saveNote}
          onRequestEdit={requestEdit} onApproveEditRequest={approveEditRequest} onRejectEditRequest={rejectEditRequest}
          onSubmitEdit={submitEdit} onApproveEdit={approveEdit} onRejectEdit={rejectEdit}
          currentUser={currentUser} isMobile={isMobile}/>}
        {tab==="nuevo"      && <Nuevo products={pricedProducts} vendors={vendors} onAdd={addOrder} onDone={()=>setTab("central")} currentUser={currentUser} isMobile={isMobile} clients={clients} onSaveClient={saveClient}/>}
        {tab==="clientes"   && <ClientesPanel clients={clients} onSave={saveClient} onDelete={deleteClient} onRequestDelete={requestDeleteClient} currentUser={currentUser} isMobile={isMobile} orders={orders}/>}
        {tab==="cotizacion" && <Cotizaciones quotes={quotes} products={pricedProducts} vendors={vendors} onAdd={addQuote} onDel={delQuote} onConvert={convertQuoteToOrder} onExtend={extendQuote} onTabChange={setTab} currentUser={currentUser} isMobile={isMobile} clients={clients} onSaveClient={saveClient}/>}
        {tab==="precios"    && <Precios products={pricedProducts}/>}
        {tab==="stock"      && <>
              {isTestUser && (
                <div style={{background:"#f5eef8",border:"1.5px solid #9b59b6",borderRadius:10,padding:"10px 16px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
                  <div>
                    <span style={{fontWeight:800,color:"#6c3483",fontSize:13}}>🧪 Modo Sandbox activo</span>
                    <span style={{color:"#888",fontSize:12,marginLeft:8}}>El stock que ves es una copia paralela. No afecta el stock real.</span>
                  </div>
                  <button onClick={()=>{const sb={};products.forEach(p=>{sb[p.id]=p.stock;});updateSandboxStock(sb);}}
                    style={{padding:"6px 14px",borderRadius:8,border:"1.5px solid #9b59b6",background:"#fff",color:"#6c3483",fontWeight:700,fontSize:12,cursor:"pointer"}}>
                    🔄 Reiniciar sandbox
                  </button>
                </div>
              )}
              <Stock products={pricedProducts} onUpd={updProd} onDel={pid=>setProducts(p=>p.filter(x=>x.id!==pid))} onAdjust={(pid,qty)=>setProducts(p=>p.map(x=>x.id===pid?{...x,stock:x.stock+qty}:x))} isAdmin={isAdmin} addLog={addLog} stockLog={stockLog} setStockLog={setStockLog} isMobile={isMobile}/>
            </>}
        {tab==="compras"    && <Compras products={products} onStock={addStock} isMobile={isMobile} canScan={currentUser.role==="admin"||isTestOrder(currentUser.vendedor||currentUser.name)||currentUser.barcodeEnabled}/>}
        {tab==="solicitud"  && <SolicitudCompra products={products} currentUser={currentUser} isAdmin={isAdmin} purchaseOrders={purchaseOrders} setPurchaseOrders={setPurchaseOrders} isMobile={isMobile} onStockExternal={addStock} addLog={addLog}
          onCreated={async(po)=>{
            await sendCrossNotif(db, setNotifs, {
              title:"📋 Nueva solicitud de compra",
              body:`${po.vendedor} creó una solicitud con ${po.items.length} producto${po.items.length!==1?"s":""}`,
              tag:`po-new-${po.id}`,
              para:"admin",
              de:po.vendedor
            });
          }}
        />}
        {tab==="admin"      && isAdmin && <AdminPanel users={users} setUsers={setUsers} vendors={vendors} setVendors={setVendors} products={products} setProducts={setProducts} stockLog={stockLog} setStockLog={setStockLog} notifs={notifs} setNotifs={setNotifs} activity={activity} setActivity={setActivity} orders={orders} priceLists={priceLists} setPriceLists={setPriceLists} isMobile={isMobile} sandboxStock={sandboxStock} setSandboxStock={setSandboxStock}/>}
      </div>
    </div>
  );
}


// ─── CHANGE PASSWORD MODAL ───────────────────────────────────────────────────
function ChangePasswordModal({currentUser, users, setUsers, onClose}) {
  const [current, setCurrent] = useState("");
  const [newPass, setNewPass] = useState("");
  const [confirm, setConfirm] = useState("");
  const [showCurrent, setShowCurrent] = useState(false);
  const [showNew, setShowNew] = useState(false);
  const [error, setError] = useState("");
  const [ok, setOk] = useState(false);

  const save = async () => {
    setError("");
    if(current !== currentUser.password) { setError("La contraseña actual es incorrecta"); return; }
    if(newPass.length < 4) { setError("La nueva contraseña debe tener al menos 4 caracteres"); return; }
    if(newPass !== confirm) { setError("Las contraseñas no coinciden"); return; }
    const updated = {...currentUser, password: newPass};
    setUsers(us=>us.map(u=>u.id===currentUser.id?updated:u));
    await db.saveUser(updated);
    // Update currentUser in parent — force re-login with new password
    setOk(true);
    setTimeout(()=>onClose(updated), 1500);
  };

  return (
    <div style={{position:"fixed",inset:0,background:"#0007",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2000,padding:16}}>
      <div style={{background:"#fff",borderRadius:20,padding:32,maxWidth:400,width:"100%",boxShadow:"0 24px 64px #0005"}}>
        {ok
          ? <div style={{textAlign:"center",padding:"16px 0"}}>
              <div style={{fontSize:52,marginBottom:8}}>✅</div>
              <div style={{fontWeight:800,fontSize:18,color:"#1e8449"}}>Contraseña actualizada</div>
            </div>
          : <>
            <div style={{fontWeight:800,fontSize:18,marginBottom:6}}>🔑 Cambiar contraseña</div>
            <div style={{fontSize:12,color:"#888",marginBottom:20}}>Usuario: <strong>{currentUser.name}</strong></div>

            <Field label="Contraseña actual">
              <div style={{position:"relative"}}>
                <input type={showCurrent?"text":"password"} value={current} onChange={e=>setCurrent(e.target.value)}
                  placeholder="Tu contraseña actual" style={{...inputStyle,paddingRight:40}}/>
                <button onClick={()=>setShowCurrent(s=>!s)} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:15,color:"#aaa"}}>{showCurrent?"🙈":"👁️"}</button>
              </div>
            </Field>
            <Field label="Nueva contraseña">
              <div style={{position:"relative"}}>
                <input type={showNew?"text":"password"} value={newPass} onChange={e=>setNewPass(e.target.value)}
                  placeholder="Mínimo 4 caracteres" style={{...inputStyle,paddingRight:40}}/>
                <button onClick={()=>setShowNew(s=>!s)} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:15,color:"#aaa"}}>{showNew?"🙈":"👁️"}</button>
              </div>
            </Field>
            <Field label="Confirmar nueva contraseña">
              <input type="password" value={confirm} onChange={e=>setConfirm(e.target.value)}
                placeholder="Repetí la nueva contraseña" style={{
                  ...inputStyle,
                  borderColor: confirm && newPass && confirm!==newPass ? "#e74c3c" : confirm && confirm===newPass ? "#1e8449" : undefined
                }}/>
              {confirm && newPass && confirm===newPass && <div style={{fontSize:11,color:"#1e8449",marginTop:3}}>✅ Las contraseñas coinciden</div>}
            </Field>

            {error && <div style={{background:"#fdecea",color:"#c0392b",borderRadius:8,padding:"8px 12px",fontSize:13,marginBottom:12}}>{error}</div>}

            <div style={{display:"flex",gap:8,marginTop:4}}>
              <button onClick={save}
                disabled={!current||!newPass||!confirm}
                style={{flex:1,padding:"10px",borderRadius:10,border:"none",background:(!current||!newPass||!confirm)?"#e5e5e5":"linear-gradient(135deg,#922b21,#c0392b)",color:(!current||!newPass||!confirm)?"#aaa":"#fff",fontWeight:800,fontSize:14,cursor:(!current||!newPass||!confirm)?"not-allowed":"pointer"}}>
                Guardar contraseña
              </button>
              <button onClick={()=>onClose(null)} style={{padding:"10px 16px",borderRadius:10,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontWeight:600,cursor:"pointer"}}>
                Cancelar
              </button>
            </div>
          </>
        }
      </div>
    </div>
  );
}

// ─── COMP POPUP ───────────────────────────────────────────────────────────────
// Shown when a pedido is confirmed - shows Comp number and print button
function CompPopup({order, onClose}) {
  const RED = "#c0392b";
  return (
    <div style={{position:"fixed",inset:0,background:"#0007",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2000,padding:16}}>
      <div style={{background:"#fff",borderRadius:20,padding:32,maxWidth:420,width:"100%",boxShadow:"0 24px 64px #0005",textAlign:"center"}}>
        <div style={{fontSize:52,marginBottom:8}}>✅</div>
        <div style={{fontWeight:800,fontSize:20,color:"#1a5276",marginBottom:6}}>Pedido Confirmado</div>
        <div style={{fontSize:13,color:"#666",marginBottom:20}}>Se generó el comprobante para <strong>{order.client}</strong></div>

        <div style={{background:"#d6eaf8",borderRadius:14,padding:"18px 24px",marginBottom:24}}>
          <div style={{fontSize:11,color:"#1a5276",fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:6}}>Número de Comprobante</div>
          <div style={{fontSize:32,fontWeight:900,color:"#1a5276",letterSpacing:-0.5}}>{order.compNum}</div>
          <div style={{fontSize:12,color:"#888",marginTop:4}}>{order.client} . {fARS(order.total)}</div>
        </div>

        <div style={{display:"flex",gap:10}}>
          <button onClick={()=>{ printDoc(order,"confirmado"); onClose(); }}
            style={{flex:1,padding:"12px",borderRadius:10,border:"none",background:"linear-gradient(135deg,#1a5276,#2980b9)",color:"#fff",fontWeight:800,fontSize:14,cursor:"pointer"}}>
            🖨️️ Imprimir Comp.
          </button>
          <button onClick={onClose}
            style={{padding:"12px 16px",borderRadius:10,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontWeight:600,fontSize:13,cursor:"pointer"}}>
            Cerrar
          </button>
        </div>
      </div>
    </div>
  );
}


// ─── CENTRAL ──────────────────────────────────────────────────────────────────
function Central({orders,products,onStage,onDel,onSaveNote,onRequestEdit,onApproveEditRequest,onRejectEditRequest,onSubmitEdit,onApproveEdit,onRejectEdit,currentUser,isMobile}) {
  const [fStage,setFStage]=useState("todos");
  const [fVendedor,setFVendedor]=useState("todos");
  const [search,setSearch]=useState("");
  const [expanded,setExpanded]=useState(null);
  const getP = id=>products.find(p=>p.id===id);
  const vendedores = useMemo(()=>[...new Set(orders.map(o=>o.vendedor).filter(Boolean))].sort(),[orders]);
  const filtered = orders.filter(o=>{
    if(fStage!=="todos"&&o.stage!==fStage) return false;
    if(fVendedor!=="todos"&&o.vendedor!==fVendedor) return false;
    if(search&&!norm(o.client).includes(norm(search))&&!o.id.toLowerCase().includes(search.toLowerCase())) return false;
    return true;
  });
  // Si todos los pedidos son sandbox (usuario Prueba), contar normalmente
  const allSandbox = orders.length>0 && orders.every(o=>o.isSandbox);
  const deliv = orders.filter(o=>o.stage==="entregado"&&(allSandbox||!o.isSandbox)).reduce((s,o)=>s+o.total,0);
  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(135px,1fr))",gap:12,marginBottom:20}}>
        {STAGES.map(s=>{const c=SCFG[s],cnt=orders.filter(o=>o.stage===s&&(allSandbox||!o.isSandbox)).length;return <div key={s} onClick={()=>setFStage(fStage===s?"todos":s)} style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 6px #0001",borderLeft:`4px solid ${c.color}`,cursor:"pointer",outline:fStage===s?`2px solid ${c.color}`:"none"}}><div style={{fontSize:26,fontWeight:800,color:c.color}}>{cnt}</div><div style={{fontSize:12,color:"#666",fontWeight:600}}>{c.icon} {c.label}</div></div>;})}
        <div style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 6px #0001",borderLeft:`4px solid ${RED}`}}>
          <div style={{fontSize:14,fontWeight:800,color:RED}}>{fARS(deliv)}</div>
          <div style={{fontSize:12,color:"#666",fontWeight:600}}>💰 Entregado</div>
        </div>
      </div>
      <div style={{background:"#fff",borderRadius:12,padding:14,marginBottom:14,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center",boxShadow:"0 1px 4px #0001"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar cliente o No pedido..." style={{flex:1,minWidth:180,padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none"}}/>
        {vendedores.length>1&&<select value={fVendedor} onChange={e=>setFVendedor(e.target.value)}
          style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",cursor:"pointer",background:"#fff"}}>
          <option value="todos">👤 Todos los vendedores</option>
          {vendedores.map(v=><option key={v} value={v}>{v}</option>)}
        </select>}
        <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
          {["todos",...STAGES].map(s=>{const c=SCFG[s];return <button key={s} onClick={()=>setFStage(s)} style={{padding:"5px 11px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:11,fontWeight:600,borderColor:fStage===s?(c?.color||RED):"#e5e5e5",background:fStage===s?(c?.bg||"#fdecea"):"#fff",color:fStage===s?(c?.color||RED):"#666"}}>{s==="todos"?"Todos":c.label}</button>;})}
        </div>
      </div>
      {filtered.length===0
        ? <div style={{textAlign:"center",padding:60,color:"#aaa"}}><div style={{fontSize:48}}>📭</div><div style={{marginTop:8}}>No hay pedidos. !Creá uno desde "Nuevo Pedido"!</div></div>
        : filtered.map(o=><OCard key={o.id} o={o} exp={expanded===o.id} toggle={()=>setExpanded(expanded===o.id?null:o.id)} getP={getP} onStage={onStage} onDel={onDel} onSaveNote={onSaveNote} onRequestEdit={onRequestEdit} onApproveEditRequest={onApproveEditRequest} onRejectEditRequest={onRejectEditRequest} onSubmitEdit={onSubmitEdit} onApproveEdit={onApproveEdit} onRejectEdit={onRejectEdit} currentUser={currentUser} products={products}/>)
      }
    </div>
  );
}

function DelBtn({onConfirm}) {
  const [confirm, setConfirm] = useState(false);
  if(confirm) return (
    <div style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:6,background:"#fdecea",borderRadius:8,padding:"6px 10px",border:"1.5px solid #fcc"}}>
      <span style={{fontSize:12,color:RED,fontWeight:600}}>?Eliminar?</span>
      <button onClick={onConfirm} style={{padding:"4px 10px",borderRadius:6,border:"none",background:RED,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:12}}>Sí</button>
      <button onClick={()=>setConfirm(false)} style={{padding:"4px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12}}>No</button>
    </div>
  );
  return <button onClick={()=>setConfirm(true)} style={{marginLeft:"auto",padding:"8px 12px",borderRadius:8,border:"1.5px solid #fcc",cursor:"pointer",background:"#fff",color:RED,fontWeight:600,fontSize:13}}>🗑 Eliminar</button>;
}

function OCard({o,exp,toggle,getP,onStage,onDel,onSaveNote,onRequestEdit,onApproveEditRequest,onRejectEditRequest,onSubmitEdit,onApproveEdit,onRejectEdit,currentUser,products}) {
  const isAdmin = currentUser?.role === "admin";
  const idx=STAGES.indexOf(o.stage), next=STAGES[idx+1];
  const [editNote,setEditNote]=useState(false);
  const [noteVal,setNoteVal]=useState(o.internalNote||"");

  // Edit request states
  const [showReqForm, setShowReqForm]     = useState(false);
  const [reqReason,   setReqReason]       = useState("");
  const [showRejectForm, setShowRejectForm] = useState(false);
  const [rejectReason, setRejectReason]   = useState("");
  const [showEditMode, setShowEditMode]   = useState(false);
  const [editItems,   setEditItems]       = useState([]);
  const [showFinalReject, setShowFinalReject] = useState(false);
  const [finalRejectReason, setFinalRejectReason] = useState("");
  const [saving, setSaving]               = useState(false);

  const es = o.editStatus || "";

  // Edit badge
  const EditBdg = () => {
    if(!es) return null;
    const cfg = {
      "solicitada":        {bg:"#fef9e7",color:"#b7770d",border:"#f0d080",label:"✏️ Edición solicitada"},
      "aprobada":          {bg:"#eafaf1",color:"#1e8449",border:"#a9dfbf",label:"✅ Podés editar"},
      "rechazada":         {bg:"#fdecea",color:"#c0392b",border:"#f1948a",label:"❌ Edición rechazada"},
      "en revisión":       {bg:"#eaf4fc",color:"#1a5276",border:"#aed6f1",label:"👀 Cambios en revisión"},
      "cambios rechazados":{bg:"#fdecea",color:"#c0392b",border:"#f1948a",label:"❌ Cambios rechazados"},
    };
    const c = cfg[es]; if(!c) return null;
    return <span style={{background:c.bg,color:c.color,border:`1px solid ${c.border}`,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>{c.label}</span>;
  };

  const startEditMode = () => {
    setEditItems(o.items.map(it=>({...it})));
    setShowEditMode(true);
  };
  const updEditItem = (pid,qty) => setEditItems(prev=>prev.map(it=>it.pid===pid?{...it,qty:Math.max(1,qty)}:it));
  const remEditItem = pid => setEditItems(prev=>prev.filter(it=>it.pid!==pid));
  const editTotal = editItems.reduce((s,it)=>s+it.price*it.qty,0);

  return (
    <div style={{background:"#fff",borderRadius:12,boxShadow:"0 2px 8px #0000000a",overflow:"hidden",marginBottom:8,borderLeft:`4px solid ${STAGE_COLORS[o.stage]||"#ccc"}`}}>
      <div onClick={toggle} style={{padding:"13px 18px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",cursor:"pointer"}}>
        <div style={{flex:1,minWidth:100}}>
          <div style={{fontWeight:700,fontSize:14,color:"#1a1a1a"}}>{o.client}</div>
          <div style={{fontSize:11,color:"#aaa",display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
            {o.isTest&&<span style={{background:"#f1c40f",color:"#1a1a1a",borderRadius:4,padding:"1px 6px",fontSize:10,fontWeight:800}}>TEST</span>}
            {o.isSandbox&&<span style={{background:"#9b59b6",color:"#fff",borderRadius:4,padding:"1px 6px",fontSize:10,fontWeight:800}}>🧪 SANDBOX</span>}
            {o.docNum&&!o.isTest&&<span style={{fontWeight:700,color:RED,fontFamily:SERIF}}>{o.docNum}</span>}
            {o.compNum&&!o.isTest&&<span style={{fontWeight:700,color:"#1a5276",fontFamily:SERIF}}>{o.compNum}</span>}
            <span>{o.date}</span>
            {o.vendedor&&<span>· 👤 {o.vendedor}</span>}
            {o.internalNote&&<span style={{color:"#e67e22"}}>· 📝 Nota</span>}
            <EditBdg/>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
          <Bdg stage={o.stage}/>
          <span style={{fontWeight:800,color:RED,fontSize:15}}>{fARS(o.total)}</span>
          <span style={{color:"#ccc"}}>{exp?"▲":"▼"}</span>
        </div>
      </div>

      {exp && (
        <div style={{borderTop:"1px solid #f5f5f5",padding:18}}>

          {/* PROGRESS BAR */}
          <div style={{display:"flex",marginBottom:18,overflowX:"auto"}}>
            {STAGES.map((s,i)=>{const done=i<=idx,c=SCFG[s];return <div key={s} style={{display:"flex",alignItems:"center",flex:1,minWidth:65}}><div style={{textAlign:"center",flex:1}}><div style={{width:30,height:30,borderRadius:"50%",background:done?c.color:"#eee",color:done?"#fff":"#aaa",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 4px",fontSize:13}}>{done?c.icon:"○"}</div><div style={{fontSize:10,color:done?c.color:"#aaa",fontWeight:done?700:400}}>{c.label}</div></div>{i<3&&<div style={{height:2,background:i<idx?RED:"#eee",flex:1}}/>}</div>;})}
          </div>

          {/* ── EDIT STATUS BLOCKS ── */}

          {/* Admin: solicitud pendiente */}
          {isAdmin && es==="solicitada" && (
            <div style={{background:"#fef9e7",border:"1.5px solid #f0d080",borderRadius:10,padding:"12px 14px",marginBottom:14}}>
              <div style={{fontWeight:800,fontSize:13,color:"#b7770d",marginBottom:4}}>✏️ Solicitud de edición</div>
              <div style={{fontSize:12,color:"#7d6608",marginBottom:10}}>Motivo: "{o.editReason}"</div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                <button onClick={async()=>{setSaving(true);await onApproveEditRequest(o.id);setSaving(false);}}
                  style={{padding:"7px 14px",borderRadius:8,border:"none",background:"#1e8449",color:"#fff",fontWeight:700,fontSize:12,cursor:"pointer"}}>
                  ✅ Aprobar solicitud
                </button>
                {!showRejectForm && <button onClick={()=>setShowRejectForm(true)}
                  style={{padding:"7px 12px",borderRadius:8,border:"1.5px solid #f1948a",background:"#fff",color:"#c0392b",fontWeight:700,fontSize:12,cursor:"pointer"}}>
                  ❌ Rechazar
                </button>}
              </div>
              {showRejectForm && (
                <div style={{marginTop:10}} onClick={e=>e.stopPropagation()}>
                  <textarea value={rejectReason} onChange={e=>setRejectReason(e.target.value)}
                    placeholder="Motivo del rechazo..."
                    style={{width:"100%",padding:"8px 10px",borderRadius:8,border:"1.5px solid #f1948a",fontSize:13,resize:"vertical",minHeight:60,outline:"none",boxSizing:"border-box"}}/>
                  <div style={{display:"flex",gap:8,marginTop:6}}>
                    <button onClick={async()=>{if(!rejectReason.trim())return;setSaving(true);await onRejectEditRequest(o.id,rejectReason);setShowRejectForm(false);setRejectReason("");setSaving(false);}}
                      disabled={!rejectReason.trim()}
                      style={{padding:"6px 14px",borderRadius:7,border:"none",background:rejectReason.trim()?"#c0392b":"#e5e5e5",color:rejectReason.trim()?"#fff":"#aaa",fontWeight:700,fontSize:12,cursor:rejectReason.trim()?"pointer":"not-allowed"}}>
                      Confirmar rechazo
                    </button>
                    <button onClick={()=>{setShowRejectForm(false);setRejectReason("");}}
                      style={{padding:"6px 12px",borderRadius:7,border:"1px solid #e5e5e5",background:"#fff",color:"#666",fontSize:12,cursor:"pointer"}}>Cancelar</button>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Vendedor: solicitud enviada esperando */}
          {!isAdmin && es==="solicitada" && (
            <div style={{background:"#fef9e7",border:"1.5px solid #f0d080",borderRadius:10,padding:"10px 14px",marginBottom:14,fontSize:12,color:"#b7770d",fontWeight:600}}>
              ⏳ Solicitud de edición enviada al admin. Esperando aprobación...
            </div>
          )}

          {/* Vendedor: solicitud rechazada */}
          {!isAdmin && es==="rechazada" && (
            <div style={{background:"#fdecea",border:"1.5px solid #f1948a",borderRadius:10,padding:"10px 14px",marginBottom:14}}>
              <div style={{fontWeight:800,fontSize:13,color:"#c0392b",marginBottom:3}}>❌ Solicitud rechazada</div>
              <div style={{fontSize:12,color:"#922b21"}}>Motivo: "{o.editRejectReason}"</div>
              <button onClick={()=>{const u=orders||[];}}
                style={{marginTop:8,padding:"5px 12px",borderRadius:7,border:"1px solid #f1948a",background:"#fff",color:"#c0392b",fontSize:11,fontWeight:600,cursor:"pointer"}}
                onClick={async()=>{setSaving(true);await onRequestEdit(o.id,"");setSaving(false);}}>
                Volver a solicitar
              </button>
            </div>
          )}

          {/* Vendedor: aprobada — puede editar */}
          {!isAdmin && es==="aprobada" && !showEditMode && (
            <div style={{background:"#eafaf1",border:"1.5px solid #a9dfbf",borderRadius:10,padding:"10px 14px",marginBottom:14,display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
              <div>
                <div style={{fontWeight:800,fontSize:13,color:"#1e8449"}}>✅ Edición aprobada</div>
                <div style={{fontSize:12,color:"#1a5276",marginTop:2}}>El admin autorizó la edición. Hacé tus cambios y enviá para revisión final.</div>
              </div>
              <button onClick={startEditMode}
                style={{padding:"8px 16px",borderRadius:8,border:"none",background:"linear-gradient(135deg,#1a5276,#2980b9)",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}>
                ✏️ Editar pedido
              </button>
            </div>
          )}

          {/* EDIT MODE — formulario de edición */}
          {!isAdmin && es==="aprobada" && showEditMode && (
            <div style={{background:"#eaf4fc",border:"1.5px solid #aed6f1",borderRadius:10,padding:"14px",marginBottom:14}} onClick={e=>e.stopPropagation()}>
              <div style={{fontWeight:800,fontSize:13,color:"#1a5276",marginBottom:10}}>✏️ Editando pedido</div>
              {editItems.map((it)=>{
                const p = getP(it.pid);
                return (
                  <div key={it.pid} style={{background:"#fff",borderRadius:8,padding:"10px 12px",marginBottom:8,border:"1px solid #d6eaf8"}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                      <div style={{flex:1,fontSize:13,fontWeight:600,color:"#1a1a1a",marginRight:8,lineHeight:1.3}}>{p?.name||it.name}</div>
                      <button onClick={()=>remEditItem(it.pid)} style={{background:"none",border:"none",color:"#c0392b",fontSize:20,cursor:"pointer",lineHeight:1,flexShrink:0}}>×</button>
                    </div>
                    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <button onClick={()=>updEditItem(it.pid,it.qty-1)} style={{width:32,height:32,borderRadius:8,border:"1.5px solid #aed6f1",background:"#fff",fontWeight:800,cursor:"pointer",fontSize:16}}>−</button>
                        <span style={{minWidth:32,textAlign:"center",fontWeight:800,fontSize:15}}>{it.qty}</span>
                        <button onClick={()=>updEditItem(it.pid,it.qty+1)} style={{width:32,height:32,borderRadius:8,border:"1.5px solid #aed6f1",background:"#fff",fontWeight:800,cursor:"pointer",fontSize:16}}>+</button>
                      </div>
                      <span style={{fontWeight:700,color:RED,fontSize:14}}>{fARS(it.price*it.qty)}</span>
                    </div>
                  </div>
                );
              })}
              <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:15,color:RED,padding:"10px 0",borderTop:"2px solid #d6eaf8",margin:"4px 0 10px"}}>
                <span>Nuevo total</span><span>{fARS(editTotal)}</span>
              </div>
              <button onClick={async()=>{if(!editItems.length)return;setSaving(true);await onSubmitEdit(o.id,editItems,editTotal);setShowEditMode(false);setSaving(false);}}
                disabled={!editItems.length||saving}
                style={{width:"100%",padding:"10px",borderRadius:9,border:"none",background:editItems.length?"linear-gradient(135deg,#1a5276,#2980b9)":"#e5e5e5",color:editItems.length?"#fff":"#aaa",fontWeight:800,fontSize:14,cursor:editItems.length?"pointer":"not-allowed",marginBottom:8}}>
                {saving?"Enviando...":"📤 Enviar para revisión"}
              </button>
              <button onClick={()=>setShowEditMode(false)}
                style={{width:"100%",padding:"9px",borderRadius:9,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontSize:13,cursor:"pointer"}}>Cancelar</button>
            </div>
          )}

          {/* Admin: cambios en revisión — lista separada para mobile */}
          {isAdmin && es==="en revisión" && (
            <div style={{background:"#eaf4fc",border:"1.5px solid #aed6f1",borderRadius:10,padding:"14px",marginBottom:14}}>
              <div style={{fontWeight:800,fontSize:13,color:"#1a5276",marginBottom:12}}>👀 Revisión de cambios — {o.vendedor}</div>
              {/* Original */}
              <div style={{marginBottom:10}}>
                <div style={{fontSize:10,fontWeight:800,color:"#888",letterSpacing:1,textTransform:"uppercase",marginBottom:6}}>Original</div>
                {o.items.map((it,i)=>{const p=getP(it.pid);return <div key={i} style={{fontSize:12,color:"#555",padding:"4px 0",borderBottom:"1px solid #f0f0f0",display:"flex",justifyContent:"space-between"}}><span>{p?.name||it.name} × {it.qty}</span><span>{fARS(it.price*it.qty)}</span></div>;})}
                <div style={{fontWeight:700,fontSize:13,color:"#555",marginTop:6,textAlign:"right"}}>{fARS(o.total)}</div>
              </div>
              {/* Nuevo */}
              <div style={{background:"#fff",borderRadius:8,padding:"10px 12px",border:"1.5px solid #aed6f1",marginBottom:12}}>
                <div style={{fontSize:10,fontWeight:800,color:"#1a5276",letterSpacing:1,textTransform:"uppercase",marginBottom:6}}>Nuevo ✏️</div>
                {(o.editItems||[]).map((it,i)=>{
                  const p=getP(it.pid);
                  const orig=o.items.find(x=>x.pid===it.pid);
                  const changed=!orig||orig.qty!==it.qty;
                  return <div key={i} style={{fontSize:12,color:changed?"#1a5276":"#555",fontWeight:changed?700:400,padding:"4px 0",borderBottom:"1px solid #f0f0f0",display:"flex",justifyContent:"space-between"}}><span>{p?.name||it.name} × {it.qty}{changed?" ✏️":""}</span><span>{fARS(it.price*it.qty)}</span></div>;
                })}
                <div style={{fontWeight:800,fontSize:13,color:RED,marginTop:6,textAlign:"right"}}>{fARS((o.editItems||[]).reduce((s,it)=>s+it.price*it.qty,0))}</div>
              </div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                <button onClick={async()=>{setSaving(true);await onApproveEdit(o.id);setSaving(false);}}
                  style={{padding:"8px 16px",borderRadius:8,border:"none",background:"linear-gradient(135deg,#1a5e20,#1e8449)",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}>
                  ✅ Aprobar cambios
                </button>
                {!showFinalReject && <button onClick={()=>setShowFinalReject(true)}
                  style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #f1948a",background:"#fff",color:"#c0392b",fontWeight:700,fontSize:13,cursor:"pointer"}}>
                  ❌ Rechazar cambios
                </button>}
              </div>
              {showFinalReject && (
                <div style={{marginTop:10}} onClick={e=>e.stopPropagation()}>
                  <textarea value={finalRejectReason} onChange={e=>setFinalRejectReason(e.target.value)}
                    placeholder="Motivo del rechazo..."
                    style={{width:"100%",padding:"8px 10px",borderRadius:8,border:"1.5px solid #f1948a",fontSize:13,resize:"vertical",minHeight:60,outline:"none",boxSizing:"border-box"}}/>
                  <div style={{display:"flex",gap:8,marginTop:6}}>
                    <button onClick={async()=>{if(!finalRejectReason.trim())return;setSaving(true);await onRejectEdit(o.id,finalRejectReason);setShowFinalReject(false);setFinalRejectReason("");setSaving(false);}}
                      disabled={!finalRejectReason.trim()}
                      style={{padding:"6px 14px",borderRadius:7,border:"none",background:finalRejectReason.trim()?"#c0392b":"#e5e5e5",color:finalRejectReason.trim()?"#fff":"#aaa",fontWeight:700,fontSize:12,cursor:finalRejectReason.trim()?"pointer":"not-allowed"}}>
                      Confirmar rechazo
                    </button>
                    <button onClick={()=>{setShowFinalReject(false);setFinalRejectReason("");}}
                      style={{padding:"6px 12px",borderRadius:7,border:"1px solid #e5e5e5",background:"#fff",color:"#666",fontSize:12,cursor:"pointer"}}>Cancelar</button>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Vendedor: cambios en revisión */}
          {!isAdmin && es==="en revisión" && (
            <div style={{background:"#eaf4fc",border:"1.5px solid #aed6f1",borderRadius:10,padding:"10px 14px",marginBottom:14,fontSize:12,color:"#1a5276",fontWeight:600}}>
              👀 Tus cambios fueron enviados. El admin está revisando la edición final.
            </div>
          )}

          {/* Vendedor: cambios rechazados en revisión final */}
          {!isAdmin && es==="cambios rechazados" && (
            <div style={{background:"#fdecea",border:"1.5px solid #f1948a",borderRadius:10,padding:"10px 14px",marginBottom:14}}>
              <div style={{fontWeight:800,fontSize:13,color:"#c0392b",marginBottom:3}}>❌ Cambios rechazados por el admin</div>
              <div style={{fontSize:12,color:"#922b21"}}>Motivo: "{o.editRejectReason}"</div>
              <div style={{fontSize:11,color:"#888",marginTop:4}}>El pedido quedó con los datos originales.</div>
            </div>
          )}

          {/* ITEMS LIST */}
          {o.items.map((it,i)=>{const p=getP(it.pid);return <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #f9f9f9",fontSize:13}}><span style={{color:"#444"}}>{p?.name||it.name} x {it.qty}</span><span style={{fontWeight:600}}>{fARS(it.price*it.qty)}</span></div>;})}
          <div style={{display:"flex",justifyContent:"flex-end",fontWeight:800,fontSize:16,color:RED,margin:"8px 0 12px"}}>{fARS(o.total)}</div>
          {o.notes&&<div style={{background:"#f9f9f9",borderRadius:8,padding:"8px 12px",fontSize:13,color:"#555",marginBottom:12}}>💬 {o.notes}</div>}

          {/* NOTA INTERNA */}
          <div style={{background:"#fffbf0",border:"1.5px solid #f0d080",borderRadius:10,padding:"10px 14px",marginBottom:14}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6}}>
              <span style={{fontSize:12,fontWeight:700,color:"#b7770d"}}>📝 Nota interna (solo admin)</span>
              {!editNote&&<button onClick={(e)=>{e.stopPropagation();setEditNote(true);setNoteVal(o.internalNote||"");}}
                style={{padding:"3px 10px",borderRadius:6,border:"1px solid #f0d080",background:"#fff",color:"#b7770d",fontSize:11,fontWeight:700,cursor:"pointer"}}>
                {o.internalNote?"✏️ Editar":"+ Agregar"}
              </button>}
            </div>
            {editNote ? (
              <div onClick={e=>e.stopPropagation()}>
                <textarea value={noteVal} onChange={e=>setNoteVal(e.target.value)} placeholder="Escribí una nota interna..."
                  style={{width:"100%",padding:"8px 10px",borderRadius:8,border:"1.5px solid #f0d080",fontSize:13,resize:"vertical",minHeight:64,outline:"none",boxSizing:"border-box",background:"#fff"}}/>
                <div style={{display:"flex",gap:8,marginTop:6}}>
                  <button onClick={async(e)=>{e.stopPropagation();await onSaveNote(o.id,noteVal);setEditNote(false);}}
                    style={{padding:"6px 14px",borderRadius:7,border:"none",background:"#b7770d",color:"#fff",fontWeight:700,fontSize:12,cursor:"pointer"}}>💾 Guardar</button>
                  <button onClick={(e)=>{e.stopPropagation();setEditNote(false);setNoteVal(o.internalNote||"");}}
                    style={{padding:"6px 12px",borderRadius:7,border:"1px solid #e5e5e5",background:"#fff",color:"#666",fontSize:12,cursor:"pointer"}}>Cancelar</button>
                </div>
              </div>
            ) : (
              <div style={{fontSize:13,color:o.internalNote?"#5d4037":"#aaa",fontStyle:o.internalNote?"normal":"italic"}}>{o.internalNote||"Sin nota interna"}</div>
            )}
          </div>

          {/* ACTIONS */}
          <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
            {next&&!es&&<button onClick={()=>onStage(o.id,next)} style={{padding:"8px 14px",borderRadius:8,border:"none",cursor:"pointer",background:SCFG[next].color,color:"#fff",fontWeight:700,fontSize:13}}>{SCFG[next].icon} Pasar a {SCFG[next].label}</button>}
            {idx>0&&o.stage!=="entregado"&&!es&&<button onClick={()=>onStage(o.id,STAGES[idx-1])} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",cursor:"pointer",background:"#fff",color:"#666",fontWeight:600,fontSize:13}}>← Retroceder</button>}
            <button onClick={()=>printDoc(o, o.stage==="reserva"?"reserva":"confirmado")} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #d6eaf8",cursor:"pointer",background:"#fff",color:"#1a5276",fontWeight:600,fontSize:13}}>
              🖨️ {o.stage==="reserva"?(o.docNum||"Imprimir"):(o.compNum||"Imprimir")}
            </button>
            <button onClick={()=>window.open(`https://wa.me/?text=${encodeURIComponent(buildWAOrder(o))}`, "_blank")}
              style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #d5f5e3",cursor:"pointer",background:"#25D366",color:"#fff",fontWeight:700,fontSize:13,display:"flex",alignItems:"center",gap:5}}>
              <span style={{fontSize:15}}>💬</span> Enviar por WhatsApp
            </button>
            {/* Solicitar edición — solo vendedor, solo si no hay edición en curso */}
            {!isAdmin && !es && o.stage!=="entregado" && (
              !showReqForm
                ? <button onClick={(e)=>{e.stopPropagation();setShowReqForm(true);}}
                    style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #aed6f1",background:"#fff",color:"#1a5276",fontWeight:600,fontSize:13,cursor:"pointer"}}>
                    ✏️ Solicitar edición
                  </button>
                : <div style={{width:"100%",marginTop:8}} onClick={e=>e.stopPropagation()}>
                    <textarea value={reqReason} onChange={e=>setReqReason(e.target.value)}
                      placeholder="Motivo de la edición..."
                      style={{width:"100%",padding:"8px 10px",borderRadius:8,border:"1.5px solid #aed6f1",fontSize:13,resize:"vertical",minHeight:60,outline:"none",boxSizing:"border-box"}}/>
                    <div style={{display:"flex",gap:8,marginTop:6}}>
                      <button onClick={async()=>{if(!reqReason.trim())return;setSaving(true);await onRequestEdit(o.id,reqReason);setShowReqForm(false);setReqReason("");setSaving(false);}}
                        disabled={!reqReason.trim()||saving}
                        style={{padding:"6px 14px",borderRadius:7,border:"none",background:reqReason.trim()?"#1a5276":"#e5e5e5",color:reqReason.trim()?"#fff":"#aaa",fontWeight:700,fontSize:12,cursor:reqReason.trim()?"pointer":"not-allowed"}}>
                        {saving?"Enviando...":"📤 Enviar solicitud"}
                      </button>
                      <button onClick={()=>{setShowReqForm(false);setReqReason("");}}
                        style={{padding:"6px 12px",borderRadius:7,border:"1px solid #e5e5e5",background:"#fff",color:"#666",fontSize:12,cursor:"pointer"}}>Cancelar</button>
                    </div>
                  </div>
            )}
            <DelBtn onConfirm={()=>onDel(o.id)}/>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── PRECIOS ──────────────────────────────────────────────────────────────────
// ─── CLIENT SELECTOR (usado en Nuevo Pedido y Cotizaciones) ──────────────────
function ClientSelector({clients, onSelect, onSaveClient, currentUser}) {
  const [search, setSearch] = useState("");
  const [showNew, setShowNew] = useState(false);
  const [form, setForm] = useState({name:"", phone:""});
  const [saving, setSaving] = useState(false);

  const found = useMemo(()=>{
    const q = norm(search);
    return q ? clients.filter(c=>norm(c.name).includes(q)||c.phone.includes(search)||norm(c.cuit||"").includes(q)).slice(0,8) : clients.slice(0,5);
  }, [clients, search]);

  const createAndSelect = async () => {
    if(!form.name.trim()||!form.phone.trim()) { alert("Nombre y teléfono son obligatorios"); return; }
    setSaving(true);
    const newClient = {id:genId(), name:form.name.trim(), phone:form.phone.trim(), email:"", cuit:"", address:"", notes:"", deleteRequested:false, deleteReason:"", createdBy:currentUser.name, createdAt:today()};
    await onSaveClient(newClient);
    onSelect(newClient);
    setSaving(false);
  };

  return (
    <div>
      <input value={search} onChange={e=>setSearch(e.target.value)}
        placeholder="🔍 Buscar cliente por nombre o teléfono..."
        style={{width:"100%",padding:"10px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:8}}/>

      {/* Results */}
      <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:10}}>
        {found.map(c=>(
          <div key={c.id} onClick={()=>onSelect(c)}
            style={{background:"#f9f9f9",borderRadius:10,padding:"11px 14px",cursor:"pointer",border:"1.5px solid #f0f0f0",display:"flex",alignItems:"center",gap:10}}
            onMouseEnter={e=>e.currentTarget.style.borderColor="#c0392b"}
            onMouseLeave={e=>e.currentTarget.style.borderColor="#f0f0f0"}>
            <span style={{fontSize:20}}>👤</span>
            <div style={{flex:1}}>
              <div style={{fontWeight:700,fontSize:13}}>{c.name}</div>
              <div style={{fontSize:11,color:"#888"}}>📱 {c.phone}{c.email&&` · 📧 ${c.email}`}</div>
            </div>
            <span style={{color:"#c0392b",fontSize:18}}>›</span>
          </div>
        ))}
        {found.length===0&&search&&<div style={{textAlign:"center",color:"#aaa",fontSize:13,padding:"12px 0"}}>No se encontró "{search}"</div>}
      </div>

      {/* Separator */}
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
        <div style={{flex:1,height:1,background:"#e5e5e5"}}/>
        <span style={{fontSize:11,color:"#aaa",fontWeight:600}}>o</span>
        <div style={{flex:1,height:1,background:"#e5e5e5"}}/>
      </div>

      {/* Quick new client */}
      {!showNew
        ? <button onClick={()=>setShowNew(true)}
            style={{width:"100%",padding:"10px",borderRadius:10,border:"1.5px dashed #aed6f1",background:"#f0f8ff",color:"#1a5276",fontWeight:700,fontSize:13,cursor:"pointer"}}>
            ➕ Cliente nuevo
          </button>
        : <div style={{background:"#f0f8ff",border:"1.5px solid #aed6f1",borderRadius:10,padding:"12px 14px"}}>
            <div style={{fontWeight:700,fontSize:12,color:"#1a5276",marginBottom:10}}>➕ Nuevo cliente rápido</div>
            <Field label="Nombre / Razón social *">
              <input value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} placeholder="Ej: Librería Nueva" style={inputStyle}/>
            </Field>
            <Field label="Teléfono *">
              <input value={form.phone} onChange={e=>setForm(f=>({...f,phone:e.target.value}))} placeholder="+54 11 1234-5678" style={inputStyle}/>
            </Field>
            <div style={{fontSize:11,color:"#888",marginBottom:10}}>Los datos opcionales (CUIT, email, dirección) se completan desde la sección Clientes.</div>
            <div style={{display:"flex",gap:8}}>
              <button onClick={createAndSelect} disabled={saving||!form.name.trim()||!form.phone.trim()}
                style={{flex:1,padding:"9px",borderRadius:8,border:"none",background:(!form.name.trim()||!form.phone.trim())?"#e5e5e5":"#1a5276",color:(!form.name.trim()||!form.phone.trim())?"#aaa":"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}>
                {saving?"Creando...":"✅ Crear y usar"}
              </button>
              <button onClick={()=>setShowNew(false)} style={{padding:"9px 14px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontSize:13,cursor:"pointer"}}>Cancelar</button>
            </div>
          </div>
      }
    </div>
  );
}

// ─── CLIENTES PANEL ───────────────────────────────────────────────────────────
// ─── CLIENT HISTORIAL ────────────────────────────────────────────────────────
function ClientHistorial({client, orders, period, setPeriod, onBack}) {
  const now = new Date();
  const days = period==="30" ? 30 : period==="180" ? 180 : 365;
  const cutoff = new Date(now.getTime() - days*24*60*60*1000);

  const parseDate = (d) => {
    if(!d) return null;
    const [day,mon,yr] = d.split("/");
    return new Date(+yr, +mon-1, +day);
  };

  const clientOrders = useMemo(()=>
    orders.filter(o=>{
      if(o.isSandbox) return false;
      const matches = o.client===client.name || o.clientId===client.id;
      if(!matches) return false;
      const dt = parseDate(o.date);
      return dt && dt >= cutoff;
    }).sort((a,b)=>(parseDate(b.date)||0)-(parseDate(a.date)||0))
  ,[orders, client, period]);

  const totalGastado = clientOrders.filter(o=>o.stage==="entregado").reduce((s,o)=>s+o.total,0);
  const totalPedidos = clientOrders.length;

  return (
    <div>
      <button onClick={onBack} style={{display:"flex",alignItems:"center",gap:6,padding:"8px 14px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",color:"#555",fontWeight:600,fontSize:13,cursor:"pointer",marginBottom:12}}>
        ← Volver a clientes
      </button>

      {/* Header cliente */}
      <div style={{background:"#fff",borderRadius:12,padding:"16px 18px",marginBottom:12,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:16}}>{client.name}</div>
        <div style={{fontSize:12,color:"#888",marginTop:3}}>
          {client.phone&&<span>📱 {client.phone}</span>}
          {client.email&&<span style={{marginLeft:10}}>📧 {client.email}</span>}
          {client.cuit&&<span style={{marginLeft:10}}>🪪 {client.cuit}</span>}
        </div>
      </div>

      {/* Filtro período */}
      <div style={{display:"flex",gap:8,marginBottom:12}}>
        {[["30","30 días"],["180","6 meses"],["365","1 año"]].map(([val,label])=>(
          <button key={val} onClick={()=>setPeriod(val)}
            style={{padding:"7px 14px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:12,fontWeight:700,
              borderColor:period===val?"#c0392b":"#e5e5e5",
              background:period===val?"#fdecea":"#fff",
              color:period===val?"#c0392b":"#666"}}>
            {label}
          </button>
        ))}
      </div>

      {/* KPIs */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12}}>
        <div style={{background:"#fff",borderRadius:12,padding:"12px 14px",boxShadow:"0 1px 4px #0001",borderLeft:"3px solid #c0392b"}}>
          <div style={{fontSize:22,fontWeight:900,color:"#c0392b"}}>{totalPedidos}</div>
          <div style={{fontSize:10,color:"#888",fontWeight:600,marginTop:2}}>Pedidos</div>
        </div>
        <div style={{background:"#fff",borderRadius:12,padding:"12px 14px",boxShadow:"0 1px 4px #0001",borderLeft:"3px solid #1e8449"}}>
          <div style={{fontSize:16,fontWeight:900,color:"#1e8449"}}>{fARS(totalGastado)}</div>
          <div style={{fontSize:10,color:"#888",fontWeight:600,marginTop:2}}>Total entregado</div>
        </div>
        <div style={{background:"#fff",borderRadius:12,padding:"12px 14px",boxShadow:"0 1px 4px #0001",borderLeft:"3px solid #1a5276"}}>
          <div style={{fontSize:16,fontWeight:900,color:"#1a5276"}}>{totalPedidos>0?fARS(totalGastado/totalPedidos):"$0"}</div>
          <div style={{fontSize:10,color:"#888",fontWeight:600,marginTop:2}}>Ticket prom.</div>
        </div>
      </div>

      {/* Lista de pedidos */}
      {clientOrders.length===0
        ? <div style={{textAlign:"center",padding:40,background:"#fff",borderRadius:12,color:"#aaa"}}>
            <div style={{fontSize:36,marginBottom:8}}>📋</div>
            <div style={{fontWeight:600}}>Sin pedidos en los últimos {period==="30"?"30 días":period==="180"?"6 meses":"año"}</div>
          </div>
        : clientOrders.map(o=>(
          <div key={o.id} style={{background:"#fff",borderRadius:12,padding:"13px 16px",marginBottom:8,boxShadow:"0 1px 4px #0001"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:4}}>
              <div>
                <div style={{fontWeight:700,fontSize:13}}>{o.docNum||o.compNum||o.id.slice(0,8)}</div>
                <div style={{fontSize:11,color:"#aaa"}}>{o.date} · 👤 {o.vendedor}</div>
              </div>
              <div style={{textAlign:"right"}}>
                <div style={{fontWeight:800,fontSize:14,color:"#c0392b"}}>{fARS(o.total)}</div>
                <Bdg stage={o.stage}/>
              </div>
            </div>
            <div style={{borderTop:"1px solid #f5f5f5",paddingTop:6,marginTop:4}}>
              {o.items.slice(0,3).map((it,i)=>(
                <div key={i} style={{fontSize:11,color:"#888",display:"flex",justifyContent:"space-between"}}>
                  <span>{it.name} × {it.qty}</span>
                  <span>{fARS(it.price*it.qty)}</span>
                </div>
              ))}
              {o.items.length>3&&<div style={{fontSize:10,color:"#aaa",marginTop:2}}>+{o.items.length-3} productos más</div>}
            </div>
          </div>
        ))
      }
    </div>
  );
}

function ClientesPanel({clients, onSave, onDelete, onRequestDelete, currentUser, isMobile, orders=[]}) {
  const isAdmin = currentUser.role === "admin";
  const [view, setView] = useState("lista");
  const [search, setSearch] = useState("");
  const [editing, setEditing] = useState(null);
  const [form, setForm] = useState({name:"",phone:"",email:"",cuit:"",address:"",notes:""});
  const [delForm, setDelForm] = useState({id:null,reason:""});
  const [saving, setSaving] = useState(false);
  const [historialClient, setHistorialClient] = useState(null);
  const [historialPeriod, setHistorialPeriod] = useState("30"); // days

  const filtered = useMemo(()=>{
    const q = norm(search);
    return q ? clients.filter(c=>norm(c.name).includes(q)||c.phone.includes(search)||norm(c.cuit||"").includes(q)) : clients;
  }, [clients, search]);

  const startEdit = (c) => {
    setEditing(c.id);
    setForm({name:c.name,phone:c.phone,email:c.email||"",cuit:c.cuit||"",address:c.address||"",notes:c.notes||""});
    setView("form");
  };
  const cancelEdit = () => { setEditing(null); setForm({name:"",phone:"",email:"",cuit:"",address:"",notes:""}); setView("lista"); };

  const save = async () => {
    if(!form.name.trim()||!form.phone.trim()){alert("Nombre y teléfono son obligatorios");return;}
    setSaving(true);
    const client = editing
      ? {...clients.find(c=>c.id===editing), ...form}
      : {id:genId(), ...form, deleteRequested:false, deleteReason:"", createdBy:currentUser.name, createdAt:today()};
    await onSave(client);
    setSaving(false);
    cancelEdit();
  };

  return (
    <div>
      {/* Tabs */}
      <div style={{background:"#fff",borderRadius:12,padding:4,marginBottom:14,display:"flex",gap:4,boxShadow:"0 1px 4px #0001"}}>
        <button onClick={()=>{cancelEdit();setView("lista");}} style={{flex:1,padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="lista"?`linear-gradient(135deg,#922b21,#c0392b)`:"transparent",color:view==="lista"?"#fff":"#555"}}>
          👥 Clientes ({clients.length})
        </button>
        <button onClick={()=>setView("form")} style={{flex:1,padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="form"?`linear-gradient(135deg,#922b21,#c0392b)`:"transparent",color:view==="form"?"#fff":"#555"}}>
          {editing?"✏️ Editando":"+ Nuevo cliente"}
        </button>
      </div>

      {/* HISTORIAL */}
      {view==="historial" && historialClient && (
        <ClientHistorial client={historialClient} orders={orders} period={historialPeriod} setPeriod={setHistorialPeriod} onBack={()=>setView("lista")}/>
      )}

      {/* LISTA */}
      {view==="lista" && <>
        <input value={search} onChange={e=>setSearch(e.target.value)}
          placeholder="🔍 Buscar por nombre, teléfono o CUIT..."
          style={{width:"100%",padding:"10px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:12}}/>
        {filtered.length===0
          ? <div style={{textAlign:"center",padding:40,color:"#aaa",background:"#fff",borderRadius:12}}>
              <div style={{fontSize:40,marginBottom:8}}>👥</div>
              <div>{search?"No se encontraron clientes":"No hay clientes aún. ¡Agregá el primero!"}</div>
            </div>
          : filtered.map(c=>(
            <div key={c.id} style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 4px #0001",marginBottom:8,border:"1.5px solid #f0f0f0"}}>
              <div style={{display:"flex",alignItems:"flex-start",gap:10}}>
                <span style={{fontSize:26,flexShrink:0}}>👤</span>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:700,fontSize:14}}>{c.name}</div>
                  <div style={{fontSize:12,color:"#555",marginTop:3}}>📱 {c.phone}</div>
                  {c.email&&<div style={{fontSize:12,color:"#888"}}>📧 {c.email}</div>}
                  {c.cuit&&<div style={{fontSize:12,color:"#888"}}>🪪 CUIT: {c.cuit}</div>}
                  {c.address&&<div style={{fontSize:12,color:"#888"}}>📍 {c.address}</div>}
                  {c.notes&&<div style={{fontSize:11,color:"#aaa",marginTop:4,fontStyle:"italic"}}>"{c.notes}"</div>}
                  {c.deleteRequested&&<div style={{background:"#fef9e7",border:"1px solid #f0d080",borderRadius:6,padding:"4px 8px",fontSize:11,color:"#b7770d",fontWeight:600,marginTop:6}}>
                    ⏳ Solicitud de baja pendiente — "{c.deleteReason}"
                    {isAdmin&&<button onClick={()=>onDelete(c.id)} style={{marginLeft:8,padding:"2px 8px",borderRadius:4,border:"none",background:"#c0392b",color:"#fff",fontSize:10,fontWeight:700,cursor:"pointer"}}>Aprobar baja</button>}
                  </div>}
                </div>
                <div style={{display:"flex",flexDirection:"column",gap:5,flexShrink:0}}>
                  <button onClick={()=>{setHistorialClient(c);setView("historial");}} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #eaf4fc",background:"#fff",color:"#1a5276",cursor:"pointer",fontSize:11,fontWeight:600}}>📋</button>
            <button onClick={()=>startEdit(c)} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:11,fontWeight:600}}>✏️</button>
                  {isAdmin
                    ? <button onClick={()=>{if(window.confirm(`¿Eliminar a "${c.name}"?`))onDelete(c.id);}} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #fcc",background:"#fff",color:"#c0392b",cursor:"pointer",fontSize:11}}>🗑</button>
                    : !c.deleteRequested&&<button onClick={()=>setDelForm({id:c.id,reason:""})} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #fcc",background:"#fff",color:"#c0392b",cursor:"pointer",fontSize:10,fontWeight:600}}>Solicitar baja</button>
                  }
                </div>
              </div>
              {delForm.id===c.id&&<div style={{marginTop:10,background:"#fef9e7",border:"1.5px solid #f0d080",borderRadius:8,padding:"10px 12px"}} onClick={e=>e.stopPropagation()}>
                <div style={{fontSize:12,fontWeight:700,color:"#b7770d",marginBottom:6}}>Motivo de la solicitud de baja</div>
                <input value={delForm.reason} onChange={e=>setDelForm(f=>({...f,reason:e.target.value}))} placeholder="Ej: Cliente duplicado..." style={{...inputStyle,fontSize:12,marginBottom:8}}/>
                <div style={{display:"flex",gap:6}}>
                  <button onClick={async()=>{if(!delForm.reason.trim()){alert("Escribí un motivo");return;}await onRequestDelete(delForm.id,delForm.reason);setDelForm({id:null,reason:""}); }}
                    style={{flex:1,padding:"7px",borderRadius:7,border:"none",background:"#b7770d",color:"#fff",fontWeight:700,fontSize:12,cursor:"pointer"}}>Enviar solicitud</button>
                  <button onClick={()=>setDelForm({id:null,reason:""})} style={{padding:"7px 12px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontSize:12,cursor:"pointer"}}>Cancelar</button>
                </div>
              </div>}
            </div>
          ))
        }
        <button onClick={()=>setView("form")} style={{width:"100%",padding:"12px",borderRadius:12,border:"2px dashed #e5e5e5",background:"#fafafa",color:"#888",fontWeight:700,fontSize:13,cursor:"pointer",marginTop:4}}>
          + Agregar nuevo cliente
        </button>
      </>}

      {/* FORM */}
      {view==="form" && (
        <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 1px 4px #0001"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14}}>{editing?"✏️ Editar cliente":"+ Nuevo cliente"}</div>
          <div style={{background:"#fdecea",borderRadius:8,padding:"7px 12px",fontSize:11,color:"#922b21",fontWeight:600,marginBottom:12}}>
            Los campos con <span style={{color:"#c0392b"}}>*</span> son obligatorios
          </div>
          <Field label="Nombre / Razón social *"><input value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} placeholder="Ej: Papelería El Centro" style={inputStyle}/></Field>
          <Field label="Teléfono *"><input value={form.phone} onChange={e=>setForm(f=>({...f,phone:e.target.value}))} placeholder="+54 11 1234-5678" type="tel" style={inputStyle}/></Field>
          <div style={{display:"flex",alignItems:"center",gap:10,margin:"14px 0 12px"}}>
            <div style={{flex:1,height:1,background:"#e5e5e5"}}/>
            <span style={{fontSize:10,color:"#aaa",fontWeight:700,letterSpacing:1}}>DATOS OPCIONALES</span>
            <div style={{flex:1,height:1,background:"#e5e5e5"}}/>
          </div>
          <Field label="CUIT / DNI"><input value={form.cuit} onChange={e=>setForm(f=>({...f,cuit:e.target.value}))} placeholder="Ej: 30-71234567-8" style={inputStyle}/></Field>
          <Field label="Email"><input type="email" value={form.email} onChange={e=>setForm(f=>({...f,email:e.target.value}))} placeholder="contacto@empresa.com" style={inputStyle}/></Field>
          <Field label="Dirección"><input value={form.address} onChange={e=>setForm(f=>({...f,address:e.target.value}))} placeholder="Ej: Av. Corrientes 1234, CABA" style={inputStyle}/></Field>
          <Field label="Notas internas"><textarea value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} placeholder="Condiciones especiales, contacto preferido..." style={{...inputStyle,resize:"vertical",minHeight:60}}/></Field>
          <div style={{display:"flex",gap:8,marginTop:8}}>
            <button onClick={save} disabled={saving} style={{flex:1,padding:"11px",borderRadius:10,border:"none",background:"#c0392b",color:"#fff",fontWeight:800,cursor:"pointer",fontSize:14}}>{saving?"Guardando...":editing?"💾 Guardar cambios":"✅ Crear cliente"}</button>
            <button onClick={cancelEdit} style={{padding:"11px 16px",borderRadius:10,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontWeight:600,color:"#666"}}>Cancelar</button>
          </div>
        </div>
      )}
    </div>
  );
}

function Precios({products}) {
  const [search,setSearch]=useState("");
  const [sortBy,setSortBy]=useState("name");
  const shown=useMemo(()=>{
    const q=norm(search);
    return products
      .filter(p=>!q||norm(p.name).includes(q)||normSKU(p.id).includes(normSKU(search)))
      .sort((a,b)=>sortBy==="name"?a.name.localeCompare(b.name):b.salePrice-a.salePrice)
      .slice(0,200);
  },[products,search,sortBy]);
  return (
    <div>
      <div style={{background:"#fff",borderRadius:12,padding:16,marginBottom:14,boxShadow:"0 1px 4px #0001"}}>
        <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar producto o código..."
            style={{flex:1,minWidth:160,padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none"}}/>
          <div style={{display:"flex",gap:6}}>
            {[{v:"name",l:"A-Z"},{v:"price",l:"Precio"}].map(opt=>(
              <button key={opt.v} onClick={()=>setSortBy(opt.v)} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid",cursor:"pointer",fontSize:12,fontWeight:600,borderColor:sortBy===opt.v?RED:"#e5e5e5",background:sortBy===opt.v?"#fdecea":"#fff",color:sortBy===opt.v?RED:"#666"}}>{opt.l}</button>
            ))}
          </div>
          <div style={{fontSize:12,color:"#aaa",whiteSpace:"nowrap"}}>{shown.length} producto{shown.length!==1?"s":""}{shown.length===200?" (máx 200)":""}</div>
        </div>
      </div>
      {shown.length===0
        ? <div style={{textAlign:"center",padding:60,color:"#aaa",background:"#fff",borderRadius:12}}><div style={{fontSize:48,marginBottom:8}}>🔍</div><div>No se encontraron productos</div></div>
        : <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px #0001",overflow:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr style={{background:"#f9f9f9",position:"sticky",top:0}}>{["Código","Descripción","Categoría","Precio"].map(h=>(<th key={h} style={{padding:"11px 14px",textAlign:"left",fontWeight:700,color:"#888",fontSize:11,textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>))}</tr></thead>
              <tbody>
                {shown.map(p=>(<tr key={p.id} style={{borderTop:"1px solid #f5f5f5"}} onMouseEnter={e=>e.currentTarget.style.background="#fafafa"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                  <td style={{padding:"10px 14px",color:"#999",fontSize:11,whiteSpace:"nowrap"}}>{p.id}</td>
                  <td style={{padding:"10px 14px",fontWeight:600,color:"#1a1a1a",maxWidth:320}}>{p.name}</td>
                  <td style={{padding:"10px 14px"}}><span style={{background:"#f5f5f5",color:"#666",borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:500}}>{p.category}</span></td>
                  <td style={{padding:"10px 14px",whiteSpace:"nowrap"}}><span style={{fontWeight:800,fontSize:15,color:RED}}>{fARS(p.salePrice)}</span></td>
                </tr>))}
              </tbody>
            </table>
          </div>
      }
    </div>
  );
}

// ─── NUEVO PEDIDO ─────────────────────────────────────────────────────────────
function ProductSelector({products,cart,setCart,isMobile}) {
  const [search,setSearch]=useState("");
  const [cat,setCat]=useState("todos");
  const [catOpen,setCatOpen]=useState(false);
  const [soloStock,setSoloStock]=useState(false);
  const CATS=useMemo(()=>["todos",...new Set(products.map(p=>p.category))].sort(),[products]);
  const shown=useMemo(()=>{
    const q=search.toLowerCase();
    return products.filter(p=>{
      if(cat!=="todos"&&p.category!==cat)return false;
      if(soloStock&&p.stock<=0)return false;
      if(q)return norm(p.name).includes(norm(q))||normSKU(p.id).includes(normSKU(q));
      return true;
    }).slice(0,80);
  },[products,search,cat,soloStock]);
  const addC=p=>setCart(c=>{const ex=c.find(i=>i.pid===p.id);return ex?c.map(i=>i.pid===p.id?{...i,qty:i.qty+1}:i):[...c,{pid:p.id,qty:1,price:p.salePrice,name:p.name}];});
  const setQ=(pid,qty)=>{if(qty<=0)setCart(c=>c.filter(i=>i.pid!==pid));else setCart(c=>c.map(i=>i.pid===pid?{...i,qty}:i));};
  return (
    <div>
      <div style={{background:"#fff",borderRadius:12,padding:16,marginBottom:12,boxShadow:"0 1px 4px #0001"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar por nombre o código..." style={{width:"100%",padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",marginBottom:10,boxSizing:"border-box"}}/>
        <div style={{position:"relative"}}>
          <button onClick={()=>setCatOpen(o=>!o)} style={{display:"flex",alignItems:"center",justifyContent:"space-between",width:"100%",padding:"8px 12px",borderRadius:8,border:`1.5px solid ${catOpen?RED:"#e5e5e5"}`,background:cat!=="todos"?"#fdecea":"#fff",color:cat!=="todos"?RED:"#666",cursor:"pointer",fontSize:13,fontWeight:600}}>
            <span>🏷️ {cat==="todos"?"Todas las categorías":cat}</span><span style={{fontSize:10,marginLeft:6}}>{catOpen?"▲":"▼"}</span>
          </button>
          {catOpen&&(<div style={{position:"absolute",top:"calc(100% + 4px)",left:0,right:0,background:"#fff",borderRadius:10,border:"1.5px solid #e5e5e5",boxShadow:"0 8px 24px #0002",zIndex:50,padding:8,display:"flex",flexWrap:"wrap",gap:5,maxHeight:220,overflowY:"auto"}}>
            {CATS.map(c=><button key={c} onClick={()=>{setCat(c);setCatOpen(false);}} style={{padding:"4px 11px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:11,fontWeight:600,borderColor:cat===c?RED:"#e5e5e5",background:cat===c?"#fdecea":"#fff",color:cat===c?RED:"#666"}}>{c==="todos"?"Todos":c}</button>)}
          </div>)}
        </div>
        {search&&<div style={{fontSize:11,color:"#aaa",marginTop:6}}>{shown.length} resultados</div>}
        {/* Filtro solo con stock */}
        <button onClick={()=>setSoloStock(s=>!s)}
          style={{marginTop:8,padding:"7px 14px",borderRadius:8,border:"1.5px solid",cursor:"pointer",fontSize:12,fontWeight:700,
            borderColor:soloStock?"#1e8449":"#e5e5e5",
            background:soloStock?"#eafaf1":"#fff",
            color:soloStock?"#1e8449":"#888",
            display:"flex",alignItems:"center",gap:6}}>
          📦 {soloStock?"Solo con stock ✓":"Mostrar solo con stock"}
        </button>
      </div>
      {isMobile
        ? <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {shown.map(p=>{const ic=cart.find(i=>i.pid===p.id);return (
              <div key={p.id} style={{background:"#fff",borderRadius:10,padding:"12px 14px",display:"flex",alignItems:"center",gap:10,boxShadow:"0 1px 4px #0001",border:ic?`2px solid ${RED}`:"2px solid transparent"}}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:700,fontSize:12,color:"#1a1a1a",lineHeight:1.3,marginBottom:2}}>{p.name}</div>
                  <div style={{fontSize:11,color:"#666"}}>{p.id} · {p.category}</div>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginTop:4}}>
                    <span style={{fontSize:15,fontWeight:800,color:RED}}>{fARS(p.salePrice)}</span>
                    <SPill n={p.stock}/>
                  </div>
                </div>
                {ic
                  ? <div style={{display:"flex",alignItems:"center",gap:6,flexShrink:0}}>
                      <button onClick={()=>setQ(p.id,ic.qty-1)} style={{width:30,height:30,borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:16,fontWeight:700}}>−</button>
                      <span style={{minWidth:24,textAlign:"center",fontWeight:800,fontSize:14}}>{ic.qty}</span>
                      <button onClick={()=>setQ(p.id,ic.qty+1)} style={{width:30,height:30,borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:16,fontWeight:700}}>+</button>
                    </div>
                  : <button onClick={()=>addC(p)} style={{padding:"7px 12px",borderRadius:7,border:"none",cursor:"pointer",background:RED,color:"#fff",fontWeight:700,fontSize:12,flexShrink:0}}>+ Agregar</button>
                }
              </div>
            );})}
          </div>
        : <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(195px,1fr))",gap:10}}>
            {shown.map(p=>{const ic=cart.find(i=>i.pid===p.id);return <div key={p.id} style={{background:"#fff",borderRadius:10,padding:14,border:ic?`2px solid ${RED}`:"2px solid transparent",boxShadow:"0 1px 4px #0001"}}>
              <div style={{fontWeight:700,fontSize:12,color:"#1a1a1a",marginBottom:3,lineHeight:1.3}}>{p.name}</div>
              <div style={{fontSize:12,color:"#666",marginBottom:7,fontWeight:500}}>{p.id} · {p.category}</div>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                <span style={{fontSize:17,fontWeight:800,color:RED}}>{fARS(p.salePrice)}</span><SPill n={p.stock}/>
              </div>
              {ic?<div style={{display:"flex",alignItems:"center",gap:5}}>
                <button onClick={()=>setQ(p.id,ic.qty-1)} style={{width:27,height:27,borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:15,fontWeight:700}}>−</button>
                <input type="number" value={ic.qty} onChange={e=>setQ(p.id,+e.target.value||0)} style={{width:40,textAlign:"center",padding:3,borderRadius:6,border:`1.5px solid ${RED}`,fontWeight:700,fontSize:13,outline:"none"}}/>
                <button onClick={()=>setQ(p.id,ic.qty+1)} style={{width:27,height:27,borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:15,fontWeight:700}}>+</button>
                <span style={{color:"#1e8449",fontSize:12,fontWeight:700}}>✓</span>
              </div>:<button onClick={()=>addC(p)} style={{width:"100%",padding:"7px",borderRadius:7,border:"none",cursor:"pointer",background:RED,color:"#fff",fontWeight:700,fontSize:12}}>+ Agregar</button>}
            </div>;})}
          </div>
      }
    </div>
  );
}

function Nuevo({products,vendors,onAdd,onDone,currentUser,isMobile,clients,onSaveClient}) {
  const [selectedClient, setSelectedClient] = useState(null);
  const [notes,setNotes]=useState("");
  const [vendedor,setVendedor]=useState(currentUser.role==="vendedor"?(currentUser.vendedor||currentUser.name):"");
  const [cart,setCart]=useState([]);
  const [globalDisc,setGlobalDisc]=useState({type:"%",value:""});
  const [ok,setOk]=useState(false);
  const [saving,setSaving]=useState(false);
  const [mStep,setMStep]=useState(1);

  const subtotal=cart.reduce((s,i)=>s+applyItemDiscount(i.price,i.qty,i.disc),0);
  const total=applyGlobalDiscount(subtotal,globalDisc);
  const globalDiscAmt=subtotal-total;

  const submit=async ()=>{
    if(!selectedClient){alert("Seleccioná un cliente");return;}
    if(!vendedor&&currentUser.role==="admin"){alert("Seleccioná un vendedor");return;}
    if(!cart.length){alert("Agregá productos");return;}
    setSaving(true);
    await onAdd({id:genId(),client:selectedClient.name,clientId:selectedClient.id,notes,vendedor:vendedor||currentUser.vendedor||currentUser.name,items:cart,total,subtotal,globalDisc,stage:"reserva",date:today()});
    setSaving(false);
    setOk(true); setTimeout(()=>onDone(),1400);
  };

  if(saving) return <SaveSpinner label="Registrando pedido..." color={RED}/>;
  if(ok) return <div style={{textAlign:"center",padding:80}}><div style={{fontSize:60}}>✅</div><div style={{fontWeight:800,color:"#1e8449",fontSize:20,marginTop:12}}>¡Pedido registrado!</div></div>;

  if(isMobile) {
    return (
      <div style={{display:"flex",flexDirection:"column",height:"100%"}}>
        {/* Steps header */}
        <div style={{display:"flex",background:"#fff",borderRadius:10,padding:"10px 14px",marginBottom:12,boxShadow:"0 1px 4px #0001",gap:4}}>
          {[{n:1,l:"Datos"},{n:2,l:"Productos"},{n:3,l:"Confirmar"}].map(s=>(
            <div key={s.n} onClick={()=>mStep>s.n&&setMStep(s.n)}
              style={{flex:1,textAlign:"center",padding:"6px 4px",borderRadius:8,background:mStep===s.n?"#fdecea":mStep>s.n?"#f9f9f9":"transparent",cursor:mStep>s.n?"pointer":"default"}}>
              <div style={{fontWeight:800,fontSize:13,color:mStep===s.n?RED:mStep>s.n?"#888":"#ccc"}}>{s.n}</div>
              <div style={{fontSize:10,color:mStep===s.n?RED:mStep>s.n?"#888":"#ccc",fontWeight:600}}>{s.l}</div>
            </div>
          ))}
        </div>

        {mStep===1 && (
          <div style={{background:"#fff",borderRadius:12,padding:16,boxShadow:"0 1px 4px #0001"}}>
            <div style={{fontWeight:800,fontSize:14,marginBottom:12}}>👥 Seleccioná el cliente</div>
            {selectedClient
              ? <div style={{background:"#fdecea",border:"1.5px solid #f5b7b1",borderRadius:10,padding:"12px 14px",marginBottom:12,display:"flex",alignItems:"center",gap:10}}>
                  <span style={{fontSize:24}}>✅</span>
                  <div style={{flex:1}}>
                    <div style={{fontWeight:700,fontSize:14,color:"#922b21"}}>{selectedClient.name}</div>
                    <div style={{fontSize:12,color:"#c0392b"}}>📱 {selectedClient.phone}</div>
                  </div>
                  <button onClick={()=>setSelectedClient(null)} style={{padding:"4px 10px",borderRadius:7,border:"1px solid #f5b7b1",background:"#fff",color:"#c0392b",fontSize:11,fontWeight:600,cursor:"pointer"}}>Cambiar</button>
                </div>
              : <ClientSelector clients={clients} onSelect={c=>{setSelectedClient(c);}} onSaveClient={onSaveClient} currentUser={currentUser}/>
            }
            {currentUser.role==="admin"&&<Field label="Vendedor *"><select value={vendedor} onChange={e=>setVendedor(e.target.value)} style={{...inputStyle,cursor:"pointer",color:vendedor?"#1a1a1a":"#aaa",marginBottom:8}}><option value="">— Seleccioná vendedor —</option>{vendors.map(v=><option key={v} value={v}>{v}</option>)}</select></Field>}
            <Field label="Notas"><textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Observaciones..." style={{...inputStyle,resize:"vertical",minHeight:55,fontSize:12}}/></Field>
            <button onClick={()=>{if(!selectedClient){alert("Seleccioná un cliente");return;}if(!vendedor&&currentUser.role==="admin"){alert("Seleccioná un vendedor");return;}setMStep(2);}}
              style={{width:"100%",padding:"12px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:selectedClient?`linear-gradient(135deg,${REDD},${RED})`:"#e5e5e5",color:selectedClient?"#fff":"#aaa",marginTop:8}}>
              Siguiente → Productos
            </button>
          </div>
        )}

        {mStep===2 && (
          <div style={{flex:1,overflow:"auto",paddingBottom:cart.length>0?72:0}}>
            <ProductSelector products={products} cart={cart} setCart={setCart} isMobile={true}/>
          </div>
        )}
        {mStep===2 && cart.length>0&&(
          <div style={{position:"fixed",bottom:0,left:0,right:0,background:`linear-gradient(135deg,${REDD},${RED})`,color:"#fff",padding:"13px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",zIndex:200,boxShadow:"0 -3px 16px #0003"}}>
            <div>
              <div style={{fontWeight:800,fontSize:15}}>{fARS(cart.reduce((s,i)=>s+i.price*i.qty,0))}</div>
              <div style={{fontSize:11,opacity:.85}}>{cart.length} producto{cart.length!==1?"s":""} seleccionado{cart.length!==1?"s":""}</div>
            </div>
            <button onClick={()=>setMStep(3)} style={{padding:"10px 20px",borderRadius:10,border:"none",background:"#fff",color:RED,fontWeight:800,fontSize:14,cursor:"pointer",boxShadow:"0 2px 8px #0002"}}>
              Ver resumen →
            </button>
          </div>
        )}

        {mStep===3 && (
          <div style={{background:"#fff",borderRadius:12,padding:16,boxShadow:"0 1px 4px #0001"}}>
            <div style={{fontWeight:800,fontSize:15,marginBottom:14,color:"#1a1a1a"}}>📋 Confirmar pedido</div>
            <div style={{fontSize:13,color:"#555",marginBottom:4}}>👤 <strong>{selectedClient?.name}</strong> · {vendedor}</div>
            <div style={{borderTop:"1px solid #f5f5f5",margin:"8px 0",paddingTop:8}}>
              {cart.map(i=><div key={i.pid} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"3px 0",color:"#555"}}><span style={{flex:1,marginRight:6}}>{i.name} × {i.qty}</span><span style={{fontWeight:600}}>{fARS(i.price*i.qty)}</span></div>)}
            </div>
            <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:17,color:RED,padding:"8px 0",borderTop:"2px solid #f5f5f5",marginBottom:14}}><span>Total</span><span>{fARS(total)}</span></div>
            <button onClick={submit} style={{width:"100%",padding:"12px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:`linear-gradient(135deg,${REDD},${RED})`,color:"#fff"}}>
              ✅ Registrar como Reserva
            </button>
            <button onClick={()=>setMStep(2)} style={{width:"100%",padding:"10px",borderRadius:10,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontWeight:600,fontSize:13,cursor:"pointer",marginTop:8}}>← Volver a productos</button>
          </div>
        )}
      </div>
    );
  }

  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 330px",gap:18,alignItems:"start"}}>
      <div>
        <div style={{fontWeight:800,fontSize:15,marginBottom:12}}>🛒 Nuevo Pedido — Seleccioná productos</div>
        <ProductSelector products={products} cart={cart} setCart={setCart}/>
      </div>
      <div style={{position:"sticky",top:16}}>
        <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 2px 12px #0002"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14}}>📋 Resumen del Pedido</div>
          {/* Client selector */}
          {selectedClient
            ? <div style={{background:"#fdecea",border:"1.5px solid #f5b7b1",borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",gap:10}}>
                <span style={{fontSize:20}}>✅</span>
                <div style={{flex:1}}>
                  <div style={{fontWeight:700,fontSize:13,color:"#922b21"}}>{selectedClient.name}</div>
                  <div style={{fontSize:11,color:"#c0392b"}}>📱 {selectedClient.phone}</div>
                </div>
                <button onClick={()=>setSelectedClient(null)} style={{padding:"3px 8px",borderRadius:6,border:"1px solid #f5b7b1",background:"#fff",color:"#c0392b",fontSize:11,cursor:"pointer"}}>Cambiar</button>
              </div>
            : <div style={{marginBottom:12}}>
                <div style={{fontSize:11,fontWeight:700,color:"#666",marginBottom:6}}>CLIENTE *</div>
                <ClientSelector clients={clients} onSelect={setSelectedClient} onSaveClient={onSaveClient} currentUser={currentUser}/>
              </div>
          }
          {currentUser.role==="admin"&&<Field label="Vendedor *"><select value={vendedor} onChange={e=>setVendedor(e.target.value)} style={{...inputStyle,color:vendedor?"#1a1a1a":"#aaa",cursor:"pointer"}}><option value="">— Seleccioná vendedor —</option>{vendors.map(v=><option key={v} value={v}>{v}</option>)}</select></Field>}
          <Field label="Notas"><textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Observaciones..." style={{...inputStyle,resize:"vertical",minHeight:55,fontSize:12}}/></Field>
          <div style={{borderTop:"1px solid #f5f5f5",margin:"4px 0 8px",paddingTop:10}}>
            {cart.length===0?<div style={{textAlign:"center",color:"#aaa",fontSize:12,padding:"10px 0"}}>Agregá productos al pedido</div>
            :cart.map(i=><div key={i.pid} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"3px 0",color:"#555"}}><span style={{flex:1,marginRight:6,lineHeight:1.3}}>{i.name} × {i.qty}</span><span style={{fontWeight:600,whiteSpace:"nowrap"}}>{fARS(i.price*i.qty)}</span></div>)}
          </div>
          <div style={{background:"#f9fdf9",border:"1.5px solid #e5e5e5",borderRadius:8,padding:"10px 12px",marginBottom:8}}>
            <div style={{fontSize:11,color:"#555",fontWeight:700,marginBottom:6}}>DESCUENTO GLOBAL</div>
            <div style={{display:"flex",gap:6,alignItems:"center"}}>
              <select value={globalDisc.type} onChange={e=>setGlobalDisc(d=>({...d,type:e.target.value}))}
                style={{padding:"5px 6px",borderRadius:6,border:"1.5px solid #e5e5e5",fontSize:13,fontWeight:700,background:"#fff",cursor:"pointer",width:48}}>
                <option value="%">%</option><option value="$">$</option>
              </select>
              <input type="number" min="0" value={globalDisc.value} onChange={e=>setGlobalDisc(d=>({...d,value:e.target.value}))}
                placeholder="0" style={{flex:1,padding:"5px 8px",borderRadius:6,border:"1.5px solid #ccc",fontSize:13,fontWeight:700,outline:"none",textAlign:"center"}}/>
              {globalDiscAmt>0&&<span style={{fontSize:11,color:"#1e8449",fontWeight:700,whiteSpace:"nowrap"}}>−{fARS(globalDiscAmt)}</span>}
            </div>
          </div>
          <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:17,color:RED,padding:"8px 0",borderTop:"2px solid #f5f5f5",marginBottom:14}}><span>Total</span><span>{fARS(total)}</span></div>
          <button onClick={submit} disabled={!cart.length||!selectedClient||(!vendedor&&currentUser.role==="admin")} style={{width:"100%",padding:"11px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:(!cart.length||!selectedClient||(!vendedor&&currentUser.role==="admin"))?"#e5e5e5":`linear-gradient(135deg,${REDD},${RED})`,color:(!cart.length||!selectedClient||(!vendedor&&currentUser.role==="admin"))?"#aaa":"#fff"}}>
            ✅ Registrar como Reserva
          </button>
        </div>
      </div>
    </div>
  );
}

function Cotizaciones({quotes,products,vendors,onAdd,onDel,onConvert,onExtend,onTabChange,currentUser,isMobile,clients,onSaveClient}) {
  const [view,setView]=useState("lista");
  const [expanded,setExpanded]=useState(null);
  const getP=id=>products.find(p=>p.id===id);
  return (
    <div>
      <div style={{background:"#fff",borderRadius:12,padding:4,marginBottom:16,display:"flex",gap:4,boxShadow:"0 1px 4px #0001"}}>
        <button onClick={()=>setView("lista")} style={{flex:1,padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="lista"?`linear-gradient(135deg,${REDD},${RED})`:"transparent",color:view==="lista"?"#fff":"#555"}}>📄 Lista de Cotizaciones ({quotes.filter(q=>!q.convertida).length})</button>
        <button onClick={()=>setView("nueva")} style={{flex:1,padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="nueva"?`linear-gradient(135deg,${REDD},${RED})`:"transparent",color:view==="nueva"?"#fff":"#555"}}>➕ Nueva Cotización</button>
      </div>
      {view==="nueva" && <NuevaCotizacion products={products} vendors={vendors} onAdd={async(q)=>{await onAdd(q);setView("lista");}} currentUser={currentUser} isMobile={isMobile} clients={clients} onSaveClient={onSaveClient}/>}
      {view==="lista" && (quotes.length===0
        ? <div style={{textAlign:"center",padding:60,color:"#aaa"}}><div style={{fontSize:48}}>📄</div><div style={{marginTop:8}}>No hay cotizaciones aún</div></div>
        : quotes.map(q=><QuoteCard key={q.id} q={q} exp={expanded===q.id} toggle={()=>setExpanded(expanded===q.id?null:q.id)} getP={getP} onDel={onDel} onConvert={async(qt)=>{await onConvert(qt);onTabChange("central");}} onExtend={onExtend}/>)
      )}
    </div>
  );
}


function QuoteCard({q,exp,toggle,getP,onDel,onConvert,onExtend}) {
  const PURPLE = "#6c3483"; const PURPLEBG = "#e8daef";
  const [showExtForm, setShowExtForm] = useState(false);
  const [extReason, setExtReason] = useState("");
  const [saving, setSaving] = useState(false);

  const qs = quoteStatus(q);
  const status    = typeof qs === "object" ? qs.status : qs;
  const hoursLeft = typeof qs === "object" ? qs.hoursLeft : null;
  const isVencida = status === "vencida";
  const canConvert = !isVencida && !q.convertida;

  const handleExtend = async () => {
    if(!extReason.trim()) return;
    setSaving(true);
    await onExtend(q.id, extReason.trim());
    setShowExtForm(false); setExtReason(""); setSaving(false);
  };

  const StatusBdg = () => {
    if(q.convertida) return <span style={{background:"#d5f5e3",color:"#1e8449",border:"1px solid #1e844944",borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>✅ Convertida</span>;
    if(status==="vencida")   return <span style={{background:"#fdecea",color:"#c0392b",border:"1px solid #f1948a",borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>⏰ Vencida</span>;
    if(status==="extendida") return <span style={{background:"#fef9e7",color:"#b7770d",border:"1px solid #f0d080",borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>🔄 Extendida</span>;
    return <span style={{background:"#d5f5e3",color:"#1e8449",border:"1px solid #a9dfbf",borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>✅ Vigente</span>;
  };

  // Color del borde lateral según estado
  const borderColor = q.convertida ? "#1e8449" : isVencida ? "#c0392b" : status==="extendida" ? "#b7770d" : "#1e8449";

  return (
    <div style={{background:"#fff",borderRadius:12,boxShadow:"0 2px 8px #0000000a",overflow:"hidden",marginBottom:8,borderLeft:`4px solid ${borderColor}`}}>
      <div onClick={toggle} style={{padding:"13px 18px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",cursor:"pointer"}}>
        <div style={{flex:1,minWidth:100}}>
          <div style={{fontWeight:700,fontSize:14,color:"#1a1a1a"}}>{q.client}</div>
          <div style={{fontSize:11,color:"#aaa",display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
            {q.isTest&&<span style={{background:"#f1c40f",color:"#1a1a1a",borderRadius:4,padding:"1px 6px",fontSize:10,fontWeight:800}}>TEST</span>}
            {q.docNum&&!q.isTest&&<span style={{fontWeight:700,color:PURPLE,fontFamily:SERIF}}>{q.docNum}</span>}
            <span>{q.date}</span>
            {q.vendedor&&<span>· 👤 {q.vendedor}</span>}
            <StatusBdg/>
            {hoursLeft>0&&<span style={{background:"#eaf4fc",color:"#1a5276",border:"1px solid #aed6f1",borderRadius:20,padding:"2px 8px",fontSize:10,fontWeight:700}}>⏳ {hoursLeft}hs</span>}
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
          <span style={{fontWeight:800,color:PURPLE,fontSize:15}}>{fARS(q.total)}</span>
          <span style={{color:"#ccc"}}>{exp?"▲":"▼"}</span>
        </div>
      </div>
      {exp && (
        <div style={{borderTop:"1px solid #f5f5f5",padding:18}}>
          {/* VENCIDA sin extensión */}
          {isVencida && !showExtForm && !q.extendida && (
            <div style={{background:"#fdecea",border:"1.5px solid #f1948a",borderRadius:10,padding:"12px 14px",marginBottom:14}}>
              <div style={{fontWeight:800,fontSize:13,color:"#c0392b",marginBottom:4}}>⏰ Cotización vencida</div>
              <div style={{fontSize:12,color:"#922b21",marginBottom:10}}>Las cotizaciones vencen a las 48 hs. Podés extenderla 48 hs más por única vez.</div>
              <button onClick={(e)=>{e.stopPropagation();setShowExtForm(true);}}
                style={{padding:"8px 16px",borderRadius:8,border:"none",cursor:"pointer",background:"linear-gradient(135deg,#e67e22,#d35400)",color:"#fff",fontWeight:700,fontSize:13}}>
                🔄 Extender 48 horas
              </button>
            </div>
          )}

          {/* FORMULARIO EXTENSIÓN */}
          {showExtForm && (
            <div style={{background:"#fffbf0",border:"1.5px solid #f0d080",borderRadius:10,padding:"12px 14px",marginBottom:14}} onClick={e=>e.stopPropagation()}>
              <div style={{fontSize:12,fontWeight:700,color:"#b7770d",marginBottom:6}}>📝 Razón de la extensión (obligatorio)</div>
              <textarea value={extReason} onChange={e=>setExtReason(e.target.value)}
                placeholder="Ej: Cliente solicitó más tiempo para confirmar con su jefe..."
                style={{width:"100%",padding:"8px 10px",borderRadius:8,border:`1.5px solid ${extReason.trim()?"#f0d080":"#f1948a"}`,fontSize:13,resize:"vertical",minHeight:70,outline:"none",boxSizing:"border-box",fontFamily:"inherit"}}/>
              <div style={{display:"flex",gap:8,marginTop:8}}>
                <button onClick={handleExtend} disabled={!extReason.trim()||saving}
                  style={{padding:"7px 14px",borderRadius:7,border:"none",background:extReason.trim()?"#b7770d":"#e5e5e5",color:extReason.trim()?"#fff":"#aaa",fontWeight:700,fontSize:12,cursor:extReason.trim()?"pointer":"not-allowed"}}>
                  {saving?"Guardando...":"✅ Confirmar extensión"}
                </button>
                <button onClick={(e)=>{e.stopPropagation();setShowExtForm(false);setExtReason("");}}
                  style={{padding:"7px 12px",borderRadius:7,border:"1px solid #e5e5e5",background:"#fff",color:"#666",fontSize:12,cursor:"pointer"}}>
                  Cancelar
                </button>
              </div>
            </div>
          )}

          {/* INFO EXTENSIÓN USADA */}
          {q.extendida && (
            <div style={{background:"#fffbf0",border:"1.5px solid #f0d080",borderRadius:10,padding:"10px 14px",marginBottom:14,fontSize:12,color:"#7d6608"}}>
              🔄 <strong>Extensión aplicada</strong> {q.extendDate&&`(${q.extendDate})`}<br/>
              Motivo: "{q.extendReason}"<br/>
              <span style={{color:"#aaa",fontSize:11}}>Extensión ya usada — no se puede extender de nuevo</span>
            </div>
          )}

          {/* VIGENTE */}
          {status==="vigente" && !q.convertida && hoursLeft>0 && (
            <div style={{background:"#eafaf1",border:"1.5px solid #a9dfbf",borderRadius:10,padding:"9px 14px",marginBottom:14,fontSize:12,color:"#1e8449",fontWeight:600}}>
              ✅ Vigente — {hoursLeft} hs restantes
            </div>
          )}

          {q.validity&&<div style={{background:"#fef9e7",borderRadius:8,padding:"7px 12px",fontSize:12,color:"#7d6608",marginBottom:12}}>⏳ Validez: <strong>{q.validity}</strong></div>}
          {q.items.map((it,i)=>{const p=getP(it.pid);return <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #f9f9f9",fontSize:13}}><span style={{color:"#444"}}>{p?.name||it.name} x {it.qty}</span><span style={{fontWeight:600}}>{fARS(it.price*it.qty)}</span></div>;})}
          <div style={{display:"flex",justifyContent:"flex-end",fontWeight:800,fontSize:16,color:PURPLE,margin:"8px 0 12px"}}>{fARS(q.total)}</div>
          {q.notes&&<div style={{background:"#f9f9f9",borderRadius:8,padding:"8px 12px",fontSize:13,color:"#555",marginBottom:12}}>💬 {q.notes}</div>}
          <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
            <button onClick={()=>printDoc(q,"cotizacion")} style={{padding:"8px 12px",borderRadius:8,border:`1.5px solid ${PURPLEBG}`,cursor:"pointer",background:"#fff",color:PURPLE,fontWeight:600,fontSize:13}}>🖨️ Imprimir</button>
            <button onClick={()=>window.open(`https://wa.me/?text=${encodeURIComponent(buildWAQuote(q))}`, "_blank")}
              style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #25D366",cursor:"pointer",background:"#25D366",color:"#fff",fontWeight:700,fontSize:13,display:"flex",alignItems:"center",gap:5}}>
              <span style={{fontSize:15}}>💬</span> Enviar por WhatsApp
            </button>
            {!q.convertida
              ? <button onClick={()=>onConvert(q)} disabled={isVencida}
                  style={{padding:"8px 14px",borderRadius:8,border:"none",cursor:isVencida?"not-allowed":"pointer",background:isVencida?"#e5e5e5":"linear-gradient(135deg,#922b21,#c0392b)",color:isVencida?"#aaa":"#fff",fontWeight:700,fontSize:13}}>
                  🛒 Pasar a Reserva
                </button>
              : <span style={{background:"#d5f5e3",color:"#1e8449",borderRadius:8,padding:"6px 12px",fontSize:12,fontWeight:700}}>✅ Convertida a Reserva</span>
            }
            <div style={{marginLeft:"auto",display:"flex",gap:8}}>
              <QuoteDelBtn onConfirm={()=>onDel(q.id)}/>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function QuoteDelBtn({onConfirm}) {
  const [c,setC]=useState(false);
  if(c) return <div style={{display:"flex",alignItems:"center",gap:6,background:"#fdecea",borderRadius:8,padding:"6px 10px",border:"1.5px solid #fcc"}}>
    <span style={{fontSize:12,color:RED,fontWeight:600}}>?Eliminar?</span>
    <button onClick={onConfirm} style={{padding:"4px 10px",borderRadius:6,border:"none",background:RED,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:12}}>Sí</button>
    <button onClick={()=>setC(false)} style={{padding:"4px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12}}>No</button>
  </div>;
  return <button onClick={()=>setC(true)} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #fcc",cursor:"pointer",background:"#fff",color:RED,fontWeight:600,fontSize:13}}>🗑 Eliminar</button>;
}

function NuevaCotizacion({products,vendors,onAdd,currentUser,isMobile,clients,onSaveClient}) {
  const PURPLE="#6c3483"; const PURPLEG="linear-gradient(135deg,#6c3483,#9b59b6)";
  const [selectedClient,setSelectedClient]=useState(null);
  const [notes,setNotes]=useState("");
  const [vendedor,setVendedor]=useState(currentUser?.role!=="admin" ? (currentUser?.vendedor||currentUser?.name||"") : "");
  const [validity,setValidity]=useState("48 horas");
  const [cart,setCart]=useState([]);
  const [globalDisc,setGlobalDisc]=useState({type:"%",value:""});
  const [saving,setSaving]=useState(false);
  const [ok,setOk]=useState(false);
  const [mStep,setMStep]=useState(1);

  const subtotal = cart.reduce((s,i)=>s+applyItemDiscount(i.price,i.qty,i.disc),0);
  const total    = applyGlobalDiscount(subtotal, globalDisc);
  const globalDiscAmt = subtotal - total;

  const submit=async()=>{
    if(!selectedClient){alert("Seleccioná un cliente");return;}
    if(!cart.length){alert("Agregá productos");return;}
    setSaving(true);
    const q={id:genId(),client:selectedClient.name,clientId:selectedClient.id,notes,vendedor,validity,items:cart,total,subtotal,globalDisc,date:today()};
    await onAdd(q);
    setSaving(false);
    setOk(true);
  };

  if(saving) return <SaveSpinner label="Guardando cotización..." color={PURPLE}/>;
  if(ok) return <div style={{textAlign:"center",padding:80}}><div style={{fontSize:60}}>📄</div><div style={{fontWeight:800,color:PURPLE,fontSize:20,marginTop:12}}>¡Cotización guardada!</div></div>;

  // ── MOBILE — flujo 3 pasos ──
  if(isMobile) return (
    <div style={{display:"flex",flexDirection:"column",minHeight:"100%"}}>
      <div style={{display:"flex",background:"#fff",borderRadius:10,padding:"10px 14px",marginBottom:12,boxShadow:"0 1px 4px #0001",gap:4}}>
        {[{n:1,l:"Cliente"},{n:2,l:"Productos"},{n:3,l:"Confirmar"}].map(s=>(
          <div key={s.n} onClick={()=>mStep>s.n&&setMStep(s.n)}
            style={{flex:1,textAlign:"center",padding:"6px 4px",borderRadius:8,
              background:mStep===s.n?"#f5eef8":mStep>s.n?"#f9f9f9":"transparent",
              cursor:mStep>s.n?"pointer":"default"}}>
            <div style={{fontWeight:800,fontSize:13,color:mStep===s.n?PURPLE:mStep>s.n?"#888":"#ccc"}}>{s.n}</div>
            <div style={{fontSize:10,color:mStep===s.n?PURPLE:mStep>s.n?"#888":"#ccc",fontWeight:600}}>{s.l}</div>
          </div>
        ))}
      </div>

      {/* Paso 1 — Cliente */}
      {mStep===1 && (
        <div style={{background:"#fff",borderRadius:12,padding:16,boxShadow:"0 1px 4px #0001"}}>
          <div style={{background:"#f5eef8",borderRadius:8,padding:"8px 12px",fontSize:12,color:PURPLE,marginBottom:12}}>
            ℹ️ Las cotizaciones <strong>no descuentan stock</strong>. Son solo presupuestos.
          </div>
          <div style={{fontWeight:800,fontSize:14,marginBottom:12}}>👥 Seleccioná el cliente</div>
          {selectedClient
            ? <div style={{background:"#f5eef8",border:"1.5px solid #d7bde2",borderRadius:10,padding:"12px 14px",marginBottom:12,display:"flex",alignItems:"center",gap:10}}>
                <span style={{fontSize:24}}>✅</span>
                <div style={{flex:1}}>
                  <div style={{fontWeight:700,fontSize:14,color:PURPLE}}>{selectedClient.name}</div>
                  <div style={{fontSize:12,color:"#888"}}>📱 {selectedClient.phone}</div>
                </div>
                <button onClick={()=>setSelectedClient(null)} style={{padding:"4px 10px",borderRadius:7,border:"1px solid #d7bde2",background:"#fff",color:PURPLE,fontSize:11,fontWeight:600,cursor:"pointer"}}>Cambiar</button>
              </div>
            : <ClientSelector clients={clients||[]} onSelect={setSelectedClient} onSaveClient={onSaveClient} currentUser={currentUser}/>
          }
          {currentUser?.role==="admin"&&<Field label="Vendedor"><select value={vendedor} onChange={e=>setVendedor(e.target.value)} style={{...inputStyle,cursor:"pointer"}}><option value="">- Sin asignar -</option>{vendors.map(v=><option key={v} value={v}>{v}</option>)}</select></Field>}
          <Field label="Válida hasta"><input value={validity} onChange={e=>setValidity(e.target.value)} placeholder="Ej: 48 horas" style={inputStyle}/></Field>
          <Field label="Notas"><textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Observaciones, condiciones..." style={{...inputStyle,resize:"vertical",minHeight:55,fontSize:12}}/></Field>
          <button onClick={()=>{if(!selectedClient){alert("Seleccioná un cliente");return;}setMStep(2);}}
            style={{width:"100%",padding:"12px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:selectedClient?PURPLEG:"#e5e5e5",color:selectedClient?"#fff":"#aaa",marginTop:8}}>
            Siguiente → Productos
          </button>
        </div>
      )}

      {/* Paso 2 — Productos */}
      {mStep===2 && (
        <div style={{flex:1,overflow:"auto",paddingBottom:cart.length>0?72:0}}>
          <ProductSelector products={products} cart={cart} setCart={setCart} isMobile={true}/>
        </div>
      )}
      {mStep===2 && cart.length>0 && (
        <div style={{position:"fixed",bottom:0,left:0,right:0,background:PURPLEG,color:"#fff",padding:"13px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",zIndex:200,boxShadow:"0 -3px 16px #0003"}}>
          <div>
            <div style={{fontWeight:800,fontSize:15}}>{fARS(total)}</div>
            <div style={{fontSize:11,opacity:.85}}>{cart.length} producto{cart.length!==1?"s":""} seleccionado{cart.length!==1?"s":""}</div>
          </div>
          <button onClick={()=>setMStep(3)} style={{padding:"10px 20px",borderRadius:10,border:"none",background:"#fff",color:PURPLE,fontWeight:800,fontSize:14,cursor:"pointer",boxShadow:"0 2px 8px #0002"}}>
            Ver resumen →
          </button>
        </div>
      )}

      {/* Paso 3 — Confirmar */}
      {mStep===3 && (
        <div style={{background:"#fff",borderRadius:12,padding:16,boxShadow:"0 1px 4px #0001"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14,color:PURPLE}}>📄 Confirmar cotización</div>
          <div style={{fontSize:13,color:"#555",marginBottom:8}}>👤 <strong>{selectedClient?.name}</strong> · {vendedor} · {validity}</div>
          <div style={{borderTop:"1px solid #f5f5f5",paddingTop:8,marginBottom:8}}>
            {cart.map(i=>{
              const lineTotal=applyItemDiscount(i.price,i.qty,i.disc);
              const hasD=parseFloat(i.disc?.value)>0;
              return <div key={i.pid} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"3px 0",color:"#555"}}>
                <span style={{flex:1,marginRight:6}}>{i.name} × {i.qty}</span>
                <span style={{fontWeight:600,color:hasD?PURPLE:undefined}}>{fARS(lineTotal)}</span>
              </div>;
            })}
          </div>
          {/* Descuento global mobile */}
          <div style={{background:"#f5eef8",border:"1.5px solid #e8daef",borderRadius:8,padding:"10px 12px",marginBottom:8}}>
            <div style={{fontSize:11,color:PURPLE,fontWeight:700,marginBottom:6}}>DESCUENTO GLOBAL</div>
            <div style={{display:"flex",gap:6,alignItems:"center"}}>
              <select value={globalDisc.type} onChange={e=>setGlobalDisc(d=>({...d,type:e.target.value}))}
                style={{padding:"5px 6px",borderRadius:6,border:"1.5px solid #e5e5e5",fontSize:13,fontWeight:700,background:"#fff",cursor:"pointer",width:48}}>
                <option value="%">%</option><option value="$">$</option>
              </select>
              <input type="number" min="0" value={globalDisc.value} onChange={e=>setGlobalDisc(d=>({...d,value:e.target.value}))}
                placeholder="0" style={{flex:1,padding:"5px 8px",borderRadius:6,border:"1.5px solid #ccc",fontSize:13,fontWeight:700,outline:"none",textAlign:"center"}}/>
              {globalDiscAmt>0&&<span style={{fontSize:11,color:PURPLE,fontWeight:700,whiteSpace:"nowrap"}}>-{fARS(globalDiscAmt)}</span>}
            </div>
          </div>
          <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:17,color:PURPLE,padding:"8px 0",borderTop:"2px solid #f5f5f5",marginBottom:14}}>
            <span>Total</span><span>{fARS(total)}</span>
          </div>
          <button onClick={submit} style={{width:"100%",padding:"12px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:PURPLEG,color:"#fff"}}>
            📄 Guardar Cotización
          </button>
          <button onClick={()=>setMStep(2)} style={{width:"100%",padding:"10px",borderRadius:10,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontWeight:600,fontSize:13,cursor:"pointer",marginTop:8}}>
            ← Volver a productos
          </button>
        </div>
      )}
    </div>
  );

  // ── DESKTOP ──
  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 330px",gap:18,alignItems:"start"}}>
      <div>
        <div style={{fontWeight:800,fontSize:15,marginBottom:12}}>📄 Nueva Cotización - Seleccioná productos</div>
        <ProductSelector products={products} cart={cart} setCart={setCart}/>
      </div>
      <div style={{position:"sticky",top:16}}>
        <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 2px 12px #0002",border:"2px solid #e8daef"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14,color:PURPLE}}>📄 Resumen de Cotización</div>
          <div style={{background:"#e8daef",borderRadius:8,padding:"7px 12px",fontSize:12,color:PURPLE,marginBottom:14}}>
            ℹ️ Las cotizaciones <strong>no descuentan stock</strong>. Son solo presupuestos para el cliente.
          </div>
          {selectedClient
            ? <div style={{background:"#f5eef8",border:"1.5px solid #d7bde2",borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",gap:10}}>
                <span style={{fontSize:20}}>✅</span>
                <div style={{flex:1}}><div style={{fontWeight:700,fontSize:13,color:PURPLE}}>{selectedClient.name}</div><div style={{fontSize:11,color:"#888"}}>📱 {selectedClient.phone}</div></div>
                <button onClick={()=>setSelectedClient(null)} style={{padding:"3px 8px",borderRadius:6,border:"1px solid #d7bde2",background:"#fff",color:PURPLE,fontSize:11,cursor:"pointer"}}>Cambiar</button>
              </div>
            : <div style={{marginBottom:12}}><div style={{fontSize:11,fontWeight:700,color:"#666",marginBottom:6}}>CLIENTE *</div>
                <ClientSelector clients={clients||[]} onSelect={setSelectedClient} onSaveClient={onSaveClient} currentUser={currentUser}/>
              </div>
          }
          {currentUser?.role==="admin"
            ? <Field label="Vendedor">
                <select value={vendedor} onChange={e=>setVendedor(e.target.value)} style={{...inputStyle,color:vendedor?"#1a1a1a":"#aaa",cursor:"pointer"}}>
                  <option value="">- Sin asignar -</option>
                  {vendors.map(v=><option key={v} value={v}>{v}</option>)}
                </select>
              </Field>
            : <div style={{marginBottom:12,background:"#f5f5f5",borderRadius:8,padding:"9px 12px",fontSize:13,color:"#555"}}>
                <span style={{fontSize:11,fontWeight:700,color:"#888",display:"block",marginBottom:2}}>VENDEDOR</span>
                <span style={{fontWeight:700}}>{vendedor||currentUser?.name}</span>
              </div>
          }
          <Field label="Válida hasta"><input value={validity} onChange={e=>setValidity(e.target.value)} placeholder="Ej: 48 horas" style={inputStyle}/></Field>
          <Field label="Notas"><textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Observaciones, condiciones..." style={{...inputStyle,resize:"vertical",minHeight:55,fontSize:12}}/></Field>
          <div style={{borderTop:"1px solid #f5f5f5",margin:"4px 0 8px",paddingTop:10}}>
            {cart.length===0?<div style={{textAlign:"center",color:"#aaa",fontSize:12,padding:"10px 0"}}>Agregá productos a la cotización</div>
            :cart.map(i=>{
              const lineTotal=applyItemDiscount(i.price,i.qty,i.disc);
              const hasD=parseFloat(i.disc?.value)>0;
              return <div key={i.pid} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"4px 0",borderBottom:"1px solid #f9f9f9",color:"#555",gap:6}}>
                <span style={{flex:1,lineHeight:1.3}}>{i.name} x {i.qty}</span>
                <div style={{textAlign:"right",whiteSpace:"nowrap"}}>
                  {hasD&&<div style={{fontSize:10,color:"#aaa",textDecoration:"line-through"}}>{fARS(i.price*i.qty)}</div>}
                  <span style={{fontWeight:600,color:hasD?PURPLE:undefined}}>{fARS(lineTotal)}</span>
                  {hasD&&<span style={{fontSize:10,color:PURPLE,marginLeft:3}}>{fmtDisc(i.disc)}</span>}
                </div>
              </div>;
            })}
          </div>
          <div style={{background:"#f5f0fa",border:"1.5px solid #e8daef",borderRadius:8,padding:"10px 12px",marginBottom:8}}>
            <div style={{fontSize:11,color:PURPLE,fontWeight:700,marginBottom:6}}>DESCUENTO GLOBAL</div>
            <div style={{display:"flex",gap:6,alignItems:"center"}}>
              <select value={globalDisc.type} onChange={e=>setGlobalDisc(d=>({...d,type:e.target.value}))}
                style={{padding:"5px 6px",borderRadius:6,border:"1.5px solid #e5e5e5",fontSize:13,fontWeight:700,background:"#fff",cursor:"pointer",width:48}}>
                <option value="%">%</option><option value="$">$</option>
              </select>
              <input type="number" min="0" value={globalDisc.value} onChange={e=>setGlobalDisc(d=>({...d,value:e.target.value}))}
                placeholder="0" style={{flex:1,padding:"5px 8px",borderRadius:6,border:"1.5px solid #ccc",fontSize:13,fontWeight:700,outline:"none",textAlign:"center"}}/>
              {globalDiscAmt>0&&<span style={{fontSize:11,color:PURPLE,fontWeight:700,whiteSpace:"nowrap"}}>-{fARS(globalDiscAmt)}</span>}
            </div>
          </div>
          {globalDiscAmt>0&&<div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:"#888",padding:"2px 0"}}>
            <span>Subtotal</span><span>{fARS(subtotal)}</span>
          </div>}
          <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:17,color:PURPLE,padding:"8px 0",borderTop:"2px solid #f5f5f5",margin:"6px 0 14px"}}><span>Total</span><span>{fARS(total)}</span></div>
          <button onClick={submit} disabled={!cart.length||!selectedClient||saving} style={{width:"100%",padding:"11px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:(!cart.length||!selectedClient)?"#e5e5e5":PURPLEG,color:(!cart.length||!selectedClient)?"#aaa":"#fff"}}>
            {saving?"Guardando...":"📄 Guardar Cotización"}
          </button>
        </div>
      </div>
    </div>
  );
}


// ─── STOCK ────────────────────────────────────────────────────────────────────
function StockAlert({low}) {
  const [open, setOpen] = useState(false);
  return (
    <div style={{background:"#fef9e7",border:"1px solid #f1c40f",borderRadius:12,marginBottom:14,overflow:"hidden"}}>
      <div onClick={()=>setOpen(o=>!o)} style={{padding:"11px 16px",display:"flex",gap:10,alignItems:"center",cursor:"pointer",userSelect:"none"}}>
        <span style={{fontSize:20}}>🎉️</span>
        <div style={{flex:1}}>
          <span style={{fontWeight:700,color:"#7d6608",fontSize:13}}>Stock bajo en {low.length} producto(s)</span>
          {!open && <span style={{fontSize:11,color:"#9a7d0a",marginLeft:8}}>- hacé clic para ver el detalle</span>}
        </div>
        <span style={{fontSize:12,color:"#9a7d0a",fontWeight:700}}>{open?"^":"v"}</span>
      </div>
      {open && (
        <div style={{borderTop:"1px solid #f1c40f22",padding:"4px 16px 14px"}}>
          {low.map(p=>(
            <div key={p.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:"1px solid #f1c40f22"}}>
              <div>
                <span style={{fontWeight:600,fontSize:13,color:"#555"}}>{p.name}</span>
                <span style={{fontSize:11,color:"#aaa",marginLeft:8}}>{p.id}</span>
              </div>
              <span style={{background:p.stock===0?"#fdecea":"#fff3cd",color:p.stock===0?RED:"#856404",borderRadius:8,padding:"2px 8px",fontSize:12,fontWeight:700,border:`1px solid ${p.stock===0?"#f5c6cb":"#ffc107"}`}}>
                {p.stock===0?"Sin stock":`🎉 ${p.stock} u.`}
              </span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function StockLog({log, onClear}) {
  const [filter,setFilter] = useState("todos");
  const [search,setSearch] = useState("");

  const TIPOS = {
    "ENTRADA": {color:"#1e8449", bg:"#d5f5e3", icon:"📥"},
    "SALIDA":  {color:"#e67e22", bg:"#fef9e7", icon:"📤"},
    "BAJA":    {color:"#c0392b", bg:"#fdecea", icon:"🗑"},
  };

  const filtered = log.filter(e=>{
    if(filter!=="todos" && e.tipo!==filter) return false;
    if(search){
      const q=search.toLowerCase();
      return norm(e.producto).includes(norm(q))||norm(e.usuario).includes(norm(q))||norm(e.motivo).includes(norm(q))||e.productoId?.toLowerCase().includes(q);
    }
    return true;
  });

  const [confirmClear, setConfirmClear] = useState(false);

  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:10,marginBottom:14}}>
        {Object.entries(TIPOS).map(([tipo,cfg])=>{
          const cnt = log.filter(e=>e.tipo===tipo).length;
          return <div key={tipo} style={{background:"#fff",borderRadius:10,padding:"12px 14px",boxShadow:"0 1px 4px #0001",borderLeft:`4px solid ${cfg.color}`}}>
            <div style={{fontSize:22,fontWeight:800,color:cfg.color}}>{cnt}</div>
            <div style={{fontSize:12,color:"#666",fontWeight:600}}>{cfg.icon} {tipo === "ENTRADA" ? "Entradas" : tipo === "SALIDA" ? "Salidas" : "Bajas"}</div>
          </div>;
        })}
        <div style={{background:"#fff",borderRadius:10,padding:"12px 14px",boxShadow:"0 1px 4px #0001",borderLeft:"4px solid #1a5276"}}>
          <div style={{fontSize:22,fontWeight:800,color:"#1a5276"}}>{log.length}</div>
          <div style={{fontSize:12,color:"#666",fontWeight:600}}>📜 Total movimientos</div>
        </div>
      </div>
      <div style={{background:"#fff",borderRadius:12,padding:12,marginBottom:12,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center",boxShadow:"0 1px 4px #0001"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar por producto, usuario o motivo..."
          style={{flex:1,minWidth:200,padding:"7px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none"}}/>
        <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
          {["todos","ENTRADA","SALIDA","BAJA"].map(t=>{
            const cfg = TIPOS[t];
            return <button key={t} onClick={()=>setFilter(t)} style={{padding:"5px 12px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:11,fontWeight:600,borderColor:filter===t?(cfg?.color||RED):"#e5e5e5",background:filter===t?(cfg?.bg||"#fdecea"):"#fff",color:filter===t?(cfg?.color||RED):"#666"}}>
              {t==="todos"?"Todos":`${cfg.icon} ${t}`}
            </button>;
          })}
        </div>
        {log.length>0&&(
          confirmClear
            ? <div style={{display:"flex",gap:6,alignItems:"center",background:"#fdecea",borderRadius:8,padding:"5px 10px",border:`1px solid ${RED}44`}}>
                <span style={{fontSize:12,color:RED,fontWeight:600}}>?Limpiar todo el historial?</span>
                <button onClick={()=>{onClear();setConfirmClear(false);}} style={{padding:"3px 10px",borderRadius:6,border:"none",background:RED,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:12}}>Sí</button>
                <button onClick={()=>setConfirmClear(false)} style={{padding:"3px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12}}>No</button>
              </div>
            : <button onClick={()=>setConfirmClear(true)} style={{padding:"5px 12px",borderRadius:8,border:"1.5px solid #fcc",background:"#fff",color:RED,cursor:"pointer",fontSize:11,fontWeight:600,whiteSpace:"nowrap"}}>🗑 Limpiar historial</button>
        )}
      </div>
      {filtered.length===0
        ? <div style={{textAlign:"center",padding:50,color:"#aaa",background:"#fff",borderRadius:12}}>
            <div style={{fontSize:40,marginBottom:8}}>📭</div>
            <div>{log.length===0?"No hay movimientos registrados aún.":"No hay movimientos que coincidan con el filtro."}</div>
          </div>
        : <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px #0001",overflow:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead>
                <tr style={{background:"#f9f9f9"}}>
                  {["Fecha","Tipo","Producto","Código","Antes","Cambio","Después","Usuario","Motivo"].map(h=>
                    <th key={h} style={{padding:"10px 12px",textAlign:"left",fontWeight:700,color:"#888",fontSize:11,textTransform:"uppercase",letterSpacing:.4,whiteSpace:"nowrap"}}>{h}</th>
                  )}
                </tr>
              </thead>
              <tbody>
                {filtered.map(e=>{
                  const cfg = TIPOS[e.tipo]||{color:"#666",bg:"#f5f5f5",icon:"-"};
                  return <tr key={e.id} style={{borderTop:"1px solid #f5f5f5"}}>
                    <td style={{padding:"9px 12px",fontSize:11,color:"#888",whiteSpace:"nowrap"}}>{e.fecha}</td>
                    <td style={{padding:"9px 12px"}}>
                      <span style={{background:cfg.bg,color:cfg.color,borderRadius:8,padding:"2px 8px",fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>
                        {cfg.icon} {e.tipo}
                      </span>
                    </td>
                    <td style={{padding:"9px 12px",fontWeight:600,color:"#1a1a1a",maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={e.producto}>{e.producto}</td>
                    <td style={{padding:"9px 12px",color:"#aaa",fontSize:11}}>{e.productoId}</td>
                    <td style={{padding:"9px 12px",color:"#666",textAlign:"center"}}>{e.stockAntes}</td>
                    <td style={{padding:"9px 12px",textAlign:"center",fontWeight:800,color:e.cambio>0?"#1e8449":RED}}>
                      {e.cambio>0?"+":""}{e.cambio}
                    </td>
                    <td style={{padding:"9px 12px",textAlign:"center",fontWeight:700,color:e.stockDespues===0?RED:"#1a1a1a"}}>{e.stockDespues}</td>
                    <td style={{padding:"9px 12px",whiteSpace:"nowrap"}}>
                      <span style={{fontWeight:600,fontSize:12}}>{e.usuario}</span>
                      <span style={{fontSize:10,color:e.rol==="admin"?RED:"#1a5276",marginLeft:4,fontWeight:600}}>({e.rol})</span>
                    </td>
                    <td style={{padding:"9px 12px",color:"#555",fontSize:12,maxWidth:220,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={e.motivo}>{e.motivo}</td>
                  </tr>;
                })}
              </tbody>
            </table>
          </div>
      }
    </div>
  );
}

function StockAdjust({products,onDel,onAdjust,addLog}) {
  const [search,setSearch]=useState("");
  const [selected,setSelected]=useState(null);
  const [mode,setMode]=useState("ajuste");
  const [qty,setQty]=useState(0);
  const [reason,setReason]=useState("");
  const [done,setDone]=useState(null);

  const q=search.toLowerCase();
  const found = useMemo(()=>
    q.length>1 ? products.filter(p=>norm(p.name).includes(norm(q))||normSKU(p.id).includes(normSKU(q))).slice(0,40) : []
  ,[products,search,q]);

  const reset = () => { setSelected(null); setQty(0); setReason(""); setSearch(""); };

  const confirmAction = async () => {
    if(!selected) return;
    if(!reason.trim()){alert("El motivo es obligatorio para registrar el movimiento");return;}
    if(mode==="baja") {
      await onDel(selected.id);
      addLog({
        tipo: "BAJA",
        productoId:   selected.id,
        producto:     selected.name,
        stockAntes:   selected.stock,
        stockDespues: 0,
        cambio:       -selected.stock,
        motivo:       reason.trim(),
      });
      setDone({msg:`✅ Producto "${selected.name}" dado de baja.`, color:"#c0392b"});
    } else {
      if(qty===0){alert("El ajuste no puede ser 0");return;}
      const stockDespues = Math.max(0, selected.stock + qty);
      onAdjust(selected.id, qty);
      addLog({
        tipo:         qty > 0 ? "ENTRADA" : "SALIDA",
        productoId:   selected.id,
        producto:     selected.name,
        stockAntes:   selected.stock,
        stockDespues,
        cambio:       qty,
        motivo:       reason.trim(),
      });
      const signo = qty>0?"+":"";
      setDone({msg:`✅ Stock de "${selected.name}" ajustado en ${signo}${qty} unidades. Nuevo stock: ${stockDespues}`, color:"#1e8449"});
    }
    reset();
    setTimeout(()=>setDone(null),5000);
  };

  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,alignItems:"start"}}>
      <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:15,marginBottom:12}}>🔍 Buscar producto</div>
        <input value={search} onChange={e=>{setSearch(e.target.value);setSelected(null);}}
          placeholder="Nombre o código del producto..."
          style={{width:"100%",padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:10}}/>
        {done&&<div style={{background:done.color==="#c0392b"?"#fdecea":"#d5f5e3",color:done.color,borderRadius:8,padding:"9px 12px",fontSize:13,fontWeight:600,marginBottom:10}}>{done.msg}</div>}
        <div style={{maxHeight:380,overflowY:"auto"}}>
          {search.length<=1
            ? <div style={{color:"#aaa",fontSize:13,textAlign:"center",padding:20}}>Escribí al menos 2 caracteres para buscar</div>
            : found.length===0
              ? <div style={{color:"#aaa",fontSize:13,textAlign:"center",padding:20}}>Sin resultados</div>
              : found.map(p=>(
                  <div key={p.id} onClick={()=>setSelected(p)}
                    style={{padding:"10px 12px",borderRadius:10,border:`1.5px solid ${selected?.id===p.id?RED:"#f0f0f0"}`,background:selected?.id===p.id?"#fdecea":"#fafafa",cursor:"pointer",marginBottom:6}}>
                    <div style={{fontWeight:600,fontSize:13,color:"#1a1a1a"}}>{p.name}</div>
                    <div style={{fontSize:11,color:"#888",marginTop:2,display:"flex",gap:12}}>
                      <span>{p.id}</span>
                      <span>Stock: <strong style={{color:p.stock===0?RED:p.stock<=5?"#e67e22":"#1e8449"}}>{p.stock===0?"Sin stock":p.stock}</strong></span>
                      <span style={{color:"#aaa"}}>{p.category}</span>
                    </div>
                  </div>
                ))
          }
        </div>
      </div>
      <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:15,marginBottom:14}}>⚙️ Acción sobre el producto</div>
        {!selected
          ? <div style={{color:"#aaa",fontSize:13,textAlign:"center",padding:30}}>Seleccioná un producto de la lista</div>
          : <>
              <div style={{background:"#f9f9f9",borderRadius:10,padding:"12px 14px",marginBottom:16}}>
                <div style={{fontWeight:700,fontSize:14,color:"#1a1a1a"}}>{selected.name}</div>
                <div style={{fontSize:12,color:"#888",marginTop:3,display:"flex",gap:16}}>
                  <span>{selected.id}</span>
                  <span>Stock actual: <strong style={{color:selected.stock===0?RED:"#1e8449"}}>{selected.stock}</strong></span>
                  <span>Venta: <strong style={{color:RED}}>{fARS(selected.salePrice)}</strong></span>
                </div>
              </div>
              <div style={{display:"flex",gap:8,marginBottom:16}}>
                {[{v:"ajuste",l:"⚖️ Ajustar stock",c:"#1a5276",bg:"#d6eaf8"},{v:"baja",l:"🗑 Dar de baja",c:RED,bg:"#fdecea"}].map(opt=>(
                  <button key={opt.v} onClick={()=>setMode(opt.v)} style={{flex:1,padding:"9px",borderRadius:9,border:`2px solid ${mode===opt.v?opt.c:"#e5e5e5"}`,background:mode===opt.v?opt.bg:"#fff",color:mode===opt.v?opt.c:"#666",fontWeight:700,fontSize:12,cursor:"pointer"}}>
                    {opt.l}
                  </button>
                ))}
              </div>
              {mode==="ajuste" && <>
                <div style={{marginBottom:12}}>
                  <label style={{fontSize:12,fontWeight:600,color:"#666",display:"block",marginBottom:4}}>
                    Cantidad a ajustar <span style={{color:"#aaa",fontWeight:400}}>(positivo para agregar, negativo para restar)</span>
                  </label>
                  <input type="number" value={qty} onChange={e=>setQty(+e.target.value||0)}
                    style={{width:"100%",padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:14,outline:"none",boxSizing:"border-box",textAlign:"center",fontWeight:700}}/>
                  {qty!==0&&<div style={{fontSize:12,color:qty>0?"#1e8449":RED,marginTop:6,fontWeight:600,textAlign:"center"}}>
                    Stock resultante: {Math.max(0, selected.stock + qty)} unidades
                  </div>}
                </div>
                <div style={{marginBottom:14}}>
                  <label style={{fontSize:12,fontWeight:600,color:"#666",display:"block",marginBottom:4}}>Motivo <span style={{color:RED}}>*</span></label>
                  <input value={reason} onChange={e=>setReason(e.target.value)} placeholder="Ej: corrección de inventario, mercadería dañada..."
                    style={{width:"100%",padding:"8px 12px",borderRadius:8,border:`1.5px solid ${!reason.trim()?"#e5e5e5":"#1e8449"}`,fontSize:13,outline:"none",boxSizing:"border-box"}}/>
                  <div style={{fontSize:11,color:"#aaa",marginTop:3}}>Obligatorio - quedará registrado en el historial de movimientos</div>
                </div>
                <button onClick={confirmAction} disabled={qty===0||!reason.trim()} style={{width:"100%",padding:"10px",borderRadius:9,border:"none",background:(qty===0||!reason.trim())?"#e5e5e5":"linear-gradient(135deg,#1a5e20,#1e8449)",color:(qty===0||!reason.trim())?"#aaa":"#fff",fontWeight:800,fontSize:14,cursor:(qty===0||!reason.trim())?"not-allowed":"pointer"}}>
                  ✅ Confirmar ajuste
                </button>
              </>}
              {mode==="baja" && <>
                <div style={{background:"#fdecea",border:"1px solid #f5c6cb",borderRadius:10,padding:"12px 14px",marginBottom:14}}>
                  <div style={{fontWeight:700,color:RED,fontSize:13,marginBottom:4}}>🎉️ Dar de baja elimina el producto del catálogo</div>
                  <div style={{fontSize:12,color:"#666"}}>Esta acción no se puede deshacer. El producto dejará de aparecer en Nuevo Pedido y Stock.</div>
                </div>
                <div style={{marginBottom:14}}>
                  <label style={{fontSize:12,fontWeight:600,color:"#666",display:"block",marginBottom:4}}>Motivo de la baja *</label>
                  <input value={reason} onChange={e=>setReason(e.target.value)} placeholder="Ej: producto descontinuado, error de carga..."
                    style={{width:"100%",padding:"8px 12px",borderRadius:8,border:`1.5px solid ${reason?"#e5e5e5":RED+"88"}`,fontSize:13,outline:"none",boxSizing:"border-box"}}/>
                </div>
                <BajaConfirm onConfirm={confirmAction} disabled={!reason.trim()}/>
              </>}
            </>
        }
      </div>
    </div>
  );
}

function BajaConfirm({onConfirm,disabled}) {
  const [step,setStep]=useState(false);
  if(step) return (
    <div style={{background:"#fdecea",borderRadius:9,padding:"12px 14px",border:`1.5px solid ${RED}44`}}>
      <div style={{fontWeight:700,color:RED,fontSize:13,marginBottom:10,textAlign:"center"}}>?Confirmás la baja definitiva?</div>
      <div style={{display:"flex",gap:8}}>
        <button onClick={onConfirm} style={{flex:1,padding:"9px",borderRadius:8,border:"none",background:RED,color:"#fff",fontWeight:800,cursor:"pointer",fontSize:13}}>Sí, dar de baja</button>
        <button onClick={()=>setStep(false)} style={{flex:1,padding:"9px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontWeight:600,color:"#666",fontSize:13}}>Cancelar</button>
      </div>
    </div>
  );
  return (
    <button onClick={()=>setStep(true)} disabled={disabled} style={{width:"100%",padding:"10px",borderRadius:9,border:"none",background:disabled?"#e5e5e5":RED,color:disabled?"#aaa":"#fff",fontWeight:800,fontSize:14,cursor:disabled?"not-allowed":"pointer"}}>
      🗑 Dar de baja este producto
    </button>
  );
}

function Stock({products,onUpd,onDel,onAdjust,isAdmin,addLog,stockLog,setStockLog,isMobile}) {
  const [search,setSearch]=useState("");
  const [cat,setCat]=useState("todos");
  const [editing,setEditing]=useState(null);
  const [stockTab,setStockTab]=useState("lista");
  const [page,setPage]=useState(0);
  const PAGE_SIZE = 100;

  const [soloConStock, setSoloConStock] = useState(false);
  const CATS=useMemo(()=>["todos",...new Set(products.map(p=>p.category))].sort(),[products]);
  const q=search.toLowerCase();

  // Reset page when filters change
  useEffect(()=>setPage(0),[search,cat,soloConStock]);

  const filtered=useMemo(()=>products.filter(p=>{
    if(soloConStock && p.stock<=0) return false;
    if(cat!=="todos"&&p.category!==cat) return false;
    if(q) return norm(p.name).includes(norm(q))||normSKU(p.id).includes(normSKU(q));
    return true;
  }),[products,cat,q,soloConStock]);

  const totalPages = Math.ceil(filtered.length / PAGE_SIZE);
  const shown = filtered.slice(page*PAGE_SIZE, (page+1)*PAGE_SIZE);
  const low=products.filter(p=>p.stock>0&&p.stock<=5);

  return (
    <div>
      {isAdmin && (
        <div style={{background:"#fff",borderRadius:12,padding:4,marginBottom:14,display:"flex",gap:4,boxShadow:"0 1px 4px #0001"}}>
          {[{k:"lista",l:"📋 Lista de Stock"},{k:"ajuste",l:"⚖️ Ajuste / Baja"},{k:"log",l:"📜 Movimientos"}].map(t=>(
            <button key={t.k} onClick={()=>setStockTab(t.k)} style={{flex:1,padding:"9px 14px",borderRadius:9,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:stockTab===t.k?`linear-gradient(135deg,${REDD},${RED})`:"transparent",color:stockTab===t.k?"#fff":"#555"}}>
              {t.l}
            </button>
          ))}
        </div>
      )}
      {isAdmin && stockTab==="ajuste" && <StockAdjust products={products} onDel={onDel} onAdjust={onAdjust} addLog={addLog}/>}
      {isAdmin && stockTab==="log" && <StockLog log={stockLog} onClear={async()=>{setStockLog([]);await db.clearStockLog();}}/>}
      {stockTab==="lista" && <>
        {low.length>0&&<StockAlert low={low}/>}
        <div style={{background:"#fff",borderRadius:12,padding:14,marginBottom:14,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center",boxShadow:"0 1px 4px #0001"}}>
          <div style={{fontWeight:800,fontSize:15,flex:1}}>
            📦 Stock — <span style={{color:RED}}>{filtered.length.toLocaleString("es-AR")}</span> productos
            {q||cat!=="todos" ? <span style={{fontSize:12,color:"#888",fontWeight:400,marginLeft:6}}>de {products.length.toLocaleString("es-AR")} total</span> : null}
          </div>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar por nombre o código..."
            style={{padding:"7px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:12,outline:"none",width:220}} autoFocus={false}/>
          <button onClick={()=>setSoloConStock(s=>!s)}
            style={{padding:"7px 14px",borderRadius:8,border:`1.5px solid ${soloConStock?"#1e8449":"#e5e5e5"}`,background:soloConStock?"#d5f5e3":"#fff",color:soloConStock?"#1e8449":"#666",fontWeight:700,fontSize:12,cursor:"pointer",whiteSpace:"nowrap"}}>
            {soloConStock ? "✅ Con stock" : "📦 Con stock"}
          </button>
          <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
            {CATS.map(c=><button key={c} onClick={()=>setCat(c)} style={{padding:"4px 10px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:11,fontWeight:600,borderColor:cat===c?RED:"#e5e5e5",background:cat===c?"#fdecea":"#fff",color:cat===c?RED:"#666"}}>{c==="todos"?"Todos":c}</button>)}
          </div>
        </div>

        <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px #0001",overflow:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{background:"#f9f9f9"}}>
              {(isMobile?["Producto","Stock","P. Venta"]:["Código","Producto","Categoría","Stock","P. Venta",...(isAdmin?["P. Costo","Margen"]:[]),""]).map(h=><th key={h} style={{padding:"10px 12px",textAlign:h==="Stock"||h==="P. Venta"?"right":"left",fontWeight:700,color:"#888",fontSize:11,textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>)}
            </tr></thead>
            <tbody>
              {shown.length===0
                ? <tr><td colSpan={8} style={{textAlign:"center",padding:40,color:"#aaa"}}>Sin resultados.</td></tr>
                : shown.map(p=>{
                    const m=p.costPrice>0?((p.salePrice-p.costPrice)/p.costPrice*100).toFixed(0):"-";
                    return isMobile ? (
                      <tr key={p.id} style={{borderTop:"1px solid #f5f5f5"}}>
                        <td style={{padding:"10px 12px"}}>
                          <div style={{fontWeight:600,fontSize:13}}>{p.name}</div>
                          <div style={{fontSize:10,color:"#aaa",marginTop:2}}>{p.id} · {p.category}</div>
                          {isAdmin&&<button onClick={()=>setEditing(p)} style={{padding:"3px 8px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:11,fontWeight:600,marginTop:4}}>✏️ Editar</button>}
                        </td>
                        <td style={{padding:"10px 12px",textAlign:"right"}}><SPill n={p.stock}/></td>
                        <td style={{padding:"10px 12px",fontWeight:700,color:RED,textAlign:"right",whiteSpace:"nowrap"}}>{fARS(p.salePrice)}</td>
                      </tr>
                    ) : (
                    <tr key={p.id} style={{borderTop:"1px solid #f5f5f5"}}>
                      <td style={{padding:"9px 12px",color:"#aaa",fontSize:11}}>{p.id}</td>
                      <td style={{padding:"9px 12px",fontWeight:600,color:"#1a1a1a",maxWidth:260}}>{p.name}</td>
                      <td style={{padding:"9px 12px",color:"#aaa",fontSize:11}}>{p.category}</td>
                      <td style={{padding:"9px 12px"}}><SPill n={p.stock}/></td>
                      <td style={{padding:"9px 12px",fontWeight:700,color:RED}}>{fARS(p.salePrice)}</td>
                      {isAdmin&&<td style={{padding:"9px 12px",color:"#666"}}>{fARS(p.costPrice)}</td>}
                      {isAdmin&&<td style={{padding:"9px 12px",fontWeight:700,color:+m>=40?"#1e8449":"#e67e22"}}>{m}%</td>}
                      <td style={{padding:"9px 12px"}}>{isAdmin&&<button onClick={()=>setEditing(p)} style={{padding:"4px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:11,fontWeight:600}}>✏️</button>}</td>
                    </tr>
                    );
                  })
              }
            </tbody>
          </table>
        </div>

        {/* Paginación */}
        {totalPages > 1 && (
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"12px 4px",flexWrap:"wrap",gap:8}}>
            <div style={{fontSize:12,color:"#888"}}>
              Mostrando {page*PAGE_SIZE+1}–{Math.min((page+1)*PAGE_SIZE, filtered.length)} de {filtered.length.toLocaleString("es-AR")} productos
            </div>
            <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
              <button onClick={()=>setPage(0)} disabled={page===0}
                style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:page===0?"not-allowed":"pointer",color:page===0?"#ccc":"#555",fontSize:12,fontWeight:600}}>
                «
              </button>
              <button onClick={()=>setPage(p=>Math.max(0,p-1))} disabled={page===0}
                style={{padding:"5px 12px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:page===0?"not-allowed":"pointer",color:page===0?"#ccc":"#555",fontSize:12,fontWeight:600}}>
                ← Anterior
              </button>
              {/* Page numbers — show at most 7 around current */}
              {Array.from({length:totalPages},(_,i)=>i)
                .filter(i=>i===0||i===totalPages-1||Math.abs(i-page)<=2)
                .reduce((acc,i,idx,arr)=>{
                  if(idx>0&&i-arr[idx-1]>1) acc.push("...");
                  acc.push(i);
                  return acc;
                },[])
                .map((item,i)=>
                  item==="..." ? <span key={`e${i}`} style={{padding:"5px 4px",fontSize:12,color:"#aaa"}}>…</span>
                  : <button key={item} onClick={()=>setPage(item)}
                      style={{padding:"5px 10px",borderRadius:7,border:`1.5px solid ${page===item?RED:"#e5e5e5"}`,background:page===item?"#fdecea":"#fff",cursor:"pointer",color:page===item?RED:"#555",fontSize:12,fontWeight:page===item?800:600,minWidth:34}}>
                      {item+1}
                    </button>
                )
              }
              <button onClick={()=>setPage(p=>Math.min(totalPages-1,p+1))} disabled={page===totalPages-1}
                style={{padding:"5px 12px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:page===totalPages-1?"not-allowed":"pointer",color:page===totalPages-1?"#ccc":"#555",fontSize:12,fontWeight:600}}>
                Siguiente →
              </button>
              <button onClick={()=>setPage(totalPages-1)} disabled={page===totalPages-1}
                style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:page===totalPages-1?"not-allowed":"pointer",color:page===totalPages-1?"#ccc":"#555",fontSize:12,fontWeight:600}}>
                »
              </button>
            </div>
          </div>
        )}

        {editing&&<EditModal p={editing} onSave={p=>{onUpd(p);setEditing(null);}} onClose={()=>setEditing(null)}/>}
      </>}
    </div>
  );
}

function EditModal({p,onSave,onClose}) {
  const [cost,setCost]=useState(p.costPrice);
  const [sale,setSale]=useState(p.salePrice);
  const [stock,setStock]=useState(p.stock);
  const m=cost>0?((sale-cost)/cost*100).toFixed(1):"-";
  return (
    <div style={{position:"fixed",inset:0,background:"#0007",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999,padding:16}}>
      <div style={{background:"#fff",borderRadius:16,padding:24,width:"100%",maxWidth:440,boxShadow:"0 20px 60px #0003"}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:16}}>
          <div style={{fontWeight:800,fontSize:14,color:"#1a1a1a",flex:1,marginRight:8,lineHeight:1.3}}>{p.name}</div>
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:22,color:"#aaa"}}>x</button>
        </div>
        {[["Precio Costo ($)",cost,setCost],["Precio Venta ($)",sale,setSale],["Stock",stock,setStock]].map(([lbl,val,set])=>(
          <div key={lbl} style={{marginBottom:12}}>
            <label style={{fontSize:12,fontWeight:600,color:"#666",display:"block",marginBottom:4}}>{lbl}</label>
            <input type="number" value={val} onChange={e=>set(+e.target.value||0)} style={{...inputStyle}}/>
          </div>
        ))}
        <div style={{background:"#f0fdf4",borderRadius:8,padding:"8px 12px",fontSize:12,color:"#1e8449",marginBottom:14}}>
          Margen: <strong>{m}%</strong> . Ganancia/u: <strong>{fARS(sale-cost)}</strong>
        </div>
        <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
          <button onClick={onClose} style={{padding:"8px 14px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontWeight:600,color:"#666"}}>Cancelar</button>
          <button onClick={()=>onSave({...p,costPrice:cost,salePrice:sale,stock})} style={{padding:"8px 14px",borderRadius:8,border:"none",background:RED,color:"#fff",cursor:"pointer",fontWeight:700}}>Guardar</button>
        </div>
      </div>
    </div>
  );
}

// ─── INGRESAR STOCK DESDE SOLICITUD ──────────────────────────────────────────
function IngresarDesdeSolicitud({po, products, onStock, onDone}) {
  const [items, setItems] = useState(
    po.items.map(i => ({
      ...i,
      qtyRecibida: i.qty, // default = lo que se pidió
      cost: products.find(p=>p.id===i.id)?.costPrice || 0,
      incluir: true,
    }))
  );
  const [saving, setSaving] = useState(false);
  const [ok, setOk] = useState(false);

  const update = (pid, field, val) =>
    setItems(prev => prev.map(i => i.pid===pid ? {...i,[field]:val} : i));

  const submit = async () => {
    const toIngresar = items.filter(i => i.incluir && i.qtyRecibida > 0);
    if(!toIngresar.length) { alert("Seleccioná al menos un producto para ingresar"); return; }
    setSaving(true);
    for(const it of toIngresar) {
      await onStock(it.pid, +it.qtyRecibida, +it.cost);
    }
    setOk(true);
    setSaving(false);
    setTimeout(()=>onDone(), 1500);
  };

  if(ok) return (
    <div style={{textAlign:"center",padding:"20px 0"}}>
      <div style={{fontSize:40,marginBottom:6}}>✅</div>
      <div style={{fontWeight:800,color:"#1e8449",fontSize:15}}>¡Stock actualizado!</div>
    </div>
  );

  return (
    <div>
      <div style={{background:"#fff",borderRadius:10,overflow:"auto",marginBottom:12}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:400}}>
          <thead><tr style={{background:"#f5f5f5"}}>
            <th style={{padding:"8px 10px",textAlign:"left",fontSize:10,color:"#888",fontWeight:700,textTransform:"uppercase",width:32}}>
              <input type="checkbox" checked={items.every(i=>i.incluir)} onChange={e=>setItems(prev=>prev.map(i=>({...i,incluir:e.target.checked})))} style={{cursor:"pointer"}}/>
            </th>
            <th style={{padding:"8px 10px",textAlign:"left",fontSize:10,color:"#888",fontWeight:700,textTransform:"uppercase"}}>Producto</th>
            <th style={{padding:"8px 10px",textAlign:"center",fontSize:10,color:"#888",fontWeight:700,textTransform:"uppercase"}}>Pedido</th>
            <th style={{padding:"8px 10px",textAlign:"center",fontSize:10,color:"#888",fontWeight:700,textTransform:"uppercase"}}>Recibido</th>
            <th style={{padding:"8px 10px",textAlign:"center",fontSize:10,color:"#888",fontWeight:700,textTransform:"uppercase"}}>P. Costo</th>
          </tr></thead>
          <tbody>
            {items.map((it,i)=>(
              <tr key={it.pid} style={{borderTop:"1px solid #f0f0f0",background:it.incluir?"#fff":"#f9f9f9",opacity:it.incluir?1:.5}}>
                <td style={{padding:"8px 10px",textAlign:"center"}}>
                  <input type="checkbox" checked={it.incluir} onChange={e=>update(it.pid,"incluir",e.target.checked)} style={{cursor:"pointer"}}/>
                </td>
                <td style={{padding:"8px 10px"}}>
                  <div style={{fontWeight:600,fontSize:12,lineHeight:1.3}}>{it.name}</div>
                  <div style={{fontSize:10,color:"#aaa"}}>{it.id}</div>
                </td>
                <td style={{padding:"8px 10px",textAlign:"center",color:"#888",fontSize:13}}>{it.qty}</td>
                <td style={{padding:"8px 10px",textAlign:"center"}}>
                  <input type="number" min="0" value={it.qtyRecibida}
                    onChange={e=>update(it.pid,"qtyRecibida",+e.target.value||0)}
                    style={{width:56,textAlign:"center",padding:"4px",borderRadius:6,border:`1.5px solid ${it.qtyRecibida<it.qty?"#e67e22":it.qtyRecibida>it.qty?"#3498db":"#1e8449"}`,fontWeight:700,fontSize:13,outline:"none"}}/>
                  {it.qtyRecibida<it.qty&&<div style={{fontSize:9,color:"#e67e22",marginTop:2}}>Falta {it.qty-it.qtyRecibida}</div>}
                  {it.qtyRecibida>it.qty&&<div style={{fontSize:9,color:"#3498db",marginTop:2}}>Extra +{it.qtyRecibida-it.qty}</div>}
                </td>
                <td style={{padding:"8px 10px",textAlign:"center"}}>
                  <input type="number" min="0" value={it.cost}
                    onChange={e=>update(it.pid,"cost",+e.target.value||0)}
                    placeholder="$0"
                    style={{width:72,textAlign:"center",padding:"4px",borderRadius:6,border:"1.5px solid #e5e5e5",fontSize:12,outline:"none"}}/>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div style={{display:"flex",gap:8,alignItems:"center"}}>
        <div style={{flex:1,fontSize:12,color:"#1e8449"}}>
          <strong>{items.filter(i=>i.incluir).length}</strong> productos · <strong>{items.filter(i=>i.incluir).reduce((s,i)=>s+i.qtyRecibida,0)}</strong> unidades a ingresar
        </div>
        <button onClick={submit} disabled={saving}
          style={{padding:"10px 20px",borderRadius:10,border:"none",background:saving?"#e5e5e5":"linear-gradient(135deg,#1a5e20,#1e8449)",color:saving?"#aaa":"#fff",fontWeight:800,fontSize:13,cursor:saving?"not-allowed":"pointer",whiteSpace:"nowrap"}}>
          {saving?"Ingresando...":"📦 Ingresar al Stock"}
        </button>
      </div>
    </div>
  );
}

// ─── SOLICITUD DE COMPRA ──────────────────────────────────────────────────────
const ESTADO_CFG = {
  abierta:   {label:"Abierta",    color:"#1a5276", bg:"#d6eaf8",  icon:"📋"},
  revisando: {label:"Revisando",  color:"#e67e22", bg:"#fef9e7",  icon:"🔍"},
  cerrada:   {label:"Cerrada",    color:"#1e8449", bg:"#d5f5e3",  icon:"✅"},
};

function printSolicitudPDF(po, logoSrc) {
  const rows = po.items.map(it=>`
    <tr>
      <td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;font-size:14px;">${it.name}</td>
      <td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;font-size:14px;">${it.id||""}</td>
      <td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;text-align:center;font-size:15px;font-weight:700;">${it.qty}</td>
      <td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;font-size:13px;color:#888;">${it.notas||""}</td>
    </tr>`).join("");

  const html = `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"/>
<title>Solicitud de Compra</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0;font-family:'Segoe UI',Arial,sans-serif;}
  body{background:#fdfbf7;}
  @media print{.no-print{display:none!important;}@page{margin:3mm;size:A4;}}
  .btn{display:block;margin:12px auto;padding:9px 28px;background:linear-gradient(135deg,#7b1a1a,#1a5276);color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;}
  .header-pdf{display:flex;align-items:center;gap:16px;padding:18px 22px;background:linear-gradient(135deg,#7b1a1a,#1a5276);}
  .logo-circle-pdf{width:64px;height:64px;border-radius:50%;overflow:hidden;border:2px solid #ffffff33;flex-shrink:0;background:#fff;display:flex;align-items:center;justify-content:center;}
  .logo-circle-pdf img{width:88%;height:88%;object-fit:contain;}
  .brand-name-pdf{color:#fff;font-family:Georgia,'Times New Roman',serif;font-size:24px;font-weight:700;letter-spacing:0.3px;}
  .brand-sub-pdf{color:#bcd9ee;font-size:9px;letter-spacing:2px;text-transform:uppercase;margin-top:3px;}
  .gold-line{height:3px;background:linear-gradient(90deg,#c9a96a,#e8d4a8,#c9a96a);}
  .content{padding:16px 20px 0;}
  .doc-meta{display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;padding-bottom:12px;border-bottom:2px solid #f0ece2;}
  .badge{display:inline-block;padding:6px 18px;border-radius:8px;font-size:14px;font-weight:800;background:#1a5276;color:#fff;}
  .info-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px 16px;margin-bottom:16px;}
  .info-box{background:#fdfbf7;border-radius:8px;padding:9px 12px;border-left:3px solid #e5ddd0;}
  .info-lbl{font-size:8px;color:#999;text-transform:uppercase;letter-spacing:.7px;margin-bottom:3px;font-weight:700;}
  .info-val{font-size:13px;font-weight:700;font-family:Georgia,serif;}
  table{width:100%;border-collapse:collapse;margin-bottom:16px;}
  thead tr{background:#fdfbf7;}
  th{padding:9px 12px;text-align:left;font-size:10px;color:#999;text-transform:uppercase;font-weight:700;border-bottom:2px solid #f0ece2;}
  th.c{text-align:center;}
  tbody tr:nth-child(even){background:#fdfbf7;}
  .footer{border-top:1px solid #f0ece2;padding-top:10px;margin:0;font-size:10px;color:#bbb;display:flex;justify-content:space-between;}
  .footer-brand{color:#7b1a1a;font-weight:700;font-family:Georgia,serif;}
  .notes{background:#fdfbf7;border-left:3px solid #1a5276;padding:8px 12px;border-radius:0 6px 6px 0;font-size:12px;color:#555;margin-bottom:14px;}
</style></head><body>
<button class="no-print btn" onclick="window.print()">🖨️ Imprimir / Guardar PDF</button>
<div>
  <div class="header-pdf">
    <div class="logo-circle-pdf"><img src="${logoSrc}" alt="LM" onerror="this.style.display='none'"/></div>
    <div>
      <div class="brand-name-pdf">Libreria Madrid</div>
      <div class="brand-sub-pdf">Sistema de Gestión Comercial</div>
    </div>
  </div>
  <div class="gold-line"></div>
  <div class="content">
    <div class="doc-meta">
      <div><div style="font-size:11px;color:#999;margin-bottom:5px">Documento de</div><div class="badge">SOLICITUD DE COMPRA</div></div>
      <div style="text-align:right">
        <div style="font-size:22px;font-weight:900;">#${po.id.slice(-6).toUpperCase()}</div>
        <div style="font-size:11px;color:#888">${po.fecha}</div>
      </div>
    </div>
    <div class="info-grid">
      <div class="info-box" style="border-left-color:#1a5276"><div class="info-lbl">Vendedor</div><div class="info-val">${po.vendedor}</div></div>
      <div class="info-box"><div class="info-lbl">Fecha</div><div class="info-val">${po.fecha}</div></div>
      <div class="info-box"><div class="info-lbl">Estado</div><div class="info-val">${ESTADO_CFG[po.estado]?.label||po.estado}</div></div>
    </div>
    ${po.notas?`<div class="notes"><strong>Notas:</strong> ${po.notas}</div>`:""}
    <table>
      <thead><tr>
        <th style="width:40%">Producto</th>
        <th style="width:20%">Código</th>
        <th class="c" style="width:15%">Cantidad</th>
        <th style="width:25%">Observaciones</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
    <div style="text-align:right;font-size:13px;color:#888;margin-bottom:16px;">
      Total items: <strong style="color:#1a5276;font-size:16px">${po.items.length}</strong> productos · 
      <strong style="color:#1a5276;font-size:16px">${po.items.reduce((s,i)=>s+i.qty,0)}</strong> unidades
    </div>
    <div class="footer">
      <span><span class="footer-brand">Libreria Madrid</span> — madrid.libreria · +54 9 11 2502-0640</span>
      <span>Emitido el ${new Date().toLocaleString("es-AR")}</span>
    </div>
  </div>
</div>
</body></html>`;

  const w = window.open("","_blank","width=820,height=750");
  if(w){ w.document.write(html); w.document.close(); }
}

function exportSolicitudXLSX(po) {
  const rows = [
    ["Solicitud de Compra — Libreria Madrid"],
    [],
    ["ID", po.id.slice(-6).toUpperCase()],
    ["Vendedor", po.vendedor],
    ["Fecha", po.fecha],
    ["Estado", ESTADO_CFG[po.estado]?.label||po.estado],
    ["Notas", po.notas||""],
    [],
    ["Producto","Código","Cantidad","Observaciones"],
    ...po.items.map(it=>[it.name, it.id||"", it.qty, it.notas||""]),
    [],
    ["Total productos", po.items.length],
    ["Total unidades", po.items.reduce((s,i)=>s+i.qty,0)],
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{wch:40},{wch:15},{wch:12},{wch:30}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Solicitud");
  XLSX.writeFile(wb, `Solicitud_${po.id.slice(-6).toUpperCase()}_${po.fecha.replace(/\//g,"-")}.xlsx`);
}

function SolicitudCompra({products,currentUser,isAdmin,purchaseOrders,setPurchaseOrders,isMobile,onStockExternal,addLog,onCreated}) {
  const [view, setView] = useState("lista"); // lista | nueva | detalle
  const [selected, setSelected] = useState(null);
  const [showAddProduct, setShowAddProduct] = useState(false);
  const [addSearch, setAddSearch] = useState("");
  const [cart, setCart] = useState([]);
  const [notas, setNotas] = useState("");
  const [saving, setSaving] = useState(false);
  const [itemNotas, setItemNotas] = useState({});

  const myOrders = isAdmin ? purchaseOrders : purchaseOrders.filter(po=>po.vendedor===currentUser.vendedor||po.vendedor===currentUser.name);

  const savePO = async (po) => {
    setPurchaseOrders(prev=>prev.find(x=>x.id===po.id)?prev.map(x=>x.id===po.id?po:x):[po,...prev]);
    await db.savePurchaseOrder(po);
  };

  const createNew = async () => {
    if(!cart.length){alert("Agregá al menos un producto");return;}
    setSaving(true);
    const po = {
      id: genId(),
      fecha: today(),
      vendedor: currentUser.vendedor||currentUser.name,
      estado: "abierta",
      items: cart.map(i=>({...i, notas: itemNotas[i.pid]||""})),
      notas,
      fechaCierre: "",
    };
    await savePO(po);
    if(onCreated) await onCreated(po);
    setCart([]); setNotas(""); setItemNotas({});
    setSaving(false);
    setView("lista");
  };

  const changeEstado = async (po, estado) => {
    const updated = {...po, estado, fechaCierre: estado==="cerrada" ? today() : po.fechaCierre};
    await savePO(updated);
    if(selected?.id===po.id) setSelected(updated);
  };

  const deletePO = async (id) => {
    if(!window.confirm("¿Eliminar esta solicitud?")) return;
    setPurchaseOrders(prev=>prev.filter(x=>x.id!==id));
    await db.deletePurchaseOrder(id);
    if(selected?.id===id){setSelected(null);setView("lista");}
  };

  const openDetail = (po) => { setSelected(po); setView("detalle"); };

  // ── LISTA ──
  if(view==="lista") return (
    <div style={{padding:isMobile?"0":"0"}}>
      <div style={{background:"#fff",borderRadius:12,padding:4,marginBottom:14,display:"flex",gap:4,boxShadow:"0 1px 4px #0001",margin:isMobile?"0 12px 14px":"0 0 14px"}}>
        <button onClick={()=>setView("nueva")} style={{flex:1,padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:`linear-gradient(135deg,#1a5e20,#1e8449)`,color:"#fff"}}>
          + Nueva Solicitud
        </button>
      </div>

      {myOrders.length===0
        ? <div style={{textAlign:"center",padding:"50px 20px",color:"#aaa",background:"#fff",borderRadius:12,margin:isMobile?"0 12px":"0"}}>
            <div style={{fontSize:48,marginBottom:8}}>📋</div>
            <div style={{fontWeight:700,fontSize:15}}>No hay solicitudes aún</div>
            <div style={{fontSize:13,marginTop:4}}>Creá una para pedir mercadería al proveedor</div>
          </div>
        : myOrders.map(po=>{
            const cfg = ESTADO_CFG[po.estado]||{};
            return (
              <div key={po.id} style={{background:"#fff",borderRadius:12,marginBottom:8,boxShadow:"0 1px 6px #0001",overflow:"hidden",margin:isMobile?"0 12px 8px":"0 0 8px"}}>
                <div style={{padding:"13px 16px",display:"flex",alignItems:"flex-start",gap:10,cursor:"pointer"}} onClick={()=>openDetail(po)}>
                  <div style={{width:10,height:10,borderRadius:"50%",background:cfg.color,marginTop:5,flexShrink:0}}/>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontWeight:800,fontSize:14}}>Solicitud #{po.id.slice(-6).toUpperCase()}</div>
                    <div style={{fontSize:11,color:"#888",marginTop:2}}>{po.fecha} · 👤 {po.vendedor}</div>
                  </div>
                  <div style={{textAlign:"right",flexShrink:0}}>
                    <span style={{background:cfg.bg,color:cfg.color,borderRadius:8,padding:"3px 10px",fontSize:11,fontWeight:700}}>{cfg.icon} {cfg.label}</span>
                    <div style={{fontSize:11,color:"#888",marginTop:4}}>{po.items.length} productos · {po.items.reduce((s,i)=>s+i.qty,0)} uds.</div>
                  </div>
                </div>
                <div style={{display:"flex",gap:6,padding:"8px 12px",background:"#fafafa",borderTop:"1px solid #f0f0f0"}}>
                  <button onClick={()=>openDetail(po)} style={{flex:1,padding:"8px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",fontSize:11,fontWeight:700,cursor:"pointer"}}>Ver detalle</button>
                  {isAdmin && po.estado==="abierta" && <button onClick={()=>changeEstado(po,"revisando")} style={{flex:1,padding:"8px",borderRadius:8,border:"none",background:"#e67e22",color:"#fff",fontSize:11,fontWeight:700,cursor:"pointer"}}>🔍 Revisar</button>}
                  {isAdmin && po.estado==="revisando" && <button onClick={()=>changeEstado(po,"cerrada")} style={{flex:1,padding:"8px",borderRadius:8,border:"none",background:"#1e8449",color:"#fff",fontSize:11,fontWeight:700,cursor:"pointer"}}>✅ Cerrar</button>}
                  {isAdmin && <button onClick={()=>deletePO(po.id)} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #fcc",background:"#fff",color:"#c0392b",fontSize:11,fontWeight:700,cursor:"pointer"}}>🗑</button>}
                </div>
              </div>
            );
          })
      }
    </div>
  );

  // ── DETALLE ──
  if(view==="detalle" && selected) {
    const cfg = ESTADO_CFG[selected.estado]||{};
    const po = purchaseOrders.find(x=>x.id===selected.id)||selected;
    return (
      <div>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14,padding:isMobile?"0 12px":"0"}}>
          <button onClick={()=>setView("lista")} style={{padding:"7px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",fontWeight:600,fontSize:12,cursor:"pointer"}}>← Volver</button>
          <div style={{fontWeight:800,fontSize:15}}>Solicitud #{po.id.slice(-6).toUpperCase()}</div>
          <span style={{background:cfg.bg,color:cfg.color,borderRadius:8,padding:"3px 10px",fontSize:11,fontWeight:700,marginLeft:"auto"}}>{cfg.icon} {cfg.label}</span>
        </div>

        <div style={{background:"#fff",borderRadius:12,padding:16,marginBottom:12,boxShadow:"0 1px 4px #0001",margin:isMobile?"0 12px 12px":"0 0 12px"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12}}>
            <div style={{background:"#f9f9f9",borderRadius:8,padding:"9px 12px",borderLeft:"3px solid #1a5276"}}>
              <div style={{fontSize:9,fontWeight:800,color:"#999",textTransform:"uppercase",marginBottom:2}}>Vendedor</div>
              <div style={{fontSize:13,fontWeight:700}}>{po.vendedor}</div>
            </div>
            <div style={{background:"#f9f9f9",borderRadius:8,padding:"9px 12px",borderLeft:"3px solid #e5e5e5"}}>
              <div style={{fontSize:9,fontWeight:800,color:"#999",textTransform:"uppercase",marginBottom:2}}>Fecha</div>
              <div style={{fontSize:13,fontWeight:700}}>{po.fecha}</div>
            </div>
            <div style={{background:"#f9f9f9",borderRadius:8,padding:"9px 12px",borderLeft:`3px solid ${cfg.color}`}}>
              <div style={{fontSize:9,fontWeight:800,color:"#999",textTransform:"uppercase",marginBottom:2}}>Estado</div>
              <div style={{fontSize:13,fontWeight:700,color:cfg.color}}>{cfg.label}</div>
            </div>
          </div>
          {po.notas&&<div style={{background:"#f9f9f9",borderLeft:"3px solid #1a5276",padding:"8px 12px",borderRadius:"0 8px 8px 0",fontSize:12,color:"#555",marginBottom:8}}>
            <strong>Notas:</strong> {po.notas}
          </div>}
        </div>

        {/* Items table — editable for admin when not closed */}
        <div style={{background:"#fff",borderRadius:12,overflow:"auto",boxShadow:"0 1px 4px #0001",marginBottom:12,margin:isMobile?"0 12px 12px":"0 0 12px"}}>
          {isAdmin && po.estado!=="cerrada" && (
            <div style={{padding:"8px 14px",background:"#fef9e7",borderBottom:"1px solid #f0e0a0",fontSize:12,color:"#7d6608",display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
              <span>✏️ Modo edición — modificá cantidades o agregá productos</span>
              <button onClick={()=>setShowAddProduct(s=>!s)}
                style={{marginLeft:"auto",padding:"5px 12px",borderRadius:7,border:"none",background:showAddProduct?"#1a5276":"#e67e22",color:"#fff",fontWeight:700,fontSize:11,cursor:"pointer",whiteSpace:"nowrap"}}>
                {showAddProduct?"✕ Cerrar buscador":"+ Agregar producto"}
              </button>
            </div>
          )}
          {isAdmin && po.estado!=="cerrada" && showAddProduct && (
            <div style={{padding:"12px 14px",borderBottom:"1px solid #f0f0f0",background:"#f9f9fb"}}>
              <input value={addSearch} onChange={e=>setAddSearch(e.target.value)} autoFocus
                placeholder="🔍 Buscar producto para agregar..."
                style={{width:"100%",padding:"8px 12px",borderRadius:8,border:"1.5px solid #1a5276",fontSize:13,outline:"none",marginBottom:8}}/>
              {addSearch.length>1 && (
                <div style={{maxHeight:220,overflowY:"auto",display:"flex",flexDirection:"column",gap:5}}>
                  {products.filter(p=>norm(p.name).includes(norm(addSearch))||normSKU(p.id).includes(normSKU(addSearch))).slice(0,20).map(p=>{
                    const already = po.items.find(i=>i.id===p.id);
                    return (
                      <div key={p.id} style={{display:"flex",alignItems:"center",gap:10,background:"#fff",borderRadius:8,padding:"8px 12px",border:"1.5px solid #e5e5e5"}}>
                        <div style={{flex:1,minWidth:0}}>
                          <div style={{fontWeight:600,fontSize:12,lineHeight:1.3}}>{p.name}</div>
                          <div style={{fontSize:10,color:"#aaa"}}>{p.id}</div>
                        </div>
                        {already
                          ? <span style={{fontSize:11,color:"#1e8449",fontWeight:700,background:"#d5f5e3",borderRadius:6,padding:"3px 8px"}}>✓ Ya está (×{already.qty})</span>
                          : <button onClick={()=>{
                              const newItems=[...po.items,{pid:p.id,id:p.id,name:p.name,qty:1,notas:""}];
                              const updated={...po,items:newItems};
                              savePO(updated); setSelected(updated);
                              setAddSearch(""); setShowAddProduct(false);
                            }} style={{padding:"5px 12px",borderRadius:7,border:"none",background:"#1a5276",color:"#fff",fontWeight:700,fontSize:12,cursor:"pointer",whiteSpace:"nowrap"}}>
                              + Agregar
                            </button>
                        }
                      </div>
                    );
                  })}
                  {products.filter(p=>norm(p.name).includes(norm(addSearch))||normSKU(p.id).includes(normSKU(addSearch))).length===0&&(
                    <div style={{textAlign:"center",color:"#aaa",fontSize:12,padding:12}}>Sin resultados</div>
                  )}
                </div>
              )}
            </div>
          )}
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,minWidth:400}}>
            <thead><tr style={{background:"#f5f5f5"}}>
              <th style={{padding:"10px 12px",textAlign:"left",fontSize:11,color:"#888",fontWeight:700,textTransform:"uppercase"}}>Producto</th>
              <th style={{padding:"10px 12px",textAlign:"left",fontSize:11,color:"#888",fontWeight:700,textTransform:"uppercase"}}>Código</th>
              <th style={{padding:"10px 12px",textAlign:"center",fontSize:11,color:"#888",fontWeight:700,textTransform:"uppercase"}}>Cant.</th>
              <th style={{padding:"10px 12px",textAlign:"left",fontSize:11,color:"#888",fontWeight:700,textTransform:"uppercase"}}>Obs.</th>
              {isAdmin && po.estado!=="cerrada" && <th style={{padding:"10px 12px",textAlign:"center",fontSize:11,color:"#888",fontWeight:700,textTransform:"uppercase"}}>Quitar</th>}
            </tr></thead>
            <tbody>
              {po.items.map((it,i)=>(
                <tr key={i} style={{borderTop:"1px solid #f5f5f5",background:i%2===0?"#fff":"#fafafa"}}>
                  <td style={{padding:"9px 12px",fontWeight:600}}>{it.name}</td>
                  <td style={{padding:"9px 12px",color:"#aaa",fontSize:11}}>{it.id||""}</td>
                  <td style={{padding:"9px 12px",textAlign:"center"}}>
                    {isAdmin && po.estado!=="cerrada"
                      ? <div style={{display:"flex",alignItems:"center",gap:4,justifyContent:"center"}}>
                          <button onClick={()=>{
                            const newItems=po.items.map((x,j)=>j===i?{...x,qty:Math.max(1,x.qty-1)}:x);
                            const updated={...po,items:newItems};
                            savePO(updated); setSelected(updated);
                          }} style={{width:24,height:24,borderRadius:5,border:"1.5px solid #ccc",background:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>-</button>
                          <input type="number" min="1" value={it.qty}
                            onChange={e=>{
                              const newItems=po.items.map((x,j)=>j===i?{...x,qty:+e.target.value||1}:x);
                              const updated={...po,items:newItems};
                              savePO(updated); setSelected(updated);
                            }}
                            style={{width:48,textAlign:"center",padding:"3px",borderRadius:6,border:"1.5px solid #1a5276",fontWeight:700,fontSize:13,outline:"none"}}/>
                          <button onClick={()=>{
                            const newItems=po.items.map((x,j)=>j===i?{...x,qty:x.qty+1}:x);
                            const updated={...po,items:newItems};
                            savePO(updated); setSelected(updated);
                          }} style={{width:24,height:24,borderRadius:5,border:"1.5px solid #ccc",background:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>+</button>
                        </div>
                      : <span style={{fontWeight:800,fontSize:15,color:"#1a5276"}}>{it.qty}</span>
                    }
                  </td>
                  <td style={{padding:"9px 12px"}}>
                    {isAdmin && po.estado!=="cerrada"
                      ? <input value={it.notas||""} onChange={e=>{
                            const newItems=po.items.map((x,j)=>j===i?{...x,notas:e.target.value}:x);
                            const updated={...po,items:newItems};
                            savePO(updated); setSelected(updated);
                          }}
                          placeholder="Obs..." style={{width:"100%",padding:"4px 8px",borderRadius:6,border:"1.5px solid #e5e5e5",fontSize:12,outline:"none"}}/>
                      : <span style={{color:"#888",fontSize:12}}>{it.notas||""}</span>
                    }
                  </td>
                  {isAdmin && po.estado!=="cerrada" && (
                    <td style={{padding:"9px 12px",textAlign:"center"}}>
                      <button onClick={()=>{
                        const newItems=po.items.filter((_,j)=>j!==i);
                        const updated={...po,items:newItems};
                        savePO(updated); setSelected(updated);
                      }} style={{background:"none",border:"none",cursor:"pointer",color:"#c0392b",fontSize:16,lineHeight:1}}>×</button>
                    </td>
                  )}
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{padding:"10px 14px",borderTop:"1px solid #f0f0f0",display:"flex",justifyContent:"flex-end",gap:20,fontSize:12,color:"#888"}}>
            <span>Total productos: <strong style={{color:"#1a5276",fontSize:14}}>{po.items.length}</strong></span>
            <span>Total unidades: <strong style={{color:"#1a5276",fontSize:14}}>{po.items.reduce((s,i)=>s+i.qty,0)}</strong></span>
          </div>
        </div>

        {/* Actions */}
        <div style={{display:"flex",gap:8,flexWrap:"wrap",padding:isMobile?"0 12px":"0"}}>
          {isAdmin && po.estado==="abierta" && <button onClick={()=>changeEstado(po,"revisando")} style={{flex:1,padding:"11px",borderRadius:10,border:"none",background:"#e67e22",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}>🔍 Marcar como Revisando</button>}
          {isAdmin && po.estado==="revisando" && <button onClick={()=>changeEstado(po,"cerrada")} style={{flex:1,padding:"11px",borderRadius:10,border:"none",background:"#1e8449",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}>✅ Cerrar solicitud</button>}
          <button onClick={()=>{
            const img=new Image();img.crossOrigin="anonymous";
            img.onload=()=>{try{const c=document.createElement("canvas");c.width=img.width;c.height=img.height;c.getContext("2d").drawImage(img,0,0);printSolicitudPDF(po,c.toDataURL("image/png"));}catch(e){printSolicitudPDF(po,"/logo.png");}};
            img.onerror=()=>printSolicitudPDF(po,"/logo.png");
            img.src="/logo.png?"+Date.now();
          }} style={{flex:1,padding:"11px",borderRadius:10,border:"1.5px solid #1a5276",background:"#fff",color:"#1a5276",fontWeight:700,fontSize:13,cursor:"pointer"}}>🖨️ PDF</button>
          <button onClick={()=>exportSolicitudXLSX(po)} style={{flex:1,padding:"11px",borderRadius:10,border:"1.5px solid #1e8449",background:"#fff",color:"#1e8449",fontWeight:700,fontSize:13,cursor:"pointer"}}>📊 Excel</button>
          {isAdmin&&<button onClick={()=>deletePO(po.id)} style={{padding:"11px 16px",borderRadius:10,border:"1.5px solid #fcc",background:"#fff",color:"#c0392b",fontWeight:700,fontSize:13,cursor:"pointer"}}>🗑</button>}
        </div>

        {/* Ingresar mercadería desde solicitud */}
        {isAdmin && po.estado==="cerrada" && (
          <div style={{background:"#d5f5e3",border:"2px solid #1e8449",borderRadius:12,padding:16,marginTop:12,margin:isMobile?"12px 12px 0":"12px 0 0"}}>
            <div style={{fontWeight:800,fontSize:14,color:"#1e8449",marginBottom:8}}>📦 ¿Llegó la mercadería?</div>
            <div style={{fontSize:12,color:"#1e8449",marginBottom:12,lineHeight:1.5}}>
              Podés ingresar el stock recibido directamente desde esta solicitud. Ajustá las cantidades si recibiste diferente a lo pedido.
            </div>
            <IngresarDesdeSolicitud po={po} products={products} onStock={onStockExternal} onDone={()=>setView("lista")}/>
          </div>
        )}
      </div>
    );
  }

  // ── NUEVA SOLICITUD ──
  const addToCart = (p) => setCart(c=>{const ex=c.find(i=>i.pid===p.id);return ex?c.map(i=>i.pid===p.id?{...i,qty:i.qty+1}:i):[...c,{pid:p.id,id:p.id,name:p.name,qty:1}];});
  const setQty = (pid,qty) => {if(qty<=0)setCart(c=>c.filter(i=>i.pid!==pid));else setCart(c=>c.map(i=>i.pid===pid?{...i,qty}:i));};

  return (
    <div style={{display:isMobile?"block":"grid",gridTemplateColumns:"1fr 340px",gap:18,alignItems:"start"}}>
      <div style={{padding:isMobile?"0":"0"}}>
        <div style={{background:"#fff",borderRadius:12,padding:isMobile?"12px":"16px",marginBottom:12,boxShadow:"0 1px 4px #0001",margin:isMobile?"0 12px 12px":"0 0 12px"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:12,color:"#1a5276"}}>📋 Nueva Solicitud de Compra</div>
          <input onChange={e=>{
              const q=e.target.value.toLowerCase();
              // handled by ProductSelector below
            }} placeholder="Usá el buscador de productos abajo" style={{...inputStyle,background:"#f9f9f9",color:"#aaa"}} readOnly/>
        </div>
        <ProductSelector products={products} cart={cart} setCart={setCart} isMobile={isMobile}/>
      </div>

      <div style={{position:isMobile?"static":"sticky",top:16,margin:isMobile?"12px":"0"}}>
        <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 2px 12px #0002",border:"2px solid #d6eaf8"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14,color:"#1a5276"}}>📋 Resumen de Solicitud</div>
          <div style={{background:"#d6eaf8",borderRadius:8,padding:"7px 12px",fontSize:12,color:"#1a5276",marginBottom:14}}>
            ℹ️ Esta solicitud <strong>no modifica el stock</strong>. Es solo para pedir al proveedor.
          </div>

          {/* Items con notas por item */}
          {cart.length===0
            ? <div style={{textAlign:"center",color:"#aaa",fontSize:12,padding:"16px 0"}}>Agregá productos del catálogo</div>
            : cart.map(i=>(
                <div key={i.pid} style={{background:"#f9f9f9",borderRadius:8,padding:"10px 12px",marginBottom:8}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                    <span style={{fontWeight:600,fontSize:12,flex:1,lineHeight:1.3,marginRight:8}}>{i.name}</span>
                    <div style={{display:"flex",alignItems:"center",gap:4}}>
                      <button onClick={()=>setQty(i.pid,i.qty-1)} style={{width:24,height:24,borderRadius:5,border:"1.5px solid #ccc",background:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>-</button>
                      <input type="number" value={i.qty} onChange={e=>setQty(i.pid,+e.target.value||0)} style={{width:40,textAlign:"center",padding:"2px",borderRadius:5,border:"1.5px solid #1a5276",fontWeight:700,fontSize:13,outline:"none"}}/>
                      <button onClick={()=>setQty(i.pid,i.qty+1)} style={{width:24,height:24,borderRadius:5,border:"1.5px solid #ccc",background:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>+</button>
                    </div>
                  </div>
                  <input value={itemNotas[i.pid]||""} onChange={e=>setItemNotas(n=>({...n,[i.pid]:e.target.value}))}
                    placeholder="Obs. (color, medida, marca...)" style={{...inputStyle,fontSize:11,padding:"4px 8px",background:"#fff"}}/>
                </div>
              ))
          }

          <Field label="Notas generales"><textarea value={notas} onChange={e=>setNotas(e.target.value)} placeholder="Urgencia, proveedor sugerido..." style={{...inputStyle,resize:"vertical",minHeight:52,fontSize:12}}/></Field>

          <div style={{display:"flex",gap:8,marginTop:8}}>
            <button onClick={createNew} disabled={!cart.length||saving}
              style={{flex:1,padding:"11px",borderRadius:10,border:"none",fontWeight:800,fontSize:14,cursor:"pointer",
                background:(!cart.length||saving)?"#e5e5e5":"linear-gradient(135deg,#1a5e20,#1e8449)",
                color:(!cart.length||saving)?"#aaa":"#fff"}}>
              {saving?"Guardando...":"📋 Crear Solicitud"}
            </button>
            <button onClick={()=>setView("lista")} style={{padding:"11px 14px",borderRadius:10,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontWeight:600,fontSize:13,cursor:"pointer"}}>Cancelar</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── ALTA DE MERCADERÍA ───────────────────────────────────────────────────────
// ─── BARCODE SCANNER ──────────────────────────────────────────────────────────
function BarcodeScanner({onDetected, onClose}) {
  const videoRef = React.useRef(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(true);
  const [lastCode, setLastCode] = useState("");
  const streamRef = React.useRef(null);
  const readerRef = React.useRef(null);

  React.useEffect(() => {
    let active = true;

    const loadZXing = () => {
      if(window.ZXing) { initScanner(); return; }
      // Try unpkg first, then jsDelivr as fallback
      const tryLoad = (url, fallback) => {
        const script = document.createElement("script");
        script.src = url;
        script.onload = () => { if(active) initScanner(); };
        script.onerror = () => {
          if(fallback) tryLoad(fallback, null);
          else { if(active) { setError("No se pudo cargar el lector. Verificá tu conexión a internet."); setLoading(false); } }
        };
        document.head.appendChild(script);
      };
      tryLoad(
        "https://unpkg.com/@zxing/library@0.19.1/umd/index.min.js",
        "https://cdn.jsdelivr.net/npm/@zxing/library@0.19.1/umd/index.min.js"
      );
    };

    const initScanner = async () => {
      try {
        // @zxing/library exposes ZXingBrowser or ZXing namespace
        const ZXingLib = window.ZXingBrowser || window.ZXing;
        if(!ZXingLib) { setError("No se pudo inicializar el lector. Recargá la pagina."); setLoading(false); return; }

        const hints = new Map();
        const BarcodeFormat = ZXingLib.BarcodeFormat;
        if(BarcodeFormat) {
          hints.set(ZXingLib.DecodeHintType.POSSIBLE_FORMATS, [
            BarcodeFormat.EAN_13, BarcodeFormat.EAN_8,
            BarcodeFormat.CODE_128, BarcodeFormat.CODE_39,
            BarcodeFormat.UPC_A, BarcodeFormat.UPC_E,
          ]);
          hints.set(ZXingLib.DecodeHintType.TRY_HARDER, true);
        }

        const ReaderClass = ZXingLib.BrowserMultiFormatReader || (ZXingLib.browser && ZXingLib.browser.BrowserMultiFormatReader);
        if(!ReaderClass) { setError("Version del lector no compatible. Intentá recargar."); setLoading(false); return; }

        const reader = new ReaderClass(hints);
        readerRef.current = reader;

        // Get camera stream
        const stream = await navigator.mediaDevices.getUserMedia({
          video: { facingMode: "environment", width: {ideal:1280}, height: {ideal:720} }
        });
        streamRef.current = stream;

        if(!active) { stream.getTracks().forEach(t=>t.stop()); return; }

        if(videoRef.current) {
          videoRef.current.srcObject = stream;
          await videoRef.current.play();
          setLoading(false);

          // Start decoding loop
          const decode = async () => {
            if(!active || !videoRef.current) return;
            try {
              const result = await reader.decodeFromVideoElement(videoRef.current);
              if(result && active) {
                const code = result.getText();
                setLastCode(code);
                // Stop and notify
                cleanup();
                onDetected(code);
              }
            } catch(e) {
              // NotFoundException is normal — keep trying
              if(active) setTimeout(decode, 150);
            }
          };
          decode();
        }
      } catch(e) {
        if(active) {
          if(e.name === "NotAllowedError") setError("Permiso de cámara denegado. Habilitalo en la configuración del navegador.");
          else setError("No se pudo acceder a la cámara: " + e.message);
          setLoading(false);
        }
      }
    };

    const cleanup = () => {
      active = false;
      if(streamRef.current) { streamRef.current.getTracks().forEach(t=>t.stop()); streamRef.current = null; }
      if(readerRef.current) { try { readerRef.current.reset(); } catch(e) {} readerRef.current = null; }
    };

    loadZXing();
    return cleanup;
  }, []);

  const handleClose = () => {
    if(streamRef.current) streamRef.current.getTracks().forEach(t=>t.stop());
    if(readerRef.current) try { readerRef.current.reset(); } catch(e) {}
    onClose();
  };

  return (
    <div style={{position:"fixed",inset:0,background:"#000",zIndex:500,display:"flex",flexDirection:"column"}}>
      {/* Header */}
      <div style={{background:"#1a1a1a",padding:"14px 20px",display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
        <div style={{fontWeight:800,fontSize:16,color:"#fff"}}>📷 Escanear código de barras</div>
        <button onClick={handleClose}
          style={{background:"#c0392b",border:"none",color:"#fff",borderRadius:10,padding:"8px 18px",cursor:"pointer",fontSize:14,fontWeight:800}}>
          ✕ Cerrar
        </button>
      </div>

      {/* Camera view */}
      <div style={{flex:1,position:"relative",display:"flex",alignItems:"center",justifyContent:"center",background:"#000"}}>
        <video ref={videoRef} style={{width:"100%",height:"100%",objectFit:"cover"}} muted playsInline/>

        {/* Scanning overlay */}
        {!error && (
          <div style={{position:"absolute",inset:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",pointerEvents:"none"}}>
            {/* Guide box */}
            <div style={{width:"80%",maxWidth:340,height:140,border:"2px solid #fff",borderRadius:12,position:"relative",boxShadow:"0 0 0 2000px #0005"}}>
              {/* Corner markers */}
              {[["0","0"],["0","auto"],["auto","0"],["auto","auto"]].map(([t,b],i)=>(
                <div key={i} style={{position:"absolute",top:t==="0"?-2:"auto",bottom:b==="auto"?-2:"auto",left:i<2?-2:"auto",right:i>=2?-2:"auto",width:20,height:20,borderTop:t==="0"?"3px solid #c0392b":"none",borderBottom:b==="auto"?"3px solid #c0392b":"none",borderLeft:i<2?"3px solid #c0392b":"none",borderRight:i>=2?"3px solid #c0392b":"none"}}/>
              ))}
              {/* Scanning line */}
              <div style={{position:"absolute",top:"50%",left:4,right:4,height:2,background:"#c0392b",boxShadow:"0 0 8px #c0392b",animation:"scan 1.5s ease-in-out infinite"}}/>
            </div>
            <div style={{color:"#fff",fontSize:13,fontWeight:600,marginTop:16,textShadow:"0 1px 4px #000"}}>
              Apuntá al código de barras
            </div>
          </div>
        )}

        {loading && !error && (
          <div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",background:"#000a"}}>
            <div style={{color:"#fff",fontWeight:700,fontSize:15}}>Iniciando cámara...</div>
          </div>
        )}

        {error && (
          <div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",padding:24}}>
            <div style={{background:"#fdecea",borderRadius:12,padding:20,textAlign:"center",maxWidth:320}}>
              <div style={{fontSize:36,marginBottom:10}}>📵</div>
              <div style={{fontWeight:700,color:"#c0392b",fontSize:14,marginBottom:8}}>{error}</div>
              <button onClick={handleClose} style={{padding:"8px 20px",borderRadius:8,border:"none",background:"#c0392b",color:"#fff",fontWeight:700,cursor:"pointer"}}>Volver</button>
            </div>
          </div>
        )}
      </div>

      {/* Bottom bar */}
      <div style={{background:"#1a1a1a",padding:"14px 20px",flexShrink:0,textAlign:"center"}}>
        <div style={{color:"#aaa",fontSize:12}}>Compatible con EAN-13, EAN-8, Code128, UPC</div>
        <button onClick={handleClose}
          style={{marginTop:10,width:"100%",padding:"12px",borderRadius:10,border:"1.5px solid #444",background:"transparent",color:"#fff",fontWeight:700,fontSize:14,cursor:"pointer"}}>
          ← Volver sin escanear
        </button>
      </div>

      <style>{`@keyframes scan{0%,100%{top:10%}50%{top:85%}}`}</style>
    </div>
  );
}

function Compras({products,onStock,isMobile,canScan}) {
  const [search,setSearch]=useState("");
  const [items,setItems]=useState([]);
  const [ok,setOk]=useState(false);
  const [showManual, setShowManual] = useState(false);
  const [manualForm, setManualForm] = useState({sku:"", name:"", qty:1, cost:0});
  const [manualError, setManualError] = useState("");
  const [mStep, setMStep] = useState(1);
  const [showScanner, setShowScanner] = useState(false);
  const [scanMsg, setScanMsg] = useState("");

  const handleBarcode = (code) => {
    setShowScanner(false);
    // Search by barcode field first, then by SKU
    const p = products.find(p=>p.barcode===code || normSKU(p.id)===normSKU(code));
    if(p) {
      if(!items.find(i=>i.pid===p.id)) {
        setItems(x=>[...x,{pid:p.id,name:p.name,qty:1,cost:p.costPrice}]);
        setScanMsg(`✅ ${p.name} agregado`);
      } else {
        setScanMsg(`ℹ️ ${p.name} ya está en la lista`);
      }
    } else {
      setScanMsg(`⚠️ Código "${code}" no encontrado en el catálogo`);
    }
    setTimeout(()=>setScanMsg(""),3000);
  }; // mobile: 1=buscar, 2=detalle

  const found=useMemo(()=>{
    const q=search.toLowerCase();
    return q?products.filter(p=>norm(p.name).includes(norm(q))||normSKU(p.id).includes(normSKU(q))).slice(0,50):[];
  },[products,search]);

  const addI=p=>{if(!items.find(i=>i.pid===p.id))setItems(x=>[...x,{pid:p.id,name:p.name,qty:1,cost:p.costPrice}]);};
  const remI=pid=>setItems(x=>x.filter(i=>i.pid!==pid));
  const updI=(pid,f,v)=>setItems(x=>x.map(i=>i.pid===pid?{...i,[f]:v}:i));
  const totalCost=items.reduce((s,i)=>s+i.qty*i.cost,0);

  const addManual = () => {
    setManualError("");
    if(!manualForm.sku.trim()) { setManualError("El SKU es obligatorio"); return; }
    if(!manualForm.name.trim()) { setManualError("El detalle es obligatorio"); return; }
    if(products.find(p=>p.id===manualForm.sku.trim())) { setManualError("Ese SKU ya existe en el catálogo. Buscalo arriba."); return; }
    if(items.find(i=>i.pid===manualForm.sku.trim())) { setManualError("Ya está en la lista"); return; }
    setItems(x=>[...x,{pid:manualForm.sku.trim(), name:manualForm.name.trim(), qty:+manualForm.qty||1, cost:+manualForm.cost||0, esNuevo:true}]);
    setManualForm({sku:"", name:"", qty:1, cost:0});
    setShowManual(false);
  };

  const submit=()=>{
    items.forEach(i=>onStock(i.pid,+i.qty,+i.cost));
    setItems([]); setSearch(""); setMStep(1);
    setOk(true); setTimeout(()=>setOk(false),2000);
  };

  if(ok) return <div style={{textAlign:"center",padding:80}}><div style={{fontSize:60}}>📦</div><div style={{fontWeight:800,color:"#1e8449",fontSize:20,marginTop:12}}>¡Stock actualizado!</div></div>;

  // ── MOBILE ──────────────────────────────────────────────────────────────────
  if(isMobile) return (
    <div style={{display:"flex",flexDirection:"column"}}>
      {showScanner && <BarcodeScanner onDetected={handleBarcode} onClose={()=>setShowScanner(false)}/>}

      {/* Steps header */}
      <div style={{display:"flex",background:"#fff",borderRadius:10,padding:"10px 14px",marginBottom:12,boxShadow:"0 1px 4px #0001",gap:4}}>
        {[{n:1,l:"Buscar productos"},{n:2,l:"Detalle y confirmar"}].map(s=>(
          <div key={s.n} onClick={()=>mStep>s.n&&setMStep(s.n)}
            style={{flex:1,textAlign:"center",padding:"6px 4px",borderRadius:8,
              background:mStep===s.n?"#eafaf1":mStep>s.n?"#f9f9f9":"transparent",
              cursor:mStep>s.n?"pointer":"default"}}>
            <div style={{fontWeight:800,fontSize:13,color:mStep===s.n?"#1e8449":mStep>s.n?"#888":"#ccc"}}>{s.n}</div>
            <div style={{fontSize:10,color:mStep===s.n?"#1e8449":mStep>s.n?"#888":"#ccc",fontWeight:600}}>{s.l}</div>
          </div>
        ))}
      </div>

      {/* Scan message */}
      {scanMsg && <div style={{background:scanMsg.startsWith("✅")?"#eafaf1":scanMsg.startsWith("ℹ️")?"#eaf4fc":"#fef9e7",border:"1.5px solid",borderColor:scanMsg.startsWith("✅")?"#a9dfbf":scanMsg.startsWith("ℹ️")?"#aed6f1":"#f0d080",borderRadius:10,padding:"10px 14px",marginBottom:10,fontSize:13,fontWeight:600,color:scanMsg.startsWith("✅")?"#1e8449":scanMsg.startsWith("ℹ️")?"#1a5276":"#b7770d"}}>
        {scanMsg}
      </div>}

      {/* Paso 1 — Buscar */}
      {mStep===1 && <>
        <div style={{background:"#fff",borderRadius:12,padding:16,marginBottom:12,boxShadow:"0 1px 4px #0001"}}>
          <div style={{fontWeight:800,fontSize:14,marginBottom:10,color:"#1a1a1a"}}>📥 Buscá los productos recibidos</div>

          {/* Scanner button — only for admin/enabled users */}
          {canScan && <button onClick={()=>setShowScanner(true)}
            style={{width:"100%",padding:"12px",borderRadius:10,border:"none",background:"linear-gradient(135deg,#1a5276,#2980b9)",color:"#fff",fontWeight:800,fontSize:14,cursor:"pointer",marginBottom:10,display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>
            📷 Escanear código de barras
          </button>}

          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 O buscá por nombre o código SKU..."
            style={{width:"100%",padding:"10px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:14,outline:"none",boxSizing:"border-box"}}/>
          {search&&<div style={{fontSize:11,color:"#aaa",marginTop:6}}>{found.length} resultados</div>}

          {/* Producto manual */}
          <div style={{marginTop:12,borderTop:"1px solid #f0f0f0",paddingTop:12}}>
            <button onClick={()=>{setShowManual(s=>!s);setManualError("");}}
              style={{width:"100%",padding:"9px",borderRadius:8,border:"1.5px solid #e67e22",background:showManual?"#fef9e7":"#fff",color:"#e67e22",fontWeight:700,fontSize:13,cursor:"pointer"}}>
              {showManual?"✕ Cancelar":"➕ Producto no está en el catálogo"}
            </button>
            {showManual && (
              <div style={{background:"#fef9e7",border:"1.5px solid #e67e22",borderRadius:10,padding:14,marginTop:10}}>
                <div style={{fontWeight:700,fontSize:13,color:"#e67e22",marginBottom:6}}>⚠️ Producto nuevo (excepcional)</div>
                <Field label="SKU *"><input value={manualForm.sku} onChange={e=>setManualForm(f=>({...f,sku:e.target.value.toUpperCase()}))} placeholder="Ej: PROD-001" style={inputStyle}/></Field>
                <Field label="Detalle *"><input value={manualForm.name} onChange={e=>setManualForm(f=>({...f,name:e.target.value}))} placeholder="Nombre completo del producto" style={inputStyle}/></Field>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                  <Field label="Cantidad"><input type="number" min={1} value={manualForm.qty} onChange={e=>setManualForm(f=>({...f,qty:e.target.value}))} style={inputStyle}/></Field>
                  <Field label="Costo ($)"><input type="number" min={0} value={manualForm.cost} onChange={e=>setManualForm(f=>({...f,cost:e.target.value}))} style={inputStyle}/></Field>
                </div>
                {manualError && <div style={{color:"#c0392b",fontSize:12,margin:"6px 0",fontWeight:600}}>⚠️ {manualError}</div>}
                <button onClick={addManual} style={{width:"100%",padding:"10px",borderRadius:8,border:"none",background:"#e67e22",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer",marginTop:4}}>
                  ➕ Agregar a la compra
                </button>
              </div>
            )}
          </div>
        </div>

        {/* Resultados */}
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {found.length>0 ? found.map(p=>{
            const inL=items.find(i=>i.pid===p.id);
            return (
              <div key={p.id} style={{background:"#fff",borderRadius:10,padding:"12px 14px",display:"flex",alignItems:"center",gap:12,boxShadow:"0 1px 4px #0001",border:inL?"2px solid #1e8449":"2px solid transparent"}}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:700,fontSize:13,color:"#1a1a1a",lineHeight:1.3}}>{p.name}</div>
                  <div style={{fontSize:11,color:"#aaa",marginTop:2}}>{p.id} · Stock: <strong>{p.stock}</strong></div>
                </div>
                <button onClick={()=>addI(p)} disabled={!!inL}
                  style={{padding:"8px 14px",borderRadius:8,border:"none",fontSize:13,fontWeight:700,
                    background:inL?"#d5f5e3":"#1e8449",color:inL?"#1a5276":"#fff",
                    cursor:inL?"not-allowed":"pointer",flexShrink:0,whiteSpace:"nowrap"}}>
                  {inL?"✓ Agregado":"+ Agregar"}
                </button>
              </div>
            );
          }) : <div style={{padding:20,color:"#aaa",fontSize:13,textAlign:"center",background:"#fff",borderRadius:10}}>
            {search?"Sin resultados":"Escribí el nombre o código del producto"}
          </div>}
        </div>

        {/* Barra flotante */}
        {items.length>0 && (
          <div style={{position:"fixed",bottom:0,left:0,right:0,background:"linear-gradient(135deg,#1a5e20,#1e8449)",color:"#fff",padding:"13px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",zIndex:200,boxShadow:"0 -3px 16px #0003"}}>
            <div>
              <div style={{fontWeight:800,fontSize:15}}>{fARS(totalCost)}</div>
              <div style={{fontSize:11,opacity:.85}}>{items.length} producto{items.length!==1?"s":""} agregado{items.length!==1?"s":""}</div>
            </div>
            <button onClick={()=>setMStep(2)} style={{padding:"10px 20px",borderRadius:10,border:"none",background:"#fff",color:"#1e8449",fontWeight:800,fontSize:14,cursor:"pointer"}}>
              Ver detalle →
            </button>
          </div>
        )}
      </>}

      {/* Paso 2 — Detalle y confirmar */}
      {mStep===2 && (
        <div style={{paddingBottom:80}}>
          <div style={{fontWeight:800,fontSize:14,marginBottom:12,color:"#1a1a1a"}}>🧾 Detalle de Compra</div>
          {items.map(it=>(
            <div key={it.pid} style={{background:"#fff",borderRadius:12,padding:14,marginBottom:10,boxShadow:"0 1px 4px #0001",border:it.esNuevo?"1.5px solid #e67e22":"none"}}>
              {it.esNuevo && <div style={{fontSize:10,color:"#e67e22",fontWeight:700,marginBottom:4}}>⚠️ NUEVO en catálogo</div>}
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
                <div style={{flex:1,marginRight:8}}>
                  <div style={{fontWeight:700,fontSize:13,lineHeight:1.3}}>{it.name}</div>
                  <div style={{fontSize:11,color:"#aaa",marginTop:2}}>SKU: {it.pid}</div>
                </div>
                <button onClick={()=>remI(it.pid)} style={{background:"none",border:"none",cursor:"pointer",color:RED,fontSize:22,lineHeight:1,padding:0}}>×</button>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                <div>
                  <div style={{fontSize:11,color:"#888",fontWeight:700,marginBottom:4}}>CANTIDAD</div>
                  <div style={{display:"flex",alignItems:"center",gap:6}}>
                    <button onClick={()=>updI(it.pid,"qty",Math.max(1,it.qty-1))} style={{width:32,height:32,borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",fontWeight:800,fontSize:16,cursor:"pointer"}}>−</button>
                    <input type="number" min={1} value={it.qty} onChange={e=>updI(it.pid,"qty",+e.target.value)}
                      style={{width:48,textAlign:"center",padding:"5px",borderRadius:7,border:"1.5px solid #e5e5e5",fontWeight:700,fontSize:14,outline:"none"}}/>
                    <button onClick={()=>updI(it.pid,"qty",it.qty+1)} style={{width:32,height:32,borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",fontWeight:800,fontSize:16,cursor:"pointer"}}>+</button>
                  </div>
                </div>
                <div>
                  <div style={{fontSize:11,color:"#888",fontWeight:700,marginBottom:4}}>P. COSTO ($)</div>
                  <input type="number" value={it.cost} onChange={e=>updI(it.pid,"cost",+e.target.value)}
                    style={{width:"100%",padding:"8px 10px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:14,fontWeight:700,outline:"none",boxSizing:"border-box"}}/>
                </div>
              </div>
              <div style={{fontSize:12,color:"#666",marginTop:8,display:"flex",gap:16}}>
                <span>Subtotal: <strong>{fARS(it.qty*it.cost)}</strong></span>
                <span>Venta est.: <strong style={{color:RED}}>{fARS(it.cost*1.5)}</strong></span>
              </div>
            </div>
          ))}

          <button onClick={()=>setMStep(1)} style={{width:"100%",padding:"10px",borderRadius:10,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontWeight:600,fontSize:13,cursor:"pointer",marginBottom:8}}>
            ← Seguir agregando productos
          </button>

          {/* Barra flotante confirmar */}
          <div style={{position:"fixed",bottom:0,left:0,right:0,background:"linear-gradient(135deg,#1a5e20,#1e8449)",color:"#fff",padding:"13px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",zIndex:200,boxShadow:"0 -3px 16px #0003"}}>
            <div>
              <div style={{fontWeight:800,fontSize:15}}>{fARS(totalCost)}</div>
              <div style={{fontSize:11,opacity:.85}}>{items.length} producto{items.length!==1?"s":""}</div>
            </div>
            <button onClick={submit} style={{padding:"10px 20px",borderRadius:10,border:"none",background:"#fff",color:"#1e8449",fontWeight:800,fontSize:14,cursor:"pointer"}}>
              📦 Ingresar al Stock
            </button>
          </div>
        </div>
      )}
    </div>
  );

  // ── DESKTOP ──────────────────────────────────────────────────────────────────
  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 340px",gap:18,alignItems:"start"}}>
      {showScanner && <BarcodeScanner onDetected={handleBarcode} onClose={()=>setShowScanner(false)}/>}
      <div>
        <div style={{background:"#fff",borderRadius:12,padding:16,marginBottom:12,boxShadow:"0 1px 4px #0001"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:12}}>📥 Ingresar al Stock</div>
          {scanMsg && <div style={{background:scanMsg.startsWith("✅")?"#eafaf1":"#fef9e7",borderRadius:8,padding:"8px 12px",fontSize:13,fontWeight:600,marginBottom:10,color:scanMsg.startsWith("✅")?"#1e8449":"#b7770d"}}>{scanMsg}</div>}
          <div style={{display:"flex",gap:8,marginBottom:10}}>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscá los productos que recibiste..."
              style={{flex:1,padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none"}}/>
            {canScan && <button onClick={()=>setShowScanner(true)}
              style={{padding:"8px 14px",borderRadius:8,border:"none",background:"linear-gradient(135deg,#1a5276,#2980b9)",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer",whiteSpace:"nowrap"}}>
              📷 Escanear
            </button>}
          </div>
          {search&&<div style={{fontSize:11,color:"#aaa",marginBottom:6}}>{found.length} resultados</div>}
          <div style={{marginTop:12,borderTop:"1px solid #f0f0f0",paddingTop:12}}>
            <button onClick={()=>{setShowManual(s=>!s);setManualError("");}}
              style={{padding:"6px 14px",borderRadius:8,border:"1.5px solid #e67e22",background:showManual?"#fef9e7":"#fff",color:"#e67e22",fontWeight:700,fontSize:12,cursor:"pointer"}}>
              {showManual?"✕ Cancelar":"➕ Producto no está en el catálogo"}
            </button>
            {showManual && (
              <div style={{background:"#fef9e7",border:"1.5px solid #e67e22",borderRadius:10,padding:14,marginTop:10}}>
                <div style={{fontWeight:700,fontSize:13,color:"#e67e22",marginBottom:10}}>⚠️ Producto nuevo (excepcional)</div>
                <div style={{fontSize:11,color:"#888",marginBottom:12,lineHeight:1.5}}>Solo usá esto si el producto no figura en el Excel.</div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
                  <div><div style={{fontSize:10,fontWeight:700,color:"#666",marginBottom:4}}>SKU *</div>
                    <input value={manualForm.sku} onChange={e=>setManualForm(f=>({...f,sku:e.target.value.toUpperCase()}))} placeholder="Ej: PROD-001" style={{...inputStyle,fontSize:12,padding:"6px 10px"}}/></div>
                  <div><div style={{fontSize:10,fontWeight:700,color:"#666",marginBottom:4}}>Cantidad *</div>
                    <input type="number" min={1} value={manualForm.qty} onChange={e=>setManualForm(f=>({...f,qty:e.target.value}))} style={{...inputStyle,fontSize:12,padding:"6px 10px"}}/></div>
                </div>
                <div style={{marginBottom:8}}><div style={{fontSize:10,fontWeight:700,color:"#666",marginBottom:4}}>Detalle *</div>
                  <input value={manualForm.name} onChange={e=>setManualForm(f=>({...f,name:e.target.value}))} placeholder="Nombre completo" style={{...inputStyle,fontSize:12,padding:"6px 10px"}}/></div>
                <div style={{marginBottom:12}}><div style={{fontSize:10,fontWeight:700,color:"#666",marginBottom:4}}>Precio de costo ($)</div>
                  <input type="number" min={0} value={manualForm.cost} onChange={e=>setManualForm(f=>({...f,cost:e.target.value}))} style={{...inputStyle,fontSize:12,padding:"6px 10px"}}/></div>
                {manualError && <div style={{color:"#c0392b",fontSize:12,marginBottom:8,fontWeight:600}}>⚠️ {manualError}</div>}
                <button onClick={addManual} style={{width:"100%",padding:"8px",borderRadius:8,border:"none",background:"#e67e22",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}>➕ Agregar a la compra</button>
              </div>
            )}
          </div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(195px,1fr))",gap:10}}>
          {found.length>0
            ? found.map(p=>{
                const inL=items.find(i=>i.pid===p.id);
                return <div key={p.id} style={{background:"#fff",borderRadius:10,padding:14,border:inL?"2px solid #1e8449":"2px solid transparent",boxShadow:"0 1px 4px #0001"}}>
                  <div style={{fontWeight:700,fontSize:12,color:"#1a1a1a",marginBottom:3,lineHeight:1.3}}>{p.name}</div>
                  <div style={{fontSize:10,color:"#aaa",marginBottom:8}}>{p.id} · Stock: <strong>{p.stock}</strong></div>
                  <button onClick={()=>addI(p)} disabled={!!inL} style={{width:"100%",padding:"7px",borderRadius:7,border:"none",fontSize:12,fontWeight:700,background:inL?"#d5f5e3":"#1e8449",color:inL?"#1a5276":"#fff",cursor:inL?"not-allowed":"pointer"}}>{inL?"✓ Agregado":"+ Agregar"}</button>
                </div>;
              })
            : <div style={{padding:20,color:"#aaa",fontSize:13}}>{search?"Sin resultados.":"Escribí el nombre del producto a ingresar."}</div>
          }
        </div>
      </div>
      <div style={{position:"sticky",top:16}}>
        <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 2px 12px #0002"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14}}>🧾 Detalle de Compra</div>
          {items.length===0
            ? <div style={{textAlign:"center",color:"#aaa",fontSize:12,padding:"16px 0"}}>Seleccioná productos</div>
            : items.map(it=>(
                <div key={it.pid} style={{background:"#f9f9f9",borderRadius:8,padding:12,marginBottom:10,border:it.esNuevo?"1.5px solid #e67e22":"none"}}>
                  {it.esNuevo && <div style={{fontSize:10,color:"#e67e22",fontWeight:700,marginBottom:4}}>⚠️ NUEVO en catálogo</div>}
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
                    <span style={{fontWeight:600,fontSize:12,flex:1,marginRight:6,lineHeight:1.3}}>{it.name}</span>
                    <button onClick={()=>remI(it.pid)} style={{background:"none",border:"none",cursor:"pointer",color:RED,fontSize:18,lineHeight:1}}>×</button>
                  </div>
                  <div style={{fontSize:10,color:"#aaa",marginBottom:6}}>SKU: {it.pid}</div>
                  <div style={{display:"flex",gap:8}}>
                    <div style={{flex:1}}><div style={{fontSize:10,color:"#aaa",marginBottom:3}}>Cantidad</div>
                      <input type="number" min={1} value={it.qty} onChange={e=>updI(it.pid,"qty",+e.target.value)} style={{...inputStyle,padding:"5px 7px",fontSize:12}}/></div>
                    <div style={{flex:1}}><div style={{fontSize:10,color:"#aaa",marginBottom:3}}>P. Costo ($)</div>
                      <input type="number" value={it.cost} onChange={e=>updI(it.pid,"cost",+e.target.value)} style={{...inputStyle,padding:"5px 7px",fontSize:12}}/></div>
                  </div>
                  <div style={{fontSize:11,color:"#666",marginTop:6}}>Subtotal: <strong>{fARS(it.qty*it.cost)}</strong> · Venta: <strong style={{color:RED}}>{fARS(it.cost*1.5)}</strong></div>
                </div>
              ))
          }
          <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:16,color:RED,padding:"10px 0",borderTop:"2px solid #f5f5f5",margin:"8px 0 14px"}}>
            <span>Total compra</span><span>{fARS(totalCost)}</span>
          </div>
          <button onClick={submit} disabled={!items.length} style={{width:"100%",padding:"11px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:!items.length?"#e5e5e5":"linear-gradient(135deg,#1a5e20,#1e8449)",color:!items.length?"#aaa":"#fff"}}>📦 Ingresar al Stock</button>
        </div>
      </div>
    </div>
  );
}

// ─── SANDBOX STOCK MANAGER ────────────────────────────────────────────────────
function SandboxStockManager({products, sandboxStock, setSandboxStock}) {
  const [search, setSearch] = useState("");
  const [edited, setEdited] = useState({});
  const [page, setPage] = useState(1);
  const PAGE_SIZE = 50;

  const filtered = useMemo(()=>{
    const q = norm(search);
    return products.filter(p=>!q||norm(p.name).includes(q)||normSKU(p.id).includes(normSKU(search)));
  }, [products, search]);

  const totalPages = Math.ceil(filtered.length / PAGE_SIZE);
  const shown = filtered.slice((page-1)*PAGE_SIZE, page*PAGE_SIZE);

  // Reset to page 1 on search change
  const handleSearch = (v) => { setSearch(v); setPage(1); };

  const getSandbox = (pid) => edited[pid] !== undefined ? edited[pid] : (sandboxStock[pid] ?? products.find(p=>p.id===pid)?.stock ?? 0);
  const setVal = (pid, val) => setEdited(e=>({...e,[pid]:Math.max(0,+val||0)}));
  const applyChanges = () => { setSandboxStock(prev=>({...prev,...edited})); setEdited({}); };
  const hasChanges = Object.keys(edited).length > 0;

  return (
    <div>
      <input value={search} onChange={e=>handleSearch(e.target.value)}
        placeholder="🔍 Buscar producto para editar su stock sandbox..."
        style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1.5px solid #d7bde2",fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:12,background:"#fff"}}/>

      {hasChanges && (
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:"#6c3483",borderRadius:10,padding:"10px 16px",marginBottom:12}}>
          <span style={{color:"#fff",fontWeight:700,fontSize:13}}>{Object.keys(edited).length} cambio{Object.keys(edited).length!==1?"s":""} sin guardar</span>
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>setEdited({})} style={{padding:"6px 12px",borderRadius:7,border:"1px solid #ffffff44",background:"transparent",color:"#fff",fontSize:12,cursor:"pointer"}}>Descartar</button>
            <button onClick={applyChanges} style={{padding:"6px 14px",borderRadius:7,border:"none",background:"#fff",color:"#6c3483",fontWeight:800,fontSize:12,cursor:"pointer"}}>✅ Aplicar</button>
          </div>
        </div>
      )}

      {/* Info y paginación arriba */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10,flexWrap:"wrap",gap:8}}>
        <span style={{fontSize:12,color:"#888"}}>{filtered.length} productos · página {page} de {totalPages}</span>
        <div style={{display:"flex",gap:6,alignItems:"center"}}>
          <button onClick={()=>setPage(p=>Math.max(1,p-1))} disabled={page===1}
            style={{padding:"5px 12px",borderRadius:7,border:"1.5px solid #d7bde2",background:page===1?"#f5f5f5":"#fff",color:page===1?"#ccc":"#6c3483",fontWeight:700,fontSize:12,cursor:page===1?"not-allowed":"pointer"}}>← Anterior</button>
          {Array.from({length:Math.min(totalPages,5)},(_,i)=>{
            // Show pages around current
            let p = page <= 3 ? i+1 : page - 2 + i;
            if(p > totalPages) return null;
            return <button key={p} onClick={()=>setPage(p)}
              style={{width:32,height:32,borderRadius:7,border:"1.5px solid",cursor:"pointer",fontSize:12,fontWeight:700,
                borderColor:page===p?"#6c3483":"#d7bde2",background:page===p?"#6c3483":"#fff",color:page===p?"#fff":"#6c3483"}}>{p}</button>;
          })}
          <button onClick={()=>setPage(p=>Math.min(totalPages,p+1))} disabled={page===totalPages}
            style={{padding:"5px 12px",borderRadius:7,border:"1.5px solid #d7bde2",background:page===totalPages?"#f5f5f5":"#fff",color:page===totalPages?"#ccc":"#6c3483",fontWeight:700,fontSize:12,cursor:page===totalPages?"not-allowed":"pointer"}}>Siguiente →</button>
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:8}}>
        {shown.map(p=>{
          const sbQty = getSandbox(p.id);
          const realQty = sandboxStock[p.id] ?? p.stock;
          const changed = edited[p.id] !== undefined && edited[p.id] !== realQty;
          return (
            <div key={p.id} style={{background:changed?"#f5eef8":"#fff",borderRadius:10,padding:"10px 12px",border:changed?"1.5px solid #9b59b6":"1.5px solid #e8daef",display:"flex",alignItems:"center",gap:10}}>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontWeight:600,fontSize:12,color:"#1a1a1a",lineHeight:1.3,marginBottom:2}}>{p.name}</div>
                <div style={{fontSize:10,color:"#aaa"}}>{p.id} · Real: {p.stock}</div>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:5,flexShrink:0}}>
                <button onClick={()=>setVal(p.id,sbQty-1)} style={{width:26,height:26,borderRadius:6,border:"1.5px solid #d7bde2",background:"#fff",fontWeight:800,cursor:"pointer",fontSize:14,color:"#6c3483"}}>−</button>
                <input type="number" value={sbQty} onChange={e=>setVal(p.id,e.target.value)}
                  style={{width:44,textAlign:"center",padding:"3px",borderRadius:6,border:`1.5px solid ${changed?"#9b59b6":"#d7bde2"}`,fontWeight:700,fontSize:13,outline:"none",color:changed?"#6c3483":"#1a1a1a",background:changed?"#f5eef8":"#fff"}}/>
                <button onClick={()=>setVal(p.id,sbQty+1)} style={{width:26,height:26,borderRadius:6,border:"1.5px solid #d7bde2",background:"#fff",fontWeight:800,cursor:"pointer",fontSize:14,color:"#6c3483"}}>+</button>
              </div>
            </div>
          );
        })}
      </div>

      {/* Paginación abajo */}
      {totalPages>1 && <div style={{display:"flex",justifyContent:"center",gap:6,marginTop:14,flexWrap:"wrap"}}>
        <button onClick={()=>setPage(p=>Math.max(1,p-1))} disabled={page===1}
          style={{padding:"5px 12px",borderRadius:7,border:"1.5px solid #d7bde2",background:page===1?"#f5f5f5":"#fff",color:page===1?"#ccc":"#6c3483",fontWeight:700,fontSize:12,cursor:page===1?"not-allowed":"pointer"}}>← Anterior</button>
        <button onClick={()=>setPage(p=>Math.min(totalPages,p+1))} disabled={page===totalPages}
          style={{padding:"5px 12px",borderRadius:7,border:"1.5px solid #d7bde2",background:page===totalPages?"#f5f5f5":"#fff",color:page===totalPages?"#ccc":"#6c3483",fontWeight:700,fontSize:12,cursor:page===totalPages?"not-allowed":"pointer"}}>Siguiente →</button>
      </div>}
    </div>
  );
}

// ─── ADMIN PANEL ──────────────────────────────────────────────────────────────
function AdminPanel({users,setUsers,vendors,setVendors,products,setProducts,stockLog,setStockLog,notifs,setNotifs,activity,setActivity,orders,priceLists,setPriceLists,isMobile,sandboxStock,setSandboxStock}) {
  const [section, setSection] = useState("home");

  const SECTIONS = [
    {k:"home",        label:"Administración",   icon:"🏠"},
    {k:"ventas",      label:"Ventas",           icon:"📈"},
    {k:"sandbox",     label:"Demo Sandbox",     icon:"🧪"},
    {k:"sandboxstock",label:"Stock Sandbox",     icon:"📦"},
    {k:"activity",    label:"Actividad",        icon:"📝"},
    {k:"vendors",     label:"Vendedores",       icon:"👥"},
    {k:"users",       label:"Usuarios",         icon:"🔐"},
    {k:"pricelists",  label:"Listas de Precio", icon:"💲"},
    {k:"excel",       label:"Importar Precios", icon:"📊"},
    {k:"notifcfg",    label:"Notificaciones",   icon:"🔔"},
  ];

  const sandboxOrders = orders.filter(o=>o.isSandbox);
  const realOrders    = orders.filter(o=>!o.isSandbox);

  // Mobile: show icon grid when no section selected, back button when inside
  if(isMobile && !section) return (
    <div style={{padding:12}}>
      <div style={{fontWeight:800,fontSize:16,marginBottom:12}}>★ Administración</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        {SECTIONS.map(s=>(
          <div key={s.k} onClick={()=>setSection(s.k)}
            style={{background:"#fff",borderRadius:14,padding:"18px 14px",textAlign:"center",boxShadow:"0 1px 6px #0001",cursor:"pointer"}}>
            <div style={{fontSize:32,marginBottom:8}}>{s.icon}</div>
            <div style={{fontSize:12,fontWeight:700,color:"#1a1a1a"}}>{s.label}</div>
            <div style={{fontSize:10,color:"#aaa",marginTop:2}}>{s.sub}</div>
          </div>
        ))}
      </div>
    </div>
  );

  return (
    <div>
      {isMobile && section && (
        <div style={{display:"flex",alignItems:"center",gap:10,padding:"12px 12px 0",marginBottom:8}}>
          <button onClick={()=>setSection(null)} style={{padding:"7px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",fontWeight:600,fontSize:12,cursor:"pointer",color:"#555"}}>← Volver</button>
          <div style={{fontWeight:800,fontSize:15}}>{SECTIONS.find(s=>s.k===section)?.icon} {SECTIONS.find(s=>s.k===section)?.label}</div>
        </div>
      )}
      {!isMobile && <div style={{background:"#fff",borderRadius:12,padding:4,marginBottom:16,display:"flex",gap:4,boxShadow:"0 1px 4px #0001",flexWrap:"wrap"}}>
        {SECTIONS.map(s=>(
          <button key={s.k} onClick={()=>setSection(s.k)} style={{flex:1,minWidth:120,padding:"10px 16px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:section===s.k?`linear-gradient(135deg,${REDD},${RED})`:"transparent",color:section===s.k?"#fff":"#555",transition:"all .15s"}}>
            {s.icon} {s.label}
          </button>
        ))}
      </div>}
      {section==="home" && (
        <div>
          <div style={{fontWeight:800,fontSize:16,marginBottom:16,color:"#1a1a1a"}}>🏠 Central de Administración</div>
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr 1fr":"repeat(auto-fill,minmax(160px,1fr))",gap:10,marginBottom:24}}>
            {SECTIONS.filter(s=>s.k!=="home").map(s=>(
              <div key={s.k} onClick={()=>setSection(s.k)}
                style={{background:"#fff",borderRadius:12,padding:"18px 14px",boxShadow:"0 1px 6px #0001",cursor:"pointer",textAlign:"center",border:"2px solid transparent",transition:"border .15s"}}
                onMouseEnter={e=>e.currentTarget.style.border="2px solid #c0392b"}
                onMouseLeave={e=>e.currentTarget.style.border="2px solid transparent"}>
                <div style={{fontSize:30,marginBottom:8}}>{s.icon}</div>
                <div style={{fontWeight:700,fontSize:13,color:"#1a1a1a"}}>{s.label}</div>
              </div>
            ))}
          </div>
        </div>
      )}
      {section==="ventas"      && <VentasPanel    orders={realOrders}/>}
      {section==="sandboxstock" && (
        <div>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10,marginBottom:16}}>
            <div>
              <div style={{fontWeight:800,fontSize:15,color:"#6c3483"}}>📦 Stock Sandbox</div>
              <div style={{fontSize:12,color:"#888",marginTop:2}}>Configurá el stock que verá el vendedor Prueba para simular pedidos</div>
            </div>
            <button onClick={()=>{const sb={};products.forEach(p=>{sb[p.id]=p.stock;});updateSandboxStock(sb);}}
              style={{padding:"7px 14px",borderRadius:8,border:"1.5px solid #9b59b6",background:"#fff",color:"#6c3483",fontWeight:700,fontSize:12,cursor:"pointer"}}>
              🔄 Resetear al stock real
            </button>
          </div>
          <SandboxStockManager products={products} sandboxStock={sandboxStock} setSandboxStock={setSandboxStock}/>
        </div>
      )}
      {section==="sandbox"     && (
        <div>
          <div style={{background:"#f5eef8",border:"1.5px solid #9b59b6",borderRadius:12,padding:"12px 18px",marginBottom:16,display:"flex",alignItems:"center",gap:12}}>
            <span style={{fontSize:28}}>🧪</span>
            <div>
              <div style={{fontWeight:800,fontSize:14,color:"#6c3483"}}>Dashboard Demo — Vendedor Prueba</div>
              <div style={{fontSize:12,color:"#888",marginTop:2}}>Mostrá a los vendedores cómo funciona el sistema. Estos datos son del sandbox y no afectan las estadísticas reales.</div>
            </div>
          </div>
          {sandboxOrders.length===0
            ? <div style={{textAlign:"center",padding:60,background:"#fff",borderRadius:12,color:"#aaa"}}>
                <div style={{fontSize:48,marginBottom:8}}>🧪</div>
                <div style={{fontWeight:700,fontSize:15}}>Sin datos de prueba aún</div>
                <div style={{fontSize:13,marginTop:4}}>Ingresá con el usuario Prueba y generá algunos pedidos para ver el demo aquí.</div>
              </div>
            : <VentasPanel orders={sandboxOrders}/>
          }
        </div>
      )}
      {section==="activity"    && <ActivityPanel  activity={activity} setActivity={setActivity}/>}
      {section==="vendors"     && <VendorsPanel   vendors={vendors} setVendors={setVendors}/>}
      {section==="users"       && <UsersPanel     users={users} setUsers={setUsers} vendors={vendors} priceLists={priceLists}/>}
      {section==="pricelists"  && <PriceListsPanel priceLists={priceLists} setPriceLists={setPriceLists}/>}
      {section==="excel"       && <ExcelPanel     products={products} setProducts={setProducts}/>}
      {section==="notifcfg"    && <NotifConfig    users={users} setUsers={setUsers} notifs={notifs} setNotifs={setNotifs}/>}
    </div>
  );
}

// ── Panel de Ventas ───────────────────────────────────────────────────────────
function VentasPanel({orders}) {
  const vendidas = orders.filter(o=>o.stage==="entregado" && !o.isSandbox);
  const [periodo, setPeriodo] = useState("mes"); // "dia"|"mes"|"vendedor"

  // ── helpers ──
  const totalGeneral = vendidas.reduce((s,o)=>s+o.total,0);
  const cantGeneral  = vendidas.length;

  // Por día (últimos 30 días)
  const byDay = useMemo(()=>{
    const map = {};
    vendidas.forEach(o=>{
      const d = o.date||"Sin fecha";
      map[d] = (map[d]||0) + o.total;
    });
    return Object.entries(map).sort((a,b)=>a[0].localeCompare(b[0])).slice(-30);
  },[vendidas]);

  // Por mes
  const byMonth = useMemo(()=>{
    const map = {};
    vendidas.forEach(o=>{
      const parts = (o.date||"").split("/");
      const key = parts.length===3 ? `${parts[1]}/${parts[2]}` : o.date||"Sin fecha";
      map[key] = (map[key]||0) + o.total;
    });
    return Object.entries(map).sort((a,b)=>a[0].localeCompare(b[0]));
  },[vendidas]);

  // Por vendedor
  const byVendedor = useMemo(()=>{
    const map = {};
    vendidas.forEach(o=>{
      const v = o.vendedor||"Sin vendedor";
      if(!map[v]) map[v]={total:0,cant:0};
      map[v].total += o.total;
      map[v].cant  += 1;
    });
    return Object.entries(map).sort((a,b)=>b[1].total-a[1].total);
  },[vendidas]);

  // ── mini bar chart ──
  const BarChart = ({data, colorFn}) => {
    const max = Math.max(...data.map(([,v])=>typeof v==="number"?v:v.total), 1);
    return (
      <div style={{display:"flex",flexDirection:"column",gap:6}}>
        {data.map(([label,val],i)=>{
          const amount = typeof val==="number" ? val : val.total;
          const cant   = typeof val==="object" ? val.cant : null;
          const pct    = (amount/max)*100;
          const color  = colorFn ? colorFn(i) : RED;
          return (
            <div key={label} style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:110,fontSize:11,color:"#555",textAlign:"right",flexShrink:0,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{label}</div>
              <div style={{flex:1,background:"#f5f5f5",borderRadius:4,overflow:"hidden",height:22}}>
                <div style={{width:`${pct}%`,background:color,height:"100%",borderRadius:4,minWidth:4,transition:"width .3s"}}/>
              </div>
              <div style={{width:130,fontSize:12,fontWeight:700,color:"#1a1a1a",whiteSpace:"nowrap"}}>
                {fARS(amount)}{cant!==null?<span style={{color:"#aaa",fontWeight:400,marginLeft:4}}>({cant})</span>:null}
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  const COLORS = ["#c0392b","#e74c3c","#e67e22","#f1c40f","#1e8449","#1a5276","#6c3483","#2980b9","#16a085","#7f8c8d"];

  if(vendidas.length===0) return (
    <div style={{textAlign:"center",padding:60,background:"#fff",borderRadius:12,color:"#aaa"}}>
      <div style={{fontSize:48,marginBottom:8}}>📈</div>
      <div style={{fontWeight:700,fontSize:15}}>No hay pedidos entregados aún</div>
      <div style={{fontSize:13,marginTop:4}}>Las ventas aparecerán aquí cuando los pedidos pasen a "Entregado"</div>
    </div>
  );

  return (
    <div>
      {/* KPIs */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:12,marginBottom:20}}>
        {[
          {label:"Total vendido",    value:fARS(totalGeneral),      color:RED,   icon:"💰"},
          {label:"Pedidos entregados",value:cantGeneral,            color:"#1a5276", icon:"📦"},
          {label:"Ticket promedio",  value:fARS(totalGeneral/Math.max(cantGeneral,1)), color:"#1e8449", icon:"📊"},
          {label:"Vendedores activos",value:byVendedor.length,      color:"#6c3483", icon:"👥"},
        ].map(k=>(
          <div key={k.label} style={{background:"#fff",borderRadius:12,padding:"16px 18px",boxShadow:"0 1px 6px #0001",borderLeft:`4px solid ${k.color}`}}>
            <div style={{fontSize:11,color:"#888",fontWeight:600,marginBottom:4}}>{k.icon} {k.label}</div>
            <div style={{fontSize:20,fontWeight:900,color:k.color}}>{k.value}</div>
          </div>
        ))}
      </div>

      {/* Tab selector */}
      <div style={{background:"#fff",borderRadius:12,padding:4,marginBottom:16,display:"flex",gap:4,boxShadow:"0 1px 4px #0001"}}>
        {[{k:"dia",l:"📅 Por día"},{k:"mes",l:"📆 Por mes"},{k:"vendedor",l:"👤 Por vendedor"}].map(t=>(
          <button key={t.k} onClick={()=>setPeriodo(t.k)} style={{flex:1,padding:"9px",borderRadius:9,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:periodo===t.k?`linear-gradient(135deg,${REDD},${RED})`:"transparent",color:periodo===t.k?"#fff":"#555"}}>
            {t.l}
          </button>
        ))}
      </div>

      {/* Chart */}
      <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 1px 4px #0001"}}>
        {periodo==="dia" && <>
          <div style={{fontWeight:800,fontSize:15,marginBottom:16}}>📅 Ventas por día (últimos 30 días)</div>
          {byDay.length===0
            ? <div style={{color:"#aaa",textAlign:"center",padding:20}}>Sin datos</div>
            : <BarChart data={byDay} colorFn={()=>RED}/>
          }
        </>}
        {periodo==="mes" && <>
          <div style={{fontWeight:800,fontSize:15,marginBottom:16}}>📆 Ventas por mes</div>
          {byMonth.length===0
            ? <div style={{color:"#aaa",textAlign:"center",padding:20}}>Sin datos</div>
            : <BarChart data={byMonth} colorFn={()=>"#1a5276"}/>
          }
        </>}
        {periodo==="vendedor" && <>
          <div style={{fontWeight:800,fontSize:15,marginBottom:16}}>👤 Dashboard por vendedor</div>
          {byVendedor.length===0
            ? <div style={{color:"#aaa",textAlign:"center",padding:20}}>Sin datos</div>
            : <>
              <BarChart data={byVendedor} colorFn={(i)=>COLORS[i%COLORS.length]}/>
              {/* Tabla detalle por vendedor */}
              <div style={{marginTop:20,overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead><tr style={{background:"#f9f9f9"}}>
                    {["#","Vendedor","Pedidos","Total","Promedio","% del total"].map(h=>(
                      <th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:700,color:"#888",fontSize:11,textTransform:"uppercase"}}>{h}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {byVendedor.map(([v,d],i)=>(
                      <tr key={v} style={{borderTop:"1px solid #f5f5f5"}}>
                        <td style={{padding:"10px 12px",fontWeight:800,color:COLORS[i%COLORS.length],width:32}}>#{i+1}</td>
                        <td style={{padding:"10px 12px",fontWeight:700}}>
                          <span style={{display:"inline-block",width:10,height:10,borderRadius:"50%",background:COLORS[i%COLORS.length],marginRight:8}}/>
                          {v}
                        </td>
                        <td style={{padding:"10px 12px",color:"#666"}}>{d.cant}</td>
                        <td style={{padding:"10px 12px",fontWeight:800,color:RED}}>{fARS(d.total)}</td>
                        <td style={{padding:"10px 12px",color:"#666"}}>{fARS(d.total/d.cant)}</td>
                        <td style={{padding:"10px 12px"}}>
                          <div style={{display:"flex",alignItems:"center",gap:8}}>
                            <div style={{flex:1,background:"#f5f5f5",borderRadius:4,height:8,overflow:"hidden"}}>
                              <div style={{width:`${(d.total/totalGeneral)*100}%`,background:COLORS[i%COLORS.length],height:"100%",borderRadius:4}}/>
                            </div>
                            <span style={{fontSize:11,color:"#888",minWidth:32}}>{((d.total/totalGeneral)*100).toFixed(1)}%</span>
                          </div>
                        </td>
                      </tr>
                    ))}
                    {/* Total row */}
                    <tr style={{borderTop:"2px solid #e5e5e5",background:"#f9f9f9"}}>
                      <td colSpan={2} style={{padding:"10px 12px",fontWeight:800,fontSize:13}}>TOTAL</td>
                      <td style={{padding:"10px 12px",fontWeight:800}}>{cantGeneral}</td>
                      <td style={{padding:"10px 12px",fontWeight:800,color:RED}}>{fARS(totalGeneral)}</td>
                      <td style={{padding:"10px 12px",fontWeight:800,color:"#666"}}>{fARS(totalGeneral/Math.max(cantGeneral,1))}</td>
                      <td style={{padding:"10px 12px",fontWeight:800}}>100%</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </>
          }
        </>}
      </div>
    </div>
  );
}

// ── Panel de Actividad ────────────────────────────────────────────────────────
function ActivityPanel({activity, setActivity}) {
  const [search, setSearch] = useState("");
  const [filterTipo, setFilterTipo] = useState("todos");
  const [confirmClear, setConfirmClear] = useState(false);

  const TIPOS = {
    "Nuevo pedido":                {color:"#1a5276", bg:"#d6eaf8",  icon:"🛒"},
    "Cambio estado: Confirmado":   {color:"#1a5276", bg:"#d6eaf8",  icon:"✅"},
    "Cambio estado: En Armado":    {color:"#6c3483", bg:"#e8daef",  icon:"📦"},
    "Cambio estado: Entregado":    {color:"#1e8449", bg:"#d5f5e3",  icon:"🎉"},
    "Cambio estado: Reserva":      {color:"#c0392b", bg:"#fdecea",  icon:"🕐"},
    "Nueva cotización":            {color:"#6c3483", bg:"#e8daef",  icon:"📄"},
    "Cotización convertida a reserva":{color:"#e67e22",bg:"#fef9e7",icon:"🔄"},
    "Pedido eliminado":            {color:"#c0392b", bg:"#fdecea",  icon:"🗑"},
    "Precio/stock editado":        {color:"#e67e22", bg:"#fef9e7",  icon:"✏️"},
  };

  const tiposUnicos = ["todos", ...new Set(activity.map(a=>a.accion))];

  const filtered = activity.filter(a=>{
    if(filterTipo!=="todos" && a.accion!==filterTipo) return false;
    if(search){
      const q=search.toLowerCase();
      return norm(a.usuario).includes(norm(q))||norm(a.accion).includes(norm(q))||norm(a.detalle).includes(norm(q));
    }
    return true;
  });

  // Resumen por usuario
  const byUser = useMemo(()=>{
    const map={};
    activity.forEach(a=>{
      if(!map[a.usuario]) map[a.usuario]=0;
      map[a.usuario]++;
    });
    return Object.entries(map).sort((a,b)=>b[1]-a[1]);
  },[activity]);

  return (
    <div>
      {/* KPIs */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:12,marginBottom:16}}>
        <div style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 4px #0001",borderLeft:`4px solid ${RED}`}}>
          <div style={{fontSize:11,color:"#888",fontWeight:600}}>📝 Total acciones</div>
          <div style={{fontSize:24,fontWeight:900,color:RED}}>{activity.length}</div>
        </div>
        {byUser.slice(0,3).map(([u,n])=>(
          <div key={u} style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 4px #0001",borderLeft:"4px solid #1a5276"}}>
            <div style={{fontSize:11,color:"#888",fontWeight:600}}>👤 {u}</div>
            <div style={{fontSize:24,fontWeight:900,color:"#1a5276"}}>{n}</div>
          </div>
        ))}
      </div>

      {/* Filtros */}
      <div style={{background:"#fff",borderRadius:12,padding:12,marginBottom:12,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center",boxShadow:"0 1px 4px #0001"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)}
          placeholder="🔍 Buscar por usuario, acción o detalle..."
          style={{flex:1,minWidth:200,padding:"7px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none"}}/>
        <select value={filterTipo} onChange={e=>setFilterTipo(e.target.value)}
          style={{padding:"7px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:12,background:"#fff",cursor:"pointer"}}>
          {tiposUnicos.map(t=><option key={t} value={t}>{t==="todos"?"Todas las acciones":t}</option>)}
        </select>
        {activity.length>0&&(
          confirmClear
            ? <div style={{display:"flex",gap:6,alignItems:"center",background:"#fdecea",borderRadius:8,padding:"5px 10px",border:`1px solid ${RED}44`}}>
                <span style={{fontSize:12,color:RED,fontWeight:600}}>¿Limpiar historial?</span>
                <button onClick={async()=>{setActivity([]);await db.clearActivity();setConfirmClear(false);}} style={{padding:"3px 10px",borderRadius:6,border:"none",background:RED,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:12}}>Sí</button>
                <button onClick={()=>setConfirmClear(false)} style={{padding:"3px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12}}>No</button>
              </div>
            : <button onClick={()=>setConfirmClear(true)} style={{padding:"6px 12px",borderRadius:8,border:"1.5px solid #fcc",background:"#fff",color:RED,cursor:"pointer",fontSize:11,fontWeight:600,whiteSpace:"nowrap"}}>🗑 Limpiar</button>
        )}
      </div>

      {/* Tabla */}
      {filtered.length===0
        ? <div style={{textAlign:"center",padding:50,color:"#aaa",background:"#fff",borderRadius:12}}>
            <div style={{fontSize:40,marginBottom:8}}>📝</div>
            <div>{activity.length===0?"No hay actividad registrada aún.":"No hay resultados para ese filtro."}</div>
          </div>
        : <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px #0001",overflow:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead>
                <tr style={{background:"#f9f9f9"}}>
                  {["Fecha","Usuario","Acción","Detalle"].map(h=>(
                    <th key={h} style={{padding:"10px 12px",textAlign:"left",fontWeight:700,color:"#888",fontSize:11,textTransform:"uppercase",letterSpacing:.4,whiteSpace:"nowrap"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.map((a,i)=>{
                  const cfg = TIPOS[a.accion]||{color:"#666",bg:"#f5f5f5",icon:"•"};
                  return (
                    <tr key={a.id||i} style={{borderTop:"1px solid #f5f5f5"}}>
                      <td style={{padding:"9px 12px",fontSize:11,color:"#888",whiteSpace:"nowrap"}}>{a.fecha}</td>
                      <td style={{padding:"9px 12px",whiteSpace:"nowrap"}}>
                        <span style={{fontWeight:700,fontSize:12}}>{a.usuario}</span>
                        <span style={{fontSize:10,color:a.rol==="admin"?RED:"#1a5276",marginLeft:4,fontWeight:600}}>({a.rol})</span>
                      </td>
                      <td style={{padding:"9px 12px"}}>
                        <span style={{background:cfg.bg,color:cfg.color,borderRadius:8,padding:"3px 9px",fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>
                          {cfg.icon} {a.accion}
                        </span>
                      </td>
                      <td style={{padding:"9px 12px",color:"#555",fontSize:12,maxWidth:300,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={a.detalle}>{a.detalle}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
      }
    </div>
  );
}

// ── Vendors ───────────────────────────────────────────────────────────────────
// FIX: DB-first pattern - primero persistir en Supabase, luego actualizar estado local.
// Esto evita que la UI quede desincronizada si Supabase falla.
// -- Price Lists -----------------------------------------------------------------
function PriceListsPanel({priceLists, setPriceLists}) {
  const [form, setForm] = useState({name:"", discount:""});
  const [editing, setEditing] = useState(null);
  const [saving, setSaving] = useState(false);

  const save = async () => {
    if(!form.name.trim()) { alert("Ingresá un nombre"); return; }
    const disc = parseFloat(form.discount)||0;
    if(disc < 0 || disc >= 100) { alert("El descuento debe ser entre 0 y 99"); return; }
    setSaving(true);
    const pl = {id: editing || genId(), name: form.name.trim(), discount: disc};
    setPriceLists(list => editing ? list.map(x=>x.id===editing?pl:x) : [...list, pl]);
    await db.savePriceList(pl);
    setForm({name:"", discount:""}); setEditing(null); setSaving(false);
  };

  const del = async (id) => {
    if(id==="default") { alert("No se puede eliminar la lista Normal"); return; }
    if(!window.confirm("¿Eliminar esta lista?")) return;
    setPriceLists(list=>list.filter(x=>x.id!==id));
    await db.deletePriceList(id);
  };

  const startEdit = (pl) => { setEditing(pl.id); setForm({name:pl.name, discount:pl.discount}); };
  const cancel = () => { setEditing(null); setForm({name:"", discount:""}); };

  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,alignItems:"start"}}>
      {/* Form */}
      <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:15,marginBottom:16}}>
          {editing ? "✏️ Editar lista" : "➕ Nueva lista de precios"}
        </div>
        <Field label="Nombre de la lista">
          <input value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}
            placeholder="Ej: Mayorista, Revendedor..." style={inputStyle}/>
        </Field>
        <Field label="Descuento sobre precio normal (%)">
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <input type="number" min="0" max="99" value={form.discount} onChange={e=>setForm(f=>({...f,discount:e.target.value}))}
              placeholder="0" style={{...inputStyle,width:100}}/>
            <span style={{fontSize:13,color:"#888"}}>%</span>
            {parseFloat(form.discount)>0&&(
              <span style={{fontSize:12,color:"#1e8449",fontWeight:600}}>
                Precio de $1000 → {fARS(1000*(1-parseFloat(form.discount)/100))}
              </span>
            )}
          </div>
          <div style={{fontSize:11,color:"#888",marginTop:4}}>0% = precios normales sin cambios</div>
        </Field>
        <div style={{display:"flex",gap:8,marginTop:8}}>
          <button onClick={save} disabled={saving||!form.name.trim()}
            style={{flex:1,padding:"10px",borderRadius:9,border:"none",fontWeight:800,fontSize:13,cursor:"pointer",
              background:(!form.name.trim())?"#e5e5e5":`linear-gradient(135deg,${REDD},${RED})`,
              color:(!form.name.trim())?"#aaa":"#fff"}}>
            {saving ? "Guardando..." : editing ? "Guardar cambios" : "Crear lista"}
          </button>
          {editing&&<button onClick={cancel} style={{padding:"10px 16px",borderRadius:9,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontWeight:600,cursor:"pointer"}}>Cancelar</button>}
        </div>
      </div>

      {/* List */}
      <div>
        <div style={{fontWeight:700,fontSize:13,color:"#888",marginBottom:10,textTransform:"uppercase",letterSpacing:.5}}>
          Listas existentes ({priceLists.length})
        </div>
        {priceLists.map(pl=>(
          <div key={pl.id} style={{background:"#fff",borderRadius:10,padding:"14px 16px",marginBottom:8,boxShadow:"0 1px 4px #0001",display:"flex",alignItems:"center",gap:12}}>
            <div style={{flex:1}}>
              <div style={{fontWeight:700,fontSize:14}}>{pl.name}</div>
              <div style={{fontSize:12,color:"#888",marginTop:2}}>
                {pl.discount===0
                  ? <span style={{color:"#1e8449",fontWeight:600}}>Precios normales (sin descuento)</span>
                  : <span style={{color:"#c0392b",fontWeight:600}}>-{pl.discount}% sobre precio normal</span>
                }
              </div>
            </div>
            <div style={{display:"flex",gap:6}}>
              {pl.id!=="default"&&<>
                <button onClick={()=>startEdit(pl)} style={{padding:"5px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12}}>✏️</button>
                <button onClick={()=>del(pl.id)} style={{padding:"5px 10px",borderRadius:6,border:"1.5px solid #fcc",background:"#fff",cursor:"pointer",fontSize:12,color:RED}}>🗑</button>
              </>}
              {pl.id==="default"&&<span style={{fontSize:11,color:"#aaa",padding:"5px 8px"}}>predeterminada</span>}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function VendorsPanel({vendors,setVendors}) {
  const [newName, setNewName] = useState("");
  const [editing, setEditing] = useState(null);
  const [editVal, setEditVal] = useState("");
  const [confirmDel, setConfirmDel] = useState(null);
  const [loading, setLoading] = useState(false);

  const add = async () => {
    const n = newName.trim();
    if(!n) return;
    if(vendors.includes(n)){alert("Ya existe ese vendedor");return;}
    setLoading(true);
    try {
      await db.addVendor(n);
      setVendors(v=>[...v,n]);
      setNewName("");
    } catch(e) {
      alert("Error al guardar el vendedor: " + e.message);
    } finally {
      setLoading(false);
    }
  };

  const doRemove = async () => {
    try {
      await db.deleteVendor(confirmDel);
      setVendors(vs=>vs.filter(x=>x!==confirmDel));
      setConfirmDel(null);
    } catch(e) {
      alert("Error al eliminar el vendedor: " + e.message);
    }
  };

  const saveEdit = async (old) => {
    const n = editVal.trim();
    if(!n) return;
    try {
      await db.updateVendor(old, n);
      setVendors(vs=>vs.map(x=>x===old?n:x));
      setEditing(null);
    } catch(e) {
      alert("Error al editar el vendedor: " + e.message);
    }
  };

  return (
    <div style={{background:"#fff",borderRadius:12,padding:24,boxShadow:"0 1px 4px #0001",maxWidth:520}}>
      <div style={{fontWeight:800,fontSize:16,marginBottom:16}}>👥 Gestión de Vendedores</div>
      <div style={{display:"flex",gap:8,marginBottom:20}}>
        <input value={newName} onChange={e=>setNewName(e.target.value)}
          onKeyDown={e=>e.key==="Enter"&&add()}
          placeholder="Nombre del nuevo vendedor"
          style={{...inputStyle,flex:1}}/>
        <button onClick={add} disabled={loading} style={{padding:"8px 18px",borderRadius:8,border:"none",background:loading?"#aaa":RED,color:"#fff",fontWeight:700,cursor:loading?"not-allowed":"pointer",fontSize:13,whiteSpace:"nowrap"}}>
          {loading?"Guardando...":"+ Agregar"}
        </button>
      </div>
      {vendors.length===0
        ? <div style={{textAlign:"center",padding:30,color:"#aaa"}}>No hay vendedores cargados.</div>
        : vendors.map(v=>(
          <div key={v} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",borderRadius:10,border:"1.5px solid #f0f0f0",marginBottom:8,background:"#fafafa"}}>
            <span style={{fontSize:20}}>👤</span>
            {editing===v
              ? <>
                  <input value={editVal} onChange={e=>setEditVal(e.target.value)}
                    onKeyDown={e=>{if(e.key==="Enter")saveEdit(v);if(e.key==="Escape")setEditing(null);}}
                    style={{...inputStyle,flex:1,fontSize:13}} autoFocus/>
                  <button onClick={()=>saveEdit(v)} style={{padding:"5px 12px",borderRadius:7,border:"none",background:"#1e8449",color:"#fff",fontWeight:700,cursor:"pointer",fontSize:12}}>v</button>
                  <button onClick={()=>setEditing(null)} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12}}>x</button>
                </>
              : <>
                  <span style={{flex:1,fontWeight:600,fontSize:14}}>{v}</span>
                  <button onClick={()=>{setEditing(v);setEditVal(v);}} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12}}>✏️️ Editar</button>
                  {confirmDel===v
                    ? <div style={{display:"flex",alignItems:"center",gap:5,background:"#fdecea",borderRadius:8,padding:"4px 8px",border:`1px solid ${RED}44`}}>
                        <span style={{fontSize:11,color:RED,fontWeight:600,whiteSpace:"nowrap"}}>?Eliminar?</span>
                        <button onClick={doRemove} style={{padding:"3px 9px",borderRadius:6,border:"none",background:RED,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:11}}>Sí</button>
                        <button onClick={()=>setConfirmDel(null)} style={{padding:"3px 8px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:11}}>No</button>
                      </div>
                    : <button onClick={()=>setConfirmDel(v)} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #fcc",background:"#fff",color:RED,cursor:"pointer",fontSize:12}}>🗑</button>
                  }
                </>
            }
          </div>
        ))
      }
    </div>
  );
}

// ── Users ─────────────────────────────────────────────────────────────────────
function UsersPanel({users,setUsers,vendors,priceLists}) {
  const [form, setForm] = useState({username:"",password:"",name:"",role:"vendedor",vendedor:"",priceList:"default",canSeeAll:true,email:"",phone:"",cargo:"",avatar:"",barcodeEnabled:false});
  const [editing, setEditing] = useState(null);
  const [showPass, setShowPass] = useState({});
  const [avatarPreview, setAvatarPreview] = useState("");

  const startEdit = (u) => {
    setEditing(u.id);
    setForm({username:u.username,password:u.password,name:u.name,role:u.role,email:u.email||"",phone:u.phone||"",cargo:u.cargo||"",vendedor:u.vendedor||"",priceList:u.priceList||"default",canSeeAll:u.canSeeAll!==false,avatar:u.avatar||"",barcodeEnabled:u.barcodeEnabled||false});
    setAvatarPreview(u.avatar||"");
  };
  const cancelEdit = () => { setEditing(null); setForm({username:"",password:"",name:"",role:"vendedor",vendedor:"",priceList:"default",canSeeAll:true,email:"",phone:"",cargo:"",avatar:"",barcodeEnabled:false}); setAvatarPreview(""); };

  const handleAvatar = (e) => {
    const file = e.target.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const b64 = ev.target.result;
      setForm(f=>({...f,avatar:b64}));
      setAvatarPreview(b64);
    };
    reader.readAsDataURL(file);
  };

  const save = async () => {
    if(!form.username.trim()||!form.password.trim()||!form.name.trim()){alert("Completá nombre, usuario y contraseña");return;}
    try {
      if(editing) {
        const updated = {...users.find(u=>u.id===editing), ...form, priceList:form.priceList||"default", canSeeAll:form.canSeeAll!==false};
        setUsers(us=>us.map(u=>u.id===editing?updated:u));
        await db.saveUser(updated);
      } else {
        if(users.find(u=>u.username===form.username.trim())){alert("Ese usuario ya existe");return;}
        const newUser = {id:genId(),username:form.username.trim(),password:form.password,name:form.name.trim(),role:form.role||"vendedor",email:form.email||"",phone:form.phone||"",cargo:form.cargo||"",vendedor:form.vendedor||"",priceList:form.priceList||"default",canSeeAll:form.canSeeAll!==false,avatar:form.avatar||"",barcodeEnabled:form.barcodeEnabled||false};
        setUsers(us=>[...us,newUser]);
        await db.saveUser(newUser);
      }
      cancelEdit();
    } catch(e) {
      alert("Error al guardar: " + (e.message||JSON.stringify(e)));
    }
  };

  const remove = async (id) => {
    if(users.filter(u=>u.role==="admin").length===1&&users.find(u=>u.id===id)?.role==="admin"){alert("Debe haber al menos un administrador");return;}
    setUsers(us=>us.filter(u=>u.id!==id));
    await db.deleteUser(id);
  };

  const Toggle = ({label,sub,val,onChange}) => (
    <div style={{background:"#f9f9f9",borderRadius:8,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",gap:12}}>
      <div>
        <div style={{fontWeight:700,fontSize:13}}>{label}</div>
        {sub&&<div style={{fontSize:11,color:"#888",marginTop:2}}>{sub}</div>}
      </div>
      <button onClick={onChange} style={{padding:"6px 14px",borderRadius:20,border:"none",cursor:"pointer",fontWeight:700,fontSize:12,background:val?"#1e8449":"#e5e5e5",color:val?"#fff":"#888",whiteSpace:"nowrap"}}>
        {val?"✅ Sí":"🔒 No"}
      </button>
    </div>
  );

  const isMobilePanel = useIsMobile();
  const [view, setView] = useState("lista");

  return (
    <div>
      <div style={{background:"#fff",borderRadius:12,padding:4,marginBottom:16,display:"flex",gap:4,boxShadow:"0 1px 4px #0001"}}>
        <button onClick={()=>{setView("lista");cancelEdit();}} style={{flex:1,padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="lista"?`linear-gradient(135deg,#922b21,${RED})`:"transparent",color:view==="lista"?"#fff":"#555"}}>
          🔐 Usuarios ({users.length})
        </button>
        <button onClick={()=>setView("form")} style={{flex:1,padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="form"?`linear-gradient(135deg,#922b21,${RED})`:"transparent",color:view==="form"?"#fff":"#555"}}>
          {editing?"✏️ Editando":"+ Nuevo usuario"}
        </button>
      </div>

      {view==="lista" && (
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {users.map(u=>(
            <div key={u.id} style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 4px #0001",border:`1.5px solid ${editing===u.id?"#c0392b":"#f0f0f0"}`}}>
              <div style={{display:"flex",alignItems:"center",gap:12}}>
                <div style={{width:48,height:48,borderRadius:"50%",background:u.avatar?"transparent":"linear-gradient(135deg,#7b1a1a,#c0392b)",overflow:"hidden",flexShrink:0,border:"2px solid #f0f0f0",display:"flex",alignItems:"center",justifyContent:"center"}}>
                  {u.avatar ? <img src={u.avatar} alt={u.name} style={{width:"100%",height:"100%",objectFit:"cover"}}/> : <span style={{fontSize:20}}>{u.role==="admin"?"👑":"👤"}</span>}
                </div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:700,fontSize:14}}>{u.name}{u.cargo&&<span style={{fontWeight:400,color:"#888",fontSize:12}}> · {u.cargo}</span>}</div>
                  <div style={{display:"flex",gap:5,flexWrap:"wrap",marginTop:4}}>
                    <span style={{background:u.role==="admin"?"#fdecea":"#eaf4fc",color:u.role==="admin"?RED:"#1a5276",borderRadius:6,padding:"1px 8px",fontWeight:700,fontSize:10}}>{u.role==="admin"?"Admin":"Vendedor"}</span>
                    {u.vendedor&&<span style={{background:"#f5eef8",color:"#6c3483",borderRadius:6,padding:"1px 8px",fontWeight:700,fontSize:10}}>👤 {u.vendedor}</span>}
                    {u.barcodeEnabled&&<span style={{background:"#eaf4fc",color:"#1a5276",borderRadius:6,padding:"1px 8px",fontWeight:700,fontSize:10}}>📷 Lector</span>}
                    {u.canSeeAll===false&&<span style={{background:"#fef9e7",color:"#b7770d",borderRadius:6,padding:"1px 8px",fontWeight:700,fontSize:10}}>🔒 Solo suyos</span>}
                  </div>
                  {(u.phone||u.email)&&<div style={{fontSize:11,color:"#aaa",marginTop:4}}>{u.phone&&<span>📱 {u.phone}</span>}{u.phone&&u.email&&" · "}{u.email&&<span>📧 {u.email}</span>}</div>}
                </div>
                <div style={{display:"flex",gap:6,flexShrink:0}}>
                  <button onClick={()=>{startEdit(u);setView("form");}} style={{padding:"6px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12,fontWeight:600}}>✏️ Editar</button>
                  <button onClick={()=>remove(u.id)} style={{padding:"6px 10px",borderRadius:8,border:"1.5px solid #fcc",background:"#fff",color:RED,cursor:"pointer",fontSize:12}}>🗑</button>
                </div>
              </div>
            </div>
          ))}
          <button onClick={()=>setView("form")} style={{width:"100%",padding:"12px",borderRadius:12,border:"2px dashed #e5e5e5",background:"#fafafa",color:"#888",fontWeight:700,fontSize:13,cursor:"pointer",marginTop:4}}>
            + Agregar nuevo usuario
          </button>
        </div>
      )}

      {view==="form" && (
        <div style={{background:"#fff",borderRadius:12,padding:24,boxShadow:"0 1px 4px #0001"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:16}}>{editing?"✏️ Editando usuario":"+ Nuevo usuario"}</div>
          <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:16,padding:"14px 16px",background:"#f9f9f9",borderRadius:12}}>
            <div style={{width:72,height:72,borderRadius:"50%",background:avatarPreview?"transparent":"linear-gradient(135deg,#7b1a1a,#c0392b)",overflow:"hidden",flexShrink:0,border:"3px solid #fff",boxShadow:"0 2px 8px #0002",display:"flex",alignItems:"center",justifyContent:"center"}}>
              {avatarPreview ? <img src={avatarPreview} alt="avatar" style={{width:"100%",height:"100%",objectFit:"cover"}}/> : <span style={{fontSize:28}}>👤</span>}
            </div>
            <div style={{flex:1}}>
              <div style={{fontWeight:700,fontSize:13,marginBottom:6}}>Foto de perfil</div>
              <label style={{display:"inline-block",padding:"8px 14px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",fontSize:12,fontWeight:600,cursor:"pointer",color:"#555"}}>
                📷 Subir foto
                <input type="file" accept="image/*" onChange={handleAvatar} style={{display:"none"}}/>
              </label>
              {avatarPreview&&<button onClick={()=>{setAvatarPreview("");setForm(f=>({...f,avatar:""}));}} style={{marginLeft:8,padding:"8px 12px",borderRadius:8,border:"none",background:"#fdecea",color:RED,fontSize:12,cursor:"pointer",fontWeight:600}}>✕ Quitar</button>}
              <div style={{fontSize:11,color:"#aaa",marginTop:6}}>Aparecerá en el PDF y lista de usuarios</div>
            </div>
          </div>
          <Field label="Nombre completo *"><input value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} placeholder="Ej: María García" style={inputStyle}/></Field>
          <Field label="Cargo / Título"><input value={form.cargo||""} onChange={e=>setForm(f=>({...f,cargo:e.target.value}))} placeholder="Ej: Representante Comercial" style={inputStyle}/></Field>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <Field label="Teléfono"><input value={form.phone||""} onChange={e=>setForm(f=>({...f,phone:e.target.value}))} placeholder="+54 11 1234-5678" style={inputStyle}/></Field>
            <Field label="Email"><input type="email" value={form.email||""} onChange={e=>setForm(f=>({...f,email:e.target.value}))} placeholder="vendedor@ejemplo.com" style={inputStyle}/></Field>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <Field label="Usuario *"><input value={form.username} onChange={e=>setForm(f=>({...f,username:e.target.value}))} placeholder="Ej: maria" style={inputStyle}/></Field>
            <Field label="Contraseña *">
              <div style={{position:"relative"}}>
                <input type={showPass.form?"text":"password"} value={form.password} onChange={e=>setForm(f=>({...f,password:e.target.value}))} placeholder="••••••••" style={{...inputStyle,paddingRight:40}}/>
                <button onClick={()=>setShowPass(s=>({...s,form:!s.form}))} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:15,color:"#aaa"}}>{showPass.form?"🙈":"👁️"}</button>
              </div>
            </Field>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <Field label="Rol"><select value={form.role} onChange={e=>setForm(f=>({...f,role:e.target.value}))} style={{...inputStyle,cursor:"pointer"}}><option value="vendedor">Vendedor</option><option value="admin">Administrador</option></select></Field>
            <Field label="Vendedor asignado"><select value={form.vendedor||""} onChange={e=>setForm(f=>({...f,vendedor:e.target.value}))} style={{...inputStyle,cursor:"pointer",color:form.vendedor?"#1a1a1a":"#aaa"}}><option value="">- Sin asignar -</option>{(vendors||[]).map(v=><option key={v} value={v}>{v}</option>)}</select></Field>
          </div>
          <Field label="Lista de precios">
            <select value={form.priceList||"default"} onChange={e=>setForm(f=>({...f,priceList:e.target.value}))} style={{...inputStyle,cursor:"pointer"}}>
              {(priceLists||[{id:"default",name:"Normal",discount:0}]).map(pl=>(<option key={pl.id} value={pl.id}>{pl.name}{pl.discount>0?` (-${pl.discount}%)`:""}</option>))}
            </select>
          </Field>
          <Toggle label="Ver todos los pedidos" sub={form.canSeeAll?"Ve todos los pedidos":"Solo sus propios pedidos"} val={form.canSeeAll} onChange={()=>setForm(f=>({...f,canSeeAll:!f.canSeeAll}))}/>
          <Toggle label="📷 Lector de código de barras" sub={form.barcodeEnabled?"Puede usar el lector":"Sin acceso al lector"} val={form.barcodeEnabled} onChange={()=>setForm(f=>({...f,barcodeEnabled:!f.barcodeEnabled}))}/>
          <div style={{display:"flex",gap:8,marginTop:8}}>
            <button onClick={save} style={{flex:1,padding:"11px",borderRadius:10,border:"none",background:RED,color:"#fff",fontWeight:800,cursor:"pointer",fontSize:14}}>{editing?"💾 Guardar cambios":"✅ Crear usuario"}</button>
            <button onClick={()=>{cancelEdit();setView("lista");}} style={{padding:"11px 16px",borderRadius:10,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontWeight:600,color:"#666"}}>Cancelar</button>
          </div>
        </div>
      )}
    </div>
  );
}
// ── Excel Import ──────────────────────────────────────────────────────────────
function ExcelPanel({products,setProducts}) {
  const fileRef = useRef();
  const [preview, setPreview] = useState(null);
  const [status, setStatus] = useState(null);
  const [mode, setMode] = useState("update");

  const parseExcel = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, {
          type:"array",
          cellFormula: false,
          cellNF: false,
          cellHTML: false,
        });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const range = XLSX.utils.decode_range(ws["!ref"] || "A1:G1000");
        const totalRows = range.e.r;

        if(totalRows < 1){setStatus({type:"error",msg:"El archivo parece estar vacío."});return;}

        const cellNum = (r, c) => {
          const addr = XLSX.utils.encode_cell({r, c});
          const cell = ws[addr];
          if(!cell || cell.v === undefined || cell.v === null || cell.v === "") return null;
          const n = typeof cell.v === "number" ? cell.v
            : parseFloat(String(cell.v).replace(/[$\s]/g,"").replace(/\./g,"").replace(",","."));
          return isNaN(n) || n === 0 ? null : n;
        };
        const cellStr = (r, c) => {
          const addr = XLSX.utils.encode_cell({r, c});
          const cell = ws[addr];
          if(!cell || cell.v === undefined) return "";
          return String(cell.v).trim();
        };

        let headerRow = 0;
        for(let r = 0; r <= Math.min(5, totalRows); r++) {
          for(let c = 0; c <= 10; c++) {
            const v = cellStr(r, c).toUpperCase()
              .normalize("NFD").replace(/[̀-ͯ]/g,"");
            if(v.includes("CODIGO") || v.includes("COD")) { headerRow = r; break; }
          }
        }

        const norm = s => String(s).toUpperCase().trim()
          .normalize("NFD").replace(/[̀-ͯ]/g,"");

        const COL = { codigo:-1, descripcion:-1, precioIVA:-1, precioOferta:-1, precioFinal:-1, fecha:-1 };
        for(let c = range.s.c; c <= range.e.c; c++) {
          const h = norm(cellStr(headerRow, c));
          if(h.includes("CODIGO") || h === "COD" || h === "ID") COL.codigo = c;
          else if(h.includes("DESCRIP") || h.includes("NOMBRE"))       COL.descripcion = c;
          else if(h.includes("CON IVA") || (h.includes("IVA") && !h.includes("OFERTA") && !h.includes("FINAL"))) COL.precioIVA = c;
          else if(h.includes("OFERTA"))  COL.precioOferta = c;
          else if(h.includes("FINAL") || h.includes("PRECIO FINAL"))   COL.precioFinal  = c;
          else if(h.includes("FECHA") || h.includes("ULTIMA") || h.includes("ACT")) COL.fecha = c;
        }

        if(COL.codigo      < 0) COL.codigo      = 0;
        if(COL.descripcion < 0) COL.descripcion = 1;
        if(COL.precioIVA   < 0) COL.precioIVA   = 2;
        if(COL.precioOferta< 0) COL.precioOferta = 3;
        if(COL.fecha       < 0) COL.fecha        = 4;
        if(COL.precioFinal < 0) COL.precioFinal  = 5;

        const colLetter = c => c < 0 ? "-" : String.fromCharCode(65 + c);
        const detectedCols = {
          "Código":        colLetter(COL.codigo),
          "Descripción":   colLetter(COL.descripcion),
          "Precio c/IVA":  colLetter(COL.precioIVA),
          "Precio Oferta": colLetter(COL.precioOferta),
          "Fecha":         colLetter(COL.fecha),
          "Precio Final":  colLetter(COL.precioFinal),
        };

        const parsed = [];
        for(let r = headerRow + 1; r <= totalRows; r++) {
          const id = cellStr(r, COL.codigo);
          if(!id) continue;
          const pIVA    = cellNum(r, COL.precioIVA);
          const pOferta = cellNum(r, COL.precioOferta);
          const pFinal  = cellNum(r, COL.precioFinal);
          parsed.push({
            id,
            name:         cellStr(r, COL.descripcion),
            precioIVA:    pIVA,
            precioOferta: pOferta,
            precioFinal:  pFinal,
            fecha:        cellStr(r, COL.fecha),
          });
        }

        if(parsed.length === 0){setStatus({type:"error",msg:"No se encontraron datos. Verificá que el archivo tenga encabezados en la primera fila."});return;}

        setPreview({rows:parsed.slice(0,10), total:parsed.length, all:parsed, cols:COL, detectedCols});
        setStatus(null);
      } catch(err) {
        setStatus({type:"error",msg:"Error al leer el archivo: "+err.message});
        setPreview(null);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const applyImport = async () => {
    if(!preview) return;
    const updated=[], notFound=[];
    const newProds = products.map(p=>({...p}));

    const resolvePrice = (row) => {
      const oferta = row.precioOferta;
      const iva    = row.precioIVA;
      if(oferta !== null && oferta !== undefined && !isNaN(oferta) && oferta > 0) return oferta;
      return iva;
    };

    preview.all.forEach(row=>{
      const idx = newProds.findIndex(p=>p.id===row.id);
      const precioVenta = resolvePrice(row);  // oferta si existe, sino IVA
      const precioCosto = row.precioIVA;      // costo = siempre precio IVA
      if(idx>=0){
        if(row.precioFinal!==null && row.precioFinal>0)
                                   newProds[idx].salePrice = row.precioFinal;
        else if(precioVenta!==null) newProds[idx].salePrice = precioVenta;
        if(precioCosto!==null)     newProds[idx].costPrice = precioCosto;  // siempre IVA como costo
        if(row.name) newProds[idx].name = row.name;
        updated.push(row.id);
      } else {
        notFound.push(row.id);
      }
    });

    if(mode==="full") {
      preview.all.forEach(row=>{
        if(!newProds.find(p=>p.id===row.id)) {
          const precio = resolvePrice(row);   // oferta si existe, sino IVA
          const costo  = row.precioIVA;       // costo = siempre IVA
          newProds.push({
            id:row.id, name:row.name||row.id,
            costPrice:costo||0,
            salePrice:(row.precioFinal&&row.precioFinal>0)?row.precioFinal:(precio||0),
            category:"Importado", stock:0
          });
        }
      });
    }

    const batchSize = 20;
    const batches = [];
    for(let i=0;i<newProds.length;i+=batchSize) {
      batches.push(newProds.slice(i,i+batchSize));
    }
    let uploadedCount = 0;
    let errorCount = 0;
    for(let i=0;i<batches.length;i++) {
      const pct = Math.round((i+1)/batches.length*100);
      setStatus({type:"progress", msg:`⏳ Subiendo... ${uploadedCount}/${newProds.length} productos (${pct}%)`});
      try {
        await db.upsertProducts(batches[i]);
        uploadedCount += batches[i].length;
      } catch(e) {
        errorCount += batches[i].length;
        console.error("Batch error:", e);
      }
      await new Promise(r=>setTimeout(r,100));
    }
    setProducts(newProds);
    setStatus({type:"success",msg:`✅ ${updated.length} productos actualizados.${notFound.length>0?` ${notFound.length} códigos no encontrados en el catálogo.`:""}`});
    setPreview(null);
    if(fileRef.current) fileRef.current.value="";
  };

  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,alignItems:"start"}}>
      <div style={{background:"#fff",borderRadius:12,padding:24,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:16,marginBottom:6}}>📊 Importar Lista de Precios</div>
        <div style={{fontSize:13,color:"#666",marginBottom:16,lineHeight:1.6}}>
          El sistema lee automáticamente las columnas:<br/>
          <strong>CÓDIGO . DESCRIPCIÓN . PRECIO CON IVA . PRECIO OFERTA . FECHA ULTIMA ACTUALIZACIÓN . PRECIO FINAL</strong>
        </div>
        <div style={{marginBottom:14}}>
          <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:6}}>Modo de importación</div>
          <div style={{display:"flex",gap:8}}>
            {[{v:"update",l:"Actualizar precios",desc:"Solo actualiza productos existentes por código"},{v:"full",l:"Importación completa",desc:"Actualiza existentes y agrega nuevos"}].map(opt=>(
              <div key={opt.v} onClick={()=>setMode(opt.v)} style={{flex:1,padding:"10px 12px",borderRadius:10,border:`2px solid ${mode===opt.v?RED:"#e5e5e5"}`,background:mode===opt.v?"#fdecea":"#fafafa",cursor:"pointer"}}>
                <div style={{fontWeight:700,fontSize:12,color:mode===opt.v?RED:"#555"}}>{opt.l}</div>
                <div style={{fontSize:10,color:"#aaa",marginTop:2}}>{opt.desc}</div>
              </div>
            ))}
          </div>
        </div>
        <div onClick={()=>fileRef.current?.click()} style={{border:"2px dashed #e5e5e5",borderRadius:12,padding:"28px 20px",textAlign:"center",cursor:"pointer",background:"#fafafa",marginBottom:14,transition:"border-color .2s"}}
          onMouseEnter={e=>e.currentTarget.style.borderColor=RED}
          onMouseLeave={e=>e.currentTarget.style.borderColor="#e5e5e5"}>
          <div style={{fontSize:36,marginBottom:8}}>📂</div>
          <div style={{fontWeight:700,color:"#555",fontSize:14}}>Hacé clic o arrastrá tu archivo Excel</div>
          <div style={{fontSize:12,color:"#aaa",marginTop:4}}>.xlsx . .xls . .csv</div>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}}
            onChange={e=>{if(e.target.files[0])parseExcel(e.target.files[0]);}}/>
        </div>
        {status&&<div style={{padding:"10px 14px",borderRadius:8,
          background:status.type==="error"?"#fdecea":status.type==="progress"?"#fef9e7":"#d5f5e3",
          color:status.type==="error"?RED:status.type==="progress"?"#7d6608":"#1e8449",
          fontSize:13,fontWeight:600,marginBottom:12}}>{status.msg}</div>}
        {preview&&(
          <button onClick={applyImport} disabled={status?.type==="progress"} style={{width:"100%",padding:"11px",borderRadius:10,border:"none",background:status?.type==="progress"?"#aaa":`linear-gradient(135deg,${REDD},${RED})`,color:"#fff",fontWeight:800,fontSize:14,cursor:status?.type==="progress"?"not-allowed":"pointer"}}>
            {status?.type==="progress"?"⏳ Importando, no cierres esta página...":"📥 Aplicar importación ("+preview.total+" productos)"}
          </button>
        )}
      </div>
      <div style={{background:"#fff",borderRadius:12,padding:24,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:15,marginBottom:12}}>
          {preview ? `👁️ Vista previa (${preview.total} filas)` : "📋 Instrucciones"}
        </div>
        {preview?.detectedCols && (
          <div style={{background:"#f0fdf4",border:"1px solid #a7f3d0",borderRadius:8,padding:"10px 12px",marginBottom:12}}>
            <div style={{fontWeight:700,fontSize:11,color:"#065f46",marginBottom:6}}>✅ Columnas detectadas:</div>
            <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
              {Object.entries(preview.detectedCols).map(([k,v])=>(
                <span key={k} style={{background:"#fff",border:"1px solid #d1fae5",borderRadius:6,padding:"3px 8px",fontSize:11,color:"#065f46"}}>
                  <strong>{v}</strong> → {k}
                </span>
              ))}
            </div>
          </div>
        )}
        {!preview
          ? <div style={{fontSize:13,color:"#666",lineHeight:1.8}}>
              <div style={{background:"#f9f9f9",borderRadius:8,padding:"12px 16px",fontFamily:"monospace",fontSize:12,lineHeight:2}}>
                <div><strong>Col A:</strong> CÓDIGO</div>
                <div><strong>Col B:</strong> DESCRIPCIÓN</div>
                <div><strong>Col C:</strong> PRECIO CON IVA</div>
                <div><strong>Col D:</strong> PRECIO OFERTA</div>
                <div><strong>Col E:</strong> FECHA ULTIMA ACTUALIZACIÓN</div>
                <div><strong>Col F:</strong> PRECIO FINAL</div>
              </div>
              <div style={{marginTop:12,fontSize:12,color:"#aaa"}}>El sistema detecta automáticamente las columnas por nombre. Las columnas no tienen que estar en un orden específico.</div>
            </div>
          : <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                <thead><tr style={{background:"#f9f9f9"}}>
                  {["Código","Descripción","P. IVA","P. Oferta","P. Final","Fecha"].map(h=><th key={h} style={{padding:"7px 8px",textAlign:"left",fontWeight:700,color:"#888",whiteSpace:"nowrap"}}>{h}</th>)}
                </tr></thead>
                <tbody>
                  {preview.rows.map((r,i)=>(
                    <tr key={i} style={{borderTop:"1px solid #f5f5f5"}}>
                      <td style={{padding:"6px 8px",fontWeight:600,color:"#444"}}>{r.id}</td>
                      <td style={{padding:"6px 8px",color:"#555",maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.name}</td>
                      <td style={{padding:"6px 8px",color:"#666"}}>{r.precioIVA!=null?fARS(r.precioIVA):"-"}</td>
                      <td style={{padding:"6px 8px",color:"#666"}}>{r.precioOferta!=null?fARS(r.precioOferta):"-"}</td>
                      <td style={{padding:"6px 8px",fontWeight:700,color:RED}}>{r.precioFinal!=null?fARS(r.precioFinal):"-"}</td>
                      <td style={{padding:"6px 8px",color:"#aaa",fontSize:10}}>{r.fecha}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {preview.total>10&&<div style={{fontSize:11,color:"#aaa",marginTop:8,textAlign:"center"}}>... y {preview.total-10} filas más</div>}
            </div>
        }
      </div>
    </div>
  );
}
