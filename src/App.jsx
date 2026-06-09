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
  getUsers:     async () => { const {data,error} = await supabase.from("lm_users").select("*").order("name"); if(error) throw error; return (data||[]).map(u=>({...u,priceList:u.price_list||"default",vendedor:u.vendedor||"",canSeeAll:u.can_see_all!==false})); },
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
      can_see_all: u.canSeeAll!==false
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
const STAGES = ["reserva","confirmado","en armado","entregado"];
const SCFG = {
  reserva:     {label:"Reserva",    color:"#c0392b", bg:"#fdecea", icon:"🕐"},
  confirmado:  {label:"Confirmado", color:"#1a5276", bg:"#d6eaf8", icon:"✅"},
  "en armado": {label:"En Armado",  color:"#6c3483", bg:"#e8daef", icon:"📦"},
  entregado:   {label:"Entregado",  color:"#1e8449", bg:"#d5f5e3", icon:"🎉"},
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
  html,body{background:#fff;color:#1a1a1a;width:210mm;}
  @media screen{body{max-width:210mm;margin:0 auto;box-shadow:0 0 20px #0002;}}
  @media print{.no-print{display:none!important;}@page{margin:3mm;size:A4 portrait;}body{width:100%;}}
  .print-btn{display:block;margin:12px auto;padding:9px 28px;background:${badgeColor};color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;}
  .pdf-wrap{padding:0 0 16px;}

  /* HEADER — imagen completa sin recorte */
  .header-img{
    width:100%;
    height:auto;
    display:block;
    border-bottom:3px solid #c0392b;
  }

  .content{padding:12px 16px 0;}

  /* DOC META row */
  .doc-meta{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;padding-bottom:10px;border-bottom:2px solid #f0f0f0;}
  .doc-left{}
  .doc-type-label{font-size:10px;color:#999;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px;}
  .doc-badge{display:inline-block;padding:5px 16px;border-radius:6px;font-size:13px;font-weight:800;letter-spacing:1px;background:${badgeColor};color:#fff;}
  .doc-right{text-align:right;}
  .doc-num{font-size:22px;font-weight:900;color:#1a1a1a;letter-spacing:-0.5px;}
  .doc-date{font-size:11px;color:#888;margin-top:3px;}

  /* INFO GRID */
  .info-grid{display:grid;grid-template-columns:1fr 1fr;gap:7px 16px;margin-bottom:12px;}
  .info-box{background:#f9f9f9;border-radius:6px;padding:7px 12px;border-left:3px solid #e5e5e5;}
  .info-box.hl{border-left-color:${badgeColor};}
  .info-label{font-size:8px;color:#999;text-transform:uppercase;letter-spacing:.7px;margin-bottom:3px;font-weight:700;}
  .info-value{font-size:13px;font-weight:700;color:#1a1a1a;}

  /* VALIDITY */
  .validity-bar{background:#fef9e7;border-left:3px solid #f1c40f;padding:6px 12px;border-radius:0 6px 6px 0;font-size:12px;color:#7d6608;margin-bottom:10px;}

  /* TABLE */
  table{width:100%;border-collapse:collapse;margin-bottom:14px;}
  thead tr{background:#f5f5f5;}
  th{padding:8px 10px;text-align:left;font-size:10px;color:#888;text-transform:uppercase;letter-spacing:.5px;font-weight:700;}
  th.r{text-align:right;}th.c{text-align:center;}
  td{padding:8px 10px;border-bottom:1px solid #f0f0f0;font-size:15px;color:#1a1a1a;vertical-align:middle;}
  td.r{text-align:right;}td.c{text-align:center;}
  tbody tr:nth-child(even){background:#fafafa;}
  tbody tr:last-child td{border-bottom:none;}

  /* TOTAL */
  .total-wrap{display:flex;justify-content:flex-end;margin-bottom:14px;}
  .total-box{background:${tipo==="cotizacion"?"#e8daef":"#fdecea"};border-radius:8px;padding:12px 20px;min-width:280px;}
  .disc-row{display:flex;justify-content:space-between;font-size:13px;margin-bottom:3px;}
  .disc-row.green{color:#1e8449;}
  .disc-row.grey{color:#888;}
  .disc-divider{border:none;border-top:1px solid #ddd;margin:6px 0;}
  .total-final{display:flex;justify-content:space-between;align-items:center;margin-top:4px;}
  .total-label{font-size:14px;color:#555;font-weight:600;}
  .total-amount{font-size:28px;font-weight:900;color:${badgeColor};}

  /* NOTES */
  .notes{background:#f9f9f9;border-left:3px solid ${badgeColor};padding:8px 12px;border-radius:0 6px 6px 0;font-size:13px;color:#555;margin-bottom:14px;}

  /* FOOTER */
  .footer{border-top:1px solid #f0f0f0;padding-top:8px;margin:0 0;font-size:10px;color:#bbb;display:flex;justify-content:space-between;align-items:center;}
  .footer-brand{color:#c0392b;font-weight:700;}
  /* TEST WATERMARK */
  .test-banner{background:#f1c40f;color:#1a1a1a;text-align:center;padding:8px;font-weight:900;font-size:13px;letter-spacing:2px;border-bottom:2px solid #d4ac0d;}
  .watermark{position:fixed;top:50%;left:50%;transform:translate(-50%,-50%) rotate(-35deg);font-size:80px;font-weight:900;color:rgba(241,196,15,0.15);pointer-events:none;white-space:nowrap;z-index:0;}
  .pdf-wrap{position:relative;z-index:1;}
</style></head><body>
<button class="no-print print-btn" onclick="window.print()">🖨️ Imprimir / Guardar PDF</button>
${doc.isTest ? `<div class="test-banner">⚠️ DOCUMENTO DE PRUEBA — NO VÁLIDO COMO COMPROBANTE</div>` : ""}
${doc.isTest ? `<div class="watermark">PRUEBA</div>` : ""}
<div class="pdf-wrap">
  <img class="header-img" src="${logoSrc}" alt="Libreria Madrid" onerror="this.style.display='none'"/>
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
  doRender("data:image/png;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCAH0DIADASIAAhEBAxEB/8QAHQABAAAHAQEAAAAAAAAAAAAAAAECAwQFBgcICf/EAGoQAAEDAgMDCAQIBwYQDAQFBQEAAgMEEQUGIRIxQQcTIlFhcYGRCDKhsRQjQlJywdHwCRUzYoKS4RZDU6Ky8RcYJDREVFZjc3WTlbPC0tMlNTY3OFd0g5SjpbQmRVV2RmSEhaTUxOLkw//EABwBAQACAwEBAQAAAAAAAAAAAAABAgMEBQYHCP/EAEcRAQACAQIDAwoDBgUCBQQCAwABAgMEEQUhMRJBUQYTMmFxgZGhsfAiwdEUM0JScuEHFSNi8VOCJDRDkrIXNaLCJnNjo9L/2gAMAwEAAhEDEQA/APZaIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiKUvYN72jvKCZFSdU0zd9REP0wpDXUY/smP9ZE7SuEVqcRoh/ZDfIqH4yov4cfqn7FG8HZnwXaK0/GVF/DfxT9ifjKi/hv4jvsTeDsz4LtFa/jGj/hv4p+xPxhR/wAMP1Sm8HZnwXSK3FdSH9/aphV0x3Tx/rKdzaVZFTE8J3SxnucFOHNO4g+KIRREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERA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3w7cx8YmPU5k/kQrJ6hpytm3D6qcG7IKyOSiqWnsY4F1+3RbBhVD6RWTy1tOajFaVm+OWeOpaR1DbO2B3WW6H0gOTaqGzU0mJa79qkaR/KWRpOXnk32WxtrKuFvAOprAeRSmDRVnemXsz6p/XdXU8U8pctOxqtB52vhfHv8AOvZ+jEYRy0Zlw8iPOvJzjFGB69TRwPLPBrhb+Ot6y1ypZFx9wjpMwU0FQdDBV/EPB6ulYE9xKzmXcxZfzNRipwbEqSviIuQxwLm97TqPELG5z5P8qZtgLMXwmB0wFmVEY2JWdzhr4LpUrnrXel4tHr/WP0eJ1GfhOXLNNRprYLd/Znfb/svG/ui0NoY5r2hzHBzSLgg3BUVwiqyByl5BeanIOY5cUw9pucOrCCbdQvofCyy+TOWymlxEYJnnC5cuYnfZD5GkRPPju79QprraxPZyx2Z9fT4ozeTOTJjnPw/JGekdezyvHtpPP4buwIpIJYp4WTQSMlieLtexwIcOsEKdbrzMxMTtIiIiBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEVCtrKaji5yplbG3hfee4cVh6nEK6siMkRbh1FxqJvWcOwffvUTOy9aTZk8RxKkoRaeTpndG3Vx8Fi6yrr54+cqZW4TSH52sr+wDf8AfirOkdfafhFPuvzmIVZ07SL/AH61LC2KWpJpY5MYrB61RNpEzuH37FSbbs1ccV+/vZVpC4RukwqlbBFbp11WdSOsXVOIQS1BNPFNjFWPWll0iYe76vJVnwsmmBr534lUj1aeHSJh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bzb+qoy+G/Y71beK2MWm1OnnfFbb28v7fNyddxrgfF69jW4O166/imPdG2SP/a7RkrlLyfm2NowzFomVBGtPOdiQeB3rM5myxgGZaN1NjWGU1Yxw0L2DaHaDvC8q5v5E8wYLTDF8r4hHjtKwbe3SOtLGOuw18lZ5L5aM8ZTmbR1szsSpozsup6y+223AO3hb8cTmn+nq6bevueTv5D49TH7VwDU9rb+GZ2tHv5THviPa71DkvNmRZjUZHxN2JYZfafg9c+4A/vbzuK2zKmdcNxub4BUxTYVi7R8ZQVY2JAfzTueO0LU8icumTcxiOnrag4PWusObqTZhPY7ct9xjBcDzJSsdWU8NUBrFPG6z2Hra8ahb2CaTG+ntvHh984eU4rXU47+b4vhmt/54jaZ9v8N/byn1sui16jp8cwS0XPvxigG4v0qIh37n+/sWbpKmGpj24nX6wRYg9RHBblbb9eTzubB2OdZi0eMfe8e9WREVmAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBEWPxTFqaicIgHT1LvVhj1ce/qSZ2TFZtO0L97msYXvcGtAuSTYBYaoxiaqe6DCIhJs+vUP0jZ29v33rG4lM98jTjMrnPcfi8Ppzcnq2iPv1dSp1hdsMZij/AIPD+9YdS+s7q2rdf3sVjmzYpiiOv3+v0TMdG6rPwZrsXr/lTP8AyMX838ylk2JKsc+52M1/CJn5GP6j7utVJIpXQNjrCMPpCOhR0/5ST6Xf2+9V3bNNShry3DaU7oo9ZpO/j5qF9/v7/JTlhMsjRic7quUepR02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BfeTwAuSsqvB/wCEgxXE5OU3L+CySSDDIMHFTDHfoGV80jXut12jYPDtQeoMD9IPkaxrFWYZQZ8w74S94YwTxywMc47gHyMa0+a6gCCAQQQdxC+M6+nvoeYrieMejnlSrxaSSadkU1OySQ3c+KKeSOPya0N/RQdbREQEREBERAREQERcb9LXkuoOUTkwrquKIR4/gtPJV4dUN0c7ZbtPhPW14Fh1Oseu4dkRfGrnJP4R3mp6aqqaaojqIKiWKaJ4fG9riC1wNwR23QfZJFqHIznGHP3JfgGbIy3nK6kaahrdzJ29GVvcHtdbsstvQEXBPTpzx+5PkSqcKpZubxDMUooI7GzhD60zu7ZAYf8ACBfOXnJP4R3mg+yqL4+5SwivzLmnCsvUDz8LxOsipIS4mwdI8NBPZqvq3yX5IwTk8yVQ5WwGHZp6Zl5JXDp1EpA25Xni5x8hYDQAINnREQEREBERAREQERYfOuWsIzhlXEMtY7TCow+vhMUrdxHEOaeDmmxB4EBBmEXyN5Tcr4hkXP8AjWUq+Z0k2GVToRJqOcZvY+3DaYWu8Vj8oySfurwjpu/r6Dj/AHxqD7BIiICIiAiKEhcGOLW7TgCQ29rnqQaRn/lc5N8h1gos15toMPrNnbNMA+aZo4EsjDnC/C41V3yf8pWRM/CX9yGZ6DFZIRtSQxuLJWDrMbwHAdtrL5SZkxTEsbzBiGL4xNJPiFZUPmqXyHpGRziTfx4LbvR3xXE8H5csmVWEySMqJMZpqdzWG3ORyyNjkYewtc4IPq0iIgIipVtLT1tHNR1kEdRTTxuimikaHNkY4WLSDvBBIsgqovlz6T3J2eS/lar8Bo3ynCqlja3DS5xJEDyRsE8S1zXNvxDQeK5hzkn8I7zQfZVFqnI1ryQZMJ/+gUP/ALdi2tAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBSTyxQROlmkbGxu9zjYBWOK4v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+PuccvYrlPNGI5bxun+D4hh07oJ2cLjiDxaRYg8QQVLlH/lZg/8A26D/AEjV3H8IFR01L6QUk8DQJKvCaaac9bxtxj+Kxq4dlI2zVhB6q6H/AEgQfYReIvwllJGzM2TK4NHOTUVTE49YY9hH+kPmvbq8W/hMf6+yH/gq/wB9Og8cruXo1ejxjPK22XGq6udg2WoJeaNUI9uWpeN7ImnSw4vOgOgB1tw1fWLkGwekwHkXyfhlGxrI2YPTSOsLbUkkYfI7xe5x8UHHav0LuTB+HGGmxrNENVs9GodUwvG11lvNAEdgt3ro3o0ZOxvk6yPWZExqdtW3CsRlOH1cbSGVFLLaRrrH1XbbpQW62I3kEE9SRB8leWqkjoOWPOlFC0Nigx+ujYBwaKh4HsW9+hF/0m8qfRrf/ZzLT/SC/wCfXPX/ANwVv+netw9CL/pN5U+jW/8As5kH0dzVi8OAZXxXHqgbUOG0U1XIL2u2NhefYF8hsdxStxvG67GcRlM1ZXVElTPIflPe4ucfMlfUj0lZHxcgOeHRgknBp2m3UW2PsJXyqQdD5A+SrGOVvO4wDDZ20VLBHz9fWvZtNp4rgaDTacSbBtxfXcASvXUvoWcmhwk08eO5nbXbOlUZ4S3a6yzmxp2Xv2rBfg1aSnblPOFc1rfhEtdTxPPHYbG4t9r3L1wg+TfLPyd4xyX59rMq4w5s5jAlpapjS1lTC71ZGg7txBGtnNIubXWlr2J+Ewo6ZmI5FxBrQKmaGuheetjDAWjwMj/NeO0H1s5Gv+aDJf8AiCh/9uxbYtT5GDfkeyWevL9B/wC3YtsQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERARQc4NaXOIAAuSeC0zMucaaN7qOjqRGAOnKBtOI/MA9507t6pkyVx17Vp2hmw4Mma3ZpG8s9i2MNpphRUcfwqudo2Nu5va77+S1LE8WpqGpMtTWQ1eJuB+MebxQW+Y0euR2A68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SHY4HivEPpy59oc6csjqLCZ2z0GAU/wASsN2yTbRdKWnqBIZ+gTuKDgi+uPJJIJuSnKMrTdr8DonA98DF8jl9KfQsz9Q5y5FMLwsTt/GuXoWYdWQE9JrGC0LwPmlgAv1tcOCDt6IsDyg5rwnI+TcTzTjcwiosPgMrhexkdubG385ziGjtKD5d8vcjZeXDPT2m4/dFXgHuqHhbf6FEgj9JvKRcbBxrG+dHOB7VyfHsSqMZxyvxerINRXVMlTLbdtvcXH2krOckmazkflLy/mzYdJHhtcyWZjfWfFe0jR2lhcB3oPqTyr4LJmTkwzRgMLS+avwmpp4QN5kdE4N/jWXyOIIJBBBG8FfYvBMUw/G8HpMXwqqjq6CshbPTzxm7ZGOFwR4L54emRyN4hkDPlZmbDKN78rYzUOmhlYLtpZnm74XfNF7lvAg23tKC59CLlcwjk4zliOD5mqRSYJjrIwap3qU08ZdsOd1MIe4E8Dsk6Ake+JM25Vjwk4vJmXBm4cG7ZqjXR81s9e3tWt4r5Aog7d6Y/Knh/KbymxuwKTnsDweA0tHNYj4Q4naklAO4E2A6w0HjZcRWz1uRseoeTmlz1X0r6XCq2vFFRGRpDqg7DnOe38wbIF9xJNtxWsIPrHyDTtqeRDI0rTe+XqEHvEDAfaCt1XAPQVz1RZn5FaTL5nb+NcuuNLURE9Iwuc50LwPm7J2O9h7F39AREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERARFYYzjOG4PBztfVMiv6rN73dzRqVW1q1je07QtSlslorWN5lfrFY7mDCsFYTW1I5212ws6UjvDh3mw7VzLN3KlPIH0+Fg0sdiC4EGQ6cTub4XPaFzepr8Rxadwa5xDztOJJ1PEknU95XJzcU7U9nTxvPj3f3en0fk1eY85q7divh3t8zvyjTYgX01M8wUoNuajd0n66FzvqHmVo0k2JV8nMxtMTXm+y0antP2lVKOiigu4ODnt9a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S00DKiWBsbg0tFuac02sSLblo39K5yE/3Df+rVv++XZkQc75PuRLkzyBj/49yjl2XC8Q5p0JkZidVIHMdva5j5S1w0B1BsQCNQF0REQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBEUsj2RRukke1jGi7nONgB1koJlb4hW0mH0zqmtqI4Im/KebXPUOs9gWkZv5S8Ow1j4cM2KmbcJXfkx3De72DtK49mLNWKYzVGSeole46C517gBo0dgXL1HFMdJ7OL8U/L79j0XD/JzUamO3l/BX19fh+rpuceVKOEPpsHaWHdzzwC/vDTo39K/cuU4jjGJYxVvLpJJHv1cS4kn6Tjw9itKWhmqXF0h2I26uJO7vKvmmCGJscbSIydAB0pT9/fxXOnHl1E9rPPLwemxRpOH17GmrvbxU6fD2EF8sjdlvrSEdEdgHE6H76q7jc0gQU0Twwn1RrJJ39Q0P7UEEkjg6pcGho6MY3MHaqpnZDF0Xc1GdHP8AlP7AtulK0jaOTTy5r5bb2neU7YAx2zKWyyt/e2n4uPtceJ0UstRHGCedtwL7exo8Vbbc04EcMZjZwb19pP38FVgjjjftRlsko3zPHRZ9EcTu1/nVt/Bh7Pj9/f3ule6eVjWBjoIn6tZ8uTiL9mquYKSKAh9RsOe3c13qs7SN3D+dQM8cGjS/nJN59aaS/wDJG79qpFkkzrTWa0a80x27te777+G5Sid59ULiSvc55MLi0nQzOF3dzRw3/wAyjHCWCx2oi/h600nf80ffVURJHAy7C1tvlke7ifvuQc/K27AYYnfLf6z/AL/e6ndG23Tkrl8MRs1rXPGuyHXDe1zlI6d0ji8fGub8oizGdg61bukpYDsEGplA9QGzR2uP38FbzVr5Ddxa7Z0AAtGzsA4qO0muPdekyPJklqOYYd77Xe76P238VLz9LTjZhhLb9Z2pH953BY90sjulctJ+WRd57hwUXRRx61MgiB+RtXee/q++ijtM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jqCyGH45i2H2FHiVTC0bmCQ7P6p0UV4vG/4q/NSnlHG/48fwn+z0CgK5bgPKVVRPbHjFM2ePjLCNl4727j7F0XB8VoMXpBU4fUsmj3G29p6iN4K6ODVYs/oTzdnS8Qwar93PPw718igFFbDdEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERARFLNLHDE6WaRkcbBdznGwaOslBMsdj2N4ZgdJ8JxKqZC032G73vPU0bytCztyq0dE19LgOzUS7jUvHQH0G/K7zYd641jeOV+LVclTWVMs0r9HPe65I6uwa7hoFy9RxKtPw4uc/L+703DvJvLn2vqPwV8O+f09/wb/njlSrsQ5ykwraoqUkjou+NePznD1e5vmVzOpqZJnXe7TgBuCpNDnuDWglx3AcVm8LwNz7SVQ7di/vP33rldnLqL9q87y9bX9l4dj7OONo+c+1jqChnrHdAbLL6vO79q2GkoaagjudXkfpH7B9irvlipwGQAXGgIHbwH33BY2pkLrl7rN42N77uPH9i3aYq4o9blZtVl1M7TyhUqaqWcmOHZaxu9x9Vg+tWzpIo2ODC8Md6zz68v2D768KM9QGgMLdBq2Me8/fikdO6RwkqC7X1WDe77Ak23lNaRWEAZJtWbLGt02vkt7B2/ftVeONkPAlx1sd57T1BHytYAQWAM0BA6LOxo4nt/nVu9znktIIvrs31Pa4/V9y6LbTPsTSzudozXhcbu4KADYxtykEjgdwVN8jY9ANp53WHu6gp6almqX3EZmcNzW+q3vO5Y7XivOZZK45npCUulqbkExxcXHe5TMe2PoUrCT8+1ye5XEkNNH/XlWHEfvUHS8zuVGbEYom83TU4Y3tOp71q31cb7Y43lu4tFNo3vyhTdA63x79gfNGr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nGKQ4dnDBJsOknaXwSbbZIpgN+w9pLSRcXF7i4uBddO9C7lXrsicpdHlytq3uy5j87aaeF7+hBO8hsczeo3s13W03PqherPTjy5TY76PeMVkkTXVWDyw11M+2rSJAx+vVsPd5DqXzaje+ORskbix7SHNcDYgjcQg+yyLE5MxQ45k/BcacLHEMPgqiOrnI2u+tZZAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERARFp+dOUHBMuB8DZBXV7dOYicLMP57tze7U9ipkyUx17V52hmwafLqL9jFXeW2zSxQROmmkZFGwXc97gA0dZJ3LmWdeVijog+ly81tVNuNTIPix9Eb3cddB3rmWcM641mWUisqNmnBuynju2NvUbfKPab9llrTr3JeSSVyM+vvfli5R49712h8nsWHa+pntT4R0989/09q+xrGcRxisfV19VLUTO0L3uubdQ4AdgsFj+7U9aja410Cr0r6Vj71DJHgbmtsL95WjWK77zL0VrWrXaK8o6RCagw+orJLRNuB6zzuatqwzCKehZzrzd9tZHDXwH33dqwzcxPhYGUtFDE1vqgkm27u6grSoxrEZ33dPsnhsAC27d5DyWxXLhx+uXPy6fV6jlP4Y+/Btckz5HGGnYR2Df3k8P2dqo1E9FhwL5pY5qjg2+jd/D78etatDHXVA2ecl2T1uNvLwCvI6CCmbtVUzGdh+xYsnEaxyiN5ZMfBf57bQnrK6aqlMjdqaQ7tNGhW8VDWTybWw0nf0je57bKpJiVLENmngdKfnP0HkrSfE6yUbPOmNnzY+iFqW1GfJ6nQx6LDjjaObIS4dG0h2I4lGwj5O8gdgCp89gsH5OCpqnDi9wjb7LlYgWvc6qa4CwWpNp/FMy260isbRyj1MqcZkZcU1LTQN4FrLu8zdWs+J1kvrVDz3GysibqLGukeGMYXOO4Diprhrv0TM1rzHyPebue495WQw3C5KhzXztkbGdWta3pP7uodqq0eFm+09zC7r2hst9mpV+aZkYN3lxPrEyD26exdLDp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8BqsXqKWOKCV7aqsZStc1ptckF3sVqVm07Qw6jLTFjm152h1DEsSlpMgV2JyuLZZaUmMcWun+LY3sLYg7zXMMg4SzGc1UVJOSKVrjNVO+bCwF7z+qCt19IHFqKLGospYO8Oo8MA+EPbuln2Q0+AAsO8rIej9lz4cznJI7nF6xlAw2/eGfG1B8WNDP010b087qYxxz2+/7PG6bURoODZNbaOzOTnHjET0+Eb29j01yc0MlHlGjfUxCKqrL1lQy3qPlO1sfoghvc0LYkVhmHE4MGwSsxSo1jpoi/Z4uPBo7SbDxXr6xGOm3dD865b31eom0R+K89PbPRforLAWzDCKZ1S7anfGHyHrcdSr1Wid43YL17Npr4CIilUREQFSnO4KqqT2uc8kBBj8awyjxjDZcOxCN0tNNYSMDy3aAINrgg7wqGEYBgmEtaMNwqjpS0W2o4QHHvdvPiVlxF1lTiNo4XVZpWZ7W3NjnFSb9uYjfx71CxPBWOMYThmMUhpcUoYKuE3s2RgOz2g7we0arL2HUpHRtPYVMxExtK1qxaNrRvDlGN8jGCVL3SYViFVh5O6N4EzB3XId5krVqjkVzE15EGKYXI3gXukYfINPvXfjEdbHuUvNu6lpX4bp7zv2dvY5GbgOhyzv2NvZO39nn9nItmgnpV+ENH+FkP8AqLN4RyJMFnYtjj3dcdLFb+M6/wDJXZebcoiNx36KteGaes77bq4/J7Q0nfs7+2Za3ljJuXMu2fhuHRif+Hl6cng47u4WC2JrSdwVQRDibqoABuC3qUrSNqxtDr4sVMVezjiIj1JGRhup1KnRFZkEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBSy+r4qZSy+r4oKSlJIUSpSSO5QsXB3oHOYeidFKbX6lAkjeFBsuGTtOh0Kqqxu1wUWyPj1BuFO5svUVKKdj9CbHtVVSqIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiweY82YBgDSMRxCNs1riCPpy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wZrxvY+tja7yJus85rXCzmgjtCsq/BsJr2bFbhtJUN6pIWu+pUtGTumG1itpP/Urb3TH02/NThx3Bagf1NjWGyk7tmpY73FXAdUyt2qeqpXN6xGXe5y1fEeS7IVcSZctUMRO8wM5o/xbKlTclGRKdwdFggBHEzvP1rHvn35xHxn9G52OF9neuS8T66Vn/wDePo2eSkxN/wD82MX+Dgb/AK11bSYRij92acTj+hBTfXEVSocoZforfBqAMt+e4/Ws1TwRU7NiJuy3qV4rM+l9Zal89Kfup39tKx+rAyZexZ+/O+PAdTYaMe6C6xuIYBDCCK/lCx2Dr266GH3MC2WrwXCax7n1WHU0znG5L4wbq2jytluM3ZgeHtPZA37FS2Lfu+ctnDr4rzm23spSPn/ZzTGsB5NJCfxzymYlV9cc2ZC/+KHaKwoG+jthMl4YsOrpr6mSGesLj+kHBdhjwLBYzePCqJp7IW/YrtlJSMADKaFtuqMLD+yc94ivw3/N1I8oY7HYtkzbeEZIrHwistOy7mXARGI8tZQxZsR3GnwoUrCOwyFgIW20NRUVDduWhkpW8GyvaX+Ibce1XQAAsAAOoItqlLVjnLg6nUY8tpmlNvXMzM/HlHyERFkaYiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICt5aGkkN308ZPXayuEVL46ZI2vG/tTFpr0lZHCqA/2OP1j9qfinD/7X/ju+1XqLB+w6b/p1+EMnnsn80/FYnCMOO+n/AI7vtUpwTDDvpv8AzHfasgifsOl/6dfhB5/LH8U/Fi3Zfwh2+k/8x/2qk/K2BP8AWob/APev+1ZlFH7Bpf8ApV+ELxq88dLz8Za/JkvLMnr4Zf8A7+T/AGlbScnuT5PXwcH/APUS/wC0tpRWjRaeOmOvwhkjiGrjplt/7p/VrNNkDJ1O8PjwCkLhuMm0/wDlErYaWmp6WBsFLBHBE3cyNoaB4BVUWetK15VjZiy6nNm/eXm3tmZERFZgEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERBY1+D4TX61uG0lQeuSFrj5kLX6/k3yfVkk4S2EnjC8t9m5bciralbdYZseoy4/QtMe9zSt5G8vykmlrayn7HWf8AYsLW8i0wJ+B4xE8f32Mt9112VFhnS4p/hblOL6ynS8uC1XJBmKK/NPo5/ovt77LGT8meaod+GF4/MeCvRqKk6LE2K8e1cdZiXmaXJGYoT8Zg1X4Rkq2flvFI/wAphlU3vhK9RIq/sNPFeOP5u+IeWvxNUt9ejnHfEgwyQb6eQf8AdL1IWtO9oPeFIYITvhjPe0J+xR4p/wA9vPWvzeYmUDh+9TDuj/aqzKN9vUn/AMmF6W+DU39rxfqBBT043QRfqBT+yR4qzxqZ/h+f9nm5lA/hDUnuaR7griPDKl+goKqT9Y/UvRgjjG6Ng7gplaNLHixzxe0/w/N57hwHEX/ksDqXf90VkKfK2YpANjBZGjtbb3ruiKY01fFjtxXJP8MONQZIzTLYGliib2yAe5ZCDk5xuRoE9dSxjquXH3LqqK8YKsVuI5p6bQ55S8mUezaqxaQji2NlgsnScneX4QOdFROR859h5LcEVoxUjuYbazPb+Jh6LLGA0luZwunv1ubtH2rKwwwwt2YYmRt6mNACnRXiIjowWva3WdxERSqIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgLH4jgmFYhc1VDC9x+Xs2d5jVZBEmN0xaazvDV58oNjJdh2I1FP1Mk6bftCsn0GY6E3NPHVs4ugfY+RW6oqdiO5mjPbv5tLix98LwyrZLA/cRNGR7VlaXFoJWg3FjxabhZyaKKZmxNEyRvU5oIWJqss4TMS+OF1M8/KheW+zcm1oT28dusbLiKpif6rweziqwIKwkmXsQhJNJiLZW/NnZ9YVLaxqi1noZHNGm1C4PHkm896OxWfRlsIsU2QN2ncsFBjsV9iQ7LhweC0+1ZKGvhf8rTr4KYmFZx2juXgBG4qNyN4VNkzHDRwKqAhSojtDiprqXRRsFKEwKCylso6ohGwupmixUtypmb0EyIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiIKU9PBO3ZnhjkH5zQVj5sv4c47UTZKd3XE8j2LKoo2iVotMdJYJ2DVsV/g9eJBwErPrCl/wCFqf8AKUpeOuNwd7Fn0Udlbzk97BsxQNOzM18bup7SFdw1sUgu14PcVkHNa4Wc0OHUQrWXDaKTU07Wnrbp7k2lG9Z7kzZWu3FThwPFWjsLa38jUys7+knwaujPRkjkHkVO8m0d0r0FTM3qxElUw9Oncfo6qvSziR5bZwIF7EJurNVyiIpQIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiIP/9k=");
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
  // ── SANDBOX STOCK — copia en memoria para el vendedor Prueba ──
  // Se inicializa con el stock real al cargar, nunca persiste en Supabase
  const [sandboxStock, setSandboxStock] = useState({}); // { pid: qty }
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
  const [loading, setLoading]   = useState(true);
  const [error, setError]       = useState(null);

  useEffect(() => {
    async function loadAll() {
      try {
        const [u,v,p,o,q,sl,act,pl,po,n] = await Promise.all([
          db.getUsers(), db.getVendors(), db.getProducts(),
          db.getOrders(), db.getQuotes(), db.getStockLog(), db.getActivity(), db.getPriceLists(), db.getPurchaseOrders(), db.getNotifs(),
        ]);
        setUsers(u); setVendors(v); setProducts(p);
        setOrders(o); setQuotes(q); setStockLog(sl); setActivity(act); setPriceLists(pl); setPurchaseOrders(po); setNotifs(n);
        // Inicializar sandbox con stock real
        const sbInit = {};
        p.forEach(prod => { sbInit[prod.id] = prod.stock; });
        setSandboxStock(sbInit);
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
  }}/>;
  return <MainApp
    currentUser={currentUser} onLogout={()=>{
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
  />;
}

// ─── MAIN APP (authenticated) ─────────────────────────────────────────────────
function MainApp({currentUser,onLogout,users,setUsers,vendors,setVendors,products,setProducts,orders,setOrders,quotes,setQuotes,stockLog,setStockLog,activity,setActivity,priceLists,setPriceLists,purchaseOrders,setPurchaseOrders,notifs,setNotifs,sandboxStock,setSandboxStock}) {
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
    // Assign Reserva-XXXXXX correlative number (skip counter for test orders)
    const test = isTestOrder(order.vendedor);
    const n = test ? 0 : await db.nextCounter("reserva");
    const orderWithNum = {...order, docNum: test ? "TEST-000000" : fmtDocNum("Reserva", n), isTest: test, isSandbox: test};

    if(test) {
      // SANDBOX: descontar del stock paralelo en memoria, no tocar Supabase
      setSandboxStock(prev => {
        const next = {...prev};
        orderWithNum.items.forEach(it => {
          next[it.pid] = Math.max(0, (next[it.pid] ?? 0) - it.qty);
        });
        return next;
      });
    } else {
      const updatedProds = products.map(x=>{const it=orderWithNum.items.find(i=>i.pid===x.id);return it?{...x,stock:Math.max(0,x.stock-it.qty)}:x;});
      setProducts(updatedProds);
      for(const p of updatedProds.filter(p=>orderWithNum.items.find(i=>i.pid===p.id))) await db.upsertProduct(p);
    }
    setOrders(o=>[orderWithNum,...o]);
    await db.upsertOrder(orderWithNum);
    const notif={id:genId(),fecha:new Date().toLocaleString("es-AR"),leida:[],tipo:"NUEVO_PEDIDO",para:"admin",icono:"🛒",titulo:"Nuevo pedido registrado",cuerpo:`${orderWithNum.client} - ${fARS(orderWithNum.total)} - ${orderWithNum.docNum}`,ref:orderWithNum.id};
    await db.addNotif(notif); setNotifs(n=>[notif,...n]);
    await logActivity("Nuevo pedido", `${orderWithNum.docNum} - ${orderWithNum.client} - ${fARS(orderWithNum.total)} - Vendedor: ${orderWithNum.vendedor||"-"}`, orderWithNum.id, "pedido");
    // Auto-print Reserva document
    setTimeout(() => printDoc(orderWithNum, "reserva"), 400);
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
    if(ord && ord.stage!=="entregado") {
      if(ord.isSandbox) {
        // SANDBOX: devolver stock al paralelo en memoria
        setSandboxStock(prev => {
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
    sendLocalNotif("✏️ Solicitud de edición", `${orders.find(o=>o.id===id)?.vendedor} quiere editar un pedido`, `edit-req-${id}`);
  };

  // Fase 2a: admin aprueba la solicitud → vendedor puede editar
  const approveEditRequest = async (id) => {
    const updated = orders.map(o=>o.id===id ? {...o, editStatus:"aprobada", editRejectReason:""} : o);
    setOrders(updated);
    await db.upsertOrder(updated.find(o=>o.id===id));
    sendLocalNotif("✅ Edición aprobada", `Tu solicitud de edición fue aprobada`, `edit-apr-${id}`);
  };

  // Fase 2b: admin rechaza la solicitud
  const rejectEditRequest = async (id, reason) => {
    const updated = orders.map(o=>o.id===id ? {...o, editStatus:"rechazada", editRejectReason:reason} : o);
    setOrders(updated);
    await db.upsertOrder(updated.find(o=>o.id===id));
    sendLocalNotif("❌ Edición rechazada", `Tu solicitud de edición fue rechazada`, `edit-rej-${id}`);
  };

  // Fase 3: vendedor guarda los cambios editados
  const submitEdit = async (id, newItems, newTotal) => {
    const updated = orders.map(o=>o.id===id ? {...o, editStatus:"en revisión", editItems:newItems, editTotal:newTotal} : o);
    setOrders(updated);
    await db.upsertOrder(updated.find(o=>o.id===id));
    sendLocalNotif("👀 Cambios para revisar", `${orders.find(o=>o.id===id)?.vendedor} editó un pedido — revisá los cambios`, `edit-sub-${id}`);
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
    sendLocalNotif("✅ Cambios aprobados", `Tu edición del pedido fue aprobada`, `edit-ok-${id}`);
    await logActivity("Edición aprobada", `Pedido ${ord.docNum||ord.compNum||""} editado`, id, "pedido");
  };

  // Fase 4b: admin rechaza los cambios finales
  const rejectEdit = async (id, reason) => {
    const updated = orders.map(o=>o.id===id ? {...o, editStatus:"cambios rechazados", editRejectReason:reason, editItems:null} : o);
    setOrders(updated);
    await db.upsertOrder(updated.find(o=>o.id===id));
    sendLocalNotif("❌ Cambios rechazados", `El admin rechazó tu edición`, `edit-no-${id}`);
  };
  const addQuote = async (quote) => {
    const test = isTestOrder(quote.vendedor);
    const n = test ? 0 : await db.nextCounter("presu");
    const quoteWithNum = {...quote, docNum: test ? "TEST-000000" : fmtDocNum("Presu", n), isTest: test};
    setQuotes(q=>[quoteWithNum,...q]);
    await db.upsertQuote(quoteWithNum);
    await logActivity("Nueva cotización", `${quoteWithNum.docNum} - ${quoteWithNum.client} - ${fARS(quoteWithNum.total)} - Vendedor: ${quoteWithNum.vendedor||"-"}`, quoteWithNum.id, "cotizacion");
  };
  const delQuote = async (id) => { setQuotes(q=>q.filter(x=>x.id!==id)); await db.deleteQuote(id); };
  const extendQuote = async (id, reason) => {
    const updated = quotes.map(q => q.id===id ? {...q, extendida:true, extendReason:reason, extendDate:today()} : q);
    setQuotes(updated);
    const quo = updated.find(x=>x.id===id);
    await db.upsertQuote({...quo, extend_reason:reason, extend_date:today()});
  };

  // Convierte una cotización en reserva — descuenta stock y arranca el circuito de ventas
  const convertQuoteToOrder = async (quote) => {
    const order = {
      id: genId(),
      client: quote.client,
      vendedor: quote.vendedor,
      notes: quote.notes,
      items: quote.items,
      total: quote.total,
      subtotal: quote.subtotal,
      globalDisc: quote.globalDisc,
      stage: "reserva",
      date: today(),
    };
    await addOrder(order);
    await logActivity("Cotización convertida a reserva", `${quote.docNum||""} - ${quote.client} - ${fARS(quote.total)}`, quote.id, "cotizacion");
    // Marcar la cotización como convertida (no eliminar, queda como historial)
    const updated = {...quote, convertida: true, ordenId: order.id};
    setQuotes(q=>q.map(x=>x.id===quote.id?updated:x));
    await db.upsertQuote({...updated});
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
    {k:"cotizacion",label:"Cotizaciones",       icon:"📄", roles:["admin","vendedor"]},
    {k:"precios",   label:"Precios",            icon:"💲", roles:["admin","vendedor"]},
    {k:"stock",     label:"Stock",              icon:"📦", roles:["admin","vendedor"]},
    {k:"compras",   label:"Alta de Mercancía",icon:"🏪", roles:["admin","vendedor"]},
    {k:"solicitud",  label:"Solicitud Compra", icon:"📋", roles:["admin","vendedor"]},
    {k:"admin",     label:"Administracion",     icon:"★",   roles:["admin"]},
  ].filter(t=>t.roles.includes(currentUser.role));

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
      // SOLICITUDES DE COMPRA
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

  // Pedir permiso de notificaciones al iniciar (solo una vez)
  useEffect(() => {
    if ("Notification" in window && Notification.permission === "default") {
      setTimeout(async () => {
        const result = await requestNotifPermission();
        setNotifPermission(result);
      }, 3000); // espera 3s para no abrumar al usuario
    }
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
    <div style={{minHeight:"100vh",background:"#f5f5f5",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>

      {/* ── HEADER ── */}
      <div style={{background:`linear-gradient(135deg,${REDD},${RED})`,boxShadow:"0 4px 16px #0004",position:"sticky",top:0,zIndex:100}}>
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
          orders={isAdmin || currentUser.canSeeAll!==false
            ? orders
            : isTestUser
              ? orders.filter(o=>o.isSandbox)
              : orders.filter(o=>o.vendedor===currentUser.vendedor||o.vendedor===currentUser.name||o.vendedor===currentUser.username)}
          products={pricedProducts} onStage={setStage} onDel={delOrder} onSaveNote={saveNote}
          onRequestEdit={requestEdit} onApproveEditRequest={approveEditRequest} onRejectEditRequest={rejectEditRequest}
          onSubmitEdit={submitEdit} onApproveEdit={approveEdit} onRejectEdit={rejectEdit}
          currentUser={currentUser} isMobile={isMobile}/>}
        {tab==="nuevo"      && <Nuevo products={pricedProducts} vendors={vendors} onAdd={addOrder} onDone={()=>setTab("central")} currentUser={currentUser} isMobile={isMobile}/>}
        {tab==="cotizacion" && <Cotizaciones quotes={quotes} products={pricedProducts} vendors={vendors} onAdd={addQuote} onDel={delQuote} onConvert={convertQuoteToOrder} onExtend={extendQuote} onTabChange={setTab} currentUser={currentUser} isMobile={isMobile}/>}
        {tab==="precios"    && <Precios products={pricedProducts}/>}
        {tab==="stock"      && <>
              {isTestUser && (
                <div style={{background:"#f5eef8",border:"1.5px solid #9b59b6",borderRadius:10,padding:"10px 16px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
                  <div>
                    <span style={{fontWeight:800,color:"#6c3483",fontSize:13}}>🧪 Modo Sandbox activo</span>
                    <span style={{color:"#888",fontSize:12,marginLeft:8}}>El stock que ves es una copia paralela. No afecta el stock real.</span>
                  </div>
                  <button onClick={()=>{const sb={};products.forEach(p=>{sb[p.id]=p.stock;});setSandboxStock(sb);}}
                    style={{padding:"6px 14px",borderRadius:8,border:"1.5px solid #9b59b6",background:"#fff",color:"#6c3483",fontWeight:700,fontSize:12,cursor:"pointer"}}>
                    🔄 Reiniciar sandbox
                  </button>
                </div>
              )}
              <Stock products={pricedProducts} onUpd={updProd} onDel={pid=>setProducts(p=>p.filter(x=>x.id!==pid))} onAdjust={(pid,qty)=>setProducts(p=>p.map(x=>x.id===pid?{...x,stock:x.stock+qty}:x))} isAdmin={isAdmin} addLog={addLog} stockLog={stockLog} setStockLog={setStockLog} isMobile={isMobile}/>
            </>}
        {tab==="compras"    && <Compras products={products} onStock={addStock} isMobile={isMobile}/>}
        {tab==="solicitud"  && <SolicitudCompra products={products} currentUser={currentUser} isAdmin={isAdmin} purchaseOrders={purchaseOrders} setPurchaseOrders={setPurchaseOrders} isMobile={isMobile} onStockExternal={addStock} addLog={addLog}/>}
        {tab==="admin"      && isAdmin && <AdminPanel users={users} setUsers={setUsers} vendors={vendors} setVendors={setVendors} products={products} setProducts={setProducts} stockLog={stockLog} setStockLog={setStockLog} notifs={notifs} setNotifs={setNotifs} activity={activity} setActivity={setActivity} orders={orders} priceLists={priceLists} setPriceLists={setPriceLists} isMobile={isMobile}/>}
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
  const deliv = orders.filter(o=>o.stage==="entregado"&&!o.isSandbox).reduce((s,o)=>s+o.total,0);
  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(135px,1fr))",gap:12,marginBottom:20}}>
        {STAGES.map(s=>{const c=SCFG[s],cnt=orders.filter(o=>o.stage===s&&!o.isSandbox).length;return <div key={s} onClick={()=>setFStage(fStage===s?"todos":s)} style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 6px #0001",borderLeft:`4px solid ${c.color}`,cursor:"pointer",outline:fStage===s?`2px solid ${c.color}`:"none"}}><div style={{fontSize:26,fontWeight:800,color:c.color}}>{cnt}</div><div style={{fontSize:12,color:"#666",fontWeight:600}}>{c.icon} {c.label}</div></div>;})}
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
    <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 6px #0001",overflow:"hidden",marginBottom:8}}>
      <div onClick={toggle} style={{padding:"13px 18px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",cursor:"pointer"}}>
        <div style={{flex:1,minWidth:100}}>
          <div style={{fontWeight:700,fontSize:14,color:"#1a1a1a"}}>{o.client}</div>
          <div style={{fontSize:11,color:"#aaa",display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
            {o.isTest&&<span style={{background:"#f1c40f",color:"#1a1a1a",borderRadius:4,padding:"1px 6px",fontSize:10,fontWeight:800}}>TEST</span>}
            {o.isSandbox&&<span style={{background:"#9b59b6",color:"#fff",borderRadius:4,padding:"1px 6px",fontSize:10,fontWeight:800}}>🧪 SANDBOX</span>}
            {o.docNum&&!o.isTest&&<span style={{fontWeight:700,color:"#c0392b"}}>{o.docNum}</span>}
            {o.compNum&&!o.isTest&&<span style={{fontWeight:700,color:"#1a5276"}}>{o.compNum}</span>}
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
              {editItems.map((it,i)=>{
                const p = getP(it.pid);
                return (
                  <div key={it.pid} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 0",borderBottom:"1px solid #d6eaf8",flexWrap:"wrap"}}>
                    <div style={{flex:1,fontSize:13,fontWeight:600,color:"#1a1a1a"}}>{p?.name||it.name}</div>
                    <div style={{display:"flex",alignItems:"center",gap:6}}>
                      <button onClick={()=>updEditItem(it.pid,it.qty-1)} style={{width:28,height:28,borderRadius:7,border:"1.5px solid #aed6f1",background:"#fff",fontWeight:800,cursor:"pointer",fontSize:14}}>−</button>
                      <span style={{minWidth:28,textAlign:"center",fontWeight:700}}>{it.qty}</span>
                      <button onClick={()=>updEditItem(it.pid,it.qty+1)} style={{width:28,height:28,borderRadius:7,border:"1.5px solid #aed6f1",background:"#fff",fontWeight:800,cursor:"pointer",fontSize:14}}>+</button>
                    </div>
                    <span style={{fontWeight:700,color:RED,minWidth:80,textAlign:"right"}}>{fARS(it.price*it.qty)}</span>
                    <button onClick={()=>remEditItem(it.pid)} style={{background:"none",border:"none",color:"#c0392b",fontSize:18,cursor:"pointer",lineHeight:1}}>×</button>
                  </div>
                );
              })}
              <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:15,color:RED,padding:"10px 0",borderTop:"2px solid #d6eaf8",margin:"8px 0"}}>
                <span>Nuevo total</span><span>{fARS(editTotal)}</span>
              </div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={async()=>{if(!editItems.length)return;setSaving(true);await onSubmitEdit(o.id,editItems,editTotal);setShowEditMode(false);setSaving(false);}}
                  disabled={!editItems.length||saving}
                  style={{padding:"8px 16px",borderRadius:8,border:"none",background:editItems.length?"linear-gradient(135deg,#1a5276,#2980b9)":"#e5e5e5",color:editItems.length?"#fff":"#aaa",fontWeight:700,fontSize:13,cursor:editItems.length?"pointer":"not-allowed"}}>
                  {saving?"Enviando...":"📤 Enviar para revisión"}
                </button>
                <button onClick={()=>setShowEditMode(false)}
                  style={{padding:"8px 12px",borderRadius:8,border:"1px solid #e5e5e5",background:"#fff",color:"#666",fontSize:13,cursor:"pointer"}}>Cancelar</button>
              </div>
            </div>
          )}

          {/* Admin: cambios en revisión — comparación original vs nuevo */}
          {isAdmin && es==="en revisión" && (
            <div style={{background:"#eaf4fc",border:"1.5px solid #aed6f1",borderRadius:10,padding:"14px",marginBottom:14}}>
              <div style={{fontWeight:800,fontSize:13,color:"#1a5276",marginBottom:12}}>👀 Revisión de cambios — {o.vendedor}</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:12}}>
                <div>
                  <div style={{fontSize:11,fontWeight:700,color:"#888",marginBottom:6}}>ORIGINAL</div>
                  {o.items.map((it,i)=>{const p=getP(it.pid);return <div key={i} style={{fontSize:12,color:"#555",padding:"3px 0",borderBottom:"1px solid #f0f0f0"}}>{p?.name||it.name} × {it.qty} — {fARS(it.price*it.qty)}</div>;})}
                  <div style={{fontWeight:700,fontSize:13,color:"#555",marginTop:6}}>{fARS(o.total)}</div>
                </div>
                <div>
                  <div style={{fontSize:11,fontWeight:700,color:"#1a5276",marginBottom:6}}>NUEVO</div>
                  {(o.editItems||[]).map((it,i)=>{
                    const p=getP(it.pid);
                    const orig=o.items.find(x=>x.pid===it.pid);
                    const changed=!orig||orig.qty!==it.qty;
                    return <div key={i} style={{fontSize:12,color:changed?"#1a5276":"#555",fontWeight:changed?700:400,padding:"3px 0",borderBottom:"1px solid #f0f0f0"}}>{p?.name||it.name} × {it.qty} — {fARS(it.price*it.qty)}{changed?" ✏️":""}</div>;
                  })}
                  <div style={{fontWeight:800,fontSize:13,color:RED,marginTop:6}}>{fARS((o.editItems||[]).reduce((s,it)=>s+it.price*it.qty,0))}</div>
                </div>
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
  const CATS=useMemo(()=>["todos",...new Set(products.map(p=>p.category))].sort(),[products]);
  const shown=useMemo(()=>{
    const q=search.toLowerCase();
    return products.filter(p=>{
      if(cat!=="todos"&&p.category!==cat)return false;
      if(q)return norm(p.name).includes(norm(q))||normSKU(p.id).includes(normSKU(q));
      return true;
    }).slice(0,80);
  },[products,search,cat]);
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

function Nuevo({products,vendors,onAdd,onDone,currentUser,isMobile}) {
  const [client,setClient]=useState("");
  const [notes,setNotes]=useState("");
  const [vendedor,setVendedor]=useState(currentUser.role==="vendedor"?(currentUser.vendedor||currentUser.name):"");
  const [cart,setCart]=useState([]);
  const [globalDisc,setGlobalDisc]=useState({type:"%",value:""});
  const [ok,setOk]=useState(false);
  const [mStep,setMStep]=useState(1);

  const subtotal=cart.reduce((s,i)=>s+applyItemDiscount(i.price,i.qty,i.disc),0);
  const total=applyGlobalDiscount(subtotal,globalDisc);
  const globalDiscAmt=subtotal-total;

  const submit=()=>{
    if(!client.trim()){alert("Ingresá el cliente");return;}
    if(!vendedor&&currentUser.role==="admin"){alert("Seleccioná un vendedor");return;}
    if(!cart.length){alert("Agregá productos");return;}
    onAdd({id:genId(),client:client.trim(),notes,vendedor:vendedor||currentUser.vendedor||currentUser.name,items:cart,total,subtotal,globalDisc,stage:"reserva",date:today()});
    setOk(true); setTimeout(()=>onDone(),1400);
  };

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
            <Field label="Cliente *"><input value={client} onChange={e=>setClient(e.target.value)} placeholder="Nombre del cliente" style={inputStyle}/></Field>
            {currentUser.role==="admin"&&<Field label="Vendedor *"><select value={vendedor} onChange={e=>setVendedor(e.target.value)} style={{...inputStyle,color:vendedor?"#1a1a1a":"#aaa",cursor:"pointer"}}><option value="">— Seleccioná vendedor —</option>{vendors.map(v=><option key={v} value={v}>{v}</option>)}</select></Field>}
            <Field label="Notas"><textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Observaciones..." style={{...inputStyle,resize:"vertical",minHeight:55,fontSize:12}}/></Field>
            <button onClick={()=>{if(!client.trim()){alert("Ingresá el cliente");return;}if(!vendedor&&currentUser.role==="admin"){alert("Seleccioná un vendedor");return;}setMStep(2);}}
              style={{width:"100%",padding:"12px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:`linear-gradient(135deg,${REDD},${RED})`,color:"#fff",marginTop:8}}>
              Siguiente → Productos
            </button>
          </div>
        )}

        {mStep===2 && (
          <div style={{flex:1,overflow:"auto"}}>
            <ProductSelector products={products} cart={cart} setCart={setCart} isMobile={true}/>
            {cart.length>0&&(
              <div style={{position:"sticky",bottom:0,background:`linear-gradient(135deg,${REDD},${RED})`,color:"#fff",padding:"12px 16px",borderRadius:"0 0 12px 12px",display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:8}}>
                <span style={{fontWeight:700}}>{cart.length} producto{cart.length!==1?"s":""} · {fARS(cart.reduce((s,i)=>s+i.price*i.qty,0))}</span>
                <button onClick={()=>setMStep(3)} style={{padding:"8px 16px",borderRadius:8,border:"none",background:"#fff",color:RED,fontWeight:800,fontSize:13,cursor:"pointer"}}>Ver resumen →</button>
              </div>
            )}
          </div>
        )}

        {mStep===3 && (
          <div style={{background:"#fff",borderRadius:12,padding:16,boxShadow:"0 1px 4px #0001"}}>
            <div style={{fontWeight:800,fontSize:15,marginBottom:14,color:"#1a1a1a"}}>📋 Confirmar pedido</div>
            <div style={{fontSize:13,color:"#555",marginBottom:4}}>👤 <strong>{client}</strong> · {vendedor}</div>
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
          <Field label="Cliente *"><input value={client} onChange={e=>setClient(e.target.value)} placeholder="Nombre del cliente" style={inputStyle}/></Field>
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
          <button onClick={submit} disabled={!cart.length||!client.trim()||(!vendedor&&currentUser.role==="admin")} style={{width:"100%",padding:"11px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:(!cart.length||!client.trim()||(!vendedor&&currentUser.role==="admin"))?"#e5e5e5":`linear-gradient(135deg,${REDD},${RED})`,color:(!cart.length||!client.trim()||(!vendedor&&currentUser.role==="admin"))?"#aaa":"#fff"}}>
            ✅ Registrar como Reserva
          </button>
        </div>
      </div>
    </div>
  );
}

function Cotizaciones({quotes,products,vendors,onAdd,onDel,onConvert,onExtend,onTabChange,currentUser}) {
  const [view,setView]=useState("lista");
  const [expanded,setExpanded]=useState(null);
  const getP=id=>products.find(p=>p.id===id);
  return (
    <div>
      <div style={{background:"#fff",borderRadius:12,padding:4,marginBottom:16,display:"flex",gap:4,boxShadow:"0 1px 4px #0001"}}>
        <button onClick={()=>setView("lista")} style={{flex:1,padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="lista"?`linear-gradient(135deg,${REDD},${RED})`:"transparent",color:view==="lista"?"#fff":"#555"}}>📄 Lista de Cotizaciones ({quotes.filter(q=>!q.convertida).length})</button>
        <button onClick={()=>setView("nueva")} style={{flex:1,padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="nueva"?`linear-gradient(135deg,${REDD},${RED})`:"transparent",color:view==="nueva"?"#fff":"#555"}}>➕ Nueva Cotización</button>
      </div>
      {view==="nueva" && <NuevaCotizacion products={products} vendors={vendors} onAdd={async(q)=>{await onAdd(q);setView("lista");}} currentUser={currentUser}/>}
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

  return (
    <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 6px #0001",overflow:"hidden",marginBottom:8}}>
      <div onClick={toggle} style={{padding:"13px 18px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",cursor:"pointer"}}>
        <div style={{flex:1,minWidth:100}}>
          <div style={{fontWeight:700,fontSize:14,color:"#1a1a1a"}}>{q.client}</div>
          <div style={{fontSize:11,color:"#aaa",display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
            {q.isTest&&<span style={{background:"#f1c40f",color:"#1a1a1a",borderRadius:4,padding:"1px 6px",fontSize:10,fontWeight:800}}>TEST</span>}
            {q.docNum&&!q.isTest&&<span style={{fontWeight:700,color:"#6c3483"}}>{q.docNum}</span>}
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

function NuevaCotizacion({products,vendors,onAdd,currentUser}) {
  const [client,setClient]=useState("");
  const [notes,setNotes]=useState("");
  const [vendedor,setVendedor]=useState(currentUser?.role!=="admin" ? (currentUser?.vendedor||currentUser?.name||"") : "");
  const [validity,setValidity]=useState("48 horas");
  const [cart,setCart]=useState([]);
  const [globalDisc,setGlobalDisc]=useState({type:"%",value:""});
  const [ok,setOk]=useState(false);

  const subtotal = cart.reduce((s,i)=>s+applyItemDiscount(i.price,i.qty,i.disc),0);
  const total    = applyGlobalDiscount(subtotal, globalDisc);
  const globalDiscAmt = subtotal - total;

  const submit=async()=>{
    if(!client.trim()){alert("Ingresá el cliente");return;}
    if(!cart.length){alert("Agregá productos");return;}
    const q={id:genId(),client:client.trim(),notes,vendedor,validity,items:cart,total,subtotal,globalDisc,date:today()};
    setOk(true);
    await onAdd(q);
  };
  if(ok) return <div style={{textAlign:"center",padding:80}}><div style={{fontSize:60}}>📄</div><div style={{fontWeight:800,color:"#6c3483",fontSize:20,marginTop:12}}>!Cotización guardada!</div></div>;
  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 330px",gap:18,alignItems:"start"}}>
      <div>
        <div style={{fontWeight:800,fontSize:15,marginBottom:12}}>📄 Nueva Cotización - Seleccioná productos</div>
        <ProductSelector products={products} cart={cart} setCart={setCart}/>
      </div>
      <div style={{position:"sticky",top:16}}>
        <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 2px 12px #0002",border:"2px solid #e8daef"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14,color:"#6c3483"}}>📄 Resumen de Cotización</div>
          <div style={{background:"#e8daef",borderRadius:8,padding:"7px 12px",fontSize:12,color:"#6c3483",marginBottom:14}}>
            ℹ️ Las cotizaciones <strong>no descuentan stock</strong>. Son solo presupuestos para el cliente.
          </div>
          <Field label="Cliente *"><input value={client} onChange={e=>setClient(e.target.value)} placeholder="Nombre del cliente" style={inputStyle}/></Field>
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
                  <span style={{fontWeight:600,color:hasD?"#6c3483":undefined}}>{fARS(lineTotal)}</span>
                  {hasD&&<span style={{fontSize:10,color:"#6c3483",marginLeft:3}}>{fmtDisc(i.disc)}</span>}
                </div>
              </div>;
            })}
          </div>
          {/* Global discount */}
          <div style={{background:"#f5f0fa",border:"1.5px solid #e8daef",borderRadius:8,padding:"10px 12px",marginBottom:8}}>
            <div style={{fontSize:11,color:"#6c3483",fontWeight:700,marginBottom:6}}>DESCUENTO GLOBAL</div>
            <div style={{display:"flex",gap:6,alignItems:"center"}}>
              <select value={globalDisc.type} onChange={e=>setGlobalDisc(d=>({...d,type:e.target.value}))}
                style={{padding:"5px 6px",borderRadius:6,border:"1.5px solid #e5e5e5",fontSize:13,fontWeight:700,background:"#fff",cursor:"pointer",width:48}}>
                <option value="%">%</option>
                <option value="$">$</option>
              </select>
              <input type="number" min="0" value={globalDisc.value} onChange={e=>setGlobalDisc(d=>({...d,value:e.target.value}))}
                placeholder="0" style={{flex:1,padding:"5px 8px",borderRadius:6,border:"1.5px solid #ccc",fontSize:13,fontWeight:700,outline:"none",textAlign:"center"}}/>
              {globalDiscAmt>0&&<span style={{fontSize:11,color:"#6c3483",fontWeight:700,whiteSpace:"nowrap"}}>-{fARS(globalDiscAmt)}</span>}
            </div>
          </div>
          {globalDiscAmt>0&&<div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:"#888",padding:"2px 0"}}>
            <span>Subtotal</span><span>{fARS(subtotal)}</span>
          </div>}
          <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:17,color:"#6c3483",padding:"8px 0",borderTop:"2px solid #f5f5f5",margin:"6px 0 14px"}}><span>Total</span><span>{fARS(total)}</span></div>
          <button onClick={submit} disabled={!cart.length||!client.trim()} style={{width:"100%",padding:"11px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:(!cart.length||!client.trim())?"#e5e5e5":"linear-gradient(135deg,#6c3483,#9b59b6)",color:(!cart.length||!client.trim())?"#aaa":"#fff"}}>
            📄 Guardar Cotización
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
  body{background:#fff;}
  @media print{.no-print{display:none!important;}@page{margin:3mm;size:A4;}}
  .btn{display:block;margin:12px auto;padding:9px 28px;background:#1a5276;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;}
  .header-img{width:100%;height:auto;display:block;border-bottom:3px solid #c0392b;}
  .content{padding:16px 20px 0;}
  .doc-meta{display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;padding-bottom:12px;border-bottom:2px solid #f0f0f0;}
  .badge{display:inline-block;padding:6px 18px;border-radius:8px;font-size:14px;font-weight:800;background:#1a5276;color:#fff;}
  .info-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px 16px;margin-bottom:16px;}
  .info-box{background:#f9f9f9;border-radius:8px;padding:9px 12px;border-left:3px solid #e5e5e5;}
  .info-lbl{font-size:8px;color:#999;text-transform:uppercase;letter-spacing:.7px;margin-bottom:3px;font-weight:700;}
  .info-val{font-size:13px;font-weight:700;}
  table{width:100%;border-collapse:collapse;margin-bottom:16px;}
  thead tr{background:#f5f5f5;}
  th{padding:9px 12px;text-align:left;font-size:10px;color:#888;text-transform:uppercase;font-weight:700;}
  th.c{text-align:center;}
  tbody tr:nth-child(even){background:#fafafa;}
  .footer{border-top:1px solid #f0f0f0;padding-top:10px;margin:0;font-size:10px;color:#bbb;display:flex;justify-content:space-between;}
  .footer-brand{color:#c0392b;font-weight:700;}
  .notes{background:#f9f9f9;border-left:3px solid #1a5276;padding:8px 12px;border-radius:0 6px 6px 0;font-size:12px;color:#555;margin-bottom:14px;}
</style></head><body>
<button class="no-print btn" onclick="window.print()">🖨️ Imprimir / Guardar PDF</button>
<div>
  <img class="header-img" src="${logoSrc}" alt="Libreria Madrid" onerror="this.style.display='none'"/>
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

function SolicitudCompra({products,currentUser,isAdmin,purchaseOrders,setPurchaseOrders,isMobile,onStockExternal,addLog}) {
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
function Compras({products,onStock,isMobile}) {
  const [search,setSearch]=useState("");
  const [items,setItems]=useState([]);
  const [ok,setOk]=useState(false);
  const [showManual, setShowManual] = useState(false);
  const [manualForm, setManualForm] = useState({sku:"", name:"", qty:1, cost:0});
  const [manualError, setManualError] = useState("");

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
    setItems([]); setOk(true); setTimeout(()=>{setOk(false);},2000);
  };
  if(ok) return <div style={{textAlign:"center",padding:80}}><div style={{fontSize:60}}>📦</div><div style={{fontWeight:800,color:"#1e8449",fontSize:20,marginTop:12}}>¡Stock actualizado!</div></div>;
  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 340px",gap:18,alignItems:"start"}}>
      <div>
        <div style={{background:"#fff",borderRadius:12,padding:16,marginBottom:12,boxShadow:"0 1px 4px #0001"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:12}}>📥 Ingresar al Stock</div>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscá los productos que recibiste..."
            style={{width:"100%",padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",boxSizing:"border-box"}}/>
          {search&&<div style={{fontSize:11,color:"#aaa",marginTop:6}}>{found.length} resultados</div>}
          {/* Producto manual */}
          <div style={{marginTop:12,borderTop:"1px solid #f0f0f0",paddingTop:12}}>
            <button onClick={()=>{setShowManual(s=>!s);setManualError("");}}
              style={{padding:"6px 14px",borderRadius:8,border:"1.5px solid #e67e22",background:showManual?"#fef9e7":"#fff",color:"#e67e22",fontWeight:700,fontSize:12,cursor:"pointer"}}>
              {showManual?"✕ Cancelar":"➕ Producto no está en el catálogo"}
            </button>
            {showManual && (
              <div style={{background:"#fef9e7",border:"1.5px solid #e67e22",borderRadius:10,padding:14,marginTop:10}}>
                <div style={{fontWeight:700,fontSize:13,color:"#e67e22",marginBottom:10}}>⚠️ Producto nuevo (excepcional)</div>
                <div style={{fontSize:11,color:"#888",marginBottom:12,lineHeight:1.5}}>
                  Solo usá esto si el producto no figura en el Excel. Generará un nuevo SKU en el catálogo.
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
                  <div>
                    <div style={{fontSize:10,fontWeight:700,color:"#666",marginBottom:4}}>SKU *</div>
                    <input value={manualForm.sku} onChange={e=>setManualForm(f=>({...f,sku:e.target.value.toUpperCase()}))}
                      placeholder="Ej: PROD-001"
                      style={{...inputStyle,fontSize:12,padding:"6px 10px"}}/>
                  </div>
                  <div>
                    <div style={{fontSize:10,fontWeight:700,color:"#666",marginBottom:4}}>Cantidad *</div>
                    <input type="number" min={1} value={manualForm.qty} onChange={e=>setManualForm(f=>({...f,qty:e.target.value}))}
                      style={{...inputStyle,fontSize:12,padding:"6px 10px"}}/>
                  </div>
                </div>
                <div style={{marginBottom:8}}>
                  <div style={{fontSize:10,fontWeight:700,color:"#666",marginBottom:4}}>Detalle del producto *</div>
                  <input value={manualForm.name} onChange={e=>setManualForm(f=>({...f,name:e.target.value}))}
                    placeholder="Nombre completo del producto"
                    style={{...inputStyle,fontSize:12,padding:"6px 10px"}}/>
                </div>
                <div style={{marginBottom:12}}>
                  <div style={{fontSize:10,fontWeight:700,color:"#666",marginBottom:4}}>Precio de costo ($)</div>
                  <input type="number" min={0} value={manualForm.cost} onChange={e=>setManualForm(f=>({...f,cost:e.target.value}))}
                    style={{...inputStyle,fontSize:12,padding:"6px 10px"}}/>
                </div>
                {manualError && <div style={{color:"#c0392b",fontSize:12,marginBottom:8,fontWeight:600}}>⚠️ {manualError}</div>}
                <button onClick={addManual}
                  style={{width:"100%",padding:"8px",borderRadius:8,border:"none",background:"#e67e22",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}>
                  ➕ Agregar a la compra
                </button>
              </div>
            )}
          </div>
        </div>
        <div style={{display:isMobile?"flex":"grid",flexDirection:isMobile?"column":undefined,gridTemplateColumns:isMobile?undefined:"repeat(auto-fill,minmax(195px,1fr))",gap:isMobile?8:10,padding:isMobile?"0 12px":0}}>
          {found.length>0
            ? found.map(p=>{
                const inL=items.find(i=>i.pid===p.id);
                return <div key={p.id} style={{background:"#fff",borderRadius:10,padding:14,border:inL?"2px solid #1e8449":"2px solid transparent",boxShadow:"0 1px 4px #0001"}}>
                  <div style={{fontWeight:700,fontSize:12,color:"#1a1a1a",marginBottom:3,lineHeight:1.3}}>{p.name}</div>
                  <div style={{fontSize:10,color:"#aaa",marginBottom:8}}>{p.id} . Stock actual: <strong>{p.stock}</strong></div>
                  <button onClick={()=>addI(p)} disabled={!!inL} style={{width:"100%",padding:"7px",borderRadius:7,border:"none",fontSize:12,fontWeight:700,background:inL?"#d5f5e3":"#1e8449",color:inL?"#1a5276":"#fff",cursor:inL?"not-allowed":"pointer"}}>{inL?"✓ Agregado":"+ Agregar a la compra"}</button>
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
                  <div style={{fontSize:11,color:"#666",marginTop:6}}>Subtotal: <strong>{fARS(it.qty*it.cost)}</strong> · Nuevo venta: <strong style={{color:RED}}>{fARS(it.cost*1.5)}</strong></div>
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

// ─── ADMIN PANEL ──────────────────────────────────────────────────────────────
function AdminPanel({users,setUsers,vendors,setVendors,products,setProducts,stockLog,setStockLog,notifs,setNotifs,activity,setActivity,orders,priceLists,setPriceLists,isMobile}) {
  const [section, setSection] = useState("vendors");

  const SECTIONS = [
    {k:"ventas",      label:"Ventas",           icon:"📈"},
    {k:"sandbox",     label:"Demo Sandbox",     icon:"🧪"},
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
      {section==="ventas"      && <VentasPanel    orders={realOrders}/>}
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
  const [form, setForm] = useState({username:"",password:"",name:"",role:"vendedor",vendedor:"",priceList:"default",canSeeAll:true});
  const [editing, setEditing] = useState(null);
  const [showPass, setShowPass] = useState({});

  const startEdit = (u) => { setEditing(u.id); setForm({username:u.username,password:u.password,name:u.name,role:u.role,email:u.email||"",vendedor:u.vendedor||"",priceList:u.priceList||"default",canSeeAll:u.canSeeAll!==false}); };
  const cancelEdit = () => { setEditing(null); setForm({username:"",password:"",name:"",role:"vendedor"}); };

  const save = async () => {
    if(!form.username.trim()||!form.password.trim()||!form.name.trim()){alert("Completa todos los campos");return;}
    try {
      if(editing) {
        const updated = {...users.find(u=>u.id===editing), ...form, priceList:form.priceList||"default", canSeeAll:form.canSeeAll!==false};
        setUsers(us=>us.map(u=>u.id===editing?updated:u));
        await db.saveUser(updated);
      } else {
        if(users.find(u=>u.username===form.username.trim())){alert("Ese usuario ya existe");return;}
        const newUser = {id:genId(),username:form.username.trim(),password:form.password,name:form.name.trim(),role:form.role||"vendedor",email:form.email||"",vendedor:form.vendedor||"",priceList:form.priceList||"default",canSeeAll:form.canSeeAll!==false};
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

  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,alignItems:"start"}}>
      <div style={{background:"#fff",borderRadius:12,padding:24,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:15,marginBottom:16}}>{editing?"✏️️ Editar usuario":"+ Nuevo usuario"}</div>
        <Field label="Nombre completo">
          <input value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} placeholder="Ej: María García" style={inputStyle}/>
        </Field>
        <Field label="Usuario (para login)">
          <input value={form.username} onChange={e=>setForm(f=>({...f,username:e.target.value}))} placeholder="Ej: maria" style={inputStyle}/>
        </Field>
        <Field label="Contraseña">
          <div style={{position:"relative"}}>
            <input type={showPass.form?"text":"password"} value={form.password}
              onChange={e=>setForm(f=>({...f,password:e.target.value}))}
              placeholder="Contraseña segura" style={{...inputStyle,paddingRight:40}}/>
            <button onClick={()=>setShowPass(s=>({...s,form:!s.form}))} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:15,color:"#aaa"}}>{showPass.form?"🙈":"👁️"}</button>
          </div>
        </Field>
        <Field label="Email (para notificaciones)">
          <input type="email" value={form.email||""} onChange={e=>setForm(f=>({...f,email:e.target.value}))}
            placeholder="correo@ejemplo.com" style={inputStyle}/>
        </Field>
        <Field label="Rol">
          <select value={form.role} onChange={e=>setForm(f=>({...f,role:e.target.value}))} style={{...inputStyle,cursor:"pointer"}}>
            <option value="vendedor">Vendedor</option>
            <option value="admin">Administrador</option>
          </select>
        </Field>
        <Field label="Vendedor asignado">
          <select value={form.vendedor||""} onChange={e=>setForm(f=>({...f,vendedor:e.target.value}))} style={{...inputStyle,cursor:"pointer",color:form.vendedor?"#1a1a1a":"#aaa"}}>
            <option value="">- Sin asignar -</option>
            {(vendors||[]).map(v=><option key={v} value={v}>{v}</option>)}
          </select>
          <div style={{fontSize:10,color:"#888",marginTop:3}}>El pedido se asignará automáticamente a este vendedor cuando inicie sesión</div>
        </Field>
        <Field label="Lista de precios">
          <select value={form.priceList||"default"} onChange={e=>setForm(f=>({...f,priceList:e.target.value}))} style={{...inputStyle,cursor:"pointer"}}>
            {(priceLists||[{id:"default",name:"Normal",discount:0}]).map(pl=>(
              <option key={pl.id} value={pl.id}>{pl.name}{pl.discount>0?` (-${pl.discount}%)`:""}</option>
            ))}
          </select>
          <div style={{fontSize:10,color:"#888",marginTop:3}}>El usuario verá los precios según esta lista</div>
        </Field>
        <div style={{background:"#f9f9f9",borderRadius:8,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",gap:12}}>
          <div>
            <div style={{fontWeight:700,fontSize:13}}>Ver todos los pedidos</div>
            <div style={{fontSize:11,color:"#888",marginTop:2}}>
              {form.canSeeAll ? "Ve todos los pedidos de todos los vendedores" : "Solo ve sus propios pedidos"}
            </div>
          </div>
          <button onClick={()=>setForm(f=>({...f,canSeeAll:!f.canSeeAll}))}
            style={{padding:"6px 14px",borderRadius:20,border:"none",cursor:"pointer",fontWeight:700,fontSize:12,
              background:form.canSeeAll?"#1e8449":"#e5e5e5",color:form.canSeeAll?"#fff":"#888",transition:"all .2s",whiteSpace:"nowrap"}}>
            {form.canSeeAll ? "✅ Habilitado" : "🔒 Solo los suyos"}
          </button>
        </div>
        <div style={{display:"flex",gap:8,marginTop:4}}>
          <button onClick={save} style={{flex:1,padding:"9px",borderRadius:8,border:"none",background:RED,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:13}}>{editing?"Guardar cambios":"Crear usuario"}</button>
          {editing&&<button onClick={cancelEdit} style={{padding:"9px 14px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontWeight:600,color:"#666",fontSize:13}}>Cancelar</button>}
        </div>
      </div>
      <div style={{background:"#fff",borderRadius:12,padding:24,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:15,marginBottom:16}}>🔐 Usuarios ({users.length})</div>
        {users.map(u=>(
          <div key={u.id} style={{padding:"12px 14px",borderRadius:10,border:`1.5px solid ${editing===u.id?"#c0392b":"#f0f0f0"}`,marginBottom:8,background:editing===u.id?"#fdecea":"#fafafa"}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <span style={{fontSize:22}}>{u.role==="admin"?"👑":"👤"}</span>
              <div style={{flex:1}}>
                <div style={{fontWeight:700,fontSize:13}}>{u.name}</div>
                <div style={{fontSize:11,color:"#888",display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
                  <span>@{u.username}</span>
                  <span style={{color:u.role==="admin"?RED:"#1a5276",fontWeight:600}}>{u.role==="admin"?"Admin":"Vendedor"}</span>
                  {u.vendedor&&<span style={{color:"#6c3483",fontWeight:600}}>. 👤 {u.vendedor}</span>}
                  {u.priceList&&u.priceList!=="default"&&<span style={{background:"#fef9e7",color:"#e67e22",borderRadius:6,padding:"1px 6px",fontWeight:700,fontSize:10}}>💲 {(priceLists||[]).find(pl=>pl.id===u.priceList)?.name||u.priceList}</span>}
                  {u.canSeeAll===false&&<span style={{background:"#fdecea",color:"#c0392b",borderRadius:6,padding:"1px 6px",fontWeight:700,fontSize:10}}>🔒 Solo sus pedidos</span>}
                  {u.email&&<span style={{color:"#aaa"}}>. {u.email}</span>}
                </div>
              </div>
              <button onClick={()=>startEdit(u)} style={{padding:"4px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:11}}>✏️️</button>
              <button onClick={()=>remove(u.id)} style={{padding:"4px 10px",borderRadius:6,border:"1.5px solid #fcc",background:"#fff",color:RED,cursor:"pointer",fontSize:11}}>🗑</button>
            </div>
          </div>
        ))}
      </div>
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
