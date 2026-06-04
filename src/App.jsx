/* eslint-disable */
import { createClient } from "@supabase/supabase-js";
import React, { useState, useMemo, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── SUPABASE CONFIG ──────────────────────────────────────────────────────────
const SUPA_URL = "https://pqwcegwadffzqecmbqbe.supabase.co";
const SUPA_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBxd2NlZ3dhZGZmenFlY21icWJlIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODA0OTgzNjgsImV4cCI6MjA5NjA3NDM2OH0.XnmgmzabW4YV4SrP1YNDtRElp7aNoGjbG37XG6VvXak";
const SUPA_SERVICE = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBxd2NlZ3dhZGZmenFlY21icWJlIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDQ5ODM2OCwiZXhwIjoyMDk2MDc0MzY4fQ.dPLccDyKEZ4YPZovD4iXrSJe-DoNf6UFDMNWCmWiK84";
const supabase = createClient(SUPA_URL, SUPA_ANON);
const supaAdmin = supabase;

const mapProduct = r => ({id:r.id,name:r.name,category:r.category,costPrice:r.cost_price,salePrice:r.sale_price,stock:r.stock});
const mapOrder = r => ({id:r.id,client:r.client,vendedor:r.vendedor,notes:r.notes,total:r.total,stage:r.stage,date:r.date,items:r.items||[],docNum:r.doc_num||"",compNum:r.comp_num||""});
const mapQuote = r => ({id:r.id,client:r.client,vendedor:r.vendedor,notes:r.notes,total:r.total,date:r.date,items:r.items||[],validity:r.validity||"",docNum:r.doc_num||""});


// ─── CORRELATIVE NUMBER HELPERS ───────────────────────────────────────────────
// Table lm_counters: { id: "reserva"|"comp"|"presu", value: number }
// SQL: CREATE TABLE lm_counters (id TEXT PRIMARY KEY, value INTEGER DEFAULT 0);
// INSERT INTO lm_counters (id,value) VALUES ('reserva',0),('comp',0),('presu',0);
const padNum = (n) => String(n).padStart(6,"0");
const fmtDocNum = (prefix, n) => `${prefix}-${padNum(n)}`;

const db = {
  getUsers:     async () => { const {data,error} = await supabase.from("lm_users").select("*").order("name"); if(error) throw error; return data||[]; },
  saveUser:     async (u) => { const {error} = await supaAdmin.from("lm_users").upsert(u); if(error) throw error; },
  deleteUser:   async (id) => { const {error} = await supaAdmin.from("lm_users").delete().eq("id",id); if(error) throw error; },

  getVendors:   async () => { const {data,error} = await supabase.from("lm_vendors").select("name").order("name"); if(error) throw error; return (data||[]).map(v=>v.name); },
  addVendor:    async (name) => { const {error} = await supaAdmin.from("lm_vendors").insert({name}); if(error) throw error; },
  deleteVendor: async (name) => { const {error} = await supaAdmin.from("lm_vendors").delete().eq("name",name); if(error) throw error; },
  updateVendor: async (old,nw) => { const {error} = await supaAdmin.from("lm_vendors").update({name:nw}).eq("name",old); if(error) throw error; },

  getProducts:  async () => { const {data,error} = await supabase.from("lm_products").select("*").order("name"); if(error) throw error; return (data||[]).map(mapProduct); },
  upsertProduct: async (p) => { const {error} = await supaAdmin.from("lm_products").upsert({id:p.id,name:p.name,category:p.category||"Importado",cost_price:p.costPrice||0,sale_price:p.salePrice||0,stock:p.stock||0}); if(error) throw error; },
  upsertProducts: async (arr) => { const {error} = await supaAdmin.from("lm_products").upsert(arr.map(p=>({id:p.id,name:p.name,category:p.category||"Importado",cost_price:p.costPrice||0,sale_price:p.salePrice||0,stock:p.stock||0}))); if(error) throw error; },
  deleteProduct: async (id) => { const {error} = await supaAdmin.from("lm_products").delete().eq("id",id); if(error) throw error; },

  getOrders:    async () => { const {data,error} = await supabase.from("lm_orders").select("*").order("date",{ascending:false}); if(error) throw error; return (data||[]).map(mapOrder); },
  upsertOrder:  async (o) => { const {error} = await supaAdmin.from("lm_orders").upsert({id:o.id,client:o.client,vendedor:o.vendedor||"",notes:o.notes||"",total:o.total,stage:o.stage,date:o.date,items:o.items,doc_num:o.docNum||"",comp_num:o.compNum||""}); if(error) throw error; },
  deleteOrder:  async (id) => { const {error} = await supaAdmin.from("lm_orders").delete().eq("id",id); if(error) throw error; },

  getQuotes:    async () => { const {data,error} = await supabase.from("lm_quotes").select("*").order("date",{ascending:false}); if(error) throw error; return (data||[]).map(mapQuote); },
  upsertQuote:  async (q) => { const {error} = await supaAdmin.from("lm_quotes").upsert({id:q.id,client:q.client,vendedor:q.vendedor||"",notes:q.notes||"",total:q.total,date:q.date,items:q.items,validity:q.validity||"",doc_num:q.docNum||""}); if(error) throw error; },
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
    // Atomically increment and return new value
    const {data,error} = await supaAdmin.rpc("increment_counter", {counter_id: id});
    if(error) {
      // Fallback: manual read+write if RPC not available
      const {data:row} = await supaAdmin.from("lm_counters").select("value").eq("id",id).single();
      const next = (row?.value||0) + 1;
      await supaAdmin.from("lm_counters").upsert({id, value:next});
      return next;
    }
    return data;
  },
};

const RED = "#c0392b", REDD = "#922b21";
const STAGES = ["reserva","confirmado","en armado","entregado"];
const SCFG = {
  reserva:     {label:"Reserva",    color:"#c0392b", bg:"#fdecea", icon:"\u{1F550}"},
  confirmado:  {label:"Confirmado", color:"#1a5276", bg:"#d6eaf8", icon:"\u2705"},
  "en armado": {label:"En Armado",  color:"#6c3483", bg:"#e8daef", icon:"\u{1F4E6}"},
  entregado:   {label:"Entregado",  color:"#1e8449", bg:"#d5f5e3", icon:"\u{1F389}"},
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
  return <span style={{background:b,color:c,borderRadius:10,padding:"2px 8px",fontSize:11,fontWeight:700}}>{n===0?"Sin stock":n<=5?`\u26a0 ${n}`:n}</span>;
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

    const itemRows = doc.items.map(it =>
      `<tr>
        <td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;font-size:15px;">${it.name||""}</td>
        <td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;text-align:center;font-size:15px;">${it.qty}</td>
        <td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;text-align:right;font-size:15px;">${fARS(it.price)}</td>
        <td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;text-align:right;font-weight:700;font-size:15px;">${fARS(it.price*it.qty)}</td>
      </tr>`
    ).join("");

    const validityHtml = (tipo==="cotizacion" && doc.validity)
      ? `<div style="background:#fef9e7;border-left:3px solid #f1c40f;padding:8px 14px;border-radius:0 8px 8px 0;font-size:13px;color:#7d6608;margin-bottom:16px;">
          ⏳ <strong>Válida hasta:</strong> ${doc.validity}
        </div>` : "";

    const html = `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"/>
<title>${docLabel} ${docNumDisplay}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0;font-family:'Segoe UI',Arial,sans-serif;}
  body{background:#fff;color:#1a1a1a;}
  @media print{.no-print{display:none!important;}@page{margin:0;size:A4;}}
  .print-btn{display:block;margin:16px auto;padding:10px 32px;background:${badgeColor};color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;}
  .pdf-wrap{padding:0 0 40px;}
  .header-img{width:100%;display:block;border-bottom:3px solid #c0392b;}
  .content{padding:24px 48px 0;}
  .doc-meta{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:22px;padding-bottom:16px;border-bottom:1px solid #f0f0f0;}
  .doc-badge{display:inline-block;padding:7px 20px;border-radius:8px;font-size:14px;font-weight:800;letter-spacing:1px;background:${badgeColor};color:#fff;}
  .doc-num{font-size:26px;font-weight:900;color:#1a1a1a;text-align:right;letter-spacing:-0.5px;}
  .doc-date{font-size:12px;color:#888;text-align:right;margin-top:4px;}
  .info-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px 24px;margin-bottom:24px;}
  .info-box{background:#f9f9f9;border-radius:8px;padding:11px 14px;border-left:3px solid #e5e5e5;}
  .info-box.hl{border-left-color:${badgeColor};}
  .info-label{font-size:9px;color:#999;text-transform:uppercase;letter-spacing:.7px;margin-bottom:4px;font-weight:700;}
  .info-value{font-size:14px;font-weight:700;color:#1a1a1a;}
  table{width:100%;border-collapse:collapse;margin-bottom:20px;}
  thead tr{background:#f5f5f5;}
  th{padding:10px 12px;text-align:left;font-size:11px;color:#888;text-transform:uppercase;letter-spacing:.5px;font-weight:700;}
  th.r{text-align:right;}th.c{text-align:center;}
  tbody tr:nth-child(even){background:#fafafa;}
  .total-wrap{display:flex;justify-content:flex-end;margin-bottom:20px;}
  .total-box{background:${tipo==="cotizacion"?"#e8daef":"#fdecea"};border-radius:10px;padding:16px 24px;display:flex;align-items:center;gap:24px;min-width:300px;}
  .total-label{font-size:14px;color:#555;font-weight:600;}
  .total-amount{font-size:28px;font-weight:900;color:${badgeColor};margin-left:auto;}
  .notes{background:#f9f9f9;border-left:3px solid ${badgeColor};padding:10px 14px;border-radius:0 8px 8px 0;font-size:13px;color:#555;margin-bottom:20px;}
  .footer{border-top:1px solid #f0f0f0;padding-top:14px;margin:0 48px;text-align:center;font-size:11px;color:#bbb;display:flex;justify-content:space-between;}
  .footer-brand{color:#c0392b;font-weight:700;}
</style></head><body>
<button class="no-print print-btn" onclick="window.print()">🖨️ Imprimir / Guardar PDF</button>
<div class="pdf-wrap">
  <img class="header-img" src="${logoSrc}" alt="Libreria Madrid" onerror="this.style.display='none'"/>
  <div class="content">
    <div class="doc-meta">
      <div>
        <div style="font-size:12px;color:#888;margin-bottom:6px;">Comprobante de</div>
        <div class="doc-badge">${docLabel}</div>
      </div>
      <div>
        <div class="doc-num">${docNumDisplay}</div>
        <div class="doc-date">${doc.date}</div>
      </div>
    </div>
    <div class="info-grid">
      <div class="info-box hl"><div class="info-label">Cliente</div><div class="info-value">${doc.client}</div></div>
      <div class="info-box"><div class="info-label">Vendedor</div><div class="info-value">${doc.vendedor||"&mdash;"}</div></div>
      <div class="info-box"><div class="info-label">Fecha</div><div class="info-value">${doc.date}</div></div>
      ${doc.validity?`<div class="info-box"><div class="info-label">V&aacute;lida hasta</div><div class="info-value" style="color:#7d6608">${doc.validity}</div></div>`:""}
    </div>
    ${validityHtml}
    <table>
      <thead><tr>
        <th style="width:50%">Descripci&oacute;n</th>
        <th class="c" style="width:10%">Cant.</th>
        <th class="r" style="width:20%">P. Unit.</th>
        <th class="r" style="width:20%">Subtotal</th>
      </tr></thead>
      <tbody>${itemRows}</tbody>
    </table>
    <div class="total-wrap">
      <div class="total-box">
        <span class="total-label">TOTAL</span>
        <span class="total-amount">${fARS(doc.total)}</span>
      </div>
    </div>
    ${doc.notes?`<div class="notes"><strong>Notas:</strong> ${doc.notes}</div>`:""}
  </div>
  <div class="footer">
    <span><span class="footer-brand">Libreria Madrid</span> &mdash; madrid.libreria &middot; +54 9 11 2502-0640</span>
    <span>Emitido el ${new Date().toLocaleString("es-AR")}</span>
  </div>
</div>
</body></html>`;

    const w = window.open("","_blank","width=820,height=750");
    if(w){ w.document.write(html); w.document.close(); }
  };

  try {
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = img.width; canvas.height = img.height;
        canvas.getContext("2d").drawImage(img, 0, 0);
        doRender(canvas.toDataURL("image/png"));
      } catch(e) { doRender(window.location.origin + "/logo.png"); }
    };
    img.onerror = () => doRender(window.location.origin + "/logo.png");
    img.src = "/logo.png?" + Date.now();
  } catch(e) { doRender("/logo.png"); }
}


// ─── NOTIFICATION TYPES CONFIG ───────────────────────────────────────────────
const NOTIF_TYPES = {
  NUEVO_PEDIDO:   {label:"Nuevo pedido",         icon:"\u{1F6D2}", color:"#1a5276", bg:"#d6eaf8"},
  CAMBIO_ESTADO:  {label:"Cambio de estado",      icon:"\u{1F4CB}", color:"#6c3483", bg:"#e8daef"},
  ALTA_MERCADERIA:{label:"Alta de mercader\u00eda",    icon:"\u{1F4E6}", color:"#1e8449", bg:"#d5f5e3"},
  PEDIDOS_PEND:   {label:"Pedidos pendientes",    icon:"\u23F0", color:"#e67e22", bg:"#fef9e7"},
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
            \u{1F514} Notificaciones {unread.length>0&&<span style={{background:"#f1c40f",color:"#1a1a1a",borderRadius:10,fontSize:11,padding:"1px 6px",marginLeft:6,fontWeight:800}}>{unread.length}</span>}
          </div>
          <div style={{display:"flex",gap:6}}>
            {unread.length>0&&<button onClick={onMarkAllRead} style={{padding:"4px 10px",borderRadius:6,border:"none",background:"#ffffff33",color:"#fff",fontSize:11,fontWeight:600,cursor:"pointer"}}>Marcar todas le\u00eddas</button>}
            <button onClick={onClose} style={{background:"#ffffff22",border:"none",color:"#fff",borderRadius:6,padding:"4px 8px",cursor:"pointer",fontSize:14}}>\u2715</button>
          </div>
        </div>
        {currentUser.role==="vendedor" && pendingOrders.filter(o=>o.vendedor===currentUser.name).length>0&&(
          <div style={{background:"#fef9e7",borderBottom:"1px solid #f1c40f22",padding:"10px 16px",display:"flex",gap:8,alignItems:"center"}}>
            <span style={{fontSize:18}}>\u23F0</span>
            <div>
              <div style={{fontWeight:700,fontSize:12,color:"#7d6608"}}>Ten\u00e9s {pendingOrders.filter(o=>o.vendedor===currentUser.name).length} pedido(s) pendiente(s)</div>
              <div style={{fontSize:11,color:"#9a7d0a"}}>Revis\u00e1 el estado de tus pedidos en Central</div>
            </div>
          </div>
        )}
        <div style={{overflowY:"auto",flex:1}}>
          {myNotifs.length===0
            ? <div style={{textAlign:"center",padding:40,color:"#aaa"}}><div style={{fontSize:36,marginBottom:8}}>\u{1F515}</div><div>No ten\u00e9s notificaciones</div></div>
            : myNotifs.map(n=>{
                const cfg = NOTIF_TYPES[n.tipo]||{icon:"\u2022",color:"#666",bg:"#f5f5f5"};
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
                    <button onClick={e=>{e.stopPropagation();delNotif(n.id);}} style={{background:"none",border:"none",color:"#ddd",cursor:"pointer",fontSize:16,flexShrink:0,padding:0,lineHeight:1}}>\u00d7</button>
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
        <strong>\u{1F514} Centro de notificaciones</strong> &mdash; Configur\u00e1 qu\u00e9 alertas recibe cada usuario dentro de la app.<br/>
        <span style={{fontSize:11,marginTop:4,display:"block",color:"#1a5276bb"}}>Para env\u00edo de emails autom\u00e1ticos conect\u00e1 <strong>EmailJS</strong> cuando la app est\u00e9 en producci\u00f3n.</span>
      </div>
      {users.map(u=>{
        const prefs = getPrefs(u);
        return (
          <div key={u.id} style={{background:"#fff",borderRadius:12,padding:18,marginBottom:12,boxShadow:"0 1px 4px #0001"}}>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
              <span style={{fontSize:24}}>{u.role==="admin"?"\u{1F451}":"\u{1F464}"}</span>
              <div>
                <div style={{fontWeight:800,fontSize:14}}>{u.name}</div>
                <div style={{fontSize:11,color:"#888"}}>@{u.username} \u00b7 <span style={{color:u.role==="admin"?RED:"#1a5276",fontWeight:600}}>{u.role==="admin"?"Admin":"Vendedor"}</span>{u.email&&<span style={{marginLeft:6,color:"#aaa"}}>\u00b7 \u{1F4E7} {u.email}</span>}{!u.email&&<span style={{marginLeft:6,color:"#e67e22",fontSize:10}}>\u26a0 Sin email configurado</span>}</div>
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
                      {active&&<span style={{color:"#fff",fontSize:11,fontWeight:800}}>\u2713</span>}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
      {notifs.length>0&&(<div style={{marginTop:8,textAlign:"right"}}><button onClick={async()=>{setNotifs([]);await db.clearNotifs();}} style={{padding:"7px 14px",borderRadius:8,border:"1.5px solid #fcc",background:"#fff",color:RED,cursor:"pointer",fontSize:12,fontWeight:600}}>\u{1F5D1} Limpiar todas las notificaciones ({notifs.length})</button></div>)}
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
      setError("Usuario o contrase\u00f1a incorrectos");
    } catch(e) {
      const u = users.find(u => u.username === username.trim() && u.password === password);
      if (u) { onLogin(u); return; }
      setError("Usuario o contrase\u00f1a incorrectos");
    } finally { setChecking(false); }
  };
  return (
    <div style={{minHeight:"100vh",background:`linear-gradient(135deg,${REDD},${RED})`,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"#fff",borderRadius:20,padding:40,width:"100%",maxWidth:380,boxShadow:"0 20px 60px #0004"}}>
        <div style={{textAlign:"center",marginBottom:28}}>
          <img src={LOGO} alt="LM" style={{width:90,height:90,borderRadius:"50%",objectFit:"cover",marginBottom:12,boxShadow:"0 4px 16px #0002"}}/>
          <div style={{fontWeight:800,fontSize:22,fontFamily:"Georgia,serif",color:"#1a1a1a"}}>Librer\u00eda LM</div>
          <div style={{fontSize:12,color:"#aaa",letterSpacing:2,textTransform:"uppercase",marginTop:2}}>Sistema de Gesti\u00f3n</div>
        </div>
        <Field label="Usuario"><input value={username} onChange={e=>setUsername(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleLogin()} placeholder="Tu usuario" style={inputStyle}/></Field>
        <Field label="Contrase\u00f1a">
          <div style={{position:"relative"}}>
            <input type={showPass?"text":"password"} value={password} onChange={e=>setPassword(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleLogin()} placeholder="Tu contrase\u00f1a" style={{...inputStyle,paddingRight:40}}/>
            <button onClick={()=>setShowPass(s=>!s)} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:16,color:"#aaa"}}>{showPass?"\u{1F648}":"\u{1F441}"}</button>
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
  const [currentUser, setCurrentUser] = useState(null);
  const [users, setUsers]       = useState([]);
  const [vendors, setVendors]   = useState([]);
  const [products, setProducts] = useState([]);
  const [orders, setOrders]     = useState([]);
  const [quotes, setQuotes]     = useState([]);
  const [stockLog, setStockLog] = useState([]);
  const [notifs, setNotifs]     = useState([]);
  const [loading, setLoading]   = useState(true);
  const [error, setError]       = useState(null);

  useEffect(() => {
    async function loadAll() {
      try {
        const [u,v,p,o,q,sl,n] = await Promise.all([
          db.getUsers(), db.getVendors(), db.getProducts(),
          db.getOrders(), db.getQuotes(), db.getStockLog(), db.getNotifs(),
        ]);
        setUsers(u); setVendors(v); setProducts(p);
        setOrders(o); setQuotes(q); setStockLog(sl); setNotifs(n);
      } catch(e) {
        setError("No se pudo conectar con la base de datos. Verific\u00e1 tu conexi\u00f3n.");
      } finally { setLoading(false); }
    }
    loadAll();
  }, []);

  if(loading) return (
    <div style={{minHeight:"100vh",background:`linear-gradient(135deg,#922b21,#c0392b)`,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:16}}>
      <img src="/logo.png" alt="LM" style={{width:80,height:80,borderRadius:"50%",objectFit:"cover"}}/>
      <div style={{color:"#fff",fontWeight:700,fontSize:16}}>Cargando Librer\u00eda LM...</div>
      <div style={{color:"#ffcccc",fontSize:13}}>Conectando con la base de datos</div>
    </div>
  );
  if(error) return (
    <div style={{minHeight:"100vh",background:"#f5f5f5",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:"#fff",borderRadius:16,padding:32,maxWidth:400,textAlign:"center",boxShadow:"0 4px 20px #0002"}}>
        <div style={{fontSize:48,marginBottom:12}}>\u26a0\ufe0f</div>
        <div style={{fontWeight:800,fontSize:16,marginBottom:8}}>Error de conexi\u00f3n</div>
        <div style={{color:"#666",fontSize:13,marginBottom:20}}>{error}</div>
        <button onClick={()=>window.location.reload()} style={{padding:"10px 24px",borderRadius:8,border:"none",background:"#c0392b",color:"#fff",fontWeight:700,cursor:"pointer"}}>Reintentar</button>
      </div>
    </div>
  );
  if(!currentUser) return <Login users={users} onLogin={u=>setCurrentUser(u)}/>;
  return <MainApp
    currentUser={currentUser} onLogout={()=>setCurrentUser(null)}
    users={users} setUsers={setUsers}
    vendors={vendors} setVendors={setVendors}
    products={products} setProducts={setProducts}
    orders={orders} setOrders={setOrders}
    quotes={quotes} setQuotes={setQuotes}
    stockLog={stockLog} setStockLog={setStockLog}
    notifs={notifs} setNotifs={setNotifs}
  />;
}

// ─── MAIN APP (authenticated) ─────────────────────────────────────────────────
function MainApp({currentUser,onLogout,users,setUsers,vendors,setVendors,products,setProducts,orders,setOrders,quotes,setQuotes,stockLog,setStockLog,notifs,setNotifs}) {
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
  const addLog = async (entry) => {
    const full = {id:genId(),fecha:new Date().toLocaleString("es-AR"),usuario:currentUser.name,rol:currentUser.role,...entry};
    setStockLog(l=>[full,...l]); await db.addStockLog(full);
  };
  const addOrder = async (order) => {
    // Assign Reserva-XXXXXX correlative number
    const n = await db.nextCounter("reserva");
    const orderWithNum = {...order, docNum: fmtDocNum("Reserva", n)};

    const updatedProds = products.map(x=>{const it=orderWithNum.items.find(i=>i.pid===x.id);return it?{...x,stock:Math.max(0,x.stock-it.qty)}:x;});
    setProducts(updatedProds); setOrders(o=>[orderWithNum,...o]);
    await db.upsertOrder(orderWithNum);
    for(const p of updatedProds.filter(p=>orderWithNum.items.find(i=>i.pid===p.id))) await db.upsertProduct(p);
    const notif={id:genId(),fecha:new Date().toLocaleString("es-AR"),leida:[],tipo:"NUEVO_PEDIDO",para:"admin",icono:"\u{1F6D2}",titulo:"Nuevo pedido registrado",cuerpo:`${orderWithNum.client} \u2014 ${fARS(orderWithNum.total)} \u2014 ${orderWithNum.docNum}`,ref:orderWithNum.id};
    await db.addNotif(notif); setNotifs(n=>[notif,...n]);
    // Auto-print Reserva document
    setTimeout(() => printDoc(orderWithNum, "reserva"), 400);
  };
  const [compPopup, setCompPopup] = useState(null); // {order, compNum}

  const setStage = async (id,stage) => {
    const ord = orders.find(o=>o.id===id);
    let updated = {...ord, stage};

    // When moving to "confirmado": assign Comp-XXXXXX
    if(stage === "confirmado" && !ord.compNum) {
      const n = await db.nextCounter("comp");
      updated = {...updated, compNum: fmtDocNum("Comp", n)};
    }

    setOrders(o=>o.map(x=>x.id===id?updated:x));
    await db.upsertOrder(updated);

    // Show popup with Comp number when confirmed
    if(stage === "confirmado" && updated.compNum) {
      setCompPopup(updated);
    }

    if(ord){
      const cfg=SCFG[stage]||{};
      const n1={id:genId(),fecha:new Date().toLocaleString("es-AR"),leida:[],tipo:"CAMBIO_ESTADO",para:"admin",icono:cfg.icon||"\u{1F4CB}",titulo:`Pedido paso a ${cfg.label}`,cuerpo:`${ord.client} \u2014 ${fARS(ord.total)} \u2014 ${updated.compNum||""}`,ref:id};
      await db.addNotif(n1); setNotifs(n=>[n1,...n]);
      const vendUser=users.find(u=>u.name===ord.vendedor||u.username===ord.vendedor);
      if(vendUser&&vendUser.id!==currentUser.id){
        const n2={id:genId(),fecha:new Date().toLocaleString("es-AR"),leida:[],tipo:"CAMBIO_ESTADO",para:vendUser.id,icono:cfg.icon||"\u{1F4CB}",titulo:`Tu pedido paso a ${cfg.label}`,cuerpo:`${ord.client} \u2014 ${fARS(ord.total)}`,ref:id};
        await db.addNotif(n2); setNotifs(n=>[n2,...n]);
      }
    }
  };
  const delOrder = async (id) => {
    const ord=orders.find(o=>o.id===id);
    if(ord&&ord.stage!=="entregado"){
      const updatedProds=products.map(x=>{const it=ord.items.find(i=>i.pid===x.id);return it?{...x,stock:x.stock+it.qty}:x;});
      setProducts(updatedProds);
      for(const p of updatedProds.filter(p=>ord.items.find(i=>i.pid===p.id))) await db.upsertProduct(p);
    }
    setOrders(o=>o.filter(x=>x.id!==id)); await db.deleteOrder(id);
  };
  const addQuote = async (quote) => {
    const n = await db.nextCounter("presu");
    const quoteWithNum = {...quote, docNum: fmtDocNum("Presu", n)};
    setQuotes(q=>[quoteWithNum,...q]);
    await db.upsertQuote(quoteWithNum);
  };
  const delQuote = async (id) => { setQuotes(q=>q.filter(x=>x.id!==id)); await db.deleteQuote(id); };
  const updProd = async (upd) => { setProducts(p=>p.map(x=>x.id===upd.id?upd:x)); await db.upsertProduct(upd); };
  const addStock = async (pid,qty,newCost) => {
    const prod=products.find(p=>p.id===pid);
    const updatedProds=products.map(x=>{if(x.id!==pid)return x;const u={...x,stock:x.stock+qty};if(newCost){u.costPrice=newCost;u.salePrice=Math.round(newCost*1.5*100)/100;}return u;});
    setProducts(updatedProds);
    const updProd=updatedProds.find(p=>p.id===pid);
    if(updProd)await db.upsertProduct(updProd);
    if(prod){
      const notif={id:genId(),fecha:new Date().toLocaleString("es-AR"),leida:[],tipo:"ALTA_MERCADERIA",para:"admin",icono:"\u{1F4E6}",titulo:"Alta de mercaderia",cuerpo:`${prod.name} \u2014 +${qty} unidades${newCost?` \u2014 Nuevo costo: ${fARS(newCost)}`:""}`,ref:pid};
      await db.addNotif(notif); setNotifs(n=>[notif,...n]);
    }
  };

  const pending = orders.filter(o=>o.stage!=="entregado").length;
  const TABS = [
    {k:"central",   label:"Central",           icon:"\u{1F4CB}", roles:["admin","vendedor"]},
    {k:"nuevo",     label:"Nuevo Pedido",       icon:"\u{1F6D2}", roles:["admin","vendedor"]},
    {k:"cotizacion",label:"Cotizaciones",       icon:"\u{1F4C4}", roles:["admin","vendedor"]},
    {k:"precios",   label:"Precios",            icon:"\u{1F4B2}", roles:["admin","vendedor"]},
    {k:"stock",     label:"Stock",              icon:"\u{1F4E6}", roles:["admin","vendedor"]},
    {k:"compras",   label:"Alta de Mercanc\u00eda",icon:"\u{1F3EA}", roles:["admin","vendedor"]},
    {k:"admin",     label:"Administracion",     icon:"\u2605",   roles:["admin"]},
  ].filter(t=>t.roles.includes(currentUser.role));

  return (
    <div style={{minHeight:"100vh",background:"#f5f5f5",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
      <div style={{background:`linear-gradient(135deg,${REDD},${RED})`,boxShadow:"0 4px 16px #0004"}}>
        <div style={{maxWidth:1200,margin:"0 auto",padding:"0 16px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap"}}>
          <div style={{display:"flex",alignItems:"center",gap:12,padding:"12px 0"}}>
            <img src={LOGO} alt="LM Logo" style={{width:54,height:54,borderRadius:"50%",objectFit:"cover",boxShadow:"0 2px 8px #0003"}}/>
            <div>
              <div style={{color:"#fff",fontWeight:800,fontSize:20,fontFamily:"Georgia,serif"}}>Librer\u00eda LM</div>
              <div style={{color:"#ffcccc",fontSize:10,letterSpacing:2,textTransform:"uppercase"}}>Sistema de Gesti\u00f3n</div>
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
              <span style={{color:"#ffeeee",fontSize:12}}>\u{1F464} {currentUser.name}</span>
              <div style={{position:"relative"}}>
                <button onClick={()=>setShowNotifs(s=>!s)} style={{background:"#ffffff22",border:"none",color:"#fff",borderRadius:6,padding:"5px 8px",cursor:"pointer",fontSize:16,lineHeight:1,position:"relative"}}>
                  \u{1F514}
                  {unreadCount>0&&<span style={{position:"absolute",top:-4,right:-4,background:"#f1c40f",color:"#1a1a1a",borderRadius:10,fontSize:9,padding:"1px 4px",fontWeight:800,minWidth:14,textAlign:"center"}}>{unreadCount}</span>}
                </button>
              </div>
              <button onClick={onLogout} style={{background:"#ffffff22",border:"none",color:"#fff",borderRadius:6,padding:"5px 8px",cursor:"pointer",fontSize:11,fontWeight:600}}>Salir</button>
            </div>
            {showNotifs&&<NotifPanel notifs={notifs} setNotifs={setNotifs} currentUser={currentUser} users={users} onClose={()=>setShowNotifs(false)} onMarkAllRead={markAllRead} pushNotif={pushNotif} orders={orders}/>}
          </div>
        </div>
      </div>
      {compPopup && <CompPopup order={compPopup} onClose={()=>setCompPopup(null)}/>}
      <div style={{maxWidth:1200,margin:"0 auto",padding:"20px 16px"}}>
        {tab==="central"    && <Central orders={orders} products={products} onStage={setStage} onDel={delOrder}/>}
        {tab==="nuevo"      && <Nuevo products={products} vendors={vendors} onAdd={addOrder} onDone={()=>setTab("central")}/>}
        {tab==="cotizacion" && <Cotizaciones quotes={quotes} products={products} vendors={vendors} onAdd={addQuote} onDel={delQuote}/>}
        {tab==="precios"    && <Precios products={products}/>}
        {tab==="stock"      && <Stock products={products} onUpd={updProd} onDel={pid=>setProducts(p=>p.filter(x=>x.id!==pid))} onAdjust={(pid,qty)=>setProducts(p=>p.map(x=>x.id===pid?{...x,stock:x.stock+qty}:x))} isAdmin={isAdmin} addLog={addLog} stockLog={stockLog} setStockLog={setStockLog}/>}
        {tab==="compras"    && <Compras products={products} onStock={addStock}/>}
        {tab==="admin"      && isAdmin && <AdminPanel users={users} setUsers={setUsers} vendors={vendors} setVendors={setVendors} products={products} setProducts={setProducts} stockLog={stockLog} setStockLog={setStockLog} notifs={notifs} setNotifs={setNotifs}/>}
      </div>
    </div>
  );
}


// ─── COMP POPUP ───────────────────────────────────────────────────────────────
// Shown when a pedido is confirmed — shows Comp number and print button
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
          <div style={{fontSize:12,color:"#888",marginTop:4}}>{order.client} · {fARS(order.total)}</div>
        </div>

        <div style={{display:"flex",gap:10}}>
          <button onClick={()=>{ printDoc(order,"confirmado"); onClose(); }}
            style={{flex:1,padding:"12px",borderRadius:10,border:"none",background:"linear-gradient(135deg,#1a5276,#2980b9)",color:"#fff",fontWeight:800,fontSize:14,cursor:"pointer"}}>
            🖨️ Imprimir Comp.
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
function Central({orders,products,onStage,onDel}) {
  const [fStage,setFStage]=useState("todos");
  const [search,setSearch]=useState("");
  const [expanded,setExpanded]=useState(null);
  const getP = id=>products.find(p=>p.id===id);
  const filtered = orders.filter(o=>{
    if(fStage!=="todos"&&o.stage!==fStage) return false;
    if(search&&!o.client.toLowerCase().includes(search.toLowerCase())&&!o.id.includes(search)) return false;
    return true;
  });
  const deliv = orders.filter(o=>o.stage==="entregado").reduce((s,o)=>s+o.total,0);
  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(135px,1fr))",gap:12,marginBottom:20}}>
        {STAGES.map(s=>{const c=SCFG[s],cnt=orders.filter(o=>o.stage===s).length;return <div key={s} onClick={()=>setFStage(fStage===s?"todos":s)} style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 6px #0001",borderLeft:`4px solid ${c.color}`,cursor:"pointer",outline:fStage===s?`2px solid ${c.color}`:"none"}}><div style={{fontSize:26,fontWeight:800,color:c.color}}>{cnt}</div><div style={{fontSize:12,color:"#666",fontWeight:600}}>{c.icon} {c.label}</div></div>;})}
        <div style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 6px #0001",borderLeft:`4px solid ${RED}`}}>
          <div style={{fontSize:14,fontWeight:800,color:RED}}>{fARS(deliv)}</div>
          <div style={{fontSize:12,color:"#666",fontWeight:600}}>\u{1F4B0} Entregado</div>
        </div>
      </div>
      <div style={{background:"#fff",borderRadius:12,padding:14,marginBottom:14,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center",boxShadow:"0 1px 4px #0001"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="\u{1F50D} Buscar cliente o N\u00b0 pedido..." style={{flex:1,minWidth:180,padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none"}}/>
        <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
          {["todos",...STAGES].map(s=>{const c=SCFG[s];return <button key={s} onClick={()=>setFStage(s)} style={{padding:"5px 11px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:11,fontWeight:600,borderColor:fStage===s?(c?.color||RED):"#e5e5e5",background:fStage===s?(c?.bg||"#fdecea"):"#fff",color:fStage===s?(c?.color||RED):"#666"}}>{s==="todos"?"Todos":c.label}</button>;})}
        </div>
      </div>
      {filtered.length===0
        ? <div style={{textAlign:"center",padding:60,color:"#aaa"}}><div style={{fontSize:48}}>\u{1F4ED}</div><div style={{marginTop:8}}>No hay pedidos. \u00a1Cre\u00e1 uno desde "Nuevo Pedido"!</div></div>
        : filtered.map(o=><OCard key={o.id} o={o} exp={expanded===o.id} toggle={()=>setExpanded(expanded===o.id?null:o.id)} getP={getP} onStage={onStage} onDel={onDel}/>)
      }
    </div>
  );
}

function DelBtn({onConfirm}) {
  const [confirm, setConfirm] = useState(false);
  if(confirm) return (
    <div style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:6,background:"#fdecea",borderRadius:8,padding:"6px 10px",border:"1.5px solid #fcc"}}>
      <span style={{fontSize:12,color:RED,fontWeight:600}}>\u00bfEliminar?</span>
      <button onClick={onConfirm} style={{padding:"4px 10px",borderRadius:6,border:"none",background:RED,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:12}}>S\u00ed</button>
      <button onClick={()=>setConfirm(false)} style={{padding:"4px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12}}>No</button>
    </div>
  );
  return <button onClick={()=>setConfirm(true)} style={{marginLeft:"auto",padding:"8px 12px",borderRadius:8,border:"1.5px solid #fcc",cursor:"pointer",background:"#fff",color:RED,fontWeight:600,fontSize:13}}>\u{1F5D1} Eliminar</button>;
}

function OCard({o,exp,toggle,getP,onStage,onDel}) {
  const idx=STAGES.indexOf(o.stage), next=STAGES[idx+1];
  return (
    <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 6px #0001",overflow:"hidden",marginBottom:8}}>
      <div onClick={toggle} style={{padding:"13px 18px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",cursor:"pointer"}}>
        <div style={{flex:1,minWidth:100}}>
          <div style={{fontWeight:700,fontSize:14,color:"#1a1a1a"}}>{o.client}</div>
          <div style={{fontSize:11,color:"#aaa",display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
            {o.docNum&&<span style={{fontWeight:700,color:"#c0392b"}}>{o.docNum}</span>}
            {o.compNum&&<span style={{fontWeight:700,color:"#1a5276"}}>{o.compNum}</span>}
            <span>{o.date}</span>
            {o.vendedor&&<span>\u00b7 \u{1F464} {o.vendedor}</span>}
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
          <Bdg stage={o.stage}/>
          <span style={{fontWeight:800,color:RED,fontSize:15}}>{fARS(o.total)}</span>
          <span style={{color:"#ccc"}}>{exp?"\u25b2":"\u25bc"}</span>
        </div>
      </div>
      {exp && (
        <div style={{borderTop:"1px solid #f5f5f5",padding:18}}>
          <div style={{display:"flex",marginBottom:18,overflowX:"auto"}}>
            {STAGES.map((s,i)=>{const done=i<=idx,c=SCFG[s];return <div key={s} style={{display:"flex",alignItems:"center",flex:1,minWidth:65}}><div style={{textAlign:"center",flex:1}}><div style={{width:30,height:30,borderRadius:"50%",background:done?c.color:"#eee",color:done?"#fff":"#aaa",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 4px",fontSize:13}}>{done?c.icon:"\u25cb"}</div><div style={{fontSize:10,color:done?c.color:"#aaa",fontWeight:done?700:400}}>{c.label}</div></div>{i<3&&<div style={{height:2,background:i<idx?RED:"#eee",flex:1}}/>}</div>;})}
          </div>
          {o.items.map((it,i)=>{const p=getP(it.pid);return <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #f9f9f9",fontSize:13}}><span style={{color:"#444"}}>{p?.name||it.name} \u00d7 {it.qty}</span><span style={{fontWeight:600}}>{fARS(it.price*it.qty)}</span></div>;})}
          <div style={{display:"flex",justifyContent:"flex-end",fontWeight:800,fontSize:16,color:RED,margin:"8px 0 12px"}}>{fARS(o.total)}</div>
          {o.notes&&<div style={{background:"#f9f9f9",borderRadius:8,padding:"8px 12px",fontSize:13,color:"#555",marginBottom:12}}>\u{1F4AC} {o.notes}</div>}
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {next&&<button onClick={()=>onStage(o.id,next)} style={{padding:"8px 14px",borderRadius:8,border:"none",cursor:"pointer",background:SCFG[next].color,color:"#fff",fontWeight:700,fontSize:13}}>{SCFG[next].icon} Pasar a {SCFG[next].label}</button>}
            {idx>0&&o.stage!=="entregado"&&<button onClick={()=>onStage(o.id,STAGES[idx-1])} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",cursor:"pointer",background:"#fff",color:"#666",fontWeight:600,fontSize:13}}>\u2190 Retroceder</button>}
            <button onClick={()=>printDoc(o, o.stage==="reserva" ? "reserva" : "confirmado")} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #d6eaf8",cursor:"pointer",background:"#fff",color:"#1a5276",fontWeight:600,fontSize:13}}>
              \u{1F5A8} {o.stage==="reserva" ? (o.docNum||"Imprimir") : (o.compNum||"Imprimir")}
            </button>
            <DelBtn onConfirm={()=>onDel(o.id)}/>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── BUSCADOR DE PRECIOS ──────────────────────────────────────────────────────
function Precios({products}) {
  const [search, setSearch] = useState("");
  const [cat, setCat]       = useState("todos");
  const [catOpen, setCatOpen] = useState(false);
  const [sortBy, setSortBy] = useState("nombre");
  const CATS = useMemo(()=>["todos",...new Set(products.map(p=>p.category))].sort(),[products]);
  const shown = useMemo(()=>{
    const q = search.toLowerCase();
    let list = products.filter(p=>{
      if(cat!=="todos" && p.category!==cat) return false;
      if(q) return p.name.toLowerCase().includes(q) || p.id.includes(q);
      return true;
    });
    if(sortBy==="precio_asc")  list=[...list].sort((a,b)=>a.salePrice-b.salePrice);
    if(sortBy==="precio_desc") list=[...list].sort((a,b)=>b.salePrice-a.salePrice);
    if(sortBy==="nombre")      list=[...list].sort((a,b)=>a.name.localeCompare(b.name));
    return list.slice(0,200);
  },[products,search,cat,sortBy]);
  return (
    <div>
      <div style={{background:"#fff",borderRadius:12,padding:16,marginBottom:14,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:16,marginBottom:12,color:"#1a1a1a"}}>\u{1F3F7}\ufe0f Buscador de Precios</div>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="\u{1F50D} Busc\u00e1 por nombre o c\u00f3digo..." style={{width:"100%",padding:"10px 14px",borderRadius:10,border:"1.5px solid #e5e5e5",fontSize:14,outline:"none",boxSizing:"border-box",marginBottom:10}} autoFocus/>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
          <div style={{position:"relative",flex:1,minWidth:200}}>
            <button onClick={()=>setCatOpen(o=>!o)} style={{display:"flex",alignItems:"center",justifyContent:"space-between",width:"100%",padding:"8px 12px",borderRadius:8,border:`1.5px solid ${catOpen?RED:"#e5e5e5"}`,background:cat!=="todos"?"#fdecea":"#fff",color:cat!=="todos"?RED:"#666",cursor:"pointer",fontSize:13,fontWeight:600}}>
              <span>\u{1F3F7}\ufe0f {cat==="todos"?"Todas las categor\u00edas":cat}</span><span style={{fontSize:10}}>{catOpen?"\u25b2":"\u25bc"}</span>
            </button>
            {catOpen&&(<div style={{position:"absolute",top:"calc(100% + 4px)",left:0,right:0,background:"#fff",borderRadius:10,border:"1.5px solid #e5e5e5",boxShadow:"0 8px 24px #0002",zIndex:50,padding:8,display:"flex",flexWrap:"wrap",gap:5,maxHeight:220,overflowY:"auto"}}>
              {CATS.map(c=>(<button key={c} onClick={()=>{setCat(c);setCatOpen(false);}} style={{padding:"4px 11px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:11,fontWeight:600,borderColor:cat===c?RED:"#e5e5e5",background:cat===c?"#fdecea":"#fff",color:cat===c?RED:"#666"}}>{c==="todos"?"Todos":c}</button>))}
            </div>)}
          </div>
          <div style={{display:"flex",gap:5}}>
            {[{v:"nombre",l:"A-Z"},{v:"precio_asc",l:"Precio \u2191"},{v:"precio_desc",l:"Precio \u2193"}].map(opt=>(
              <button key={opt.v} onClick={()=>setSortBy(opt.v)} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid",cursor:"pointer",fontSize:12,fontWeight:600,borderColor:sortBy===opt.v?RED:"#e5e5e5",background:sortBy===opt.v?"#fdecea":"#fff",color:sortBy===opt.v?RED:"#666"}}>{opt.l}</button>
            ))}
          </div>
          <div style={{fontSize:12,color:"#aaa",whiteSpace:"nowrap"}}>{shown.length} producto{shown.length!==1?"s":""}{shown.length===200?" (m\u00e1x 200)":""}</div>
        </div>
      </div>
      {shown.length===0
        ? <div style={{textAlign:"center",padding:60,color:"#aaa",background:"#fff",borderRadius:12}}><div style={{fontSize:48,marginBottom:8}}>\u{1F50D}</div><div>No se encontraron productos</div></div>
        : <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px #0001",overflow:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr style={{background:"#f9f9f9",position:"sticky",top:0}}>{["C\u00f3digo","Descripci\u00f3n","Categor\u00eda","Precio"].map(h=>(<th key={h} style={{padding:"11px 14px",textAlign:"left",fontWeight:700,color:"#888",fontSize:11,textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>))}</tr></thead>
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
function ProductSelector({products,cart,setCart}) {
  const [search,setSearch]=useState("");
  const [cat,setCat]=useState("todos");
  const [catOpen,setCatOpen]=useState(false);
  const CATS=useMemo(()=>["todos",...new Set(products.map(p=>p.category))].sort(),[products]);
  const shown=useMemo(()=>{
    const q=search.toLowerCase();
    return products.filter(p=>{if(cat!=="todos"&&p.category!==cat)return false;if(q)return p.name.toLowerCase().includes(q)||p.id.includes(q);return true;}).slice(0,80);
  },[products,search,cat]);
  const addC=p=>setCart(c=>{const ex=c.find(i=>i.pid===p.id);return ex?c.map(i=>i.pid===p.id?{...i,qty:i.qty+1}:i):[...c,{pid:p.id,qty:1,price:p.salePrice,name:p.name}];});
  const setQ=(pid,qty)=>{if(qty<=0)setCart(c=>c.filter(i=>i.pid!==pid));else setCart(c=>c.map(i=>i.pid===pid?{...i,qty}:i));};
  return (
    <div>
      <div style={{background:"#fff",borderRadius:12,padding:16,marginBottom:12,boxShadow:"0 1px 4px #0001"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="\u{1F50D} Buscar por nombre o c\u00f3digo..." style={{width:"100%",padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",marginBottom:10,boxSizing:"border-box"}}/>
        <div style={{position:"relative"}}>
          <button onClick={()=>setCatOpen(o=>!o)} style={{display:"flex",alignItems:"center",justifyContent:"space-between",width:"100%",padding:"8px 12px",borderRadius:8,border:`1.5px solid ${catOpen?RED:"#e5e5e5"}`,background:cat!=="todos"?"#fdecea":"#fff",color:cat!=="todos"?RED:"#666",cursor:"pointer",fontSize:13,fontWeight:600}}>
            <span>\u{1F3F7}\ufe0f {cat==="todos"?"Todas las categor\u00edas":cat}</span><span style={{fontSize:10,marginLeft:6}}>{catOpen?"\u25b2":"\u25bc"}</span>
          </button>
          {catOpen&&(<div style={{position:"absolute",top:"calc(100% + 4px)",left:0,right:0,background:"#fff",borderRadius:10,border:"1.5px solid #e5e5e5",boxShadow:"0 8px 24px #0002",zIndex:50,padding:8,display:"flex",flexWrap:"wrap",gap:5,maxHeight:220,overflowY:"auto"}}>
            {CATS.map(c=><button key={c} onClick={()=>{setCat(c);setCatOpen(false);}} style={{padding:"4px 11px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:11,fontWeight:600,borderColor:cat===c?RED:"#e5e5e5",background:cat===c?"#fdecea":"#fff",color:cat===c?RED:"#666"}}>{c==="todos"?"Todos":c}</button>)}
          </div>)}
        </div>
        {search&&<div style={{fontSize:11,color:"#aaa",marginTop:6}}>{shown.length} resultados</div>}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(195px,1fr))",gap:10}}>
        {shown.map(p=>{const ic=cart.find(i=>i.pid===p.id);return <div key={p.id} style={{background:"#fff",borderRadius:10,padding:14,border:ic?`2px solid ${RED}`:"2px solid transparent",boxShadow:"0 1px 4px #0001"}}>
          <div style={{fontWeight:700,fontSize:12,color:"#1a1a1a",marginBottom:3,lineHeight:1.3}}>{p.name}</div>
          <div style={{fontSize:12,color:"#666",marginBottom:7,fontWeight:500}}>{p.id} \u00b7 {p.category}</div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <span style={{fontSize:17,fontWeight:800,color:RED}}>{fARS(p.salePrice)}</span><SPill n={p.stock}/>
          </div>
          {ic?<div style={{display:"flex",alignItems:"center",gap:5}}>
            <button onClick={()=>setQ(p.id,ic.qty-1)} style={{width:27,height:27,borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:15,fontWeight:700}}>\u2212</button>
            <input type="number" value={ic.qty} onChange={e=>setQ(p.id,+e.target.value||0)} style={{width:40,textAlign:"center",padding:3,borderRadius:6,border:`1.5px solid ${RED}`,fontWeight:700,fontSize:13,outline:"none"}}/>
            <button onClick={()=>setQ(p.id,ic.qty+1)} style={{width:27,height:27,borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:15,fontWeight:700}}>+</button>
            <span style={{color:"#1e8449",fontSize:12,fontWeight:700}}>\u2713</span>
          </div>:<button onClick={()=>addC(p)} style={{width:"100%",padding:"7px",borderRadius:7,border:"none",cursor:"pointer",background:RED,color:"#fff",fontWeight:700,fontSize:12}}>+ Agregar</button>}
        </div>;})}
      </div>
    </div>
  );
}

function Nuevo({products,vendors,onAdd,onDone}) {
  const [client,setClient]=useState("");
  const [notes,setNotes]=useState("");
  const [vendedor,setVendedor]=useState("");
  const [cart,setCart]=useState([]);
  const [ok,setOk]=useState(false);
  const total=cart.reduce((s,i)=>s+i.price*i.qty,0);
  const submit=()=>{
    if(!client.trim()){alert("Ingres\u00e1 el cliente");return;}
    if(!vendedor){alert("Selecci\u00f3n\u00e1 un vendedor");return;}
    if(!cart.length){alert("Agreg\u00e1 productos");return;}
    onAdd({id:genId(),client:client.trim(),notes,vendedor,items:cart,total,stage:"reserva",date:today()});
    setOk(true); setTimeout(()=>onDone(),1400);
  };
  if(ok) return <div style={{textAlign:"center",padding:80}}><div style={{fontSize:60}}>\u2705</div><div style={{fontWeight:800,color:"#1e8449",fontSize:20,marginTop:12}}>\u00a1Pedido registrado!</div></div>;
  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 330px",gap:18,alignItems:"start"}}>
      <div>
        <div style={{fontWeight:800,fontSize:15,marginBottom:12}}>\u{1F6D2} Nuevo Pedido \u2014 Selecci\u00f3n\u00e1 productos</div>
        <ProductSelector products={products} cart={cart} setCart={setCart}/>
      </div>
      <div style={{position:"sticky",top:16}}>
        <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 2px 12px #0002"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14}}>\u{1F4CB} Resumen del Pedido</div>
          <Field label="Cliente *"><input value={client} onChange={e=>setClient(e.target.value)} placeholder="Nombre del cliente" style={inputStyle}/></Field>
          <Field label="Vendedor *"><select value={vendedor} onChange={e=>setVendedor(e.target.value)} style={{...inputStyle,color:vendedor?"#1a1a1a":"#aaa",cursor:"pointer"}}><option value="">\u2014 Selecci\u00f3n\u00e1 vendedor \u2014</option>{vendors.map(v=><option key={v} value={v}>{v}</option>)}</select></Field>
          <Field label="Notas"><textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Observaciones..." style={{...inputStyle,resize:"vertical",minHeight:55,fontSize:12}}/></Field>
          <div style={{borderTop:"1px solid #f5f5f5",margin:"4px 0 10px",paddingTop:10}}>
            {cart.length===0?<div style={{textAlign:"center",color:"#aaa",fontSize:12,padding:"10px 0"}}>Agreg\u00e1 productos al pedido</div>
            :cart.map(i=><div key={i.pid} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"3px 0",color:"#555"}}><span style={{flex:1,marginRight:6,lineHeight:1.3}}>{i.name} \u00d7 {i.qty}</span><span style={{fontWeight:600,whiteSpace:"nowrap"}}>{fARS(i.price*i.qty)}</span></div>)}
          </div>
          <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:17,color:RED,padding:"8px 0",borderTop:"2px solid #f5f5f5",marginBottom:14}}><span>Total</span><span>{fARS(total)}</span></div>
          <button onClick={submit} disabled={!cart.length||!client.trim()||!vendedor} style={{width:"100%",padding:"11px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:(!cart.length||!client.trim()||!vendedor)?"#e5e5e5":`linear-gradient(135deg,${REDD},${RED})`,color:(!cart.length||!client.trim()||!vendedor)?"#aaa":"#fff"}}>
            \u2705 Registrar como Reserva
          </button>
        </div>
      </div>
    </div>
  );
}


// ─── COTIZACIONES ─────────────────────────────────────────────────────────────
// IMPORTANTE: Crear tabla en Supabase antes de usar:
// CREATE TABLE lm_quotes (
//   id TEXT PRIMARY KEY, client TEXT, vendedor TEXT, notes TEXT,
//   total NUMERIC, date TEXT, items JSONB, validity TEXT
// );
function Cotizaciones({quotes,products,vendors,onAdd,onDel}) {
  const [view,setView]=useState("lista"); // "lista" | "nueva"
  const [expanded,setExpanded]=useState(null);
  const getP=id=>products.find(p=>p.id===id);
  return (
    <div>
      <div style={{background:"#fff",borderRadius:12,padding:4,marginBottom:16,display:"flex",gap:4,boxShadow:"0 1px 4px #0001"}}>
        <button onClick={()=>setView("lista")} style={{flex:1,padding:"10px 16px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="lista"?`linear-gradient(135deg,#6c3483,#9b59b6)`:"transparent",color:view==="lista"?"#fff":"#555"}}>
          \u{1F4C4} Lista de Cotizaciones ({quotes.length})
        </button>
        <button onClick={()=>setView("nueva")} style={{flex:1,padding:"10px 16px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="nueva"?`linear-gradient(135deg,#6c3483,#9b59b6)`:"transparent",color:view==="nueva"?"#fff":"#555"}}>
          \u2795 Nueva Cotizaci\u00f3n
        </button>
      </div>
      {view==="nueva" && <NuevaCotizacion products={products} vendors={vendors} onAdd={async(q)=>{await onAdd(q);setView("lista");}}/>}
      {view==="lista" && (
        quotes.length===0
          ? <div style={{textAlign:"center",padding:60,color:"#aaa",background:"#fff",borderRadius:12}}>
              <div style={{fontSize:48,marginBottom:8}}>\u{1F4C4}</div>
              <div>No hay cotizaciones. \u00a1Cre\u00e1 una!</div>
            </div>
          : quotes.map(q=><QuoteCard key={q.id} q={q} exp={expanded===q.id} toggle={()=>setExpanded(expanded===q.id?null:q.id)} getP={getP} onDel={onDel}/>)
      )}
    </div>
  );
}

function QuoteCard({q,exp,toggle,getP,onDel}) {
  const PURPLE = "#6c3483"; const PURPLEBG = "#e8daef";
  return (
    <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 6px #0001",overflow:"hidden",marginBottom:8}}>
      <div onClick={toggle} style={{padding:"13px 18px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",cursor:"pointer"}}>
        <div style={{flex:1,minWidth:100}}>
          <div style={{fontWeight:700,fontSize:14,color:"#1a1a1a"}}>{q.client}</div>
          <div style={{fontSize:11,color:"#aaa",display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
            {q.docNum&&<span style={{fontWeight:700,color:"#6c3483"}}>{q.docNum}</span>}
            <span>{q.date}</span>
            {q.vendedor&&<span>\u00b7 \u{1F464} {q.vendedor}</span>}
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
          <span style={{background:PURPLEBG,color:PURPLE,border:`1px solid ${PURPLE}44`,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>\u{1F4C4} Cotizaci\u00f3n</span>
          <span style={{fontWeight:800,color:PURPLE,fontSize:15}}>{fARS(q.total)}</span>
          <span style={{color:"#ccc"}}>{exp?"\u25b2":"\u25bc"}</span>
        </div>
      </div>
      {exp && (
        <div style={{borderTop:"1px solid #f5f5f5",padding:18}}>
          {q.validity&&<div style={{background:"#fef9e7",borderRadius:8,padding:"7px 12px",fontSize:12,color:"#7d6608",marginBottom:12}}>\u23F3 Validez: <strong>{q.validity}</strong></div>}
          {q.items.map((it,i)=>{const p=getP(it.pid);return <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #f9f9f9",fontSize:13}}><span style={{color:"#444"}}>{p?.name||it.name} \u00d7 {it.qty}</span><span style={{fontWeight:600}}>{fARS(it.price*it.qty)}</span></div>;})}
          <div style={{display:"flex",justifyContent:"flex-end",fontWeight:800,fontSize:16,color:PURPLE,margin:"8px 0 12px"}}>{fARS(q.total)}</div>
          {q.notes&&<div style={{background:"#f9f9f9",borderRadius:8,padding:"8px 12px",fontSize:13,color:"#555",marginBottom:12}}>\u{1F4AC} {q.notes}</div>}
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            <button onClick={()=>printDoc(q,"cotizacion")} style={{padding:"8px 12px",borderRadius:8,border:`1.5px solid ${PURPLEBG}`,cursor:"pointer",background:"#fff",color:PURPLE,fontWeight:600,fontSize:13}}>\u{1F5A8} Imprimir</button>
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
    <span style={{fontSize:12,color:RED,fontWeight:600}}>\u00bfEliminar?</span>
    <button onClick={onConfirm} style={{padding:"4px 10px",borderRadius:6,border:"none",background:RED,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:12}}>S\u00ed</button>
    <button onClick={()=>setC(false)} style={{padding:"4px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12}}>No</button>
  </div>;
  return <button onClick={()=>setC(true)} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #fcc",cursor:"pointer",background:"#fff",color:RED,fontWeight:600,fontSize:13}}>\u{1F5D1} Eliminar</button>;
}

function NuevaCotizacion({products,vendors,onAdd}) {
  const [client,setClient]=useState("");
  const [notes,setNotes]=useState("");
  const [vendedor,setVendedor]=useState("");
  const [validity,setValidity]=useState("48 horas");
  const [cart,setCart]=useState([]);
  const [ok,setOk]=useState(false);
  const total=cart.reduce((s,i)=>s+i.price*i.qty,0);
  const submit=async()=>{
    if(!client.trim()){alert("Ingres\u00e1 el cliente");return;}
    if(!cart.length){alert("Agreg\u00e1 productos");return;}
    const q={id:genId(),client:client.trim(),notes,vendedor,validity,items:cart,total,date:today()};
    setOk(true);
    await onAdd(q);
  };
  if(ok) return <div style={{textAlign:"center",padding:80}}><div style={{fontSize:60}}>\u{1F4C4}</div><div style={{fontWeight:800,color:"#6c3483",fontSize:20,marginTop:12}}>\u00a1Cotizaci\u00f3n guardada!</div></div>;
  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 330px",gap:18,alignItems:"start"}}>
      <div>
        <div style={{fontWeight:800,fontSize:15,marginBottom:12}}>\u{1F4C4} Nueva Cotizaci\u00f3n \u2014 Selecci\u00f3n\u00e1 productos</div>
        <ProductSelector products={products} cart={cart} setCart={setCart}/>
      </div>
      <div style={{position:"sticky",top:16}}>
        <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 2px 12px #0002",border:"2px solid #e8daef"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14,color:"#6c3483"}}>\u{1F4C4} Resumen de Cotizaci\u00f3n</div>
          <div style={{background:"#e8daef",borderRadius:8,padding:"7px 12px",fontSize:12,color:"#6c3483",marginBottom:14}}>
            \u2139\ufe0f Las cotizaciones <strong>no descuentan stock</strong>. Son solo presupuestos para el cliente.
          </div>
          <Field label="Cliente *"><input value={client} onChange={e=>setClient(e.target.value)} placeholder="Nombre del cliente" style={inputStyle}/></Field>
          <Field label="Vendedor">
            <select value={vendedor} onChange={e=>setVendedor(e.target.value)} style={{...inputStyle,color:vendedor?"#1a1a1a":"#aaa",cursor:"pointer"}}>
              <option value="">\u2014 Selecci\u00f3n\u00e1 vendedor \u2014</option>
              {vendors.map(v=><option key={v} value={v}>{v}</option>)}
            </select>
          </Field>
          <Field label="V\u00e1lida hasta (opcional)"><input value={validity} onChange={e=>setValidity(e.target.value)} placeholder="Ej: 30/06/2025" style={inputStyle}/></Field>
          <Field label="Notas"><textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Observaciones, condiciones..." style={{...inputStyle,resize:"vertical",minHeight:55,fontSize:12}}/></Field>
          <div style={{borderTop:"1px solid #f5f5f5",margin:"4px 0 10px",paddingTop:10}}>
            {cart.length===0?<div style={{textAlign:"center",color:"#aaa",fontSize:12,padding:"10px 0"}}>Agreg\u00e1 productos a la cotizaci\u00f3n</div>
            :cart.map(i=><div key={i.pid} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"3px 0",color:"#555"}}><span style={{flex:1,marginRight:6,lineHeight:1.3}}>{i.name} \u00d7 {i.qty}</span><span style={{fontWeight:600,whiteSpace:"nowrap"}}>{fARS(i.price*i.qty)}</span></div>)}
          </div>
          <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:17,color:"#6c3483",padding:"8px 0",borderTop:"2px solid #f5f5f5",marginBottom:14}}><span>Total</span><span>{fARS(total)}</span></div>
          <button onClick={submit} disabled={!cart.length||!client.trim()} style={{width:"100%",padding:"11px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:(!cart.length||!client.trim())?"#e5e5e5":"linear-gradient(135deg,#6c3483,#9b59b6)",color:(!cart.length||!client.trim())?"#aaa":"#fff"}}>
            \u{1F4C4} Guardar Cotizaci\u00f3n
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
        <span style={{fontSize:20}}>⚠️</span>
        <div style={{flex:1}}>
          <span style={{fontWeight:700,color:"#7d6608",fontSize:13}}>Stock bajo en {low.length} producto(s)</span>
          {!open && <span style={{fontSize:11,color:"#9a7d0a",marginLeft:8}}>— hacé clic para ver el detalle</span>}
        </div>
        <span style={{fontSize:12,color:"#9a7d0a",fontWeight:700}}>{open?"▲":"▼"}</span>
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
                {p.stock===0?"Sin stock":`⚠ ${p.stock} u.`}
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
      return e.producto?.toLowerCase().includes(q)||e.usuario?.toLowerCase().includes(q)||e.motivo?.toLowerCase().includes(q)||e.productoId?.toLowerCase().includes(q);
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
                <span style={{fontSize:12,color:RED,fontWeight:600}}>¿Limpiar todo el historial?</span>
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
                  const cfg = TIPOS[e.tipo]||{color:"#666",bg:"#f5f5f5",icon:"•"};
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
    q.length>1 ? products.filter(p=>p.name.toLowerCase().includes(q)||p.id.includes(q)).slice(0,40) : []
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
                  <div style={{fontSize:11,color:"#aaa",marginTop:3}}>Obligatorio — quedará registrado en el historial de movimientos</div>
                </div>
                <button onClick={confirmAction} disabled={qty===0||!reason.trim()} style={{width:"100%",padding:"10px",borderRadius:9,border:"none",background:(qty===0||!reason.trim())?"#e5e5e5":"linear-gradient(135deg,#1a5e20,#1e8449)",color:(qty===0||!reason.trim())?"#aaa":"#fff",fontWeight:800,fontSize:14,cursor:(qty===0||!reason.trim())?"not-allowed":"pointer"}}>
                  ✅ Confirmar ajuste
                </button>
              </>}
              {mode==="baja" && <>
                <div style={{background:"#fdecea",border:"1px solid #f5c6cb",borderRadius:10,padding:"12px 14px",marginBottom:14}}>
                  <div style={{fontWeight:700,color:RED,fontSize:13,marginBottom:4}}>⚠️ Dar de baja elimina el producto del catálogo</div>
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
      <div style={{fontWeight:700,color:RED,fontSize:13,marginBottom:10,textAlign:"center"}}>¿Confirmás la baja definitiva?</div>
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

function Stock({products,onUpd,onDel,onAdjust,isAdmin,addLog,stockLog,setStockLog}) {
  const [search,setSearch]=useState("");
  const [cat,setCat]=useState("todos");
  const [editing,setEditing]=useState(null);
  const [stockTab,setStockTab]=useState("lista");
  const CATS=useMemo(()=>["todos",...new Set(products.map(p=>p.category))].sort(),[products]);
  const q=search.toLowerCase();
  const shown=useMemo(()=>products.filter(p=>{
    if(cat!=="todos"&&p.category!==cat) return false;
    if(q) return p.name.toLowerCase().includes(q)||p.id.includes(q);
    return true;
  }).slice(0,150),[products,cat,q]);
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
          <div style={{fontWeight:800,fontSize:15,flex:1}}>📦 Stock ({products.filter(p=>p.stock>0).length} productos con stock)</div>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar..."
            style={{padding:"7px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:12,outline:"none",width:180}}/>
          <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
            {CATS.map(c=><button key={c} onClick={()=>setCat(c)} style={{padding:"4px 10px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:11,fontWeight:600,borderColor:cat===c?RED:"#e5e5e5",background:cat===c?"#fdecea":"#fff",color:cat===c?RED:"#666"}}>{c==="todos"?"Todos":c}</button>)}
          </div>
        </div>
        <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px #0001",overflow:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{background:"#f9f9f9"}}>
              {["Código","Producto","Categoría","Stock","P. Venta",...(isAdmin?["P. Costo","Margen"]:[]),""].map(h=><th key={h} style={{padding:"10px 12px",textAlign:"left",fontWeight:700,color:"#888",fontSize:11,textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>)}
            </tr></thead>
            <tbody>
              {shown.length===0
                ? <tr><td colSpan={8} style={{textAlign:"center",padding:40,color:"#aaa"}}>Sin resultados.</td></tr>
                : shown.map(p=>{
                    const m=p.costPrice>0?((p.salePrice-p.costPrice)/p.costPrice*100).toFixed(0):"—";
                    return <tr key={p.id} style={{borderTop:"1px solid #f5f5f5"}}>
                      <td style={{padding:"9px 12px",color:"#aaa",fontSize:11}}>{p.id}</td>
                      <td style={{padding:"9px 12px",fontWeight:600,color:"#1a1a1a",maxWidth:260}}>{p.name}</td>
                      <td style={{padding:"9px 12px",color:"#aaa",fontSize:11}}>{p.category}</td>
                      <td style={{padding:"9px 12px"}}><SPill n={p.stock}/></td>
                      <td style={{padding:"9px 12px",fontWeight:700,color:RED}}>{fARS(p.salePrice)}</td>
                      {isAdmin&&<td style={{padding:"9px 12px",color:"#666"}}>{fARS(p.costPrice)}</td>}
                      {isAdmin&&<td style={{padding:"9px 12px",fontWeight:700,color:+m>=40?"#1e8449":"#e67e22"}}>{m}%</td>}
                      <td style={{padding:"9px 12px"}}>{isAdmin&&<button onClick={()=>setEditing(p)} style={{padding:"4px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:11,fontWeight:600}}>✏️</button>}</td>
                    </tr>;
                  })
              }
            </tbody>
          </table>
        </div>
        {editing&&<EditModal p={editing} onSave={p=>{onUpd(p);setEditing(null);}} onClose={()=>setEditing(null)}/>}
      </>}
    </div>
  );
}

function EditModal({p,onSave,onClose}) {
  const [cost,setCost]=useState(p.costPrice);
  const [sale,setSale]=useState(p.salePrice);
  const [stock,setStock]=useState(p.stock);
  const m=cost>0?((sale-cost)/cost*100).toFixed(1):"—";
  return (
    <div style={{position:"fixed",inset:0,background:"#0007",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999,padding:16}}>
      <div style={{background:"#fff",borderRadius:16,padding:24,width:"100%",maxWidth:440,boxShadow:"0 20px 60px #0003"}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:16}}>
          <div style={{fontWeight:800,fontSize:14,color:"#1a1a1a",flex:1,marginRight:8,lineHeight:1.3}}>{p.name}</div>
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:22,color:"#aaa"}}>×</button>
        </div>
        {[["Precio Costo ($)",cost,setCost],["Precio Venta ($)",sale,setSale],["Stock",stock,setStock]].map(([lbl,val,set])=>(
          <div key={lbl} style={{marginBottom:12}}>
            <label style={{fontSize:12,fontWeight:600,color:"#666",display:"block",marginBottom:4}}>{lbl}</label>
            <input type="number" value={val} onChange={e=>set(+e.target.value||0)} style={{...inputStyle}}/>
          </div>
        ))}
        <div style={{background:"#f0fdf4",borderRadius:8,padding:"8px 12px",fontSize:12,color:"#1e8449",marginBottom:14}}>
          Margen: <strong>{m}%</strong> · Ganancia/u: <strong>{fARS(sale-cost)}</strong>
        </div>
        <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
          <button onClick={onClose} style={{padding:"8px 14px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontWeight:600,color:"#666"}}>Cancelar</button>
          <button onClick={()=>onSave({...p,costPrice:cost,salePrice:sale,stock})} style={{padding:"8px 14px",borderRadius:8,border:"none",background:RED,color:"#fff",cursor:"pointer",fontWeight:700}}>Guardar</button>
        </div>
      </div>
    </div>
  );
}

// ─── ALTA DE MERCADERÍA ───────────────────────────────────────────────────────
function Compras({products,onStock}) {
  const [search,setSearch]=useState("");
  const [items,setItems]=useState([]);
  const [ok,setOk]=useState(false);
  const found=useMemo(()=>{
    const q=search.toLowerCase();
    return q?products.filter(p=>p.name.toLowerCase().includes(q)||p.id.includes(q)).slice(0,50):[];
  },[products,search]);
  const addI=p=>{if(!items.find(i=>i.pid===p.id))setItems(x=>[...x,{pid:p.id,name:p.name,qty:1,cost:p.costPrice}]);};
  const remI=pid=>setItems(x=>x.filter(i=>i.pid!==pid));
  const updI=(pid,f,v)=>setItems(x=>x.map(i=>i.pid===pid?{...i,[f]:v}:i));
  const totalCost=items.reduce((s,i)=>s+i.qty*i.cost,0);
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
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(195px,1fr))",gap:10}}>
          {found.length>0
            ? found.map(p=>{
                const inL=items.find(i=>i.pid===p.id);
                return <div key={p.id} style={{background:"#fff",borderRadius:10,padding:14,border:inL?"2px solid #1e8449":"2px solid transparent",boxShadow:"0 1px 4px #0001"}}>
                  <div style={{fontWeight:700,fontSize:12,color:"#1a1a1a",marginBottom:3,lineHeight:1.3}}>{p.name}</div>
                  <div style={{fontSize:10,color:"#aaa",marginBottom:8}}>{p.id} · Stock actual: <strong>{p.stock}</strong></div>
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
                <div key={it.pid} style={{background:"#f9f9f9",borderRadius:8,padding:12,marginBottom:10}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
                    <span style={{fontWeight:600,fontSize:12,flex:1,marginRight:6,lineHeight:1.3}}>{it.name}</span>
                    <button onClick={()=>remI(it.pid)} style={{background:"none",border:"none",cursor:"pointer",color:RED,fontSize:18,lineHeight:1}}>×</button>
                  </div>
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
// ─── ALTA DE MERCADERÍA ───────────────────────────────────────────────────────
function Compras({products,onStock}) {
  const [search,setSearch]=useState("");
  const [items,setItems]=useState([]);
  const [ok,setOk]=useState(false);
  const found=useMemo(()=>{
    const q=search.toLowerCase();
    return q?products.filter(p=>p.name.toLowerCase().includes(q)||p.id.includes(q)).slice(0,50):[];
  },[products,search]);
  const addI=p=>{if(!items.find(i=>i.pid===p.id))setItems(x=>[...x,{pid:p.id,name:p.name,qty:1,cost:p.costPrice}]);};
  const remI=pid=>setItems(x=>x.filter(i=>i.pid!==pid));
  const updI=(pid,f,v)=>setItems(x=>x.map(i=>i.pid===pid?{...i,[f]:v}:i));
  const totalCost=items.reduce((s,i)=>s+i.qty*i.cost,0);
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
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(195px,1fr))",gap:10}}>
          {found.length>0
            ? found.map(p=>{
                const inL=items.find(i=>i.pid===p.id);
                return <div key={p.id} style={{background:"#fff",borderRadius:10,padding:14,border:inL?"2px solid #1e8449":"2px solid transparent",boxShadow:"0 1px 4px #0001"}}>
                  <div style={{fontWeight:700,fontSize:12,color:"#1a1a1a",marginBottom:3,lineHeight:1.3}}>{p.name}</div>
                  <div style={{fontSize:10,color:"#aaa",marginBottom:8}}>{p.id} · Stock actual: <strong>{p.stock}</strong></div>
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
                <div key={it.pid} style={{background:"#f9f9f9",borderRadius:8,padding:12,marginBottom:10}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
                    <span style={{fontWeight:600,fontSize:12,flex:1,marginRight:6,lineHeight:1.3}}>{it.name}</span>
                    <button onClick={()=>remI(it.pid)} style={{background:"none",border:"none",cursor:"pointer",color:RED,fontSize:18,lineHeight:1}}>×</button>
                  </div>
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
function AdminPanel({users,setUsers,vendors,setVendors,products,setProducts,stockLog,setStockLog,notifs,setNotifs}) {
  const [section, setSection] = useState("vendors");

  const SECTIONS = [
    {k:"vendors", label:"Vendedores",       icon:"👥"},
    {k:"users",   label:"Usuarios",         icon:"🔐"},
    {k:"excel",   label:"Lista de Precios", icon:"📊"},
    {k:"notifcfg",label:"Notificaciones",   icon:"🔔"},
  ];

  return (
    <div>
      <div style={{background:"#fff",borderRadius:12,padding:4,marginBottom:16,display:"flex",gap:4,boxShadow:"0 1px 4px #0001",flexWrap:"wrap"}}>
        {SECTIONS.map(s=>(
          <button key={s.k} onClick={()=>setSection(s.k)} style={{flex:1,minWidth:120,padding:"10px 16px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:section===s.k?`linear-gradient(135deg,${REDD},${RED})`:"transparent",color:section===s.k?"#fff":"#555",transition:"all .15s"}}>
            {s.icon} {s.label}
          </button>
        ))}
      </div>
      {section==="vendors"  && <VendorsPanel vendors={vendors} setVendors={setVendors}/>}
      {section==="users"    && <UsersPanel   users={users}     setUsers={setUsers}/>}
      {section==="excel"    && <ExcelPanel   products={products} setProducts={setProducts}/>}
      {section==="notifcfg" && <NotifConfig  users={users} setUsers={setUsers} notifs={notifs} setNotifs={setNotifs}/>}
    </div>
  );
}

// ── Vendors ───────────────────────────────────────────────────────────────────
// FIX: DB-first pattern — primero persistir en Supabase, luego actualizar estado local.
// Esto evita que la UI quede desincronizada si Supabase falla.
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
                  <button onClick={()=>saveEdit(v)} style={{padding:"5px 12px",borderRadius:7,border:"none",background:"#1e8449",color:"#fff",fontWeight:700,cursor:"pointer",fontSize:12}}>✓</button>
                  <button onClick={()=>setEditing(null)} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12}}>✕</button>
                </>
              : <>
                  <span style={{flex:1,fontWeight:600,fontSize:14}}>{v}</span>
                  <button onClick={()=>{setEditing(v);setEditVal(v);}} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12}}>✏️ Editar</button>
                  {confirmDel===v
                    ? <div style={{display:"flex",alignItems:"center",gap:5,background:"#fdecea",borderRadius:8,padding:"4px 8px",border:`1px solid ${RED}44`}}>
                        <span style={{fontSize:11,color:RED,fontWeight:600,whiteSpace:"nowrap"}}>¿Eliminar?</span>
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
function UsersPanel({users,setUsers}) {
  const [form, setForm] = useState({username:"",password:"",name:"",role:"vendedor"});
  const [editing, setEditing] = useState(null);
  const [showPass, setShowPass] = useState({});

  const startEdit = (u) => { setEditing(u.id); setForm({username:u.username,password:u.password,name:u.name,role:u.role}); };
  const cancelEdit = () => { setEditing(null); setForm({username:"",password:"",name:"",role:"vendedor"}); };

  const save = async () => {
    if(!form.username.trim()||!form.password.trim()||!form.name.trim()){alert("Completa todos los campos");return;}
    if(editing) {
      const updated = {...users.find(u=>u.id===editing), ...form};
      setUsers(us=>us.map(u=>u.id===editing?updated:u));
      await db.saveUser(updated);
    } else {
      if(users.find(u=>u.username===form.username.trim())){alert("Ese usuario ya existe");return;}
      const newUser = {id:genId(),...form,username:form.username.trim(),name:form.name.trim()};
      setUsers(us=>[...us,newUser]);
      await db.saveUser(newUser);
    }
    cancelEdit();
  };
  const remove = async (id) => {
    if(users.filter(u=>u.role==="admin").length===1&&users.find(u=>u.id===id)?.role==="admin"){alert("Debe haber al menos un administrador");return;}
    setUsers(us=>us.filter(u=>u.id!==id));
    await db.deleteUser(id);
  };

  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,alignItems:"start"}}>
      <div style={{background:"#fff",borderRadius:12,padding:24,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:15,marginBottom:16}}>{editing?"✏️ Editar usuario":"➕ Nuevo usuario"}</div>
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
            <button onClick={()=>setShowPass(s=>({...s,form:!s.form}))} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:15,color:"#aaa"}}>{showPass.form?"🙈":"👁"}</button>
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
                <div style={{fontSize:11,color:"#888"}}>@{u.username} · <span style={{color:u.role==="admin"?RED:"#1a5276",fontWeight:600}}>{u.role==="admin"?"Admin":"Vendedor"}</span>{u.email&&<span style={{color:"#aaa",marginLeft:6}}>· {u.email}</span>}</div>
              </div>
              <button onClick={()=>startEdit(u)} style={{padding:"4px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:11}}>✏️</button>
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
              .normalize("NFD").replace(/[\u0300-\u036f]/g,"");
            if(v.includes("CODIGO") || v.includes("COD")) { headerRow = r; break; }
          }
        }

        const norm = s => String(s).toUpperCase().trim()
          .normalize("NFD").replace(/[\u0300-\u036f]/g,"");

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

        const colLetter = c => c < 0 ? "—" : String.fromCharCode(65 + c);
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
      const precioCosto = resolvePrice(row);
      if(idx>=0){
        if(row.precioFinal!==null && row.precioFinal>0)
                                   newProds[idx].salePrice = row.precioFinal;
        else if(precioCosto!==null) newProds[idx].salePrice = precioCosto;
        if(precioCosto!==null)     newProds[idx].costPrice = precioCosto;
        if(row.name) newProds[idx].name = row.name;
        updated.push(row.id);
      } else {
        notFound.push(row.id);
      }
    });

    if(mode==="full") {
      preview.all.forEach(row=>{
        if(!newProds.find(p=>p.id===row.id)) {
          const precio = resolvePrice(row);
          const costo = resolvePrice(row);
          newProds.push({
            id:row.id, name:row.name||row.id,
            costPrice:costo||0,
            salePrice:(row.precioFinal&&row.precioFinal>0)?row.precioFinal:(costo||0),
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
          <strong>CÓDIGO · DESCRIPCIÓN · PRECIO CON IVA · PRECIO OFERTA · FECHA ULTIMA ACTUALIZACIÓN · PRECIO FINAL</strong>
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
          <div style={{fontSize:12,color:"#aaa",marginTop:4}}>.xlsx · .xls · .csv</div>
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
          {preview ? `👁 Vista previa (${preview.total} filas)` : "📋 Instrucciones"}
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
                      <td style={{padding:"6px 8px",color:"#666"}}>{r.precioIVA!=null?fARS(r.precioIVA):"—"}</td>
                      <td style={{padding:"6px 8px",color:"#666"}}>{r.precioOferta!=null?fARS(r.precioOferta):"—"}</td>
                      <td style={{padding:"6px 8px",fontWeight:700,color:RED}}>{r.precioFinal!=null?fARS(r.precioFinal):"—"}</td>
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
