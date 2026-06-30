/* eslint-disable */
import { createClient } from "@supabase/supabase-js";
import React, { useState, useMemo, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
// ─── Iconos propios (SVG inline, sin dependencias externas) ──────────────────
// Reemplaza a lucide-react: mismo nombre/props (size, color, strokeWidth) para no tener
// que tocar nada mas en el codigo. No requiere "npm install" ni tocar package.json.
function Ico({ children, size = 16, strokeWidth = 2, color, style }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color || "currentColor"}
      strokeWidth={strokeWidth} strokeLinecap="round" strokeLinejoin="round" style={style}>
      {children}
    </svg>
  );
}
const Clock = (p) => <Ico {...p}><circle cx="12" cy="12" r="9"/><polyline points="12 7 12 12 16 14"/></Ico>;
const CheckCircle2 = (p) => <Ico {...p}><circle cx="12" cy="12" r="9"/><polyline points="8 12 11 15 16 9"/></Ico>;
const Package = (p) => <Ico {...p}><path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/></Ico>;
const PartyPopper = (p) => <Ico {...p}><path d="M5 19l3-8 8-3-3 8-8 3z"/><path d="M14 5l1 2"/><path d="M19 9l1 2"/><path d="M17 4l2 1"/></Ico>;
const Wallet = (p) => <Ico {...p}><rect x="3" y="6" width="18" height="13" rx="2"/><path d="M3 10h18"/><circle cx="16" cy="14" r="1" fill={p.color||"currentColor"} stroke="none"/></Ico>;
const Search = (p) => <Ico {...p}><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.5" y2="16.5"/></Ico>;
const ShieldCheck = (p) => <Ico {...p}><path d="M12 3l7 3v6c0 5-3 8-7 9-4-1-7-4-7-9V6l7-3z"/><polyline points="9 12 11.5 14.5 15 10"/></Ico>;
const Paperclip = (p) => <Ico {...p}><path d="M19 11l-7.5 7.5a4 4 0 1 1-5.5-5.5l8-8a2.7 2.7 0 1 1 3.8 3.8L11 15.5"/></Ico>;
const Banknote = (p) => <Ico {...p}><rect x="2" y="6" width="20" height="12" rx="2"/><circle cx="12" cy="12" r="3"/></Ico>;
const ClipboardList = (p) => <Ico {...p}><rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 3h6v3H9z"/><line x1="9" y1="11" x2="15" y2="11"/><line x1="9" y1="15" x2="15" y2="15"/></Ico>;
const ShoppingCart = (p) => <Ico {...p}><circle cx="9" cy="20" r="1"/><circle cx="18" cy="20" r="1"/><path d="M3 4h2l2.4 12.2a2 2 0 0 0 2 1.6h7.8a2 2 0 0 0 2-1.6L21 8H6"/></Ico>;
const Users = (p) => <Ico {...p}><circle cx="9" cy="8" r="3.5"/><path d="M3 20c0-3.5 3-6 6-6s6 2.5 6 6"/><circle cx="17.5" cy="9.5" r="2.8"/><path d="M22 20c0-2.8-2-5-4.5-5.6"/></Ico>;
const FileText = (p) => <Ico {...p}><path d="M7 3h7l4 4v13a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z"/><path d="M14 3v4h4"/><line x1="9" y1="13" x2="15" y2="13"/><line x1="9" y1="17" x2="15" y2="17"/></Ico>;
const CircleDollarSign = (p) => <Ico {...p}><circle cx="12" cy="12" r="9"/><path d="M9 15.5c.5 1 1.5 1.5 3 1.5 2 0 3-.9 3-2.2 0-3-6-1.4-6-4.4 0-1.3 1.2-2.2 3-2.2 1.5 0 2.5.5 3 1.5"/><line x1="12" y1="6.5" x2="12" y2="17.5"/></Ico>;
const Store = (p) => <Ico {...p}><path d="M3 9l1.5-5h15L21 9"/><path d="M4 9v11h16V9"/><path d="M9 20v-6h6v6"/></Ico>;
const ListChecks = (p) => <Ico {...p}><polyline points="3 7 4.5 8.5 7 6"/><line x1="11" y1="7" x2="21" y2="7"/><polyline points="3 13 4.5 14.5 7 12"/><line x1="11" y1="13" x2="21" y2="13"/><polyline points="3 19 4.5 20.5 7 18"/><line x1="11" y1="19" x2="21" y2="19"/></Ico>;
const Star = (p) => <Ico {...p}><path d="M12 3l2.9 6 6.6.6-5 4.4 1.5 6.5L12 17l-5.9 3.5L7.6 14l-5-4.4 6.6-.6z"/></Ico>;
const CheckCircle = (p) => <Ico {...p}><circle cx="12" cy="12" r="9"/><polyline points="8 12 11 15 16 9"/></Ico>;
const XCircle = (p) => <Ico {...p}><circle cx="12" cy="12" r="9"/><line x1="9" y1="9" x2="15" y2="15"/><line x1="15" y1="9" x2="9" y2="15"/></Ico>;
const AlertTriangle = (p) => <Ico {...p}><path d="M12 3l9.5 17H2.5z"/><line x1="12" y1="10" x2="12" y2="14.5"/><circle cx="12" cy="17.3" r="0.4" fill={p.color||"currentColor"} stroke="none"/></Ico>;
const XIcon = (p) => <Ico {...p}><line x1="6" y1="6" x2="18" y2="18"/><line x1="18" y1="6" x2="6" y2="18"/></Ico>;
const Pencil = (p) => <Ico {...p}><path d="M4 20l1-4 11-11 3 3-11 11-4 1Z"/><path d="M14 6l3 3"/></Ico>;
const Truck = (p) => <Ico {...p}><rect x="1" y="7" width="13" height="9" rx="1"/><path d="M14 10h4l3 3v3h-7z"/><circle cx="6" cy="18.5" r="1.6"/><circle cx="17" cy="18.5" r="1.6"/></Ico>;
const TrendUp = (p) => <Ico {...p}><polyline points="3 17 9 11 13 15 21 7"/><polyline points="14 7 21 7 21 14"/></Ico>;
const TrendDown = (p) => <Ico {...p}><polyline points="3 7 9 13 13 9 21 17"/><polyline points="14 17 21 17 21 10"/></Ico>;
const Home = (p) => <Ico {...p}><path d="M3 11l9-8 9 8"/><path d="M5 10v10h14V10"/></Ico>;
const CreditCard = (p) => <Ico {...p}><rect x="2" y="5" width="20" height="14" rx="2"/><line x1="2" y1="10" x2="22" y2="10"/></Ico>;
const Gift = (p) => <Ico {...p}><rect x="3" y="9" width="18" height="11" rx="1"/><path d="M12 9v11"/><path d="M8 9c-2 0-3-1.5-3-3s2-3 4 0c2-3 4-3 4 0s-1 3-3 3"/></Ico>;
const Tag = (p) => <Ico {...p}><path d="M3 11l9-8 9 8-9 8z"/><circle cx="9" cy="9" r="1.3" fill={p.color||"currentColor"} stroke="none"/></Ico>;
const Beaker = (p) => <Ico {...p}><path d="M9 3h6"/><path d="M10 3v6l-5 9a2 2 0 0 0 2 3h10a2 2 0 0 0 2-3l-5-9V3"/></Ico>;
const Lock = (p) => <Ico {...p}><rect x="5" y="10" width="14" height="10" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3"/></Ico>;
const BarChart = (p) => <Ico {...p}><line x1="5" y1="20" x2="5" y2="11"/><line x1="12" y1="20" x2="12" y2="5"/><line x1="19" y1="20" x2="19" y2="14"/></Ico>;
const Bell = (p) => <Ico {...p}><path d="M6 9a6 6 0 1 1 12 0c0 3 1 5 2 6H4c1-1 2-3 2-6Z"/><path d="M10 19a2 2 0 0 0 4 0"/></Ico>;
const ChevronDown = (p) => <Ico {...p}><polyline points="6 9 12 15 18 9"/></Ico>;
const MessageCircle = (p) => <Ico {...p}><path d="M21 11.5a8.5 8.5 0 1 1-3.8-7.1"/><path d="M3 21l1.9-4.7A8.5 8.5 0 0 1 12 3a8.5 8.5 0 0 1 8.5 8.5c0 1.6-.4 3-1.1 4.3"/></Ico>;
const Printer = (p) => <Ico {...p}><path d="M6 9V3h12v6"/><rect x="4" y="9" width="16" height="8" rx="1.5"/><path d="M6 17v4h12v-4"/></Ico>;
const RefreshCw = (p) => <Ico {...p}><path d="M3 12a9 9 0 0 1 15-6.7L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-15 6.7L3 16"/><path d="M3 21v-5h5"/></Ico>;
const Save = (p) => <Ico {...p}><path d="M5 4h11l3 3v13H5z"/><path d="M8 4v5h7V4"/><path d="M8 14h8v6H8z"/></Ico>;
const Send = (p) => <Ico {...p}><path d="M22 2L11 13"/><path d="M22 2l-7 20-4-9-9-4z"/></Ico>;
const Camera = (p) => <Ico {...p}><path d="M4 8h3l1.5-2h7L17 8h3a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V9a1 1 0 0 1 1-1z"/><circle cx="12" cy="14" r="3.5"/></Ico>;
const Trash = (p) => <Ico {...p}><polyline points="4 7 20 7"/><path d="M6 7l1 13a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1l1-13"/><path d="M9 7V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v3"/></Ico>;
const LogOut = (p) => <Ico {...p}><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></Ico>;
const Key = (p) => <Ico {...p}><circle cx="7" cy="12" r="4"/><line x1="11" y1="12" x2="21" y2="12"/><line x1="21" y1="12" x2="21" y2="16"/><line x1="17" y1="12" x2="17" y2="15"/></Ico>;
const Menu = (p) => <Ico {...p}><line x1="4" y1="7" x2="20" y2="7"/><line x1="4" y1="12" x2="20" y2="12"/><line x1="4" y1="17" x2="20" y2="17"/></Ico>;
const Scale = (p) => <Ico {...p}><path d="M12 3v18"/><path d="M5 7h14"/><path d="M5 7l-3 7a3.5 3.5 0 0 0 7 0z"/><path d="M19 7l-3 7a3.5 3.5 0 0 0 7 0z"/></Ico>;
const Phone = (p) => <Ico {...p}><path d="M5 4h3l1.5 4.5L7 10.5a11 11 0 0 0 6.5 6.5l1.5-2.5L19.5 16v3a1.5 1.5 0 0 1-1.6 1.5C10.5 20 4 13.5 3.5 6.1A1.5 1.5 0 0 1 5 4z"/></Ico>;
const Mail = (p) => <Ico {...p}><rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3 6.5l9 6 9-6"/></Ico>;
const MapPin = (p) => <Ico {...p}><path d="M12 21s7-6.5 7-12a7 7 0 1 0-14 0c0 5.5 7 12 7 12z"/><circle cx="12" cy="9" r="2.5"/></Ico>;
const IdCard = (p) => <Ico {...p}><rect x="2" y="5" width="20" height="14" rx="2"/><circle cx="8" cy="11" r="2"/><path d="M5 17c0-1.5 1.3-2.5 3-2.5s3 1 3 2.5"/><line x1="14" y1="9" x2="19" y2="9"/><line x1="14" y1="13" x2="19" y2="13"/></Ico>;
const Plus = (p) => <Ico {...p}><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></Ico>;
const Info = (p) => <Ico {...p}><circle cx="12" cy="12" r="9"/><line x1="12" y1="11" x2="12" y2="16"/><circle cx="12" cy="7.5" r="0.4" fill={p.color||"currentColor"} stroke="none"/></Ico>;
const Lightbulb = (p) => <Ico {...p}><path d="M9 18h6"/><path d="M10 21h4"/><path d="M12 3a6 6 0 0 0-3.7 10.7c.5.4.7 1 .7 1.6V16h6v-.7c0-.6.2-1.2.7-1.6A6 6 0 0 0 12 3z"/></Ico>;
const Crown = (p) => <Ico {...p}><path d="M3 17h18l-1.5-8-4 3-2.5-6-2.5 6-4-3z"/><path d="M5 20h14"/></Ico>;
const EyeOff = (p) => <Ico {...p}><path d="M2 12s3.5-7 10-7c2 0 3.7.6 5.1 1.4M22 12s-1.2 2.4-3.3 4.2M9.5 9.7a3 3 0 0 0 4.2 4.2"/><line x1="3" y1="3" x2="21" y2="21"/></Ico>;
const Wrench = (p) => <Ico {...p}><path d="M14.7 6.3a4 4 0 1 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.5 2.5-2-2z"/></Ico>;
const Folder = (p) => <Ico {...p}><path d="M3 7a1 1 0 0 1 1-1h5l2 2h9a1 1 0 0 1 1 1v9a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1z"/></Ico>;
const ImageIcon = (p) => <Ico {...p}><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><path d="M21 15l-5-5L5 21"/></Ico>;
const UserMinus = (p) => <Ico {...p}><circle cx="9" cy="8" r="3.5"/><path d="M2 20c0-3.5 3-6 7-6s7 2.5 7 6"/><line x1="16" y1="10" x2="22" y2="10"/></Ico>;
const ArrowLeftIcon = (p) => <Ico {...p}><line x1="20" y1="12" x2="4" y2="12"/><polyline points="11 5 4 12 11 19"/></Ico>;
const ArrowRightIcon = (p) => <Ico {...p}><line x1="4" y1="12" x2="20" y2="12"/><polyline points="13 5 20 12 13 19"/></Ico>;
const Eye = (p) => <Ico {...p}><path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7Z"/><circle cx="12" cy="12" r="3"/></Ico>;
const BoxIcon = (p) => <Ico {...p}><path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/></Ico>;

// ─── Sistema global de toasts (reemplaza toast.error()) — sin Context ni prop-drilling ──
let _pushToast = null;
function ToastHost() {
  const [toasts, setToasts] = useState([]);
  const idRef = useRef(0);
  useEffect(() => {
    _pushToast = (type, msg) => {
      const id = ++idRef.current;
      setToasts(t => [...t, {id, type, msg}]);
      setTimeout(() => setToasts(t => t.filter(x => x.id !== id)), 4200);
    };
    return () => { _pushToast = null; };
  }, []);
  const dismiss = id => setToasts(t => t.filter(x => x.id !== id));
  const CFG = {
    success: {color:"#1e8449", bg:"#eafaf1", border:"#a9dfbf", Icon:CheckCircle},
    error:   {color:"#c0392b", bg:"#fdecea", border:"#f1948a", Icon:XCircle},
    info:    {color:"#1a5276", bg:"#eaf2f8", border:"#aed6f1", Icon:AlertTriangle},
  };
  return (
    <div style={{position:"fixed",bottom:18,right:18,left:18,display:"flex",flexDirection:"column",alignItems:"flex-end",gap:8,zIndex:99999,pointerEvents:"none"}}>
      {toasts.map(t=>{
        const c = CFG[t.type]||CFG.info;
        return (
          <div key={t.id} style={{background:"#fff",border:`1.5px solid ${c.border}`,borderRadius:11,padding:"11px 14px",display:"flex",alignItems:"center",gap:10,minWidth:240,maxWidth:360,boxShadow:"0 4px 18px rgba(0,0,0,.15)",pointerEvents:"auto",animation:"lm-toast-in .2s ease"}}>
            <div style={{width:26,height:26,borderRadius:8,background:c.bg,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
              <c.Icon size={14} color={c.color} strokeWidth={2.4}/>
            </div>
            <div style={{fontSize:12.5,color:"#333",flex:1,lineHeight:1.35}}>{t.msg}</div>
            <button onClick={()=>dismiss(t.id)} style={{background:"none",border:"none",cursor:"pointer",color:"#bbb",padding:2,flexShrink:0}}><XIcon size={13} strokeWidth={2.2}/></button>
          </div>
        );
      })}
      <style>{`@keyframes lm-toast-in{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}`}</style>
    </div>
  );
}
const toast = {
  success: (msg) => _pushToast && _pushToast("success", msg),
  error:   (msg) => _pushToast && _pushToast("error", msg),
  info:    (msg) => _pushToast && _pushToast("info", msg),
};

// ─── Sistema global de confirmación (reemplaza window.confirm()) ──────────
let _askConfirm = null;
function ConfirmHost() {
  const [state, setState] = useState(null);
  useEffect(() => {
    _askConfirm = (title, msg, danger=false) => new Promise(resolve => setState({title,msg,danger,resolve}));
    return () => { _askConfirm = null; };
  }, []);
  if(!state) return null;
  const close = (val) => { state.resolve(val); setState(null); };
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(20,20,20,.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100000,padding:20}} onClick={()=>close(false)}>
      <div onClick={e=>e.stopPropagation()} style={{background:"#fff",borderRadius:14,padding:"22px 24px",maxWidth:340,width:"100%",boxShadow:"0 12px 40px rgba(0,0,0,.25)",animation:"lm-pop-in .15s ease"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
          <div style={{width:34,height:34,borderRadius:10,flexShrink:0,background:state.danger?"#fdecea":"#eaf2f8",display:"flex",alignItems:"center",justifyContent:"center"}}>
            <AlertTriangle size={17} color={state.danger?"#c0392b":"#1a5276"} strokeWidth={2.2}/>
          </div>
          <div style={{fontWeight:700,fontSize:15,color:"#1a1a1a"}}>{state.title}</div>
        </div>
        <div style={{fontSize:13,color:"#666",marginBottom:18,lineHeight:1.5}}>{state.msg}</div>
        <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
          <button onClick={()=>close(false)} style={{padding:"8px 16px",borderRadius:8,border:"1px solid #ddd",background:"#fff",color:"#666",fontWeight:600,fontSize:12.5,cursor:"pointer"}}>Cancelar</button>
          <button onClick={()=>close(true)} style={{padding:"8px 16px",borderRadius:8,border:"none",fontWeight:700,fontSize:12.5,cursor:"pointer",color:"#fff",background:state.danger?"#c0392b":"#1e8449"}}>
            {state.danger?"Sí, continuar":"Confirmar"}
          </button>
        </div>
      </div>
      <style>{`@keyframes lm-pop-in{from{opacity:0;transform:scale(.95)}to{opacity:1;transform:scale(1)}}`}</style>
    </div>
  );
}
const confirmDialog = (title, msg, danger=false) => _askConfirm ? _askConfirm(title, msg, danger) : Promise.resolve(window.confirm(`${title}\n${msg}`));

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

const mapProduct = r => ({id:r.id,name:r.name,category:r.category,costPrice:r.cost_price,salePrice:r.sale_price,stock:r.stock,multiploCompra:r.multiplo_compra||1,barcode:r.barcode||"",costPriceAnterior:r.cost_price_anterior||0,imageUrl:r.image_url||""});
const mapOrder = r => ({id:r.id,client:r.client,vendedor:r.vendedor,notes:r.notes,total:r.total,stage:r.stage,date:r.date,items:r.items||[],docNum:r.doc_num||"",compNum:r.comp_num||"",isTest:r.is_test||false,isSandbox:r.is_sandbox||false,internalNote:r.internal_note||"",editStatus:r.edit_status||"",editReason:r.edit_reason||"",editItems:r.edit_items||null,editRejectReason:r.edit_reject_reason||"",comprobanteUrl:r.comprobante_url||"",comprobanteNombre:r.comprobante_nombre||"",comprobanteFecha:r.comprobante_fecha||"",pagoTipo:r.pago_tipo||"",pagoEfectivoFecha:r.pago_efectivo_fecha||""});
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

// ─── COMPROBANTES DE PAGO — Supabase Storage ──────────────────────────────────
// Bucket: "comprobantes" (público). Las imágenes se comprimen en el cliente antes
// de subir para no acumular fotos de cámara de varios MB cada una.
function compressImage(file, maxWidth = 1600, quality = 0.8) {
  return new Promise((resolve) => {
    if(!file || !file.type || !file.type.startsWith("image/")) { resolve(file); return; }
    try {
      const reader = new FileReader();
      reader.onload = (e) => {
        const img = new Image();
        img.onload = () => {
          try {
            let { width, height } = img;
            if(width > maxWidth) { height = Math.round(height * maxWidth / width); width = maxWidth; }
            const canvas = document.createElement("canvas");
            canvas.width = width; canvas.height = height;
            const ctx = canvas.getContext("2d");
            ctx.drawImage(img, 0, 0, width, height);
            canvas.toBlob((blob) => {
              if(!blob) { resolve(file); return; }
              resolve(new File([blob], file.name.replace(/\.\w+$/, ".jpg"), { type: "image/jpeg" }));
            }, "image/jpeg", quality);
          } catch { resolve(file); }
        };
        img.onerror = () => resolve(file);
        img.src = e.target.result;
      };
      reader.onerror = () => resolve(file);
      reader.readAsDataURL(file);
    } catch { resolve(file); }
  });
}
async function uploadComprobanteFile(orderId, file) {
  const processed = await compressImage(file);
  const ext = (processed.name.split(".").pop() || "jpg").toLowerCase();
  const path = `${orderId}-${Date.now()}.${ext}`;
  const { error } = await supabase.storage.from("comprobantes").upload(path, processed, { cacheControl: "3600", upsert: false });
  if(error) throw error;
  const { data } = supabase.storage.from("comprobantes").getPublicUrl(path);
  return { url: data.publicUrl, nombre: file.name };
}

// ─── FOTOS DE PRODUCTO — Supabase Storage ─────────────────────────────────────
// Bucket: "productos" (público). Comprimidas mas chico que los comprobantes (800px)
// porque son miniaturas para el selector, no documentos a leer en detalle.
async function uploadProductImage(productId, file) {
  const processed = await compressImage(file, 800, 0.8);
  const ext = (processed.name.split(".").pop() || "jpg").toLowerCase();
  const path = `${productId}-${Date.now()}.${ext}`;
  const { error } = await supabase.storage.from("productos").upload(path, processed, { cacheControl: "3600", upsert: false });
  if(error) throw error;
  const { data } = supabase.storage.from("productos").getPublicUrl(path);
  return data.publicUrl;
}

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
  upsertProduct: async (p) => { const {error} = await supaAdmin.from("lm_products").upsert({id:p.id,name:p.name,category:p.category||"Importado",cost_price:p.costPrice||0,sale_price:p.salePrice||0,stock:p.stock||0,multiplo_compra:p.multiploCompra||1,barcode:p.barcode||"",cost_price_anterior:p.costPriceAnterior||0,image_url:p.imageUrl||""}); if(error) throw error; },
  upsertProducts: async (arr) => { const {error} = await supaAdmin.from("lm_products").upsert(arr.map(p=>({id:p.id,name:p.name,category:p.category||"Importado",cost_price:p.costPrice||0,sale_price:p.salePrice||0,stock:p.stock||0,multiplo_compra:p.multiploCompra||1,barcode:p.barcode||"",cost_price_anterior:p.costPriceAnterior||0,image_url:p.imageUrl||""}))); if(error) throw error; },
  deleteProduct: async (id) => { const {error} = await supaAdmin.from("lm_products").delete().eq("id",id); if(error) throw error; },

  getOrders:    async () => { const {data,error} = await supabase.from("lm_orders").select("*").order("date",{ascending:false}); if(error) throw error; return (data||[]).map(mapOrder); },
  upsertOrder:  async (o) => { const {error} = await supaAdmin.from("lm_orders").upsert({id:o.id,client:o.client,vendedor:o.vendedor||"",notes:o.notes||"",total:o.total,stage:o.stage,date:o.date,items:o.items,doc_num:o.docNum||"",comp_num:o.compNum||"",is_test:o.isTest||false,is_sandbox:o.isSandbox||false,internal_note:o.internalNote||"",edit_status:o.editStatus||"",edit_reason:o.editReason||"",edit_items:o.editItems||null,edit_reject_reason:o.editRejectReason||"",comprobante_url:o.comprobanteUrl||"",comprobante_nombre:o.comprobanteNombre||"",comprobante_fecha:o.comprobanteFecha||"",pago_tipo:o.pagoTipo||"",pago_efectivo_fecha:o.pagoEfectivoFecha||""}); if(error) throw error; },
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
  getPurchaseOrders: async () => { const {data,error} = await supabase.from("lm_purchase_orders").select("*").order("fecha",{ascending:false}); if(error) throw error; return (data||[]).map(po=>({...po, fechaCierre:po.fecha_cierre||"", fechaRecibido:po.fecha_recibido||""})); },
  savePurchaseOrder: async (po) => { const {error} = await supaAdmin.from("lm_purchase_orders").upsert({id:po.id,fecha:po.fecha,vendedor:po.vendedor,estado:po.estado,items:po.items,notas:po.notas||"",fecha_cierre:po.fechaCierre||"",fecha_recibido:po.fechaRecibido||""}); if(error) throw error; },
  deletePurchaseOrder: async (id) => { const {error} = await supaAdmin.from("lm_purchase_orders").delete().eq("id",id); if(error) throw error; },
  // Activity log
  // SQL: CREATE TABLE lm_activity (id TEXT PRIMARY KEY, fecha TEXT, usuario TEXT, rol TEXT, accion TEXT, detalle TEXT, ref_id TEXT, ref_tipo TEXT);
  getActivity:  async () => { const {data,error} = await supabase.from("lm_activity").select("*").order("fecha",{ascending:false}).limit(500); if(error) throw error; return data||[]; },
  addActivity:  async (a) => { try { await supaAdmin.from("lm_activity").insert(a); } catch(e) {} }, // silent fail - non critical
  clearActivity: async () => { const {error} = await supaAdmin.from("lm_activity").delete().neq("id","none"); if(error) throw error; },
  // Ofertas y combos (promos)
  // SQL: CREATE TABLE lm_promos (id TEXT PRIMARY KEY, tipo TEXT NOT NULL, nombre TEXT NOT NULL, activa BOOLEAN DEFAULT true, vigencia_desde TEXT DEFAULT '', vigencia_hasta TEXT DEFAULT '', data JSONB DEFAULT '{}', created_at TEXT DEFAULT '');
  getPromos:    async () => { try { const {data,error} = await supabase.from("lm_promos").select("*").order("created_at",{ascending:false}); if(error) throw error; return (data||[]).map(r=>({id:r.id,tipo:r.tipo,nombre:r.nombre,activa:r.activa!==false,vigenciaDesde:r.vigencia_desde||"",vigenciaHasta:r.vigencia_hasta||"",data:r.data||{},createdAt:r.created_at||""})); } catch(e) { console.warn("getPromos:", e); return []; } },
  savePromo:    async (p) => { const {error} = await supaAdmin.from("lm_promos").upsert({id:p.id,tipo:p.tipo,nombre:p.nombre,activa:p.activa!==false,vigencia_desde:p.vigenciaDesde||"",vigencia_hasta:p.vigenciaHasta||"",data:p.data||{},created_at:p.createdAt||new Date().toISOString()}); if(error) throw error; },
  deletePromo:  async (id) => { const {error} = await supaAdmin.from("lm_promos").delete().eq("id",id); if(error) throw error; },
  // Configuracion global de la app (un solo registro con id="global")
  // SQL: CREATE TABLE lm_settings (id TEXT PRIMARY KEY DEFAULT 'global', exigir_pago_confirmado BOOLEAN DEFAULT false);
  getSettings:  async () => {
    try {
      const {data,error} = await supabase.from("lm_settings").select("*").eq("id","global").maybeSingle();
      if(error) throw error;
      return {exigirPagoConfirmado: data?.exigir_pago_confirmado || false};
    } catch(e) { console.warn("getSettings:", e); return {exigirPagoConfirmado:false}; }
  },
  saveSettings: async (s) => { const {error} = await supaAdmin.from("lm_settings").upsert({id:"global", exigir_pago_confirmado:s.exigirPagoConfirmado||false}); if(error) throw error; },
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
// Iconos reales (solo para presentacion visual) — SCFG.icon sigue como string para notificaciones/logs
const STAGE_ICONS = { reserva: Clock, confirmado: CheckCircle2, "en armado": Package, entregado: PartyPopper };
const Dot = () => <span style={{width:3,height:3,borderRadius:99,background:"#ccc",display:"inline-block"}}/>;
const SCFG = {
  reserva:     {label:"Reserva",    color:"#b7770d", bg:"#fef9e7", icon:"🕐"},
  confirmado:  {label:"Confirmado", color:"#1e8449", bg:"#eafaf1", icon:"✅"},
  "en armado": {label:"En Armado",  color:"#1a5276", bg:"#eaf4fc", icon:"📦"},
  entregado:   {label:"Entregado",  color:"#6c3483", bg:"#f5eef8", icon:"🎉"},
};

const LOGO = "/logo.png";
const PDF_LOGO_BANNER = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAQDAwMDAgQDAwMEBAQFBgoGBgUFBgwICQcKDgwPDg4MDQ0PERYTDxAVEQ0NExoTFRcYGRkZDxIbHRsYHRYYGRj/2wBDAQQEBAYFBgsGBgsYEA0QGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBj/wAARCAF3CWADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD7+ooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAopM0Z5oAWimlwoy3H14qCTULKI4ku4VPoXGaB2LNFUP7Yss4RpJT6Rxs39KDqT/APLPTrxx6lAn8yKVx8rL9JxWbJqdwoz9jjT/AK63Cr/jUH9sOf8AlrpyH08/d/Ki6DkZs8UflWGdWmz/AMfdlj/ZSQ/0oGpXDdLpD/uWrGhyDkZt5HqKX8qxhe3R/wCW8x+lk3+NO+1Xh5864H/bif8AGlcOQ1/yorJ+13X/AD0n/Gyb/Gm/broH/WN/wK1YU7hyM2OKWscalODhpYB/vROKeNUcY+e2PtuZf5ii6DkZq0VmpqhJ5SA/7s4z+tTLfk9bWUe4w38jRcOVlyiqo1CAHDiSP/fQinreWrHCzx59CcU7isT0U0OrDgg/Q0uRQIWiijn0oAKKOfSigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKQkUALRUFxd21qm64njiX1dsVQOsiYZsLO4uQP48eWn/fTYpNjUWzWpCcDNc5Nq1w77Gv7aFs/6q0Q3En0z0FQSIzr5txazMvaTVLkRj/vgf4UuYtU31NybVtPgYLJdxlj/AAp8x/IVGdVkdd1vp9y6/wB+UCJf/HqxopmYbba5Yg8FNMtMAf8AA2pz22TvntYQ397ULoyH67F4/lSuV7NLctS6xPu2m80+AnoqEzMfwFR+ddz5IfVZx0+SNYF/M80RRzEbIbibYedtlbCID/gTU2SC3DZuRGWHe8uy5/75HFFx2RDJ5KcyxWSjubq7Mp/IZFSRScYhlXH8P2OwyP8AvpuKmiaNCPsqn/t1s8fq1PkeQcTCUA97m8CfotILkZW7IxINSYess6Qj8hUZijBxJFZY9Jrp5D+Q4p6tExwpsw3YxwvKfzPFSqbjbt3XgHqkSQj8zQGqIUtwpzGtmo/6ZWTN+pqwILgjKvcj/ct41/nULSjGJpTn/ptfY/RaYrWpbpaN/wAAkl/WgdmWdko+9dXY+tyiU1xHjL3Lf8CvM/yFRhTn93C2P+mdln/0I1IDKP4Jh/wCOP8AnQTZkebQt808R+tzJTgLM9HiI9ppaGuJF6yzr9bmIfypovPW6H/Artf6CncZKEtem5P+/ktHl2/Z1H/bWWmC7TH/AB8wn/t8P9BS/a0H/LxH/wCBTf4Uw1HAxL0uQv0nf/CnB0I/4/GP/bx/iKj+1qf+XlB9Lo/4Uv2oE83Yz7XK/wBRSsFmP2E8icEe5jb+Yo+ylj8scbfWFf8A2U0wXCnq276yRmlyrD7qn6Ih/kaYtR5glU8KQfUGRf8AGmnzCcFmJ9N6t+hFOUjosZX3CMP5Gn+cQNvmMPfew/mKAuRBCORHt+keP1U08TyIcCUr7CT+jCnhw3AZT/3w3+FO3Eg5QY9MMP8AGgQqXtwOSdw90/qpNTJqIPBQE99jA/oarbImONgz6grn+lKYlI5ZgOwYE/zz/OgVkXheQnG5ihP98YqZXVxlWBHsc1lCJsHy2Byf4Tj+WabtkQ52n8uf0xTuLkRs0VlJczJx5hI/2uf54NWFvum9B9QcfzouJxZdoqFLiJjjdg+h4qbIpk7BRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUVG00affkRfqwFAWuSUUisCMggj1FGR60AGaM0xpYkzukQY7E1QvdW+z2/m2lpLqB7x2zpu/8AHmGaTaLjTlJ2SNLNGa8x1n43eH/DjsPEPhzxdpqDkySaPK6Y9dyZH61l2/7T3wdnOD4jmhPpNaSKR+lc7xdGLtKaXzPWpcPZnWjz0cPKS7xTa/C57FmjNcDofxn+GHiO7FtpfjXTHnPAimk8lj9N4Ga7pHV4g8bKytyGByD+Na06sKivCSfoefisFicJLkxFNwf95NfmS0UUVocwUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRUU88VvEZZpUjQclnOAKAJaYZAqktgAdSegrFm115YmfT4VMK9bq5PlxD6Z5b8Kynma+XznZ9QRT/rZz9ntE+g6tU8xoqbZuya5bs5isYpL6QdfIHyr9WPArMm1S7nl8p7wRuf8Al209fOk+hfoKqLuu4toaW9iX+GH/AEa0X8erVPDC0kZhWVmjHWDTV8qIezSHk0r3NFFRIWItpdzpbWch/iuWNzcH6KOhqR4ZJ18yeC4nT/nrqM3kx/URipIGghkMVp5cbnrHYR+bIf8AekNI7rHIDItvBJ6zE3Mx/wCAjgUh9R8SSvFsjupSo4MenQiFB/wM/wA6aEsrebJjs45T1Lk3Mv8A9allDFA9yJnTs17MIUx7IvNJFKxj22rTOg6iygEKY93agNSxI8rR7pY7t17m5lFvH/3yOaiikyxFvJCrDtZWxkbH+81VTcW6zAA2Ylx0G68lx+HA/lUz/a5EHmRXRUcg3k626f8AfCc4oCxJMRkfa84PT7bdY59kX+lEMhIzbA8cH7HaYH/fbUyC1lJPktDHu5/0K13H/vt+tNm+wo3+mXPmseMXN2XP/fCfyoCw+aZASlyyjI6Xd5z/AN8LTYpQy5t1BI72lkW/8fenQyxxxk2VjKFU9YbVYV/76fn8aHuZ3kwwgBx0muWlb8kGKBDpGu8Aut0Mjnz7tIR+SU1bYvKCEsuf9mSY/rwaVVvChKuyD1gs1T9XNI6lsLNcyN/sy3hx+SCga8idbaeNiRNJGP8AplbRxf8AoVRvJAsmJtSkJ7hr0Kf++VFR/Y4GbIto3PqbdpD+bmpdqxn7xQD/AK5R/wAuaA3I3bTnPJWY/wC0ZpDUiraBRtsGP/bmT+rGk85HPM+/2892/wDQRS7UbkwFvrA7fqxoEKHRTkae6j18uJf5mn/ayo+VCP8AekiX+Qpm0L0g2/SONf5k0CQ9On1dB/IUBa4pvJs8Rp/4ED+i0v2uY9Y1/wDAg/8AxNAeQjqv/f3/AAFKGkPYn6O3+FO4W8hPtUw/gH/f8/8AxNKLucH/AFa/9/s/+y0uZuwb/vp/8KC0uOv6v/hSuHyENwzffgT/AL7U/wA1pA8WebZD9BGf8KXL+g/Nv8KUO46KPzP+FFxWG/6OzDNoc+yj+jVJ+6UZ2zRj2Dj+tMy5+8o/MUuQD/qx+lFxjt0R/wCXh/8AgZz/ADWlVYz9yaNz9F/oRTdx24wQPbP+NJuTGMfmKLgT7Zx/ED7Bj/Ig0ZkXnyznttx/Qj+VQBYlHyAAn04/lTxuAwsrZPq1FxWJfN5+cc99wPH5inLKCMqwwPTkf1FRB5gMK4I+n+FBbJG+FDj06/qKomxPuRhztY+mP8/yoMaZ4BX6GoQ0ZHIkGfxH9acrcnZMpA7Hj/P5UBYf5JHAYY9OlKDLERgsp9u9IDIMApke1OEo6Zx7HigCVLxx99Qcd+lWEuYn77T6GqnyOeVB9xwaQxD+E/nzTuTZGjmlrNDTRdCcfmKnS77MPxFFyeUt0UxZEf7rA0+mIKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiik3YoAD1pu734qnqusaZoulS6lq17BZ2kSlpJ53Cqo+pr5a+JX7W5juZNM+G9pG6glW1S8TIJ/6Zp/U/lXJisbRwsb1H8j38h4YzHPavs8DTulu3pFer/Tc+pNV1nStD057/AFjUbWwtkGWmuZBGoH1NeB+OP2tfB+iNLaeELKbxBdjgTsfKtgfXJ+ZvwH418p6nqnjv4i6ytzrWpXupyyN+7a5chAT2RP6KK9Z8Afs722oTwT+L9VgtUcjFvPOIC+ewTO9v0rxJZricU+XCxsu7P1Oj4f5JkNP6xn1f2kv5I6L07v8AA5bWvjl4v8Z6hu13W9e8vOU0rw/L9ij+hZdzn6810vhXWPiVOVPhL4LJeSHlb7XBc30n18yeQKPyr6w8O+BvBPgPRAthpunWcUS5a4eNE/Esf8aS+8eW6I0unWytZJ/rNSv5Ra2yD2ZuX/AY963p5dVj71atr5f5u54+K4xwNa9DLMuXItFzSaXzUeVN+rbPIbLRv2sPEAH27xHoPhmBhwscSM6+2FU/zrct/gx46uIi/jT46+JZYz1SwdbVQfZjn+VZHjP9p3wjoge20zVLjW7teCmlxiKEH/rq+SR9BXz14u+PnjjxRvit5o9KtmyNtuTJKR6GVyW/LFY18ZhKGjnKb9X+mh6OUcOcRZraVKhTw1N9fZxT+XMnL+tz3W8/Z80GTUXudJ+J/ji7vc8yRSmdyR2L8D8yK0LD4I/FaKUmx+M/iWxgA+X7URI5+oDkCvjddf11XLrrepKxOSRdSD+tdh4ZX4sa8yDw9rWsOD0P9p+WP1fNcNPMaEnaNJ38m/8AM+rxnBubUKXNXzCnyrT36cXb70fVS/C74+Wg/wBD+OX2gel5Y7v6Gsqb4EfEPXLmSTxfrvgbUzjKM2grvkP+267Gx7g5rzm2+HH7T8VgLseKtQgj9BrBc49cKTn6dapXtx8fdGnaGH4kPcTIMtBNdmGQfhKAD+ddzrQS9+jO3m3+TZ8xDLcTKdsLmGG5u8YKL++MdPvOzv8A9lqzn3C60eezk7XOhX/nRZ9fIuAGA9g9XND+AnxY8JgTeA/i3Nbxr/y6X0EkacfwlCWX8q8luvi3+0NoqGS71jVkjH/LVrVJU/76CkVBD+058YoGAbxFBKR1EtpH/QVzfW8BB83JKL8tPyZ7K4e4sxFPkjiaNWHaT5190oy/NH03pmvftD6Aoh8S+DdC8TwLx9o0m9FvOw9drjbn24rX0n47+Cp9YXRPEY1Dwrq2dv2TXIDBk/7Mn3WHvmvJfAv7X9rMI7L4gaMbeThTqGnruT6tGeR+Fe5pJ8M/i74ZKBtI8RWTryjbXePPt95TXsYbExrR/wBnq38n/V/zPzbOsmrZdVazjA+zj/PS0XrrzR+Xus7OG5guLZJ7eaOaJxlZI2DKw9iKlyOK8MHwb8XeANQOpfCLxbLHa5y+gaw7TWzj0Vuq+xrsPDPxQt7vU08P+MdKm8L+ID8q216f3NyfWGX7r/Tr7V2wxGvLUXK/w+8+bxWUR5XVwFVVYddLSXrH9VdeZ6LRTQwNAb2rpPEHUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFNLCgBTTZJYoo2eSRURerMcAVlXeuxiVrbTovtlwv39rYjj/336D6da56e4e+ZppZEv8AYeZZD5dlCfb/AJ6GpcjSNJvc2rjXjLE7aaieSv3ry4PlxL9O7fhWO0pnX7bK4uFB4vtQ+SBf+ucfVvamqr3O27YicKflvL5NsKe0MPf2Jqx5cULrd3DkyHhbq9G6Rv8ArlEOlS22bKKjt/X9f0iMo8+26kBlx0vNSG1B/wBc4R19s1M0UY23V2TMR9241A7VH/XOIf4UryOk/mNugkfhZLgebcv/ALsY4WmFhDc5OYrh+AZB9oum+i9EpDZLJJviE0ibox0mvz5cf/AIxyaSTMsIkuC8sQ6SXTeRAPog5b/PSq7zrDd4P7q5b1/0q7P4dEpyw3D3IkKLbyno8/8ApVz+Cj5VphYm80vbnHmy24/uYtLcfieW/wA9aihuGf5LEu47rpsWxPxlf+lJMLG3uVN5J51yfu/amNxLz/diXgU+a6uXA3wCIfwnUJMc+0Kc/nQAsdtKZMqLWCTOTsU3c34sflBouEsYm/4mE3mvngXs5dvwiShop5Iv9JnuHh9HYWkOPTaPmP40+3t0iizbJ5aYwTaxCJfxlfk0C8xVunSDbbWtwIs4DELaRfX+8fw5psbXDsDFJCpzz9jgMrf99vx+IoVoWctEI5XI+9ErXLfi5+UfUU+V3kYrIFJByBcSlz/3wnAoAY9vG7AXUhkcdrq4Mh/74j4qVI/s8eUV4V/6Zxpbj8zyaSRxbp++uGhQnplbVT+HLGoY5IHBe3tmm5xvjgL/APj8mBQFmyXMEjHAjmYdNoe4P5nAqTzJEwP3ka/7UiQD8hk1Wa6nkygMWScbWmaU/wDfMYwD+NDeZFkyTvCCMAhY7bPtk5agLE/llxuCRknuI3lI/FuKa08cRw92Yz/dEiRj8lBNVwsMzA7Tcbf4mElxj8ThaX7RHB8iskP95PMjhP1wgLUDsTb4ZBuSCWf/AIBJJ+rYFBk8tgPs6R/7TGKPH6k1CJGmO8RtN/dYRSSfq5ApymdTiNCmfvDzI4v0QE/rQKxYWW4kYqJEOO4aR/8A0FRSFZSfnJT6wj+btUZhuJRhyrJ6sZZD+pAphjgU4e4gUjoqpEpH57jQFiTMRPNxs+kkS/yBqRGUjaJJW998h/8AQVFRq8bYxcXLH0R3x/46opXjRj89tcP7vuH/AKEwoAeVVT/q5m+okP8ANhQfI7xxj/eUf1emiJNvyWMZ9Q2z/wCypyIR0tIE+hX+iCgBhe2HG2H/AMh//FU4S239yD846m2XWPlCKPq/9KN9wvWWIf8AAm/xoEMEtvj/AFUX/jn+FHm2/wDzzi/JKlEr954/w5/9mpdzn/l4X/vn/wCyoEQ74O6RD6lP8acGtz/DGPoy/wBGqUO/acZ/3f8A7KlLSf8APYH8P/r0AR/uOMc/7rn/AOKpwVD0aRfxP+Bpr+YR8xjP+8T/AI0qqo/gtvwH/wBagBxRccSP+IH9QKNnpImf93/A01kzyIVP0fH9KdglMFH4/uv/APXoAPKk28BfwJ/qKMSggbW/Ag/1oCgDBEij2UH+lKH42mQj2INAClmBIbI/3hj+lIGRgBtDfTmnCRtvDJkdMHFKRk/PEr+vAOP500JiAgZwSv408OxHVXHoR1phMY+8GT8cDH40xpbcdLlT29f5UXFYnyvdGU+q805W/uuD7HiqZu0X7uWz/d/+vimm+GceQfx4ouOzNISFT8wIPrTsxv2H8jWYLx+icUv2mUjkIfdeDRcXKzS8sg5RvwPWnLPKhwwyPes9LiU8CUj2NSCWU/f3Y9qLi5TUSaN++D71LxWOAW6MRUgeZO5Ye9VcnlNSiqaTK3BJB+lS88HnmmTaxPRUIJHenrnvQIfRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRTDxz0rz/AOIXxj8CfDYi38Q6kxv3j8xLC2QyTMvrgcKPckVFSpCnHmm7I6sHgsRjaqoYWm5zeySuz0IkDvXjnxR/aH8H/D2OXT7WRNZ1wAgWVu42xH/po3Rfp1r54+J37T/ijxjE+leFYZfD2ltlXdHzdTDpgsPuj2Xn3rG+GHwA8T/EXUWuNSuotG09GDzvctuuXB5yIs7hnn5mwD714GIzidaXssFG77n6vlPh1h8uo/2jxPVVOC15E9X62/Ja+aOa8X/EHx78W/EiQ6hNc3jSPi30uyVjGnptQdT7mu18Mfs+amjxXPjK7ktZWAKaRpqC4vGz2b+GMfU19BXFj8HP2ePB/mylEvJFIGCJL68PoO4X6YUVxuj/ALQF3ZeH7vxRd+FotJ06aQrZJORHvUf3SfmlcnGW4Ue5rkWApQnzYud57vy/r+kfRT4rx+Jw3suHcL7LDp8sXonJ/wB2/wCNk35pmjpPwbubbTC9zcWvgnSNuJHjl82+mX/bnb7ufRcfjWTefEv4MfCZJLfwhYx6tqq5DXsubiVm9dzHj8xXh3if4q3vjjxRd6t4s1HVpbN2/c6Rp85iiAHQFm6D1wMnNecTyCe6eVIEhV24jjztXngDPNc9bNIU/wDd4r1f6LoezlvA2IxjvnVWTWjcY6Rb7OTfNJru9Ox614t/aG8beIrsvavHZLk7HcCV1/3QflX8ia8z1HXde1+6Lapql/qMrHpNK0n4Af0Fav8AwgHiWDRE1fVLaHSLGQZSbUpRBuB/uqfmY+wGaS08Ya54dAt/D2siFU48+3tUjY/Riu78eK82rVqzd68nb+ul0faYDA5dhYcmU0oNrT09ZJS+419L8PeI7vQ4rbSvh8gc8tqV8h3Nn+6XIVR6cVlap4SXR4TNrHiHSFnJ5tLSX7RL+IX5V/E1i3+tavqs7Tanqt9eSPnc087OTntyelUPoAPpWM6tN6Jf18j0cNgMWpc06iinq0k2/vk3p6JG3BpNhcn9zq1jD/tXc5TP4Kp/nWsmnro6ia21rwizjkFJXmkH4lf0Ncdk+tGT61Eaqj9n8WdVbL51dJVLrtZHW6l478X39iNNuPFtzJaoflhiJjjHuNoFa+kfE74k6NYkWHiiK4tEHzwT+VKuPQq4yfoK87oq1iqifNzO/qY1ciwVSn7J0oOPZxVv+HPbrD4zeC9UiEfi7wA1nOR8+oeHbtrVyfXy87a1YdM+F3itw2ifESwErdNP8XWIjJ9hMmD+OTXz5Rn5dvauhZlN6VIqX4fkeJV4Kw8fewVWdJ+T5l90r2+TR9HXPwI02eE3F34f1ezhIyNQ8N3K6pbf7xjOJAPYZrJg+B3irTLl9V+G3j2yvbyH5hbRSPZXgx2MbY59q8k8M6/4g0bV4jonie50RieJkuHjjB9wuR+let2Xxl+KduiDVrTw/wCNIU4WZxHNIv0kjYMD+Ga7aVXCVdZQafdf8C35HzWPwPEOBfs6eJhUi/sz0uv+3r3+U16HQeGv2nvH/gvVv7B+JWgyXxhOyRnTyLpPcg/K9e9+H/iV8J/i7pP9nG8sLqSQfNpuoqEmU+wbv7qa+YvEfxw0nxbqEGlePvhtB/ZscIidYpWW8tmz9+KRwCBjHytxxXOah8KBqmmSeJfhRrg8TWEI8ySzT91qVmP9qLq2P7y12UswrU9ISVSPZ/F/wfxPm8fwhl2LjGrjKTwVaW0otOm36p2i/K8fJs+4LLQ9e8K4TRL+XVdLHTT76TdLCPSKU8kf7LfnXS2WoRXsfCyRSj70Mq7XT6j+vSvgbwX+0B8RPB0y6ZqOtXl7ZRNs8q7XzJYCO3zYLDsVJHsQa+kPB37SfhPVGtbbxYYtJln+W31JDvs5z6bzzE3qr4xXq4TNcPV0T5X2Z8HxFwDm+AfPKCqrfmhu15rr59euu57vRVe2ure9to7m1njmhkAZJYmDKw9QRwRUy1661Vz86aadmOooooAKKKKACiiigAooooAKKKKACiiigAoorjvFPxX+GvgnWk0jxd460DRL94hOttqF6kMhjJIDBWOcEqRn2NAHY0V5n/w0P8C/+it+Dv8AwaRf40f8ND/Av/orfg7/AMGkX+NAHplFcb4Y+LPwz8aa2dH8JeO/D+tagImnNrYXqTSbFIBbapzgFl59xXZUAFFFFABRXjnx7/aH8M/AvRbL7dYy6vreobjaaXBIIyUXhpJHOdiAnHQkngdCRxPwM/bD8PfFrxung3WvDz+G9ZuVZrI/ahcQXRUZMYYqpV8AkAjBwec8EA+maK8/+L3jLxz4H8Dw6z4C+H0/ja+Nysc1hBOYnSIqSZBhWLYIAwB3r5Mn/wCChHiS1upba5+ElhDNE5jkjfWJAyMDggjyOCCMUAfedFeH/s3fHu9+PGga/qN74Yt9DbSrqK3VIbtrgS749+SSi4x0717hQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFcP8WPip4Z+D/w4ufF/iZ5XiVxDbWkABlupmztjTPGeCSTwACe1AHcUV8eeAP29fD/iPx9aaH4s8GyeHrC9mWCLUo74XCQsxAUzKUXauTywJx16Zr7DoAKKKKACiiigAooooAKKK4XXvjR8JvC/iG50LxF8RvDWl6nbELPZ3eoRxyxEqGAZScjIIP40Ad1RXmf/AA0P8C/+it+Dv/BpF/jWloPxp+EvijxDbaF4e+I/hnU9TuSRBZ2moRySykAsQqg5PAJ/CgDuqKKKACiiuR8VfFP4ceBtVh0zxh440DQ7yaLz47fUL1IXdMkbgrHOMgjPtQB11FeZ/wDDQ/wL/wCit+Dv/BpF/jTk/aE+B0kionxZ8HszEKANUiySTgDrQB6VRR1GRRQAUUUUAFFFFABRRRQAUV8yfG79pP4lfB7xffxSfBhr7wvFLHFa+IZNQeOGcugODtiYI24soBOTtryM/wDBRDXgpP8AwqnTeBn/AJDUn/xigD73orO0HUm1nwrpurvCIWvLWK5MYbcELoGxnvjOM1o0AFFFFABRRWP4q8T6L4M8Gal4p8Q3i2mmadA1xcTHnCjsB3JOAB3JAoA2KK+Kbf8A4KFaK3jEQ3Pw5v4tAMm03a3qtdKmfvmHbtJxztD57Zr7I0bWNN8QeHrLXNHu47vT76BLm3uIzlZI3AZWH4GgC9RRRQAUVjeLNS1nR/A+q6p4d0Q63qttbPLaaaJfK+1SAZWPdg4z0zivjXxB+3X498J6/Lofif4IRaRqUIDSWl7qsscigjIODB0I5B70AfcdFfLHwD/a41T4zfGBfBV34Fs9HjaxmvPtUOovOcxlBt2mJeu/rntX1PQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABR3FNLKqEscAc5Nc9f+Ig8ci6Y8QijOJb6b/VR+w/vt7Ck2kVGLlsa1/qdppqBriT5m4SNBudz6AdTXOX2pXd5P9lnWVCwyum2rZmcf9NX6IPas7zWVxN5lxEZ+Bcuu+7uc9ok/wCWa+9TRWeUa2aLyk+89nBJ8x/2rib+lQ5XOiNOMPUYFNwhtxHFcJFybWBtlpB7yP8A8tD7VY8qMxrc3UscwThZZ02W8ftFF/FS+YhgUxiBoYukjLstYj6IvWRqa8uyVJ5HkErcJNOm6Z/+uUQ+6Pc1JW5LJK4dZmZomb7s9wm+Z/8ArnEPuj3pofyrkgGSO4fqF/fXb/U9Ix/Kqc99FayMjyNBK/3o0kU3D/8AXSU/LGPYVmT+IvD1pH5N7rdnGhPzWdhKSD/10kGWf8KidaEfiaRtTw1WppCLform0JiJnt7dW80/6yCybfKf+us54X6D0p8VtJ5ThpFghH34rJtq/wDbSduT+FchqHxO8K6fGtrZN9oQHpHH5UKf8BJDOfrWHe/FvQQ6lre41GRfum4ZUiT/AHYlGPxJNc7x+HWnOd1PJcdUWlJr10/M9LguYEhZNMthJGPvtB+5hH+9M3LfhSqtxPCSZm8nulr/AKPD/wAClb5m/CvJ7j42wHHkaMJ5APle4cuF/wB1AAoqjJ8bdXY7o9ItpJezyqXK/QbsD8qh5jR6X+5nRHh7Gv7KXrJHtkEMcMDfZwI4s8m2HkofrK3LfhToZAAXtF4zktbJtH4zPz+IrwKX4w+KJZC32WCSXszxhyv0ByBUR+LfjaR9zvDKR2eKNgPw24pf2lDpFl/6tYnrKP3/APAPoRFZ2LR7FYdWhXznH1kf5R+FQNc2j3G3IupwRwM3Tj8vkWvnub4o+MriQG4min29I3RWQf8AAcYof4reNCBG1xbogPESxoq/98gUf2jH+V/cP/Vuv/PH7/8AgH0HJdzyHymRUJHCTyb2/CGP+RNMleSNNtxdSRg4+Uutqp/4CuXI+teAn4reNfL2LPDCmMbY4o48/XAp6fFbxbCoKLBF6tHCisfxAz+tH9ow6xYf6t4jpOP3nvSCKH97FB5eTzIsQjz9ZJTk/UCkE/2plKKLk8ncqPcn/vpsIDXiUPxf1qPDtoVhI46zSRu7n8WY1pf8LwvmUC40WJvVmZm/IH5aazCl1v8AcZy4exi2SfzR6+ounAQkgdSsk+P/AByIf1pPIW35luI4O52RpFn8W3NmvMIfjXpMm1bzTL9gTyFuAqj6KoGfpmtq0+LvgsyqkcVxbk9X8lFA+pBLVrHHUH9o555LjY/8u2/Sz/U7fZbTDJinu++XEkoP4sVWp4/NXCwW6R+3mBePpGCf1rnIfiH4Su5QkWuW7N6SIyj/AL6krSg1uyvQFtL6C4z/AApcbyP+ArgVtGtTl8MkziqYSvT+ODXqjTaOY/NNNEo7fu8kfjIf6UKiMvNzPIO+xzj8kUD9aqg3GeE8v3Cqn+JprLIQDLOOP7zFs/mcVpcw5S0yWinc8CEjo0mD/wChMacLmJDhCi4/uE/+yiqW2IHPmMfUIoGP8/WnBULcRO3+8SaY+VdSyb4Y+Z94/wB05/VjUYvtp+QHHoNq/wAhTRFL/Dbqv+9TxFOTyyKPbn+VArREN5MT8iN+LMf5Yo864cZ8kH3IJ/macIjgh7jA9aBBHn5pWb8AR+lAaEYeYHlIx74UU5pLnHEqj6YP8qk+zxBvuy/gTj+VP8mLHEGf+BEf1pWYXRWEkpPzTr/n8KfmTPM34gn/AAqYRwoOY4gf9og0FrcHlbdfpigLkBkYHmdh+NHnZ/5bvU4ltl+68P4DNL9ogA+8n/fFArkAm/6bH86cJTjiUf5/Cll1KyhGZbiNR6FapS+I7BeIkaU5xkR4FF0tylFvoXBLJniZPxFPEzgcuv8AKsVtdmkU7Yo4x/1zz/Oq73ay/wCum3f8B/w4pcxSpPqbr6ikZ+aeL6A5NRtrgH+rjd/xrEE9kOQgbPquP/r1J9otsACOJffn+v8AhS5h+yS3RpnWbh/uQKv1G6mm7mf7xZR7fLWeLiA9JcH2Y/4U8NEx+WRjn1IP86LhyIvCRTy0pz7tmpAYiQQ4JH+z/WqAU7htckVJ5b5yGjP4kU7i5S+HQDKq3/AWxUgmA7f99VmqG6bl+gIqRZSrcyH6CncnlNASI3RI/wDgNPDEdHwfQiqAnTpvBP1NSrcR8ckD160E2Lu98c/N9BT0lOOGI9sVUWWL+8AfyqYSIQMkH60xNFpXPufpxUiyMDwxP1FVAYiMg4P51ICPVT9DTuTYtgq3XIP0qRWkTo2R71VVx0B/A1MrAD0ppmbiWkmVjycGp15IqkMMefzqxDkSAZyOadyWixRRRTJCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigBCcUm72pTXEfEL4neH/h7oF3qOqTh3gT5YFI3O5HyoPc9fYZNROpGnHmm7I6MJhK2LqxoUIuUnskZ/xg+K+lfC7wU99MFuNVuAUsbLdzI/8AePog7n8K+ALm58S/ELx808zy6lrWqT/99E9h2VQPwAFWvGnjHxD8SvHk2taq0lxdXLiO3to8sI1z8saD/OTX1f8AAr4FW3hzQG1bxLbrLqd5HtnU8+Wh/wCWAPp/fPfp0zXyVWdTNq/JHSEf6+8/obA4TBeHuVPEYm0sXUX9Jf3V17v5Hmfw/wDhHqmoX4tfBVjCDCdl5401GPzIlcffXT4iMNggjzTnOMgivRr0aX8K7e+8NfC21Oq+L513ax4j1F/N+zdy00h43c5CdB1NdR8UPiU+iWH/AAjHgqErKI2E13aqAlpEow2w/dBHQt0XOBk9Pk7xx8S59U0iPwp4bee00NTvnGcSXsp5ZnI5YZJxknPU10YipQwMXGHxf1t/nueNlGEzTimtGviVak3eze67z6tfyw0jfolvmeLdb019ekuHv7nxRrDMTdavfOTEz5+7FH/cHYn8AK5XUdV1DVroXGo3ctxIqhE3tkIo6Ko6KPYVUIIOGBBHrUsEMciyPLMsaoucdWc9gB618xUqyqSZ+7YTL6GDpRSV3Hr/AJdEvJWRNp1kt9qMUEt3DaRMwD3E33UB746mvWoovg14Lsje6L421bVPECrmKePS0ljgbH8G/C5z/EQSO1eM5x0pckjmqoV1SvaKb7nNmeUyzBxUq0owW8Y2tL1undeRc1bVtT1rVJNQ1bUbq/uZD8091IZHP4np+FU80fWkrCUm3dnrUaMKUVCCsl20CiiipNgooooAKKKKACiiigApc/Nuxz696SinclxT3NWXXbu+tkg1Yi+EYxHNL/rkHYCTqQPRsiorDWtV0rV4dV0m+msL6E7o7i0bynX8VxWfRV+0le99Tm+pUOVw5Fyvp089Nj6Y8Jal8Mvj1pMGgePVGj+OlTZHrUAWI3x5wTgbWbGMq3J7GuP8YfBPx18Mb+WS5h/tjwxMwFzd2kJmUR56yRZyrAdCDx2NeMo7JIJI3ZGU5DKcEH2NfUHwY/aZks0g8K/Emc3FowEUGrP8zJ22zeo/2vzr2cPiKGKtHEe7PpL/ADPzPOMnzXIFKvk/73DvWVGWvL503uvTp0TNb4ReLJ/hl5NudZTXvhxqE2y21KJsvpczfwTKeUUk/QfnX1DHqdm96tosyea8fmouf9Yn95fUfSvnL4j/AAnv9BuJ/iD8JUiubK+TdqegqN9vdxtyXRRx78fUVj6V8S9G03Q08F+LZNS0l7adBY3qn/SNGdvukk9YhnhuQVOCO1e9h68sL+5qbdO3yfn57H5Nm+UUs/tmOBfNJ/GklzesorqurWkt1Z6P6vD57Yp9eZeAfiBqeoeI7rwR4yghtfEVnEJ4Z4R+41S3PAnh/TcvbNem169Koqi5on53jMHUwlX2VX1T3TT2afVBRRRWhyhRRRQAUUUUAFFFFABRRRQAV5v8QvgN8K/ilrCax408Lpf6lHALZLxLiWKRYwSQo2MB1ZjyO9ekUUAfjb8YvDWl+DPj34s8KaEk0em6bqL21skspkZUABALHk9a7D9l34d+GPil8fY/C3jC3uLnTf7PnuTHDcNCS6bcZZeccnisb9o3/k7H4gf9hiX+S16J+wz/AMnXx/8AYHu//ZKAPu/wD8BfhR8MtdOt+DPCUOn6mYGtmvDPLLIY2Klly7EclV7dq9IoooAKKKiubiK0tJbmdtsUSGR2xnAAyT+QoA+M/wBt/wCCvjbxlrOifEHwhpN1rUdjZHT72xs0Mk8a+YXWVUHLr8zAheRgHGM48j/ZV+A3xD1P4+aF4w1fw3qeiaHoVz9tlutQt2t/PkVWCRRqwBYknkjgAHJycV9a/wDDY/7PR6eOWP0065/+IrX8L/tQfBTxl4x0/wAL+HvFrXeqahL5NtAbGdN7YLY3MgA4U9TQB7Bj5QDzX4rePP8Akqvif/sLXf8A6Oev2pzkA1+K3jv/AJKr4n/7C13/AOjnoA+2v+CeP/IjeOv+wnbf+iDX2hXwX+xd8QPBPwy+DfjHXvHHiWw0e2udXijgWd8yTFLcZCRrlm69hXql7+3n8E7W9MNvY+Lb6IHH2iDT0RD74kkVv0oA+oKK8h+HH7TPwf8AihqcWk+H/En2XVpeI9N1SI2s0h9Ez8rn2Via9ezgZoAKK8g8V/tPfBbwV4zv/CviPxcbTVbBxHc24sp5NjFQwGVQg8EdKxj+2R+z0FJPjl8AZ/5B1z/8RQB7xRVHTdY07VvDlpr1jcq+n3Vut3FO3yAxMoYMc9ODnmvC/G37ZvwP8F6rPpq6tf8AiK5gJEo0K3E8akdR5rMqE8dmNAH0FRVPSdRg1jQLHV7ZZEgvLeO5jWQAMFdQwBwSM4PrXC/Eb47fC34VSrbeMvFVva3zLvXT4Fa4uSMZBMaAlQexOBQB6NRXy5L+3t8FEvPJTTvF0qYz5q2EQH5GUN+levfDT45/DL4tJIngzxHFcXkS75dPuUMFyi/3vLbBK/7S5HvQB6LRWL4t8V6F4H8F6h4r8S3v2PSdPi865uNjPsXIGdqgk8kdBXkP/DY/7Pf/AEPLf+C65/8AiKAPeKK8s8CftFfCP4k+ME8L+D/FBv8AVHhe4WA2c0WUTG47nUDjI711vjj4heDPhx4dOt+NfENnpFmSVRp2+eVh/DGgyzt7AGgDpq+ev2wPhN4k+KnwTtI/CVubzVdGvhfpYhgGuU8tkdUzxvAbIB64I71t/DD9qD4c/F34kzeDPCNrrxuYrOS9+1XlqkMLIjIpA+cvnLjqo71L4i/ao+B3hbxbqXhrXPGBttT024a2uoBYzv5cinBGVQg/UUAfnd4A/Z6+Kvjrx/aeHl8Ga3pUBmVb2/1Kykt4bSPI3MzOBlsZwo5J/E1+ucMaw26RLnaihRuOTgDHNeGJ+2J+z7JKsaeN3LMQoH9nXPU8f3K92ByMigAorh/iP8Xvh98JrOyuPHmvrpYvjILVPIkmebZt37VRSeNy/nXjl9+3f8DLN9sI8T3g7NDpwQH/AL+Op/SgD6bor5ds/wBvb4J3MoWaw8W2iE/6yWwjdfr8krH9K9w+HvxW8AfFLR5NR8D+JLXVEhwJ4VzHNAT03xMAy/UjBoA7KiiigAryXxz+zV8G/iH4jvPEPiTwn5mr3hUz31vdzQyOVUKCdrbeAAOnavWqKAPxI1mwNp4l1i0s45TbWV3NEDktsRZSilj+Qz60vhzxBqfhbxfpfiXR5miv9Muo7y3bP8aMGAPscYPsTXvf7Pngqx+Iv7QPxJ8E6iq+VqmjapArt/yyk+1xmOQe6uFb8K+fNX0q/wBC1++0TVbdre/sbiS1uYmGCkiMVYfmDQB+z3gfxbp3jv4caL4w0lw1nqlpHdIP7m4fMh91bKn3FdBXxX+wL8S/tnhzWvhXqFxmawY6npyu3WFyBKg/3XIbH/TQ+9falAEVzcQ2lnLdXMqxQxIZJJHOAigZJJ9ABX4+/Gj4g3vxY+N3iPxuolaxeYRWgwcQ2iHy4c+mfvfVzX3v+2l8TT4G/Z9l8O6fc+Vq3ih20+PacMtuADO//fJCf9tK+TZvhkfCv/BOm48d39vs1HxPr1m8O4fMlnGZBGPYM29/cFaAPnjc394/nX6OfAf9mH4La38EfCXjDXPCbanqt/YQ3k0l1eTFfMPOQisABkdK/OKv18/Z1/5NU8A/9gaD/wBBoA9OAwMCiivF/Gv7VXwU8B+JL7w9rXia4l1WxlMFzZ2VjNM0cgAJUkLtzyO9AHtFFfLVx+318E4ZCkWm+Lrgg4Hl2UK5/wC+phWtof7cHwJ1i9jtru/1vRN5x5upaeRGv+80bOAPc0AfR9FU9K1XTNc0a21bR7+2v7C5QSQXNtIJI5VPQqw4Iq5QAUVyHj74o+Avhhoy6n448S2ekxSHEUchLzTH0SNcs34DivD779vP4I2tx5dvaeK72POPOh09EX64kkU/pQBb/bm/5NLu/wDsLWH/AKNr8yG/1bfQ19s/tH/tMfCr4wfs3X3h7wxe6nBq/wButbhLO/smjLqkmWw6lk4HON2a+Jm/1bfQ0AftZ4H/AOSY+Hf+wXa/+iVrfrm/CFzDZ/CXQrq4bZFFpFvI7YzhRApJ/IV5WP2x/wBnojI8csR6jTrk/wDslAHvFFeM+H/2qvgd4n8Vad4d0Xxi1xqOoXC21tCbG4TfIxwBkoAPqa9T1/xFoXhXw/ca54j1ez0vTrdd0t1dyiONfxPf0HU0AadedfHXwDqHxN/Z+8S+DNJmji1C9twbUyttRpUdZFVj2BK4z2zXD6F+2D8IvFPxV0rwH4b/ALd1K71O6+yQXsdkI7bdgkEmR1bbweQprq/Hv7Q/wk+GfjA+F/GXic6fqggS58gWk0v7tyQpyikc7T+VAH5kQfAv4xXPjEeGIvht4jXUjL5RSSzdYlOcbjMR5YT/AG923Hev1Z+FXg6b4f8AwW8M+DLq5W5uNLsI7eaVM7WkAy23PO3cTj2xXm//AA2N+z1/0PD/APguuf8A4ivZPDniHSfFnhPT/EmhXJudN1CBbm2mKMm9GGQcMAR+NAGpRRXi/jv9qv4KfD/WJtI1TxQ2oajCdstrpEDXZjP91nX5FPsWzQB7RX5e/tuf8nfap/2DrL/0Wa+nof29/gnJdeVLpvi6Bc482SxiIx64WUnH4V8fftQePPC3xK/aEufF3g+/e90y50+1QSSQvCyuisGUq4BBHH50Add+wx/ydnF/2Bbz+cVfpvX5kfsMf8nZxf8AYFvP5xV+m9ABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRSdutAATjtVPUNUtNNt/Nun254VByzn0UdzWfqniBLeR7SxVJ7pBl2dtsUA9ZG7fTrXJzX2M38l4435H2+RQJJPVbdD91f9s1EpqJvToOW5o6jqN3qFwba5ifpuGnRPggf3p3/hH+zVRd05WYSQyiP5VuXTFtB/swR/xt7+uK5nUvGHh7SbUpM6ToCW+xQSZQt/eml/jb2Ga5a9+JXiC9ZTpVqIiw2RyImCo9ELdB/uj8a8+pmFKL5U+Z9lqe7h8lxNWPNblj3en/BPVJprPSUaW5maGSQZZpZALiYf7THiNfauc1bx94d06IRPMl5tOVtbb5bdT6sxx5h/SvLpbbULyd5tV1BmlJLOA/3Cf78jZIPsOaS2sLBGBtohNKxx9omUsST02Ick89zXPLE4mfwxUV97+49CGV4OjrVm5vyVl9+50978Wr2WRf7L0zzbjOEldDLtHbYmAqfr9ayZNb+IN68hFy9q8nLu0wRh/vFeQPbNOQyQSmOGErIRlwWHmFenzt/CM9hT/sRmXNzcecucBUGyEN6jPMhqHQnP+JNv52OmNejR/gUYr1V3+JgNo9/cbml1qNgxw8safu8+m4/eP0qdPDOllFWa9vp8/eZn27j6LGOcfWt8i0gKhlHmEfIZfnf6hOiD60g1GFEYq5RG6rDxz6tJ/QU44SjH7I55liZaRk0vLT8jMi8MaRny4dGkklH3g5Z3x6noFq0mg6eiH/iW2MMA6kMvX/ac5/IVK15cPbgwWLeUxxumyqMfZPvP+PFIY75pQswhecDKKw8xkHsi/Kn1NbKEFsjldarL4pfiILKzEYRLWyWLPDCPYh/4Fjc34VI1hb+UFupFCnlUEXlIfov32+tIiX0xZ0umeVRhpIxvI+sjfKv0FOjscKZptQaI93VwmfrIxJP0UVXKuxnzvrIc8FvDGsTDy0PRZFWPP0jXLH8akNlDIgQQSY7CZliX8I0GT+JpILWzKnyklcn/AJaPKY1PvnBZvwqT7Bawn9/OQp/gjLxg/oWNVbyJcrdRps7eBRFcDAPIjL+WD/wBRuP4mnrbwBNiWccYPZ38sH/gIBY/jT0h02E7VtCN3qHTP4DLH9KUxWSvg2mCe3zJ+gJY/jRZEOb7kZ0qyPE8EQB6KmIwf0LGkbQdBRR5mmLHnkEkIPzILH8BUot7QMQto6Z7DdHn34yx/GkWztN3Muwn+EbkJ/m1Fl2Dnl/MU5fD2iSD5bTA7EYTP4tz+lVpPCWjuuY/tCHt1Yf99MQK2DaWMR4jdX/v+Zgn8wTT1tG/1jXkh9PnGfpzk/pS5E+hSrTW0jmpPAtqy/LetG392RS2f0AqnL4DcnZBdwSN/dHf8q7L7JKnzvdFx28xcke3JJ/QUvl3qgZMT7vus4K4/wC+v8Kl0YPoarHVltI8+k8D6ogLRWwdQesb/wCFUpdA1q0ODDdRH0wenrXp7SXMfE0KLj75V87vcDj+RpP7SjRBkXOT93K4GP0FZywlN6s6IZriI+Z5tba34r0tVittZvIUU5WPzGVR+HSt2x+KXjGyYeeLe9Uf89Yhn81wa6gz2dypH2W1KngFhyD+HH5mqU2h6Vd8m2Cknb+47H14zj86lYZx+CTRcswp1f49JP5Iuad8boQwXVPD5QEcvbODk/7rD+tdZpvxQ8IaiqxtqxtGx925iYAe2c4rzefwPBMG+zyuvPAmxj9c1lXfgK8iG5PJdcZyp7Vop4mHW5jLDZZW2Ti/J/53/M+gbbV9NvoBPYXi3a9nt2DD/voVM10Cy7bfce5d+f0zXzGdD1bT5jJbm7jkHG9M5/Na1bHx1410tgDfC8jXol2olx+fP61axsl/Ej9xzzyGnLWhVT9dPyv+h9Cm8cTHaIUPcgfN/SgXVw7lg7n3jT/6xryrTvjWikJrGh+R0BltT09Ttb/Gu00rx/4V1jAg12OORjgR3H7lvyPH61vDFUp6KR59fKsVQ1lDTutfyOgUXTHJEw9922k8mVmy4QH1Z81Bd6pptnD5k9yqg87txO76DNYN340s0JFnaB+Pvy4UfXvW7ml1OKFKc9kdIYum6SHHsM1BPe6faD95fKD/AHUXJ/IVwl54surnImvdqf3IflA/HisiTxDAvIZSfU/Mf8KzdZHVDAzZ382vxni1t7iQ/wB58KKz59Uu5OJLlkU/wxsBXCTeKoRkNMPbLf4VCviJ532wI8h9I1J/lmsnXTOhYGS3R25lhyTsd88bnfrQLpSCFWL8CTXL28Xie9b/AEXQ9Qkz0P2cjP4tWrb+EfHV0cf2HNGf708yoKak3shSpQjvJfeaLXKAkt5eSP42xSC8gLDmJsdNuWp9t8OvGrnDjT7f3Muf5A1px/DDxCy/v9bs0/3Q7f4VaU+xjKpRX20ZqygrxAT9TgVIHAxmKBR6lsn+VbMfwquOPO8QLn/Yts/zarcfwuhT72vXR/3YUWqUJ9jKVej/ADfgznln+bJmjHoUBzUhm3NkyzP9OK6ZPhtYKPn1jUm+hVf6VMvw50gDnUNSP/bUf4U+SRn9Ypd/wOXWRQMhJz9WFSo+TnyZT9WzXSj4eaMP+XzUP+/w/wAKd/wr/Sh0vdQ/7+j/AAp8kiXiKXf8DBW4kC4MZA9+alW7kHHlJj6EVs/8IBp4+7qN8P8AgQ/wpv8AwgduPuatdj6gGq5ZE+2pd/wKC3znAMCHHTk/yqdbsE/NAR9DU58Dyj/V6zJ/wKMH+tRnwbqKn5NShb/eQj+tFpdhOdN9Ry3EZx/D9alR2P3QD9Carnwxrsf3Jrd/oxFRtpHiCI5+y7/9yQGi77C9zozSDkAbl/XNSLKvGUIrJzrNv/rLK5H0Xd/KnDVZFYCeFx/vJincnk7G2siEdD9CKkUr2P51kxapbOfvBT9aux3EbDKOM/WmmQ4tF5CKtW5zIPpVBXUnrVy1KmcYIPBqjKSL1FFFUZhRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABTD16AU+sPxT4m0nwj4Rv/ABHrdwILGyiMkjdz6KB3JOAB6mlKSirt6F0qU6s406avJuyXds5j4s/FHSfhn4ON9cSxvqV1mKxtmP33x99h12L1Pr0718CeNvG2r+NdbN1qFzLJErM0auerH7zt/tH9BgdqteP/ABvr3xS+I82tXkcjyzuILKxiBfykzhI1HckkZPcmve/gr+zop1+21/xfbiS3sCGFu3KzXPXb7pH0J7tnsK+OxFatmlb2VLSC/q7P6NyXLMs4Dy5Y3MXfEyV7df8ADH8r/oXf2cvgabKCDx34stSt1Iu+xtpBzCp/5aEf3j29B9a3vit8dns/F0Xw58AJDPfmQRXt4wykI6lFweTjqe3PNdL8fvixF8NvBH9maQ8Z1/UkMdog/wCWCdDKR6DoPfFfDcmpyWNncQwTM97eZ+2XZOWwTkxg+/8AEe/TpnPRjMVDAQWGw+/V/wBdWeVw5kuK4vxM88zWN4vSnH7Pr/hj+LPRfir8VX1xH8M+H5I49PUIl7dQDb9tZBwgx/yyU5wO5JPSvJYJ3t5fNjxvAwCf4fce9R0V81XxE60+eR+2ZVlGGyzDrDUY2XXzfd/15CkljknJJ5NJXZ6R8NNfv/Cs3ibUGttH0hIzJHcX8gjNx/1yQ8tz36fWuPkVFlZY5PMUHhwCM++DUTpTgk5Lc6sNmGGxMpQoTUnHR26Ptfa4zIPcUZHtXLeLPDFxqNu95pN1Pb3agkxrIQsn4eteQSalqsMjI19cqwJBBc8GvYwWURxcOaFTVbqx+c8UeIdbhzEqjicI3GXwyUlZ/hv3TPojIoJA7/rXzsNW1Qkf6fc/9/DXoHhKyvI3S6vGuLudxlYnlO1R6kd61xOR/V4c86hxZH4pvOMSsPRwjXd82i+5a+h6TRTYzIYwZQob0XtTq+fas7H67TlzRTCiiikWFFFFABRRRQAUUUUAHap7S3FzdLbmZImc4VpDhSewJ7Z9aiR2Rw6nBByDTp5fOmaQoiFjkhBgZ+namrGVSMnpEddWtxZ3klrdwPDPGdro4wQaiNbUGp22p2kena45BjXZb34BZ4R2R/78f6r244qhqGm3Wm3XkXKDkBkkQ7klU9GU9wfWtJQ05o7HNRxD5vZVlaX4PzX+W6/E+n/2YvjUtuYfhx4ovMIxxplzK3AJ/wCWJJ/T8vSrf7SHhrRo/GNnf6rCBbzx7ftMZxJCCdpR/wDYJ+6x+63B4bj5Mjd4pUeJ2R1YMrKcEEdCPeveNM+Mtr4x8K6XoHju0kvNQsm+yNOibjf2coCyxsOolUYdT0JT1r3MLmCrUPq1bdbN/l/kflee8IVMtzZZ3lyfLK/tIx0ab+0vXr9+1zvvhRNca9at4D1TU1HiPw1It54d1kj5jAw+RW7tGwO0jtnHUCvpHSdcS7ht4b2P7HfuGSS1c8rIv3gD3HOR6gg18t3PhjWfB3hy28U6Hm41/wADyje6/c1fR5fmjk98KSD6FT3r3L+17L4leArbUvD86x6i8S6npdx2Lr/CSO4OUdfevfwE5QXs5br8V/wNvuPyXinDUsVVjiqL/dSbTa+zLdu3SMk1O3m+qZ6aOtLWF4S8RQeJvCdrq8KmN3BSeE/ehlU7XQ+4YEVtF8AnHAr1YtSV0fAVacqU3Ca1Wg+ikDBgCD1paZAUUUUAFFFFABRRRQAUUUUAfkJ+0b/ydj8QP+wxL/Ja9E/YZ/5Ovj/7A93/AOyV53+0b/ydj8QP+wxL/Ja9E/YZ/wCTr4/+wPd/+yUAfpvRRRQAUjKrKVYAgjBBGc0tFAH5w/t3eGvD3hv4u+Gv+Ef0PTtKF3pcktwLG3WESuJiAzBQATjvXlf7Mv8Ayd14B/7Cf/tKSvaP+ChH/JXfCH/YHl/9HmvF/wBmX/k7rwD/ANhP/wBpSUAfrgv+rX6CvxX8d/8AJVfE/wD2Frv/ANHPX7UL/q1+gr8V/Hf/ACVXxP8A9ha7/wDRz0AfUX7FHwg+HPxH0PxXq3jjwvaa5cWF5Bb2wu2Zo40aLe3yAgEk9zXv3xT/AGQvhT4q+H9/D4R8K2PhzxBHCz2N5p6mNWkAyqSLnDIx4PcZyDXnP/BPH/kRvHX/AGE7b/0Qa+0KAPw6IuLO8wfMguIJOdrFWjdT2I6EEdR6V+qn7KXxUvfip+z5Z3mt3JuNc0mZtMvpmPzTFACkp92Qrn1IY1+anxVtI9P+PPjawiUKkGvX0SqOwFw4r66/4J4X0n2Xx9pu4+WJLK4C+hKyIT+Sj8qAPYf2rvAPgm7/AGb/ABp4tuPCmjvrttZiaLU/sqC4V96KG8wAMTjjkmvy2bkEeuRX62ftS/8AJnvj3/sHD/0Ylfkp/H+NAH07rfxP+Lfxy+Gnhr4RfCjw1rU2i6ZpdrZ6rLaJt+2zpGqsJZchY4QRwpYburcYFebfFD9nP4ofCTwNa+JvGWnWEVjdS/Zf9DuvPaCQoWUSYXAztIBBPIr9Hv2aLe3t/wBkrwD5EEcXmaPDK4RQu5yMljjqSepruvGPg3w14/8ABt54V8XaVFqek3YXzraQsobawZSGUggggHINAHlfxL+KbfCH9jew8VWqxyao+l2dnp0cn3TcSRKFYjuFAZyO+3Hevy6km1nxd4yEt/fT3+r6tdqsl1cuXeaWRwoZievLCvtz/goBJFpPw4+HnhixXybJbqd0hBJwIYUjQZPXAlIr4j8Nar/YXjXRtc+yG7/s+/t7z7OG2+d5Uqvszg4ztxnB69DQB+lfjz9mH4WWP7LmseHdJ8LadDqWn6VJcW+sCEfamuYoy3mNJ1O4qcr0wxGBxX5p+GPEus+FPFOm+KfDl7JZ6nYypc20yHGGHOD6qehHQgkV9la1+3re6x4c1DST8GLuIXdtLb+Z/arNs3oVzj7MM4z0r4oj0+9jiRPsV0dqhc+Q/YY9KAP2Z8Ia3o/xP+DujeILjT7a607XdOiuZLO5iEkZDqCyMrZBAORg+lfnj+2z4R8MeDf2hNNsPCmhWGjWtzocV1NBYwiJHlM86lyo4zhVHHoK+z/2SZZpf2NvBXno6OkNxFtcEEBbqVRwfYCvkn9v7/k5XRf+xbg/9KbigDyD4EfEy2+Efxhj8b3VhNf/AGfT7q3itIzjzpZEAjUnsu4DJ5OOgJr0C5+EH7TP7RviubxxrmhTxC4H7i41mX7DbwxdVjgjbLhB7Lz1JJrB/ZHt7e6/bB8IxXMEUyA3LhZEDAMLdyDg9weQe1fq4OlAH54/sf8AhHXPAf7b2veEPEdssGp6bo11DOqNuQnzLchlbA3KQQQccg16P+3t4Y8OWfwl0jxHaaDptvrE+tpFNqENuiTyoYZCVZwMsMqDznpX1Gnw/wDCEfxTk+I8eixJ4mksv7Ok1BWYM8GQdpXO0/dHOM4AGa+cv2//APk37Qf+w/H/AOiJaAPzus/+Qlb/APXVP/QhX7hL9xfpX4eWf/ISt/8Arqn/AKEK/cNfuL9KAPlP9sX4aeKPid4t+FmieH9Hv7u2fULmC+vLeEvHZRyG3BkkYcKNoc5P9017H4T/AGfvg34MsFttF+HuhlgObi9tlu5mPqXl3HNJ8VPjz8Nfg9FCnjDWmW/nXfDplnGZ7mRM437B91cg/MxA4OM14Vef8FBfAS3Bi0zwD4ovOcIZHgi3/QB2P6UAea/tz/Cbwj4L1Tw14x8KaTa6QdWkms720s4hFC7ogdJAo4U43A4AzgH6+C/AfxvqngD9oXwtrum3MkaSX0VleRq2Fnt5XCOjDuMHcM9CoNelftM/tE2nxu8M+H9OtvBOr+Hzpt7JcmS+cMsoaLZtXCjnvXhXgz/kpPh3/sKWv/o5aAP2vHSikHT8aWgAooooA/O39jX/AJPZ8V/9eOo/+lcdUf25/hmfC/xntfHthb7dO8Sx/vyo4W8iADf99ptb3Kt71e/Y0/5PZ8Vf9eOo/wDpXHX11+0f8Mx8U/2etc0C2hD6rbJ/aGmnv9oiBYKP94bk/wCBUAfmJ8IfiBc/C/41aB41gL+VZXIF3Gv/AC1tn+WVf++SSPcCv2Ms7y21DTbe/sp0ntriNZYZUOVkRgCrA+hBBr8PCCDhlKnurDBHsRX2t8PP2lxoP/BPjWLCW/8A+Kr0QjQdODHLuswPkS/9s08z/v0v94UAcL8XdRvf2lv27LbwZolw0mkW90NGt5Y/mVLeIlrq4GOxIkwe4CetfRX7bGmWWjfsbWekabAsFnZ6pY28ESjhERXVQPoAK4j9gX4Y+TpetfFfVICZLonS9MaQZPlqQZ5Of7zbVz/sN616F+3b/wAmpD/sNWn/ALPQB+Z1fr5+zr/yap4B/wCwNB/6DX5B1+vn7Ov/ACap4B/7A0H/AKDQB6cehr4p8L/s5x/ED9tz4meIviV4S1CTwtDetNYC6WSCG/lcqAysCN6KqtkA4ywzX2sTgZNfPHjz9s74M+CdbudHt7zUfEl7bsUkGjQrJCrAkFfNdlQkY7Ej3oA9PHwa+E6eHn0RPhx4XWxePy2iGmxDIxj723dnHfOa/Kj40eC7P4d/H3xT4N015GsdPvStt5hyyxOokRSe5CuBnvivsk/8FA/DE8hXTfhf4muwOv7+LI+oTdXxr8ZPG8fxI+OOv+NotJudKTUpY3FldHMkW2JEw3A67c/jQB9Tf8E+/G2qyah4p+HtzcPLpsMEeq2kbtkQOX8uQKOwbKMR6jPc19k+OvF2m+AvhvrXjLVz/oelWj3TqDgyFR8qA+rNhR7kV8G/8E/P+S6+Jv8AsBD/ANHrX0D+3Hqs+nfsnXVrBLs/tDVLS1kH95N5kI/8cFAH52ePfHfiL4k/EC/8YeKLxp7+8ckLuJS3jz8sUY/hRRwB9SeSa+5v2Tf2bPBSfCHT/iB458O2etazrSfabaHUIhLFaWxPyARtxvYDcWIPBAGMHP55sNyMPUYr9rfBFhHpfw08PaZCu2O10y2gUegWJR/SgD5Y/bK+FHw28M/s7T+JfDvgjRNK1aPUbSFbuxtlhcI8mGX5cDBHtX58t/q2+hr9N/25v+TS7v8A7C1j/wCja/Mhv9W30NAH7V+CUWT4XeHkdQytpdsCrDII8la+Lf27/AfgvwroHhDVfDXhbSdHvL29uIriWwtVgMqrGCAwUAHk56V9p+B/+SY+Hf8AsF2v/ola+Tv+Ch3/ACI/gT/sIXP/AKKWgD4p+HviG28J/Fnw14nvhKbXS9Tt7yYRDLlEcM20dzgHFe7eL7b9oT9rnxmNZ0rwlqEHhaKQ/wBmQXD/AGaxto/75d8CWQjBZlB9FAFeCeBYo5/ip4XhmjSSKTWLJHRwCrKbhAQQeoI7V+1EcaRRLHGioijaqqMAAdAB2oA/L/4f/Cnxh8If25Ph14b8Z21tHdzX0N3FJaSmWKRCHHysVGSCpBGOK+yv2t/C/hu//Zb8Y+IL3QdNuNWsrFWtr+S2QzwkSqPlkxuH3jxnHNeqa/4A8I+KPFugeJ9c0WG71bQJmn0y7LMr27MAG+6QGBwOGyK4H9qv/kzjx7/2D1/9HJQB+TLfeNfr7+zz/wAmr+Af+wLb/wDoNfkE33m+tfr7+zz/AMmr+Af+wJb/APoNAHjn7bPxq1LwL4HsfAHhi8e01fxBG73V1E+2S3s1O0hSOhkJK57AN65Hwd8NPA978SPizoPgfT3MMmqXQiaULnyYgC0kmO+1FY/hXrn7bGqy6h+19q9o8m5NOsLO1jX+6DF5pH5yk1f/AGF7BLv9rBLl1ybPRrudD6MTFH/J2oA+5tA/Z3+C/h/wxBodr8O9BuYYkCme+tEuJ5TjlnkYElj1Nfnl+1l4R8NeCP2ndS0Hwno1rpGmrZWsy2tqu2NXdCWIHbJr9XK/L39tz/k7/Vf+wdZf+izQBb/YY/5Ozi/7At5/OKv03r8yP2GP+Ts4v+wLefzir9N6ACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoo70maAFpDQWwKzNZ1/S9CsDdandLEuPlQcu59FXqamU4xV5OxUISnJRirtmhJJHDE0ksioijLMxwAPU1xOu+LIEtC63jWVgflFwP9bc+0S+n+1XmPjD4uSX92YrVAsCHMduPnJP95u2fzriprnXvEF0bvUbqSFX7bv3jL9T91f5dhXkVc05ny4Zcz7vY+qwnDUoJVMZLlXbr9x22q/EbTrVfIt7RdqnMNq/zLu/vOvWRv94ge3auUvdT8R+ILsz3csiNIPlEp3Pj2XhUX/OKrW0FjaHbZQiWU9ZQpb8FHVm9zxVp0bZItxLjHLxq/APrI/TPX5RXK6U6zvXlfyWi/wCCexTnQw2mFhZ93q/l2K8VhYWr+bPKLqbOPNkO9FPoo/iP0FXi92ZTHHE0DuPmQnErD/aP8C89OOlVo7mOOTMLGMnI8xFzIw/2B/AOvJ5p8Ul1IhjtI4reLd8zuSwJ/m7da6IQjBWirGNWc5u83d+ZaEFnBCst5IkgHCgjbED0+VBy7e5NH295WaKESoCPmSAASEHj5m6IPampYQrOXuZZ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/wABbI/CuioAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKQ0AGabv5rnPFfjnw/4PtfM1e9AmIylrF80r/Rew9zge9fPfjX4x+IfE5ks9OY6Zpxyvkwv88g/wBt+v4DA+tceJxtOho3d9j2MsyTE493irR7v9O57D41+MPh7wwslpZMmp6kuV8qJ/3cZ/23H8hk/SvnjxT468Q+Lrzz9VvWaJTmOBPljj/3V6ficn3rnGLO2SSTT4EtzKBczFF7lV3H6AV4NfFVMQ7Tdl2Pv8vyjDZfHmpx5pfzPf0S6DYIJrm4EUMbu7fwqMk/4Cus0vwvDbxrc6kVY9k6qD6e5qpa+IdL02Ax2GmSSN/z0lYAn3OP5VBN4l1O+lHlwxemApb8OvSiEqFFXbuzStDGYl8sI8q/E6stJPEyQYhgUfPK/A/+v9BWPf6tBbI0FozIrcNKeJJfx/hFUpH1e7gH9pXwggHRAQij8BVU3umWx/dxvcv/AHsYH5nmsquZxelJXHQyOSd6sl6EKObl8vFJKMfLDEpwfqew/U1bay1WWIM8UdlBjrI4jGPqaqS67fuuyBktk9Ihg/meaovI0jmSWRnc/wATnJriqYitLS9vQ9elgacNbff/AJGstrpcLE3GpLKw4KwqXP5n+lObVNOt1xbWsjEfxORmsQtn1xVi0sJ7w5jUhM43n+g7msI4d1HZu51TqwoxvJ2Ra+3XV9OsMFuXduig/wCcfWtGHw/PcwhrqRUXqSH2jH5dPeprSwksYeDEuevyqxP1Of8A631q+iX8sXmyzLFD2OxdzH2H+TXs4fAU6a95anzeNzerO8aTsinF4atnlEUKK7Z53McD688D9a0bbwxaBzHCltNIDhpGB2J+vP0HPrVuKC/MG0yCCA9EVEDN9eeKm8y/wsMTrjoFSJeR+fP8vrXoRhFdDw6mJqy3mOh0LT7YcCCSQ9XKqfwAx0/T1NWItHt5juRIQuc7m2fj25/Q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1qZEkqD+9sJAz7c16i/xM+Hlq32eHxXpMjD/AJZWcvnn/vmMGrEHjezvSP7O0XXr1G6SJYPGn5vtpuk5e6qll2jZf5mcMbCg/avBc0n9qq5Sv5291P53OT0T4T3epKl98RvFOteIrk4Y2ksvkWqn08qPqPqTXpWn6Tp2k2KWel2cFnAgwsUMYRR+ArCk13xTJldP8FTKf717exRKf++Sx/SulhaUxIZlVXKgsFOQDjkA9xXVRpQh8KPEzDGYnEa1pK3SKtZf9urRDwPU1BcWNtd4FzH5qDnY/K/iO9WaK2POTa1QxY1RQqAKo4AHAFLt4p1FAhAuKWiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKhkbDY5qao2hRmLHOfY0ARZ96YzjPQ1N9nT1f86Psy/32pDuQbsDgLTd5xxVj7MvZmo+zL6tRYd0VSSRgE03BPOaufZ0P96lEEY7H8TRYfMU9nFJsweQBV8QoP4B+NOCIOigfhRYXMUAmTwCfoKcIG7J+fFX8CkwPSiwczKwtWPUqv05p4tV/iZj+gqeimK7I/IiAxsXHuM1lX3hTw7qWTeaRayMerBAD+YrZopOKe44zlF3izg7z4T+GpSWtBJbN7fMKxLj4Tzxf8el1FIB2JKE16vScVHsYvodEcbWX2rni83gPW7Vm/wBFnYdcxMGH5Vny6HewJiVLmMg874cmvePwpGRHGHVSPcZqHQXQ2WYz6o8E+wSggCWPPptI/OnG0uA3EkR9l/8Ar17dNpOm3C4lsYGHugqjN4V0ObGbMJjpsYipdA0WYJ7o8gNtdBuDCfYKDTjbXQGRHbL/ANs+a9Sm8EaTJ91509gwIqq/gO0IxHeOv/bMUvYstY6B5zHFeg/ci/Iin+VqH8Lxp/u5rvG8BEHMd3E3++h/xpjeBbrHF1b/AIKRR7KQ/rkH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";
const fARS = n => "$" + Number(n).toLocaleString("es-AR",{minimumFractionDigits:2,maximumFractionDigits:2});
const genId = () => Date.now().toString(36)+Math.random().toString(36).slice(2);
const today = () => new Date().toLocaleDateString("es-AR");
const todayISO = () => new Date().toISOString().slice(0,10);
// Parsea la fecha guardada en lm_stocklog ("DD/MM/YYYY, HH:MM:SS" o "DD/MM/YYYY")
const parseFechaLog = (s) => {
  if(!s) return null;
  const part = s.split(",")[0].trim();
  const [dd,mm,yyyy] = part.split("/");
  if(!dd||!mm||!yyyy) return null;
  return new Date(+yyyy, +mm-1, +dd);
};
// Pedidos viejos (de versiones anteriores de esta funcionalidad) pueden tener pagoTipo="comprobante"
// o "efectivo" sin sufijo. Los normalizamos a su equivalente confirmado para que sigan funcionando.
const normPagoTipo = (tipo) => tipo==="comprobante" ? "comprobante_confirmado" : tipo==="efectivo" ? "efectivo_confirmado" : (tipo||"");
// Notificaciones viejas (de antes de este arreglo) pueden tener leida:false en vez de una lista.
// Esto evita que romper toda la app con un .includes() sobre ese dato viejo.
const leidaArr = (n) => Array.isArray(n?.leida) ? n.leida : [];
function isPromoVigente(p) {
  if(!p || !p.activa) return false;
  const t = todayISO();
  if(p.vigenciaDesde && t < p.vigenciaDesde) return false;
  if(p.vigenciaHasta && t > p.vigenciaHasta) return false;
  return true;
}
// Promo vigente (3x2 o descuento simple) que aplica a un producto puntual. Prioridad: 3x2 > descuento.
function getProductPromo(promos, pid) {
  const list = (promos||[]).filter(p=>isPromoVigente(p) && (p.tipo==="3x2"||p.tipo==="descuento") && p.data?.productoId===pid);
  return list.find(p=>p.tipo==="3x2") || list.find(p=>p.tipo==="descuento") || null;
}
function getVigentCombos(promos) {
  return (promos||[]).filter(p=>isPromoVigente(p) && p.tipo==="combo");
}
// Calcula el descuento equivalente {type,value} según la promo y la cantidad actual del item
function computeAutoDisc(promo, qty) {
  if(!promo) return {disc:null, label:null};
  if(promo.tipo==="3x2") {
    const comprar = promo.data?.comprar||3, pagar = promo.data?.pagar||2;
    if(!comprar || qty<comprar) return {disc:null, label:null};
    const groups = Math.floor(qty/comprar);
    const remainder = qty - groups*comprar;
    const payableUnits = groups*pagar + remainder;
    const pct = qty>0 ? (1 - payableUnits/qty)*100 : 0;
    return {disc:{type:"%",value:Math.round(pct*10)/10}, label:`🏷️ ${comprar}×${pagar} aplicado`};
  }
  if(promo.tipo==="descuento") {
    const tipoValor = promo.data?.tipoValor||"%", valor = promo.data?.valor||0;
    return {disc:{type:tipoValor,value:valor}, label:`🔻 -${valor}${tipoValor==="%"?"%":""} aplicado`};
  }
  return {disc:null,label:null};
}

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
    // Llamada a la API de OneSignal — con timeout propio, nunca debe colgar indefinidamente
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 8000);
    try {
      await fetch("https://onesignal.com/api/v1/notifications", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Key ${OS_API_KEY}`,
        },
        body: JSON.stringify(payload),
        signal: controller.signal,
      });
    } finally {
      clearTimeout(timeoutId);
    }
  } catch(e) {
    console.warn("OneSignal push error:", e);
  }
}

// ─── CROSS-DEVICE NOTIFICATIONS via lm_notifs + OneSignal ────────────────────
async function sendCrossNotif(db, setNotifs, {title, body, tag, para, de}) {
  const notif = {
    id: genId(),
    fecha: new Date().toLocaleString("es-AR"),
    titulo: title, cuerpo: body,
    tag: tag||"lm",
    para, de,
    leida: [],  // lista de userIds que ya la leyeron (igual al resto de las notificaciones de la app)
  };
  // La notificacion es secundaria al hecho que la dispara (confirmar pago, aprobar edicion, etc.
  // que ya se guardo ANTES de llamar a esta funcion en todos los casos). Si el insert falla por un
  // hipo de red, no debe dejar el boton de "Guardando..." de quien la llamo trabado para siempre.
  try {
    await db.addNotif(notif);
    setNotifs(n=>[notif,...n]);
  } catch(e) {
    console.warn("No se pudo guardar la notificación (la acción principal ya se guardó igual):", e);
  }
  // Enviar push via OneSignal (cuando la app está cerrada) — best-effort, nunca debe bloquear ni trabar a quien envía
  sendPushNotif({ title, body, targetUsername: para }).catch(()=>{});
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

  /* HEADER — banner completo (logo + texto + ola) con línea dorada de acento */
  .header-img{
    width:100%;
    height:auto;
    display:block;
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
  <img class="header-img" src="${logoSrc}" alt="Libreria Madrid" onerror="this.style.display='none'"/>
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
  doRender(PDF_LOGO_BANNER);
}


// ─── NOTIFICATION TYPES CONFIG ───────────────────────────────────────────────
const NOTIF_TYPES = {
  NUEVO_PEDIDO:   {label:"Nuevo pedido",         icon:"🛒", color:"#1a5276", bg:"#d6eaf8"},
  CAMBIO_ESTADO:  {label:"Cambio de estado",      icon:"📋", color:"#6c3483", bg:"#e8daef"},
  ALTA_MERCADERIA:{label:"Alta de mercadería",    icon:"📦", color:"#1e8449", bg:"#d5f5e3"},
  PEDIDOS_PEND:   {label:"Pedidos pendientes",    icon:"⏰", color:"#e67e22", bg:"#fef9e7"},
};

// Para notificaciones creadas via sendCrossNotif (comprobantes, efectivo, ediciones, bajas, solicitudes
// de compra) que no tienen un "tipo" de la lista de arriba — se deriva del prefijo del tag.
// Tambien sirve como clave de agrupacion en el panel.
const iconForTag = (tag="") => {
  if(tag.startsWith("comp-ok"))     return {icon:"✅", color:"#1e8449", bg:"#eafaf1", cat:"Comprobantes confirmados"};
  if(tag.startsWith("comp-no"))     return {icon:"❌", color:"#c0392b", bg:"#fdecea", cat:"Comprobantes rechazados"};
  if(tag.startsWith("comp"))        return {icon:"📎", color:"#1a5276", bg:"#eaf2f8", cat:"Comprobantes nuevos"};
  if(tag.startsWith("ef-pend"))     return {icon:"💵", color:"#b7770d", bg:"#fef9e7", cat:"Efectivo pendiente"};
  if(tag.startsWith("ef-ok"))       return {icon:"✅", color:"#1e8449", bg:"#eafaf1", cat:"Efectivo confirmado"};
  if(tag.startsWith("ef-no"))       return {icon:"❌", color:"#c0392b", bg:"#fdecea", cat:"Efectivo rechazado"};
  if(tag.startsWith("edit-req"))    return {icon:"✏️", color:"#1a5276", bg:"#eaf2f8", cat:"Ediciones solicitadas"};
  if(tag.startsWith("edit"))        return {icon:"✏️", color:"#6c3483", bg:"#f5eef8", cat:"Ediciones"};
  if(tag.startsWith("del-client"))  return {icon:"👤", color:"#c0392b", bg:"#fdecea", cat:"Bajas de cliente"};
  if(tag.startsWith("po-"))         return {icon:"📋", color:"#6c3483", bg:"#f5eef8", cat:"Solicitudes de compra"};
  return {icon:"🔔", color:"#666", bg:"#f5f5f5", cat:"Otras"};
};

// ─── NOTIF PANEL ─────────────────────────────────────────────────────────────
function NotifRow({n, isRead, cfg, onClick, onDelete}) {
  return (
    <div onClick={onClick} style={{padding:"12px 16px",borderBottom:"1px solid #f9f9f9",background:isRead?"#fff":"#fafbff",cursor:"pointer",display:"flex",gap:10,alignItems:"flex-start"}}>
      <span style={{fontSize:20,flexShrink:0,marginTop:2}}>{cfg.icon}</span>
      <div style={{flex:1,minWidth:0}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:6}}>
          <div style={{fontWeight:isRead?500:700,fontSize:13,color:"#1a1a1a",lineHeight:1.3}}>{n.titulo}</div>
          {!isRead&&<span style={{width:8,height:8,borderRadius:"50%",background:RED,flexShrink:0,marginTop:4}}/>}
        </div>
        <div style={{fontSize:12,color:"#666",marginTop:2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{n.cuerpo}</div>
        <div style={{fontSize:10,color:"#aaa",marginTop:4}}>{n.fecha}</div>
      </div>
      <button onClick={e=>{e.stopPropagation();onDelete();}} style={{background:"none",border:"none",color:"#ddd",cursor:"pointer",fontSize:16,flexShrink:0,padding:0,lineHeight:1}}>x</button>
    </div>
  );
}
function NotifGroup({items, cfg, cat, currentUser, markRead, delNotif}) {
  const [open, setOpen] = useState(false);
  const unreadN = items.filter(n=>!leidaArr(n).includes(currentUser.id)).length;
  return (
    <div style={{borderBottom:"1px solid #f5f5f5"}}>
      <div onClick={()=>setOpen(o=>!o)} style={{padding:"11px 16px",display:"flex",alignItems:"center",gap:10,cursor:"pointer",background:unreadN>0?"#fafbff":"#fff"}}>
        <div style={{width:32,height:32,borderRadius:9,background:cfg.bg,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontSize:16}}>{cfg.icon}</div>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontWeight:700,fontSize:13,color:"#1a1a1a"}}>{cat}</div>
          <div style={{fontSize:11,color:"#888"}}>{items.length} notificaciones{unreadN>0?` · ${unreadN} sin leer`:""}</div>
        </div>
        <span style={{fontSize:11,color:"#bbb",transform:open?"rotate(180deg)":"none",transition:"transform .15s",flexShrink:0}}>▼</span>
      </div>
      {open && <div style={{background:"#fcfcfc"}}>
        {items.map(n=>(
          <NotifRow key={n.id} n={n} isRead={leidaArr(n).includes(currentUser.id)} cfg={cfg}
            onClick={()=>markRead(n.id)} onDelete={()=>delNotif(n.id)}/>
        ))}
      </div>}
    </div>
  );
}
function NotifPanel({notifs,setNotifs,currentUser,users,onClose,onMarkAllRead,pushNotif,orders}) {
  const myNotifs = notifs.filter(n =>
    n.para === "todos" || n.para === currentUser.role || n.para === currentUser.id
  );
  const unread = myNotifs.filter(n => !leidaArr(n).includes(currentUser.id));
  const markRead = async (id) => {
    const n = notifs.find(x=>x.id===id);
    if(!n||leidaArr(n).includes(currentUser.id)) return;
    const updated = [...leidaArr(n), currentUser.id];
    setNotifs(ns=>ns.map(x=>x.id===id?{...x,leida:updated}:x));
    await db.updateNotif(id, {leida:updated});
  };
  const delNotif = async (id) => { setNotifs(ns=>ns.filter(n=>n.id!==id)); await db.deleteNotif(id); };
  const pendingOrders = orders.filter(o=>o.stage!=="entregado");
  return (
    <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:998}} onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{position:"fixed",top:70,left:12,right:12,maxWidth:380,marginLeft:"auto",maxHeight:"80vh",background:"#fff",borderRadius:16,boxShadow:"0 12px 40px #0003",border:"1px solid #f0f0f0",display:"flex",flexDirection:"column",zIndex:999,overflow:"hidden"}}>
        <div style={{padding:"14px 18px",borderBottom:"1px solid #f5f5f5",display:"flex",alignItems:"center",justifyContent:"space-between",background:`linear-gradient(135deg,${REDD},${RED})`,borderRadius:"16px 16px 0 0"}}>
          <div style={{color:"#fff",fontWeight:800,fontSize:15,display:"flex",alignItems:"center",gap:7}}>
            <Bell size={15} strokeWidth={2.2}/>Notificaciones {unread.length>0&&<span style={{background:"#f1c40f",color:"#1a1a1a",borderRadius:10,fontSize:11,padding:"1px 6px",marginLeft:0,fontWeight:800}}>{unread.length}</span>}
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
            : (() => {
                // Agrupa por categoria (tipo viejo si existe, sino la categoria derivada del tag).
                // Si una categoria tiene 2+ notificaciones se muestra colapsada; si tiene 1 sola, suelta.
                const groups = {}; const order = [];
                myNotifs.forEach(n=>{
                  const cat = NOTIF_TYPES[n.tipo]?.label || iconForTag(n.tag).cat;
                  if(!groups[cat]) { groups[cat]=[]; order.push(cat); }
                  groups[cat].push(n);
                });
                return order.map(cat=>{
                  const items = groups[cat];
                  const cfg = NOTIF_TYPES[items[0].tipo] || iconForTag(items[0].tag);
                  if(items.length>1) {
                    return <NotifGroup key={cat} items={items} cfg={cfg} cat={cat} currentUser={currentUser} markRead={markRead} delNotif={delNotif}/>;
                  }
                  const n = items[0];
                  return <NotifRow key={n.id} n={n} isRead={leidaArr(n).includes(currentUser.id)} cfg={cfg} onClick={()=>markRead(n.id)} onDelete={()=>delNotif(n.id)}/>;
                });
              })()
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
function AppInner() {
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
  const [promos, setPromos]     = useState([]);
  const [settings, setSettings] = useState({exigirPagoConfirmado:false});
  const [loading, setLoading]   = useState(true);
  const [error, setError]       = useState(null);

  useEffect(() => {
    async function loadAll() {
      try {
        const [u,v,p,o,q,sl,act,pl,po,n,cl,pr,st] = await Promise.all([
          db.getUsers(), db.getVendors(), db.getProducts(),
          db.getOrders(), db.getQuotes(), db.getStockLog(), db.getActivity(), db.getPriceLists(), db.getPurchaseOrders(), db.getNotifs(), db.getClients(), db.getPromos(), db.getSettings(),
        ]);
        setUsers(u); setVendors(v); setProducts(p);
        setOrders(o); setQuotes(q); setStockLog(sl); setActivity(act); setPriceLists(pl); setPurchaseOrders(po); setNotifs(n); setClients(cl); setPromos(pr); setSettings(st);
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
    currentUser={currentUser} onLogout={()=>{
      // Logout de OneSignal en segundo plano, con timeout propio — nunca debe bloquear el logout real de la app
      Promise.race([
        (async()=>{ const OS=await getOneSignal(); if(OS) await OS.logout(); })(),
        new Promise(r=>setTimeout(r,3000)),
      ]).catch(()=>{});
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
    promos={promos} setPromos={setPromos}
    settings={settings} setSettings={setSettings}
  />;
}

// ─── RED DE SEGURIDAD: si algo se rompe, mostrar el error en pantalla ────────
// en vez de dejar la app en blanco. Permite mandar una foto del error directo
// desde el celular, sin necesitar conectar a una computadora.
class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { error: null }; }
  static getDerivedStateFromError(error) { return { error }; }
  componentDidCatch(error, info) { console.error("ErrorBoundary atrapó:", error, info); }
  render() {
    if(!this.state.error) return this.props.children;
    return (
      <div style={{minHeight:"100vh",background:"#fdecea",padding:20,fontFamily:"system-ui,-apple-system,sans-serif",boxSizing:"border-box"}}>
        <div style={{background:"#fff",borderRadius:14,padding:20,boxShadow:"0 1px 6px #0001",maxWidth:480,margin:"30px auto"}}>
          <div style={{fontSize:32,marginBottom:8,textAlign:"center"}}>⚠️</div>
          <div style={{fontWeight:800,fontSize:16,color:"#c0392b",marginBottom:8,textAlign:"center"}}>Algo se rompió</div>
          <div style={{fontSize:13,color:"#666",marginBottom:14,textAlign:"center"}}>Mandale una foto de esto a Claude para que lo arregle.</div>
          <div style={{background:"#1a1a1a",color:"#f5b7b1",borderRadius:8,padding:12,fontSize:12,fontFamily:"monospace",overflowX:"auto",marginBottom:14,whiteSpace:"pre-wrap",wordBreak:"break-word"}}>
            {String(this.state.error?.message || this.state.error)}
            {this.state.error?.stack && <div style={{color:"#aaa",fontSize:10,marginTop:8}}>{this.state.error.stack.split("\n").slice(0,4).join("\n")}</div>}
          </div>
          <button onClick={()=>{this.setState({error:null});window.location.reload();}}
            style={{width:"100%",padding:11,borderRadius:9,border:"none",background:"#c0392b",color:"#fff",fontWeight:700,fontSize:14,cursor:"pointer"}}>
            Recargar
          </button>
        </div>
      </div>
    );
  }
}

// ─── Wrapper raiz: monta los hosts de toasts/confirmacion una sola vez, ──────
// sin importar el estado de auth (loading/error/login/app), y sin necesitar
// pasar props a traves de toda la app (toast.success(...) y confirmDialog(...) funcionan desde cualquier componente)
export default function App() {
  return (
    <ErrorBoundary>
      <AppInner/>
      <ToastHost/>
      <ConfirmHost/>
    </ErrorBoundary>
  );
}

// ─── MAIN APP (authenticated) ─────────────────────────────────────────────────
function MainApp({currentUser,onLogout,users,setUsers,vendors,setVendors,products,setProducts,orders,setOrders,quotes,setQuotes,stockLog,setStockLog,activity,setActivity,priceLists,setPriceLists,purchaseOrders,setPurchaseOrders,notifs,setNotifs,sandboxStock,setSandboxStock,clients,setClients,promos,setPromos,settings,setSettings}) {
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
    !leidaArr(n).includes(currentUser.id) &&
    (n.para==="todos"||n.para===currentUser.role||n.para===currentUser.id)
  ).length;
  const markAllRead = async () => {
    const toUpdate = notifs.filter(n=>!leidaArr(n).includes(currentUser.id)&&(n.para==="todos"||n.para===currentUser.role||n.para===currentUser.id));
    for(const n of toUpdate) await db.updateNotif(n.id,{leida:[...leidaArr(n),currentUser.id]});
    setNotifs(ns=>ns.map(n=>leidaArr(n).includes(currentUser.id)?n:{...n,leida:[...leidaArr(n),currentUser.id]}));
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
    // Cada pedido de prueba necesita un numero UNICO (no "TEST-000000" fijo para todos),
    // sino el upsert en Supabase de un 2do/3er pedido sandbox colisiona con el anterior
    // y el dashboard de Admin nunca llega a verlo.
    const orderWithNum = {...order, docNum: test ? `TEST-${genId().slice(-6).toUpperCase()}` : fmtDocNum("Reserva", n), isTest: test, isSandbox: test};
    const isSandbox = test;

    if(isSandbox) {
      // SANDBOX: el stock usa el paralelo en memoria/localStorage (nunca toca lm_products).
      // El PEDIDO si se guarda en Supabase para que el admin pueda verlo en el Dashboard Demo
      // desde cualquier dispositivo, sin generar notificaciones ni tocar el stock real.
      updateSandboxStock(prev => {
        const next = {...prev};
        orderWithNum.items.forEach(it => {
          next[it.pid] = Math.max(0, (next[it.pid] ?? 0) - it.qty);
        });
        return next;
      });
      setOrders(o=>[{...orderWithNum, isSandbox: true},...o]);
      await db.upsertOrder({...orderWithNum, isSandbox: true});
      if(!order.fromQuote) setTimeout(() => printDoc(orderWithNum, "reserva"), 400);
      return; // ← no notifs, no log de actividad, no toca stock real
    }

    // REAL: descontar stock, guardar en Supabase, notificar
    const updatedProds = products.map(x=>{const qty=orderWithNum.items.filter(i=>i.pid===x.id).reduce((s,i)=>s+i.qty,0);return qty>0?{...x,stock:Math.max(0,x.stock-qty)}:x;});
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
      updated = {...updated, compNum: test ? `TEST-${genId().slice(-6).toUpperCase()}` : fmtDocNum("Comp", n), isTest: test};
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
        const updatedProds=products.map(x=>{const qty=ord.items.filter(i=>i.pid===x.id).reduce((s,i)=>s+i.qty,0);return qty>0?{...x,stock:x.stock+qty}:x;});
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

  // Comprobante de pago — opcional, un solo archivo por pedido. Queda "pendiente" hasta que el admin lo confirme o rechace.
  const uploadComprobante = async (id, {url, nombre, fecha}) => {
    const updated = orders.map(o=>o.id===id ? {...o, comprobanteUrl:url, comprobanteNombre:nombre, comprobanteFecha:fecha, pagoTipo:"comprobante_pendiente"} : o);
    setOrders(updated);
    const ord = updated.find(o=>o.id===id);
    await db.upsertOrder(ord);
    // Avisar al admin (salvo en pedidos de prueba/sandbox, o si quien sube ya es el admin)
    const isSandboxOrder = ord.isSandbox || isTestOrder(ord.vendedor);
    if(isSandboxOrder) { sendLocalNotif("📎 Comprobante subido", "Esperando confirmación del admin", `comp-${id}`); return; }
    if(currentUser.role==="admin") return;
    await sendCrossNotif(db, setNotifs, {title:"📎 Comprobante subido", body:`${currentUser.name} subió un comprobante de pago (${ord.client})`, tag:`comp-${id}`, para:"admin", de:currentUser.name});
  };

  // Admin confirma que el comprobante es correcto
  const confirmarComprobante = async (id) => {
    const updated = orders.map(o=>o.id===id ? {...o, pagoTipo:"comprobante_confirmado"} : o);
    setOrders(updated);
    const ord = updated.find(o=>o.id===id);
    const isSandboxOrder = ord.isSandbox || isTestOrder(ord.vendedor);
    await db.upsertOrder(ord);
    if(isSandboxOrder) return;
    await sendCrossNotif(db, setNotifs, {title:"✅ Comprobante confirmado", body:`El admin confirmó el comprobante del pedido de ${ord.client}`, tag:`comp-ok-${id}`, para:ord.vendedor||"", de:"admin"});
  };

  // Admin rechaza el comprobante (motivo obligatorio) → vuelve a sin pago, el vendedor puede resubir
  const rechazarComprobante = async (id, motivo) => {
    const updated = orders.map(o=>o.id===id ? {...o, pagoTipo:"", comprobanteUrl:"", comprobanteNombre:"", comprobanteFecha:""} : o);
    setOrders(updated);
    const ord = updated.find(o=>o.id===id);
    const isSandboxOrder = ord.isSandbox || isTestOrder(ord.vendedor);
    await db.upsertOrder(ord);
    if(isSandboxOrder) return;
    await sendCrossNotif(db, setNotifs, {title:"❌ Comprobante rechazado", body:`El admin rechazó el comprobante del pedido de ${ord.client}: "${motivo}"`, tag:`comp-no-${id}`, para:ord.vendedor||"", de:"admin"});
  };

  // ── PAGO EN EFECTIVO ─────────────────────────────────────────────────────────
  // Fase 1 (vendedor): solicita registrar efectivo → queda en "efectivo_pendiente"
  // Si lo marca el Admin (pedido propio, venta directa), queda confirmado de una — no tiene sentido que se apruebe a si mismo
  const marcarEfectivo = async (id) => {
    const esAdmin = currentUser.role === "admin";
    const updated = orders.map(o=>o.id===id ? {...o, pagoTipo: esAdmin ? "efectivo_confirmado" : "efectivo_pendiente", pagoEfectivoFecha:today()} : o);
    setOrders(updated);
    const ord = updated.find(o=>o.id===id);
    const isSandboxOrder = ord.isSandbox || isTestOrder(ord.vendedor);
    await db.upsertOrder(ord);
    if(isSandboxOrder) { sendLocalNotif(esAdmin?"💵 Efectivo confirmado":"💵 Efectivo registrado", esAdmin?"Pago en efectivo registrado":"Esperando confirmación del admin", `ef-pend-${id}`); return; }
    if(esAdmin) return; // el admin no necesita notificarse a si mismo
    await sendCrossNotif(db, setNotifs, {title:"💵 Pago en efectivo", body:`${currentUser.name} registró pago en efectivo (${ord.client})`, tag:`ef-pend-${id}`, para:"admin", de:currentUser.name});
  };

  // Fase 2a (admin): confirma el efectivo
  const confirmarEfectivo = async (id) => {
    const updated = orders.map(o=>o.id===id ? {...o, pagoTipo:"efectivo_confirmado"} : o);
    setOrders(updated);
    const ord = updated.find(o=>o.id===id);
    const isSandboxOrder = ord.isSandbox || isTestOrder(ord.vendedor);
    await db.upsertOrder(ord);
    if(isSandboxOrder) return;
    await sendCrossNotif(db, setNotifs, {title:"✅ Efectivo confirmado", body:`El admin confirmó el pago en efectivo del pedido de ${ord.client}`, tag:`ef-ok-${id}`, para:ord.vendedor||"", de:"admin"});
  };

  // Fase 2b (admin): rechaza el efectivo → vuelve a sin pago
  const rechazarEfectivo = async (id) => {
    const updated = orders.map(o=>o.id===id ? {...o, pagoTipo:"", pagoEfectivoFecha:""} : o);
    setOrders(updated);
    const ord = updated.find(o=>o.id===id);
    const isSandboxOrder = ord.isSandbox || isTestOrder(ord.vendedor);
    await db.upsertOrder(ord);
    if(isSandboxOrder) return;
    await sendCrossNotif(db, setNotifs, {title:"❌ Efectivo no confirmado", body:`El admin no pudo confirmar el pago en efectivo del pedido de ${ord.client}`, tag:`ef-no-${id}`, para:ord.vendedor||"", de:"admin"});
  };

  // ── CONFIGURACION GLOBAL: exigir pago confirmado para habilitar Entregar ────
  const toggleExigirPago = async () => {
    const next = {...settings, exigirPagoConfirmado: !settings.exigirPagoConfirmado};
    setSettings(next);
    await db.saveSettings(next);
  };

  // ── EDIT REQUEST FLOW ─────────────────────────────────────────────────────
  // Fase 1: vendedor solicita edición
  const requestEdit = async (id, reason) => {
    const updated = orders.map(o=>o.id===id ? {...o, editStatus:"solicitada", editReason:reason} : o);
    setOrders(updated);
    const ord0 = updated.find(o=>o.id===id);
    const isSandboxOrder = ord0 && (ord0.isSandbox || isTestOrder(ord0.vendedor));
    if(isSandboxOrder) { sendLocalNotif("✏️ Solicitud enviada", `Esperá la aprobación del admin`, `edit-req-${id}`); return; }
    await db.upsertOrder(ord0);
    await sendCrossNotif(db, setNotifs, {title:"✏️ Solicitud de edición", body:`${ord0?.vendedor} quiere editar un pedido (${ord0?.client})`, tag:`edit-req-${id}`, para:"admin", de:ord0?.vendedor||""});
    sendLocalNotif("✏️ Solicitud enviada", `Esperá la aprobación del admin`, `edit-req-${id}`);
  };

  // Fase 2a: admin aprueba la solicitud → vendedor puede editar
  const approveEditRequest = async (id) => {
    const updated = orders.map(o=>o.id===id ? {...o, editStatus:"aprobada", editRejectReason:""} : o);
    setOrders(updated);
    const ordApr = updated.find(o=>o.id===id);
    const isSandboxOrder = ordApr && (ordApr.isSandbox || isTestOrder(ordApr.vendedor));
    if(isSandboxOrder) return;
    await db.upsertOrder(ordApr);
    await sendCrossNotif(db, setNotifs, {title:"✅ Edición aprobada", body:`Tu solicitud para editar el pedido de ${ordApr?.client} fue aprobada`, tag:`edit-apr-${id}`, para:ordApr?.vendedor||"", de:"admin"});
  };

  // Fase 2b: admin rechaza la solicitud
  const rejectEditRequest = async (id, reason) => {
    const updated = orders.map(o=>o.id===id ? {...o, editStatus:"rechazada", editRejectReason:reason} : o);
    setOrders(updated);
    const ordRej = updated.find(o=>o.id===id);
    const isSandboxOrder = ordRej && (ordRej.isSandbox || isTestOrder(ordRej.vendedor));
    if(isSandboxOrder) return;
    await db.upsertOrder(ordRej);
    await sendCrossNotif(db, setNotifs, {title:"❌ Edición rechazada", body:`Tu solicitud para editar el pedido de ${ordRej?.client} fue rechazada. Motivo: "${reason}"`, tag:`edit-rej-${id}`, para:ordRej?.vendedor||"", de:"admin"});
  };

  // Fase 3: vendedor guarda los cambios editados
  const submitEdit = async (id, newItems, newTotal) => {
    const updated = orders.map(o=>o.id===id ? {...o, editStatus:"en revisión", editItems:newItems, editTotal:newTotal} : o);
    setOrders(updated);
    const ordSub = updated.find(o=>o.id===id);
    const isSandboxOrder = ordSub && (ordSub.isSandbox || isTestOrder(ordSub.vendedor));
    if(isSandboxOrder) { sendLocalNotif("📤 Cambios enviados", `El admin revisará tu edición`, `edit-sub-${id}`); return; }
    await db.upsertOrder(ordSub);
    await sendCrossNotif(db, setNotifs, {title:"👀 Cambios para revisar", body:`${ordSub?.vendedor} editó el pedido de ${ordSub?.client} — revisá los cambios`, tag:`edit-sub-${id}`, para:"admin", de:ordSub?.vendedor||""});
    sendLocalNotif("📤 Cambios enviados", `El admin revisará tu edición`, `edit-sub-${id}`);
  };

  // Fase 4a: admin aprueba los cambios finales
  const approveEdit = async (id) => {
    const ord = orders.find(o=>o.id===id);
    if(!ord) return;
    const isSandboxOrder = ord.isSandbox || isTestOrder(ord.vendedor);
    // Aplicar cambios: restaurar stock viejo, descontar nuevo
    if(isSandboxOrder) {
      // SANDBOX: ajustar el stock paralelo en memoria, nunca tocar Supabase
      updateSandboxStock(prev => {
        const next = {...prev};
        ord.items.forEach(it => { next[it.pid] = (next[it.pid] ?? 0) + it.qty; });
        ord.editItems.forEach(it => { next[it.pid] = Math.max(0, (next[it.pid] ?? 0) - it.qty); });
        return next;
      });
    } else {
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
    if(isSandboxOrder) return;
    const ordOk = updated.find(o=>o.id===id);
    await db.upsertOrder(ordOk);
    await sendCrossNotif(db, setNotifs, {title:"✅ Cambios aprobados", body:`Tu edición del pedido de ${ordOk?.client} fue aprobada`, tag:`edit-ok-${id}`, para:ordOk?.vendedor||"", de:"admin"});
    await logActivity("Edición aprobada", `Pedido ${ord.docNum||ord.compNum||""} editado`, id, "pedido");
  };

  // Fase 4b: admin rechaza los cambios finales
  const rejectEdit = async (id, reason) => {
    const updated = orders.map(o=>o.id===id ? {...o, editStatus:"cambios rechazados", editRejectReason:reason, editItems:null} : o);
    setOrders(updated);
    const ordNo = updated.find(o=>o.id===id);
    const isSandboxOrder = ordNo && (ordNo.isSandbox || isTestOrder(ordNo.vendedor));
    if(isSandboxOrder) return;
    await db.upsertOrder(ordNo);
    await sendCrossNotif(db, setNotifs, {title:"❌ Cambios rechazados", body:`El admin rechazó tu edición del pedido de ${ordNo?.client}. Motivo: "${reason}"`, tag:`edit-no-${id}`, para:ordNo?.vendedor||"", de:"admin"});
  };
  const addQuote = async (quote) => {
    const test = isTestOrder(quote.vendedor);
    const n = test ? 0 : await db.nextCounter("presu");
    const quoteWithNum = {...quote, docNum: test ? `TEST-${genId().slice(-6).toUpperCase()}` : fmtDocNum("Presu", n), isTest: test};
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
    if(updProd)await db.upsertProduct(updProd); // esto es lo importante: el stock real. Si esto falla, sí queremos que el error se propague.
    if(prod){
      // La notificación es secundaria — si falla (ej: hipo de red), NO debe tirar abajo
      // toda la operación ni dejar el boton de "Guardando..." trabado para siempre,
      // porque el stock de arriba ya quedó guardado correctamente.
      try {
        const notif={id:genId(),fecha:new Date().toLocaleString("es-AR"),leida:[],tipo:"ALTA_MERCADERIA",para:"admin",icono:"📦",titulo:"Alta de mercaderia",cuerpo:`${prod.name} - +${qty} unidades${newCost?` - Nuevo costo: ${fARS(newCost)}`:""}`,ref:pid};
        await db.addNotif(notif); setNotifs(n=>[notif,...n]);
      } catch(e) {
        console.warn("No se pudo crear la notificacion de alta de mercaderia (el stock SI se actualizo):", e);
      }
    }
  };

  const pending = orders.filter(o=>o.stage!=="entregado"&&!o.isSandbox).length;
  const TABS = [
    {k:"central",   label:"Central",           icon:ClipboardList, roles:["admin","vendedor"]},
    {k:"nuevo",     label:"Nuevo Pedido",       icon:ShoppingCart,  roles:["admin","vendedor"]},
    {k:"clientes",  label:"Clientes",           icon:Users,         roles:["admin","vendedor"]},
    {k:"cotizacion",label:"Cotizaciones",       icon:FileText,      roles:["admin","vendedor"]},
    {k:"precios",   label:"Precios",            icon:CircleDollarSign, roles:["admin","vendedor"]},
    {k:"stock",     label:"Stock",              icon:Package,       roles:["admin","vendedor"]},
    {k:"compras",   label:"Alta Mercancía",icon:Store,        roles:["admin","vendedor"]},
    {k:"solicitud",  label:"Solicitudes", icon:ListChecks,    roles:["admin","vendedor"]},
    {k:"admin",     label:"Admin",     icon:Star,        roles:["admin"]},
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
  // El admin rechaza la baja solicitada — el cliente queda activo, sin necesidad de eliminarlo
  const rejectDeleteClient = async (id) => {
    const client = clients.find(c=>c.id===id);
    if(!client) return;
    const updated = {...client, deleteRequested:false, deleteReason:""};
    setClients(prev=>prev.map(c=>c.id===id?updated:c));
    await db.saveClient(updated);
    await sendCrossNotif(db, setNotifs, {
      title:"❌ Baja de cliente rechazada",
      body:`El admin no aprobó la baja de "${client.name}" — sigue activo`,
      tag:`del-client-rej-${id}`, para:client.createdBy||"", de:"admin"
    });
  };

  // ── BANDEJA DE APROBACIONES — deep-link a solicitudes de compra desde Central ──
  const [deepLinkPOId, setDeepLinkPOId] = useState(null);
  // Deep-link a una seccion puntual de Admin (ej: "Ver en Pagos" desde un pedido bloqueado)
  const [deepLinkAdminSection, setDeepLinkAdminSection] = useState(null);
  const [deepLinkOfertaProductId, setDeepLinkOfertaProductId] = useState(null);
  // Deep-link desde "Pedir reposicion" en stock bajo -> abre Solicitud de Compra con el producto precargado
  const [deepLinkSolicitudItem, setDeepLinkSolicitudItem] = useState(null);
  const quickReviewPO = async (id) => {
    const po = purchaseOrders.find(p=>p.id===id);
    if(!po) return;
    const updated = {...po, estado:"revisando"};
    setPurchaseOrders(prev=>prev.map(x=>x.id===id?updated:x));
    await db.savePurchaseOrder(updated);
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
          sendLocalNotif(n.titulo, n.cuerpo, n.tag);
          setNotifs(prev => prev.find(x=>x.id===n.id) ? prev : [n,...prev]);
        }
      })
      .on("postgres_changes", {event:"INSERT", schema:"public", table:"lm_purchase_orders"}, (payload) => {
        const po = payload.new;
        setPurchaseOrders(prev => prev.find(x=>x.id===po.id) ? prev : [{...po, fechaCierre:po.fecha_cierre, fechaRecibido:po.fecha_recibido},...prev]);
        if(po.vendedor !== currentUser.vendedor && po.vendedor !== currentUser.name) {
          sendLocalNotif("📋 Nueva solicitud de compra", `${po.vendedor} creó una solicitud`, `po-${po.id}`);
        }
      })
      .on("postgres_changes", {event:"UPDATE", schema:"public", table:"lm_purchase_orders"}, (payload) => {
        const po = payload.new;
        setPurchaseOrders(prev => prev.map(x => x.id===po.id ? {...po, fechaCierre:po.fecha_cierre, fechaRecibido:po.fecha_recibido} : x));
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
                <div style={{color:"#ffcccc",fontSize:11,display:"flex",alignItems:"center",gap:4}}><Users size={10} strokeWidth={2.5}/>{currentUser.name}{activeList.discount>0&&<span style={{marginLeft:6,background:"#f1c40f",color:"#1a1a1a",borderRadius:4,padding:"1px 5px",fontSize:10,fontWeight:800}}>{activeList.name}</span>}</div>
              </div>
              <div style={{display:"flex",gap:6,flexShrink:0}}>
                <button onClick={()=>setShowNotifs(s=>!s)} style={{width:34,height:34,borderRadius:9,background:"#ffffff1a",border:"1px solid #ffffff2a",color:"#fff",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",position:"relative"}}>
                  <Bell size={16} strokeWidth={2.1}/>{unreadCount>0&&<span style={{position:"absolute",top:-3,right:-3,background:"#f1c40f",color:"#1a1a1a",borderRadius:"50%",width:15,height:15,fontSize:8,fontWeight:900,display:"flex",alignItems:"center",justifyContent:"center"}}>{unreadCount}</span>}
                  {notifPermission==="default"&&<span style={{position:"absolute",top:-3,right:-3,background:"#e67e22",borderRadius:"50%",width:10,height:10}}/>}
                </button>
                <button onClick={()=>setMobileMenu(o=>!o)} style={{width:34,height:34,borderRadius:9,background:"#ffffff1a",border:"1px solid #ffffff2a",color:"#fff",fontSize:18,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer"}}><Menu size={17} strokeWidth={2.3}/></button>
              </div>
            </div>
            <div style={{display:"flex",overflowX:"auto",gap:1,padding:"0 10px",scrollbarWidth:"none",alignItems:"center"}}>
              {TABS.map(t=>(
                <button key={t.k} onClick={()=>setTab(t.k)}
                  style={{padding:"7px 13px",border:"none",cursor:"pointer",fontSize:11,color:tab===t.k?"#fff":"#ffbbbb",fontWeight:tab===t.k?700:600,borderRadius:"8px 8px 0 0",background:tab===t.k?"#ffffff18":"transparent",borderBottom:tab===t.k?"3px solid #fff":"3px solid transparent",whiteSpace:"nowrap",flexShrink:0,display:"inline-flex",alignItems:"center",gap:5}}>
                  <t.icon size={13} strokeWidth={2.25}/> {t.label}
                </button>
              ))}
              <button onClick={async()=>{if(await confirmDialog("¿Seguro que querés salir?","Vas a cerrar tu sesión.")) onLogout();}}
                style={{padding:"5px 10px",border:"1px solid #ffffff33",cursor:"pointer",fontSize:11,color:"#ffbbbb",fontWeight:600,borderRadius:7,background:"transparent",whiteSpace:"nowrap",flexShrink:0,marginLeft:6,display:"inline-flex",alignItems:"center",gap:4}}>
                <LogOut size={11} strokeWidth={2.4}/> Salir
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
                  <button key={t.k} onClick={()=>setTab(t.k)} style={{padding:"11px 10px",border:"none",cursor:"pointer",fontSize:14,background:tab===t.k?"#fff":"transparent",color:tab===t.k?RED:"#ffcccc",fontWeight:tab===t.k?800:600,borderRadius:"8px 8px 0 0",position:"relative",display:"inline-flex",alignItems:"center",gap:6}}>
                    <t.icon size={14} strokeWidth={2.25}/> {t.label}
                    {t.k==="central"&&pending>0&&<span style={{position:"absolute",top:5,right:3,background:"#fff",color:RED,borderRadius:10,fontSize:10,padding:"1px 5px",fontWeight:800,border:`1.5px solid ${RED}`}}>{pending}</span>}
                  </button>
                ))}
              </nav>
              <div style={{display:"flex",alignItems:"center",gap:10,padding:"0 8px",borderLeft:"1px solid #ffffff33",marginLeft:4,marginTop:14}}>
                {isAdmin && priceLists.length>1 && (
                  <select value={previewListId||"default"} onChange={e=>setPreviewListId(e.target.value==="default"?null:e.target.value)}
                    style={{background:"#ffffff18",border:"1px solid #ffffff40",color:"#fff",borderRadius:6,padding:"6px 10px",cursor:"pointer",fontSize:12,fontWeight:600}}>
                    {priceLists.map(pl=><option key={pl.id} value={pl.id} style={{color:"#1a1a1a"}}>{pl.name}{pl.discount>0?` (-${pl.discount}%)`:""}</option>)}
                  </select>
                )}
                <button onClick={()=>setShowChangePass(true)}
                  style={{background:"#ffffff18",border:"1px solid #ffffff40",color:"#fff",borderRadius:6,padding:"6px 10px",cursor:"pointer",fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:5}}>
                  <Users size={11} strokeWidth={2.5}/> {currentUser.name}
                  {activeList.discount>0&&<span style={{background:"#f1c40f",color:"#1a1a1a",borderRadius:4,padding:"1px 5px",fontSize:10,fontWeight:800,marginLeft:3}}>{activeList.name}</span>}
                  <Key size={11} strokeWidth={2.4} style={{opacity:.7}}/>
                </button>
                <div style={{position:"relative"}}>
                  <button onClick={()=>setShowNotifs(s=>!s)} style={{background:"#ffffff18",border:"1px solid #ffffff40",color:"#fff",borderRadius:6,padding:"6px 9px",cursor:"pointer",fontSize:16,lineHeight:1,position:"relative",display:"flex",alignItems:"center"}}>
                    <Bell size={15} strokeWidth={2.1}/>
                    {unreadCount>0&&<span style={{position:"absolute",top:-4,right:-4,background:"#f1c40f",color:"#1a1a1a",borderRadius:10,fontSize:9,padding:"1px 4px",fontWeight:800,minWidth:14,textAlign:"center"}}>{unreadCount}</span>}
                  </button>
                </div>
                <button onClick={onLogout} style={{background:"transparent",border:"1.5px solid #ffffff55",color:"#ffe5e5",borderRadius:6,padding:"5px 10px",cursor:"pointer",fontSize:13,fontWeight:700,display:"inline-flex",alignItems:"center",gap:5}}><LogOut size={12} strokeWidth={2.4}/> Salir</button>
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
                <span style={{width:28,display:"flex",justifyContent:"center"}}><t.icon size={18} strokeWidth={2.25}/></span>{t.label}
              </div>
            ))}
            <div style={{borderTop:"2px solid #e5e5e5",marginTop:216,paddingTop:8}}>
              {/* Botón de notificaciones OneSignal */}
              <div onClick={async()=>{
                try {
                  const OS = await getOneSignal();
                  if(!OS) {
                    toast.error("El sistema de notificaciones no está disponible. Verificá tu conexión y recargá la app.");
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
                  toast.error("Hubo un problema al activar las notificaciones. Intentá de nuevo.");
                }
                setMobileMenu(false);
              }} style={{display:"flex",alignItems:"center",gap:14,padding:"13px 20px",fontSize:14,fontWeight:600,
                color:notifPermission==="granted"?"#1e8449":notifPermission==="denied"?"#aaa":"#e67e22",
                cursor:"pointer",
                background:notifPermission==="granted"?"#eafaf1":"transparent"}}>
                <span style={{width:28,display:"flex",justifyContent:"center"}}>
                  <Bell size={18} strokeWidth={2.1} color={notifPermission==="granted"?"#1e8449":notifPermission==="denied"?"#aaa":"#e67e22"}/>
                </span>
                {notifPermission==="granted"
                  ? "Notificaciones activas ✓"
                  : notifPermission==="denied"
                    ? "Notificaciones bloqueadas"
                    : "Activar notificaciones push"
                }
              </div>
              <div onClick={async()=>{if(await confirmDialog("¿Seguro que querés salir?","Vas a cerrar la sesión y volver a la pantalla de inicio.")) onLogout();}} style={{display:"flex",alignItems:"center",gap:14,padding:"13px 20px",fontSize:14,fontWeight:600,color:RED,cursor:"pointer",background:"#fdecea"}}>
                <span style={{width:28,display:"flex",justifyContent:"center"}}><LogOut size={18} strokeWidth={2.1}/></span>Salir
              </div>
              <div onClick={()=>{setShowChangePass(true);setMobileMenu(false);}} style={{display:"flex",alignItems:"center",gap:14,padding:"13px 20px",fontSize:14,fontWeight:600,color:"#333",cursor:"pointer"}}>
                <span style={{width:28,display:"flex",justifyContent:"center"}}><Key size={18} strokeWidth={2.1}/></span>Cambiar contraseña
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
          onUploadComprobante={uploadComprobante}
          onConfirmarComprobante={confirmarComprobante} onRechazarComprobante={rechazarComprobante}
          onMarcarEfectivo={marcarEfectivo} onConfirmarEfectivo={confirmarEfectivo} onRechazarEfectivo={rechazarEfectivo}
          exigirPagoConfirmado={settings.exigirPagoConfirmado}
          clients={clients} purchaseOrders={purchaseOrders}
          onDeleteClient={deleteClient} onRejectDeleteClient={rejectDeleteClient}
          onQuickReviewPO={quickReviewPO} onViewPO={(id)=>{setDeepLinkPOId(id);setTab("solicitud");}}
          onGoToPagos={()=>{setDeepLinkAdminSection("pagos");setTab("admin");}}
          currentUser={currentUser} isMobile={isMobile}/>}
        {tab==="nuevo"      && <Nuevo products={pricedProducts} vendors={vendors} onAdd={addOrder} onDone={()=>setTab("central")} currentUser={currentUser} isMobile={isMobile} clients={clients} onSaveClient={saveClient} promos={promos}/>}
        {tab==="clientes"   && <ClientesPanel clients={clients} onSave={saveClient} onDelete={deleteClient} onRequestDelete={requestDeleteClient} onRejectDelete={rejectDeleteClient} currentUser={currentUser} isMobile={isMobile} orders={orders}/>}
        {tab==="cotizacion" && <Cotizaciones quotes={quotes} products={pricedProducts} vendors={vendors} onAdd={addQuote} onDel={delQuote} onConvert={convertQuoteToOrder} onExtend={extendQuote} onTabChange={setTab} currentUser={currentUser} isMobile={isMobile} clients={clients} onSaveClient={saveClient}/>}
        {tab==="precios"    && <Precios products={pricedProducts} canScan={currentUser.role==="admin"||isTestOrder(currentUser.vendedor||currentUser.name)||currentUser.barcodeEnabled}/>}
        {tab==="stock"      && <>
              {isTestUser && (
                <div style={{background:"#f5eef8",border:"1.5px solid #9b59b6",borderRadius:10,padding:"10px 16px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
                  <div>
                    <span style={{fontWeight:800,color:"#6c3483",fontSize:13,display:"inline-flex",alignItems:"center",gap:5}}><Beaker size={13} strokeWidth={2.4}/>Modo Sandbox activo</span>
                    <span style={{color:"#888",fontSize:12,marginLeft:8}}>El stock que ves es una copia paralela. No afecta el stock real.</span>
                  </div>
                  <button onClick={()=>{const sb={};products.forEach(p=>{sb[p.id]=p.stock;});updateSandboxStock(sb);}}
                    style={{padding:"6px 14px",borderRadius:8,border:"1.5px solid #9b59b6",background:"#fff",color:"#6c3483",fontWeight:700,fontSize:12,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:5}}>
                    <RefreshCw size={12} strokeWidth={2.4}/> Reiniciar sandbox
                  </button>
                </div>
              )}
              <Stock products={pricedProducts} onUpd={updProd} onDel={pid=>setProducts(p=>p.filter(x=>x.id!==pid))} onAdjust={(pid,qty)=>setProducts(p=>p.map(x=>x.id===pid?{...x,stock:x.stock+qty}:x))} isAdmin={isAdmin} addLog={addLog} stockLog={stockLog} setStockLog={setStockLog} isMobile={isMobile}
                onCrearOferta={(pid)=>{setDeepLinkOfertaProductId(pid);setDeepLinkAdminSection("ofertas");setTab("admin");}}
                onPedirReposicion={(pid,qty)=>{setDeepLinkSolicitudItem({pid,qty});setTab("solicitud");}}/>
            </>}
        {tab==="compras"    && <Compras products={products} onStock={addStock} isMobile={isMobile} canScan={currentUser.role==="admin"||isTestOrder(currentUser.vendedor||currentUser.name)||currentUser.barcodeEnabled}/>}
        {tab==="solicitud"  && <SolicitudCompra products={products} currentUser={currentUser} isAdmin={isAdmin} purchaseOrders={purchaseOrders} setPurchaseOrders={setPurchaseOrders} isMobile={isMobile} onStockExternal={addStock} addLog={addLog}
          autoOpenId={deepLinkPOId} onConsumedAutoOpen={()=>setDeepLinkPOId(null)}
          prefillItem={deepLinkSolicitudItem} onConsumedPrefillItem={()=>setDeepLinkSolicitudItem(null)}
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
        {tab==="admin"      && isAdmin && <AdminPanel users={users} setUsers={setUsers} vendors={vendors} setVendors={setVendors} products={products} setProducts={setProducts} stockLog={stockLog} setStockLog={setStockLog} notifs={notifs} setNotifs={setNotifs} activity={activity} setActivity={setActivity} orders={orders} purchaseOrders={purchaseOrders} priceLists={priceLists} setPriceLists={setPriceLists} isMobile={isMobile} sandboxStock={sandboxStock} setSandboxStock={setSandboxStock} updateSandboxStock={updateSandboxStock} promos={promos} setPromos={setPromos} settings={settings} onToggleExigirPago={toggleExigirPago} onConfirmarComprobante={confirmarComprobante} onRechazarComprobante={rechazarComprobante} onConfirmarEfectivo={confirmarEfectivo} onRechazarEfectivo={rechazarEfectivo}
          autoSection={deepLinkAdminSection} onConsumedAutoSection={()=>setDeepLinkAdminSection(null)}
          prefillProductId={deepLinkOfertaProductId} onConsumedPrefill={()=>setDeepLinkOfertaProductId(null)}/>}
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


// ─── CONSOLA DE APROBACIONES — todo lo que el admin necesita revisar, en un solo lugar ──
function ConsolaAprobAccionBtn({color,bg="#fff",children,onClick,disabled}) {
  return <button disabled={disabled} onClick={onClick} style={{padding:"6px 12px",borderRadius:7,border:`1px solid ${color}55`,background:bg,color,fontWeight:700,fontSize:11.5,cursor:disabled?"default":"pointer",opacity:disabled?0.6:1}}>{children}</button>;
}
function ConsolaAprobItem({children}) {
  return <div style={{background:"#fafafa",borderRadius:9,padding:"10px 12px",marginBottom:8}}>{children}</div>;
}
function ConsolaAprobSection({icon:Icon,color,title,n,children}) {
  if(n===0) return null;
  return (
    <div style={{marginBottom:16}}>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:9}}>
        <div style={{width:24,height:24,borderRadius:7,background:color+"18",display:"flex",alignItems:"center",justifyContent:"center"}}><Icon size={13} color={color} strokeWidth={2.3}/></div>
        <div style={{fontWeight:700,fontSize:12.5,color:"#1a1a1a"}}>{title}</div>
        <span style={{background:color,color:"#fff",borderRadius:20,fontSize:10,fontWeight:800,padding:"1px 7px"}}>{n}</span>
      </div>
      {children}
    </div>
  );
}
function ConsolaAprobEdicion({o,onApproveEditRequest,onRejectEditRequest,onApproveEdit,onRejectEdit,onViewOrder}) {
  const [busy,setBusy]=useState(false);
  const [rejecting,setRejecting]=useState(false);
  const [motivo,setMotivo]=useState("");
  const esSolicitud = o.editStatus==="solicitada";
  const aprobar = esSolicitud ? onApproveEditRequest : onApproveEdit;
  const rechazar = esSolicitud ? onRejectEditRequest : onRejectEdit;
  return (
    <ConsolaAprobItem>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
        <div>
          <div style={{fontWeight:700,fontSize:12.5,color:"#1a1a1a"}}>{o.client}</div>
          <div style={{fontSize:11,color:"#999",margin:"2px 0 5px"}}>👤 {o.vendedor} · {esSolicitud?"pide permiso para editar":"envió cambios para revisar"}</div>
        </div>
        <button onClick={()=>onViewOrder(o.id)} style={{display:"flex",alignItems:"center",gap:4,background:"none",border:"none",color:"#1a5276",fontWeight:700,fontSize:11,cursor:"pointer",padding:0,flexShrink:0}}><Eye size={12} strokeWidth={2.3}/>Ver pedido</button>
      </div>
      {o.editReason&&<div style={{fontSize:11.5,color:"#666",marginBottom:8,fontStyle:"italic"}}>"{o.editReason}"</div>}
      {!rejecting ? (
        <div style={{display:"flex",gap:6}}>
          <ConsolaAprobAccionBtn color="#1e8449" disabled={busy} onClick={async()=>{setBusy(true);try{await aprobar(o.id);toast.success(`Edición ${esSolicitud?"habilitada":"aprobada"} — ${o.client}`);}catch(e){console.warn(e);toast.error("No se pudo aprobar. Probá de nuevo.");}finally{setBusy(false);}}}>✅ Aprobar</ConsolaAprobAccionBtn>
          <ConsolaAprobAccionBtn color="#c0392b" disabled={busy} onClick={()=>setRejecting(true)}>❌ Rechazar</ConsolaAprobAccionBtn>
        </div>
      ) : (
        <div>
          <input value={motivo} onChange={e=>setMotivo(e.target.value)} placeholder="Motivo del rechazo..." style={{width:"100%",padding:"6px 9px",borderRadius:7,border:"1px solid #e5b4b0",fontSize:11.5,boxSizing:"border-box",marginBottom:6}}/>
          <div style={{display:"flex",gap:6}}>
            <ConsolaAprobAccionBtn color="#fff" bg="#c0392b" disabled={!motivo.trim()||busy} onClick={async()=>{setBusy(true);try{await rechazar(o.id,motivo);toast.info(`Edición rechazada — ${o.client}`);setRejecting(false);setMotivo("");}catch(e){console.warn(e);toast.error("No se pudo rechazar. Probá de nuevo.");}finally{setBusy(false);}}}>Confirmar rechazo</ConsolaAprobAccionBtn>
            <ConsolaAprobAccionBtn color="#999" onClick={()=>{setRejecting(false);setMotivo("");}}>Cancelar</ConsolaAprobAccionBtn>
          </div>
        </div>
      )}
    </ConsolaAprobItem>
  );
}
function ConsolaAprobPago({o,onConfirmarComprobante,onRechazarComprobante,onConfirmarEfectivo,onRechazarEfectivo,onViewOrder}) {
  const [busy,setBusy]=useState(false);
  const [rejecting,setRejecting]=useState(false);
  const [motivo,setMotivo]=useState("");
  const pt = normPagoTipo(o.pagoTipo);
  const esComprobante = pt==="comprobante_pendiente";
  return (
    <ConsolaAprobItem>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
        <div>
          <div style={{fontWeight:700,fontSize:12.5,color:"#1a1a1a"}}>{o.client}</div>
          <div style={{fontSize:11,color:"#999"}}>{esComprobante?"📎 Comprobante":"💵 Efectivo"} · 👤 {o.vendedor}</div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8,flexShrink:0}}>
          <div style={{fontWeight:800,fontSize:13,color:"#c0392b"}}>{fARS(o.total)}</div>
          <button onClick={()=>onViewOrder(o.id)} style={{display:"flex",alignItems:"center",gap:4,background:"none",border:"none",color:"#1a5276",fontWeight:700,fontSize:11,cursor:"pointer",padding:0}}><Eye size={12} strokeWidth={2.3}/>Ver</button>
        </div>
      </div>
      {!rejecting ? (
        <div style={{display:"flex",gap:6,marginTop:8}}>
          <ConsolaAprobAccionBtn color="#1e8449" disabled={busy} onClick={async()=>{setBusy(true);try{await (esComprobante?onConfirmarComprobante(o.id):onConfirmarEfectivo(o.id));toast.success(`Pago de ${o.client} confirmado`);}catch(e){console.warn(e);toast.error("No se pudo confirmar. Probá de nuevo.");}finally{setBusy(false);}}}>✅ Confirmar</ConsolaAprobAccionBtn>
          {esComprobante
            ? <ConsolaAprobAccionBtn color="#c0392b" disabled={busy} onClick={()=>setRejecting(true)}>❌ Rechazar</ConsolaAprobAccionBtn>
            : <ConsolaAprobAccionBtn color="#c0392b" disabled={busy} onClick={async()=>{setBusy(true);try{await onRechazarEfectivo(o.id);toast.info(`Pago de ${o.client} rechazado`);}catch(e){console.warn(e);toast.error("No se pudo rechazar. Probá de nuevo.");}finally{setBusy(false);}}}>❌ Rechazar</ConsolaAprobAccionBtn>}
        </div>
      ) : (
        <div style={{marginTop:8}}>
          <input value={motivo} onChange={e=>setMotivo(e.target.value)} placeholder="Motivo del rechazo..." style={{width:"100%",padding:"6px 9px",borderRadius:7,border:"1px solid #e5b4b0",fontSize:11.5,boxSizing:"border-box",marginBottom:6}}/>
          <div style={{display:"flex",gap:6}}>
            <ConsolaAprobAccionBtn color="#fff" bg="#c0392b" disabled={!motivo.trim()||busy} onClick={async()=>{setBusy(true);try{await onRechazarComprobante(o.id,motivo);toast.info(`Comprobante de ${o.client} rechazado`);setRejecting(false);setMotivo("");}catch(e){console.warn(e);toast.error("No se pudo rechazar. Probá de nuevo.");}finally{setBusy(false);}}}>Confirmar rechazo</ConsolaAprobAccionBtn>
            <ConsolaAprobAccionBtn color="#999" onClick={()=>{setRejecting(false);setMotivo("");}}>Cancelar</ConsolaAprobAccionBtn>
          </div>
        </div>
      )}
    </ConsolaAprobItem>
  );
}
function ConsolaAprobSolicitud({po,onQuickReviewPO,onViewPO}) {
  const [busy,setBusy]=useState(false);
  return (
    <ConsolaAprobItem>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <div style={{fontWeight:700,fontSize:12.5,color:"#1a1a1a"}}>Solicitud de compra</div>
          <div style={{fontSize:11,color:"#999"}}>{po.items?.length||0} productos · 👤 {po.vendedor}</div>
        </div>
        <button onClick={()=>onViewPO(po.id)} style={{display:"flex",alignItems:"center",gap:4,background:"none",border:"none",color:"#1a5276",fontWeight:700,fontSize:11,cursor:"pointer",padding:0}}><Eye size={12} strokeWidth={2.3}/>Ver detalle</button>
      </div>
      <div style={{marginTop:8}}>
        <ConsolaAprobAccionBtn color="#b7770d" disabled={busy} onClick={async()=>{setBusy(true);try{await onQuickReviewPO(po.id);toast.info("Marcada como revisando");}catch(e){console.warn(e);toast.error("No se pudo actualizar.");}finally{setBusy(false);}}}>🔍 Marcar revisando</ConsolaAprobAccionBtn>
      </div>
    </ConsolaAprobItem>
  );
}
function ConsolaAprobBaja({c,onDeleteClient,onRejectDeleteClient}) {
  const [busy,setBusy]=useState(false);
  return (
    <ConsolaAprobItem>
      <div style={{fontWeight:700,fontSize:12.5,color:"#1a1a1a"}}>{c.name}</div>
      <div style={{fontSize:11,color:"#999",margin:"2px 0 6px"}}>👤 solicitado por {c.createdBy||"—"}</div>
      {c.deleteReason&&<div style={{fontSize:11.5,color:"#666",marginBottom:8,fontStyle:"italic"}}>"{c.deleteReason}"</div>}
      <div style={{display:"flex",gap:6}}>
        <ConsolaAprobAccionBtn color="#c0392b" disabled={busy} onClick={async()=>{
          const ok = await confirmDialog("¿Eliminar cliente?",`Vas a eliminar a "${c.name}" definitivamente. Esta acción no se puede deshacer.`,true);
          if(!ok) return;
          setBusy(true);try{await onDeleteClient(c.id);toast.success("Cliente eliminado");}catch(e){console.warn(e);toast.error("No se pudo eliminar.");}finally{setBusy(false);}
        }}>🗑️ Aprobar baja</ConsolaAprobAccionBtn>
        <ConsolaAprobAccionBtn color="#999" disabled={busy} onClick={async()=>{setBusy(true);try{await onRejectDeleteClient(c.id);toast.info("Baja rechazada, el cliente sigue activo");}catch(e){console.warn(e);toast.error("No se pudo actualizar.");}finally{setBusy(false);}}}>Rechazar</ConsolaAprobAccionBtn>
      </div>
    </ConsolaAprobItem>
  );
}
function ConsolaAprobaciones({orders,clients,purchaseOrders,onApproveEditRequest,onRejectEditRequest,onApproveEdit,onRejectEdit,onConfirmarComprobante,onRechazarComprobante,onConfirmarEfectivo,onRechazarEfectivo,onDeleteClient,onRejectDeleteClient,onQuickReviewPO,onViewPO,onViewOrder}) {
  const ediciones = orders.filter(o=>o.editStatus==="solicitada"||o.editStatus==="en revisión");
  const pagos = orders.filter(o=>{const pt=normPagoTipo(o.pagoTipo); return pt==="comprobante_pendiente"||pt==="efectivo_pendiente";});
  const solicitudes = (purchaseOrders||[]).filter(po=>po.estado==="abierta");
  const bajas = (clients||[]).filter(c=>c.deleteRequested);
  const total = ediciones.length+pagos.length+solicitudes.length+bajas.length;
  const [open,setOpen] = useState(true);

  if(total===0) return null; // no molestar si no hay nada pendiente

  return (
    <div style={{background:"#fff",borderRadius:14,padding:"16px 18px",marginBottom:18,boxShadow:"0 1px 4px #0001"}}>
      <div onClick={()=>setOpen(o=>!o)} style={{display:"flex",alignItems:"center",justifyContent:"space-between",cursor:"pointer",marginBottom:open?14:0}}>
        <div style={{fontWeight:800,fontSize:14.5,color:"#1a1a1a"}}>📋 Necesita tu aprobación</div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{background:"#c0392b",color:"#fff",borderRadius:20,fontSize:11,fontWeight:800,padding:"2px 9px"}}>{total}</span>
          <ChevronDown size={14} strokeWidth={2.3} color="#bbb" style={{transform:open?"rotate(180deg)":"none",transition:"transform .15s"}}/>
        </div>
      </div>
      {open && <>
        <ConsolaAprobSection icon={Pencil} color="#1a5276" title="Ediciones de pedido" n={ediciones.length}>
          {ediciones.map(o=><ConsolaAprobEdicion key={o.id} o={o} onApproveEditRequest={onApproveEditRequest} onRejectEditRequest={onRejectEditRequest} onApproveEdit={onApproveEdit} onRejectEdit={onRejectEdit} onViewOrder={onViewOrder}/>)}
        </ConsolaAprobSection>
        <ConsolaAprobSection icon={Banknote} color="#b7770d" title="Pagos por confirmar" n={pagos.length}>
          {pagos.map(o=><ConsolaAprobPago key={o.id} o={o} onConfirmarComprobante={onConfirmarComprobante} onRechazarComprobante={onRechazarComprobante} onConfirmarEfectivo={onConfirmarEfectivo} onRechazarEfectivo={onRechazarEfectivo} onViewOrder={onViewOrder}/>)}
        </ConsolaAprobSection>
        <ConsolaAprobSection icon={Truck} color="#6c3483" title="Solicitudes de compra" n={solicitudes.length}>
          {solicitudes.map(po=><ConsolaAprobSolicitud key={po.id} po={po} onQuickReviewPO={onQuickReviewPO} onViewPO={onViewPO}/>)}
        </ConsolaAprobSection>
        <ConsolaAprobSection icon={UserMinus} color="#c0392b" title="Bajas de cliente" n={bajas.length}>
          {bajas.map(c=><ConsolaAprobBaja key={c.id} c={c} onDeleteClient={onDeleteClient} onRejectDeleteClient={onRejectDeleteClient}/>)}
        </ConsolaAprobSection>
      </>}
    </div>
  );
}

// ─── CENTRAL ──────────────────────────────────────────────────────────────────
function Central({orders,products,onStage,onDel,onSaveNote,onRequestEdit,onApproveEditRequest,onRejectEditRequest,onSubmitEdit,onApproveEdit,onRejectEdit,onUploadComprobante,onConfirmarComprobante,onRechazarComprobante,onMarcarEfectivo,onConfirmarEfectivo,onRechazarEfectivo,exigirPagoConfirmado,clients,purchaseOrders,onDeleteClient,onRejectDeleteClient,onQuickReviewPO,onViewPO,onGoToPagos,currentUser,isMobile}) {
  const isAdmin = currentUser?.role === "admin";
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
      {isAdmin && <ConsolaAprobaciones
        orders={orders} clients={clients} purchaseOrders={purchaseOrders}
        onApproveEditRequest={onApproveEditRequest} onRejectEditRequest={onRejectEditRequest}
        onApproveEdit={onApproveEdit} onRejectEdit={onRejectEdit}
        onConfirmarComprobante={onConfirmarComprobante} onRechazarComprobante={onRechazarComprobante}
        onConfirmarEfectivo={onConfirmarEfectivo} onRechazarEfectivo={onRechazarEfectivo}
        onDeleteClient={onDeleteClient} onRejectDeleteClient={onRejectDeleteClient}
        onQuickReviewPO={onQuickReviewPO} onViewPO={onViewPO}
        onViewOrder={(id)=>{setExpanded(id);setSearch("");setFStage("todos");setFVendedor("todos");
          setTimeout(()=>{document.getElementById(`order-${id}`)?.scrollIntoView({behavior:"smooth",block:"center"});},80);}}
      />}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:12,marginBottom:20}}>
        {STAGES.map(s=>{
          const c=SCFG[s], Icon=STAGE_ICONS[s], cnt=orders.filter(o=>o.stage===s&&(allSandbox||!o.isSandbox)).length;
          return (
            <div key={s} onClick={()=>setFStage(fStage===s?"todos":s)} style={{
              background:"#fff",borderRadius:12,padding:"14px 16px",
              boxShadow:"0 1px 2px rgba(20,20,20,.04), 0 4px 14px rgba(20,20,20,.06)",
              borderLeft:`3px solid ${c.color}`,cursor:"pointer",
              outline:fStage===s?`2px solid ${c.color}`:"none",
            }}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                <div style={{fontFamily:SERIF,fontSize:26,fontWeight:700,color:"#1a1a1a"}}>{cnt}</div>
                <div style={{width:30,height:30,borderRadius:9,background:c.bg,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                  <Icon size={15} color={c.color} strokeWidth={2.25}/>
                </div>
              </div>
              <div style={{fontSize:11.5,color:"#7a7a7a",marginTop:4,fontWeight:600}}>{c.label}</div>
            </div>
          );
        })}
        <div style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 2px rgba(20,20,20,.04), 0 4px 14px rgba(20,20,20,.06)",borderLeft:`3px solid ${RED}`}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
            <div style={{fontFamily:SERIF,fontSize:18,fontWeight:700,color:RED}}>{fARS(deliv)}</div>
            <div style={{width:30,height:30,borderRadius:9,background:"#fdecea",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
              <Wallet size={15} color={RED} strokeWidth={2.25}/>
            </div>
          </div>
          <div style={{fontSize:11.5,color:"#7a7a7a",marginTop:4,fontWeight:600}}>Total entregado</div>
        </div>
      </div>
      <div style={{background:"#fff",borderRadius:12,padding:14,marginBottom:14,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center",boxShadow:"0 1px 4px #0001"}}>
        <div style={{position:"relative",flex:1,minWidth:180}}>
          <Search size={15} color="#aaa" strokeWidth={2.25} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)"}}/>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar cliente o N° pedido..." style={{width:"100%",padding:"8px 12px 8px 32px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",boxSizing:"border-box"}}/>
        </div>
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
        : filtered.map(o=><OCard key={o.id} o={o} exp={expanded===o.id} toggle={()=>setExpanded(expanded===o.id?null:o.id)} getP={getP} onStage={onStage} onDel={onDel} onSaveNote={onSaveNote} onRequestEdit={onRequestEdit} onApproveEditRequest={onApproveEditRequest} onRejectEditRequest={onRejectEditRequest} onSubmitEdit={onSubmitEdit} onApproveEdit={onApproveEdit} onRejectEdit={onRejectEdit} onUploadComprobante={onUploadComprobante} onConfirmarComprobante={onConfirmarComprobante} onRechazarComprobante={onRechazarComprobante} onMarcarEfectivo={onMarcarEfectivo} onConfirmarEfectivo={onConfirmarEfectivo} onRechazarEfectivo={onRechazarEfectivo} exigirPagoConfirmado={exigirPagoConfirmado} onGoToPagos={onGoToPagos} currentUser={currentUser} products={products}/>)
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
  return <button onClick={()=>setConfirm(true)} style={{marginLeft:"auto",padding:"8px 12px",borderRadius:8,border:"1.5px solid #fcc",cursor:"pointer",background:"#fff",color:RED,fontWeight:600,fontSize:13,display:"inline-flex",alignItems:"center",gap:5}}><Trash size={12} strokeWidth={2.4}/> Eliminar</button>;
}

function OCard({o,exp,toggle,getP,onStage,onDel,onSaveNote,onRequestEdit,onApproveEditRequest,onRejectEditRequest,onSubmitEdit,onApproveEdit,onRejectEdit,onUploadComprobante,onConfirmarComprobante,onRechazarComprobante,onMarcarEfectivo,onConfirmarEfectivo,onRechazarEfectivo,exigirPagoConfirmado,onGoToPagos,currentUser,products}) {
  const isAdmin = currentUser?.role === "admin";
  const idx=STAGES.indexOf(o.stage), next=STAGES[idx+1];
  const pt = normPagoTipo(o.pagoTipo); // normaliza valores viejos ("comprobante"/"efectivo" sin sufijo)
  const pagoConfirmado = pt==="comprobante_confirmado" || pt==="efectivo_confirmado";
  const entregaBloqueada = !!exigirPagoConfirmado && next==="entregado" && !pagoConfirmado;
  const [editNote,setEditNote]=useState(false);
  const [noteVal,setNoteVal]=useState(o.internalNote||"");
  const [uploadingComp, setUploadingComp] = useState(false);
  const compFileRef = useRef();
  const compFileRefDoc = useRef();

  // Edit request states
  const [showReqForm, setShowReqForm]     = useState(false);
  const [reqReason,   setReqReason]       = useState("");
  const [showRejectForm, setShowRejectForm] = useState(false);
  const [rejectReason, setRejectReason]   = useState("");
  const [showEditMode, setShowEditMode]   = useState(false);
  const [editItems,   setEditItems]       = useState([]);
  const [editSearch,  setEditSearch]      = useState("");
  const [showFinalReject, setShowFinalReject] = useState(false);
  const [finalRejectReason, setFinalRejectReason] = useState("");
  const [saving, setSaving]               = useState(false);
  const [savingEfectivo, setSavingEfectivo] = useState(false);
  const [showRejectComp, setShowRejectComp] = useState(false);
  const [rejectCompMotivo, setRejectCompMotivo] = useState("");

  const es = o.editStatus || "";

  // Edit badge
  const EditBdg = () => {
    if(!es) return null;
    const cfg = {
      "solicitada":        {bg:"#fef9e7",color:"#b7770d",border:"#f0d080",label:"Edición solicitada",Icon:Pencil},
      "aprobada":          {bg:"#eafaf1",color:"#1e8449",border:"#a9dfbf",label:"Podés editar",Icon:CheckCircle},
      "rechazada":         {bg:"#fdecea",color:"#c0392b",border:"#f1948a",label:"Edición rechazada",Icon:XCircle},
      "en revisión":       {bg:"#eaf4fc",color:"#1a5276",border:"#aed6f1",label:"Cambios en revisión",Icon:Eye},
      "cambios rechazados":{bg:"#fdecea",color:"#c0392b",border:"#f1948a",label:"Cambios rechazados",Icon:XCircle},
    };
    const c = cfg[es]; if(!c) return null;
    return <span style={{display:"inline-flex",alignItems:"center",gap:4,background:c.bg,color:c.color,border:`1px solid ${c.border}`,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}><c.Icon size={11} strokeWidth={2.5}/>{c.label}</span>;
  };

  const handleCompFile = async (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = ""; // permitir volver a elegir el mismo archivo después
    if(!file) return;
    setUploadingComp(true);
    try {
      const { url, nombre } = await uploadComprobanteFile(o.id, file);
      await onUploadComprobante(o.id, { url, nombre, fecha: today() });
    } catch(err) {
      console.warn(err);
      toast.error("No se pudo subir el comprobante. Probá de nuevo.");
    } finally {
      setUploadingComp(false);
    }
  };

  const startEditMode = () => {
    setEditItems(o.items.map(it=>({...it})));
    setEditSearch("");
    setShowEditMode(true);
  };
  const updEditItem = (pid,qty) => setEditItems(prev=>prev.map(it=>it.pid===pid?{...it,qty:Math.max(1,qty)}:it));
  const remEditItem = pid => setEditItems(prev=>prev.filter(it=>it.pid!==pid));
  const addEditItem = (p) => setEditItems(prev=>{
    const ex = prev.find(it=>it.pid===p.id);
    if(ex) return prev.map(it=>it.pid===p.id?{...it,qty:it.qty+1}:it);
    return [...prev, {pid:p.id, name:p.name, price:p.salePrice, qty:1}];
  });
  const editSearchResults = editSearch.trim()
    ? (products||[]).filter(p=>norm(p.name).includes(norm(editSearch))||normSKU(p.id).includes(normSKU(editSearch))).slice(0,6)
    : [];
  const editTotal = editItems.reduce((s,it)=>s+it.price*it.qty,0);

  return (
    <div id={`order-${o.id}`} style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 2px rgba(20,20,20,.04), 0 4px 14px rgba(20,20,20,.06)",overflow:"hidden",marginBottom:8,borderLeft:`3px solid ${STAGE_COLORS[o.stage]||"#ccc"}`}}>
      <div onClick={toggle} style={{padding:"13px 18px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",cursor:"pointer"}}>
        <div style={{flex:1,minWidth:100}}>
          <div style={{fontWeight:700,fontSize:14,color:"#1a1a1a",letterSpacing:-0.1}}>{o.client}</div>
          <div style={{fontSize:11,color:"#999",display:"flex",gap:7,flexWrap:"wrap",alignItems:"center",marginTop:3}}>
            {o.isTest&&<span style={{background:"#f1c40f",color:"#1a1a1a",borderRadius:4,padding:"1px 6px",fontSize:10,fontWeight:800}}>TEST</span>}
            {o.isSandbox&&<span style={{display:"inline-flex",alignItems:"center",gap:3,background:"#9b59b6",color:"#fff",borderRadius:4,padding:"1px 6px",fontSize:10,fontWeight:800}}><Beaker size={9} strokeWidth={2.6}/>SANDBOX</span>}
            {o.docNum&&!o.isTest&&<span style={{fontWeight:700,color:RED,fontFamily:SERIF}}>{o.docNum}</span>}
            {o.compNum&&!o.isTest&&<span style={{fontWeight:700,color:"#1a5276",fontFamily:SERIF}}>{o.compNum}</span>}
            <span>{o.date}</span>
            {o.vendedor&&<><Dot/><span>{o.vendedor}</span></>}
            {o.internalNote&&<><Dot/><span style={{display:"inline-flex",alignItems:"center",gap:3,color:"#e67e22"}}><Pencil size={10} strokeWidth={2.5}/>Nota</span></>}
            {/* Pago — badge inline */}
            {pt==="comprobante_pendiente"  && <span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#eaf2f8",color:"#1a5276",border:"1px solid #aed6f1",borderRadius:20,padding:"2.5px 9px 2.5px 7px",fontSize:10.5,fontWeight:700}}><Paperclip size={10} strokeWidth={2.5}/>Comprobante · pendiente</span>}
            {pt==="efectivo_pendiente"     && <span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#fef9e7",color:"#b7770d",border:"1px solid #f0d080",borderRadius:20,padding:"2.5px 9px 2.5px 7px",fontSize:10.5,fontWeight:700}}><Banknote size={10} strokeWidth={2.5}/>Efectivo · pendiente</span>}
            {pagoConfirmado                && <span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#eafaf1",color:"#1e8449",border:"1px solid #a9dfbf",borderRadius:20,padding:"2.5px 9px 2.5px 7px",fontSize:10.5,fontWeight:700}}><ShieldCheck size={10.5} strokeWidth={2.5}/>Pagado</span>}
            <EditBdg/>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
          <Bdg stage={o.stage}/>
          <span style={{fontFamily:SERIF,fontWeight:700,color:RED,fontSize:16}}>{fARS(o.total)}</span>
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
              <div style={{fontWeight:800,fontSize:13,color:"#b7770d",marginBottom:4,display:"flex",alignItems:"center",gap:6}}><Pencil size={13} strokeWidth={2.4}/>Solicitud de edición</div>
              <div style={{fontSize:12,color:"#7d6608",marginBottom:10}}>Motivo: "{o.editReason}"</div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                <button onClick={async()=>{setSaving(true);try{await onApproveEditRequest(o.id);}catch(e){console.warn(e);toast.error("No se pudo aprobar. Probá de nuevo.");}finally{setSaving(false);}}}
                  style={{padding:"7px 14px",borderRadius:8,border:"none",background:"#1e8449",color:"#fff",fontWeight:700,fontSize:12,cursor:"pointer"}}>
                  <CheckCircle size={12} strokeWidth={2.4}/> Aprobar solicitud
                </button>
                {!showRejectForm && <button onClick={()=>setShowRejectForm(true)}
                  style={{padding:"7px 12px",borderRadius:8,border:"1.5px solid #f1948a",background:"#fff",color:"#c0392b",fontWeight:700,fontSize:12,cursor:"pointer"}}>
                  <XCircle size={12} strokeWidth={2.4}/> Rechazar
                </button>}
              </div>
              {showRejectForm && (
                <div style={{marginTop:10}} onClick={e=>e.stopPropagation()}>
                  <textarea value={rejectReason} onChange={e=>setRejectReason(e.target.value)}
                    placeholder="Motivo del rechazo..."
                    style={{width:"100%",padding:"8px 10px",borderRadius:8,border:"1.5px solid #f1948a",fontSize:13,resize:"vertical",minHeight:60,outline:"none",boxSizing:"border-box"}}/>
                  <div style={{display:"flex",gap:8,marginTop:6}}>
                    <button onClick={async()=>{if(!rejectReason.trim())return;setSaving(true);try{await onRejectEditRequest(o.id,rejectReason);setShowRejectForm(false);setRejectReason("");}catch(e){console.warn(e);toast.error("No se pudo rechazar. Probá de nuevo.");}finally{setSaving(false);}}}
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
              <div style={{fontWeight:800,fontSize:13,color:"#c0392b",marginBottom:3,display:"flex",alignItems:"center",gap:6}}><XCircle size={13} strokeWidth={2.4}/>Solicitud rechazada</div>
              <div style={{fontSize:12,color:"#922b21"}}>Motivo: "{o.editRejectReason}"</div>
              <button
                style={{marginTop:8,padding:"5px 12px",borderRadius:7,border:"1px solid #f1948a",background:"#fff",color:"#c0392b",fontSize:11,fontWeight:600,cursor:"pointer"}}
                onClick={async()=>{setSaving(true);try{await onRequestEdit(o.id,"");}catch(e){console.warn(e);toast.error("No se pudo enviar la solicitud. Probá de nuevo.");}finally{setSaving(false);}}}>
                Volver a solicitar
              </button>
            </div>
          )}

          {/* Vendedor: aprobada — puede editar */}
          {!isAdmin && es==="aprobada" && !showEditMode && (
            <div style={{background:"#eafaf1",border:"1.5px solid #a9dfbf",borderRadius:10,padding:"10px 14px",marginBottom:14,display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
              <div>
                <div style={{fontWeight:800,fontSize:13,color:"#1e8449",display:"flex",alignItems:"center",gap:6}}><CheckCircle size={13} strokeWidth={2.4}/>Edición aprobada</div>
                <div style={{fontSize:12,color:"#1a5276",marginTop:2}}>El admin autorizó la edición. Hacé tus cambios y enviá para revisión final.</div>
              </div>
              <button onClick={startEditMode}
                style={{padding:"8px 16px",borderRadius:8,border:"none",background:"linear-gradient(135deg,#1a5276,#2980b9)",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}>
                <Pencil size={12} strokeWidth={2.4}/> Editar pedido
              </button>
            </div>
          )}

          {/* EDIT MODE — formulario de edición */}
          {!isAdmin && es==="aprobada" && showEditMode && (
            <div style={{background:"#eaf4fc",border:"1.5px solid #aed6f1",borderRadius:10,padding:"14px",marginBottom:14}} onClick={e=>e.stopPropagation()}>
              <div style={{fontWeight:800,fontSize:13,color:"#1a5276",marginBottom:10,display:"flex",alignItems:"center",gap:6}}><Pencil size={13} strokeWidth={2.4}/>Editando pedido</div>
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
              {/* Agregar producto nuevo */}
              <div style={{position:"relative",marginBottom:10}}>
                <Search size={13} color="#aaa" strokeWidth={2.3} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)"}}/>
                <input value={editSearch} onChange={e=>setEditSearch(e.target.value)}
                  placeholder="Buscar producto para agregar..."
                  style={{width:"100%",padding:"8px 12px 8px 32px",borderRadius:8,border:"1.5px solid #aed6f1",fontSize:13,outline:"none",boxSizing:"border-box"}}/>
                {editSearch.trim() && (
                  <div style={{background:"#fff",borderRadius:8,border:"1.5px solid #d6eaf8",marginTop:4,maxHeight:220,overflowY:"auto"}}>
                    {editSearchResults.length===0
                      ? <div style={{padding:"10px 12px",fontSize:12,color:"#aaa"}}>Sin resultados</div>
                      : editSearchResults.map(p=>(
                        <div key={p.id} onClick={()=>{addEditItem(p);setEditSearch("");}}
                          style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 12px",borderBottom:"1px solid #f5f5f5",cursor:"pointer"}}
                          onMouseEnter={e=>e.currentTarget.style.background="#eaf4fc"}
                          onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontSize:13,fontWeight:600,color:"#1a1a1a"}}>{p.name}</div>
                            <div style={{fontSize:11,color:"#888"}}>{p.id} · {fARS(p.salePrice)}</div>
                          </div>
                          <span style={{color:"#1a5276",fontWeight:800,fontSize:16,flexShrink:0,marginLeft:8}}>+</span>
                        </div>
                      ))
                    }
                  </div>
                )}
              </div>
              <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:15,color:RED,padding:"10px 0",borderTop:"2px solid #d6eaf8",margin:"4px 0 10px"}}>
                <span>Nuevo total</span><span>{fARS(editTotal)}</span>
              </div>
              <button onClick={async()=>{if(!editItems.length)return;setSaving(true);try{await onSubmitEdit(o.id,editItems,editTotal);setShowEditMode(false);}catch(e){console.warn(e);toast.error("No se pudo enviar la edición. Probá de nuevo.");}finally{setSaving(false);}}}
                disabled={!editItems.length||saving}
                style={{width:"100%",padding:"10px",borderRadius:9,border:"none",background:editItems.length?"linear-gradient(135deg,#1a5276,#2980b9)":"#e5e5e5",color:editItems.length?"#fff":"#aaa",fontWeight:800,fontSize:14,cursor:editItems.length?"pointer":"not-allowed",marginBottom:8}}>
                {saving?"Enviando...":<><Send size={12} strokeWidth={2.4}/> Enviar para revisión</>}
              </button>
              <button onClick={()=>setShowEditMode(false)}
                style={{width:"100%",padding:"9px",borderRadius:9,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontSize:13,cursor:"pointer"}}>Cancelar</button>
            </div>
          )}

          {/* Admin: cambios en revisión — lista separada para mobile */}
          {isAdmin && es==="en revisión" && (
            <div style={{background:"#eaf4fc",border:"1.5px solid #aed6f1",borderRadius:10,padding:"14px",marginBottom:14}}>
              <div style={{fontWeight:800,fontSize:13,color:"#1a5276",marginBottom:12,display:"flex",alignItems:"center",gap:6}}><Eye size={13} strokeWidth={2.4}/>Revisión de cambios — {o.vendedor}</div>
              {/* Original */}
              <div style={{marginBottom:10}}>
                <div style={{fontSize:10,fontWeight:800,color:"#888",letterSpacing:1,textTransform:"uppercase",marginBottom:6}}>Original</div>
                {o.items.map((it,i)=>{const p=getP(it.pid);return <div key={i} style={{fontSize:12,color:"#555",padding:"4px 0",borderBottom:"1px solid #f0f0f0",display:"flex",justifyContent:"space-between"}}><span>{p?.name||it.name} × {it.qty}</span><span>{fARS(it.price*it.qty)}</span></div>;})}
                <div style={{fontWeight:700,fontSize:13,color:"#555",marginTop:6,textAlign:"right"}}>{fARS(o.total)}</div>
              </div>
              {/* Nuevo */}
              <div style={{background:"#fff",borderRadius:8,padding:"10px 12px",border:"1.5px solid #aed6f1",marginBottom:12}}>
                <div style={{fontSize:10,fontWeight:800,color:"#1a5276",letterSpacing:1,textTransform:"uppercase",marginBottom:6,display:"flex",alignItems:"center",gap:5}}>Nuevo <Pencil size={10} strokeWidth={2.5}/></div>
                {(o.editItems||[]).map((it,i)=>{
                  const p=getP(it.pid);
                  const orig=o.items.find(x=>x.pid===it.pid);
                  const changed=!orig||orig.qty!==it.qty;
                  return <div key={i} style={{fontSize:12,color:changed?"#1a5276":"#555",fontWeight:changed?700:400,padding:"4px 0",borderBottom:"1px solid #f0f0f0",display:"flex",justifyContent:"space-between"}}><span>{p?.name||it.name} × {it.qty}{changed&&<Pencil size={9} strokeWidth={2.6} style={{display:"inline",marginLeft:4,verticalAlign:"-1px"}}/>}</span><span>{fARS(it.price*it.qty)}</span></div>;
                })}
                <div style={{fontWeight:800,fontSize:13,color:RED,marginTop:6,textAlign:"right"}}>{fARS((o.editItems||[]).reduce((s,it)=>s+it.price*it.qty,0))}</div>
              </div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                <button onClick={async()=>{setSaving(true);try{await onApproveEdit(o.id);}catch(e){console.warn(e);toast.error("No se pudo aprobar. Probá de nuevo.");}finally{setSaving(false);}}}
                  style={{padding:"8px 16px",borderRadius:8,border:"none",background:"linear-gradient(135deg,#1a5e20,#1e8449)",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}>
                  <CheckCircle size={12} strokeWidth={2.4}/> Aprobar cambios
                </button>
                {!showFinalReject && <button onClick={()=>setShowFinalReject(true)}
                  style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #f1948a",background:"#fff",color:"#c0392b",fontWeight:700,fontSize:13,cursor:"pointer"}}>
                  <XCircle size={12} strokeWidth={2.4}/> Rechazar cambios
                </button>}
              </div>
              {showFinalReject && (
                <div style={{marginTop:10}} onClick={e=>e.stopPropagation()}>
                  <textarea value={finalRejectReason} onChange={e=>setFinalRejectReason(e.target.value)}
                    placeholder="Motivo del rechazo..."
                    style={{width:"100%",padding:"8px 10px",borderRadius:8,border:"1.5px solid #f1948a",fontSize:13,resize:"vertical",minHeight:60,outline:"none",boxSizing:"border-box"}}/>
                  <div style={{display:"flex",gap:8,marginTop:6}}>
                    <button onClick={async()=>{if(!finalRejectReason.trim())return;setSaving(true);try{await onRejectEdit(o.id,finalRejectReason);setShowFinalReject(false);setFinalRejectReason("");}catch(e){console.warn(e);toast.error("No se pudo rechazar. Probá de nuevo.");}finally{setSaving(false);}}}
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
              <Eye size={12} strokeWidth={2.4} style={{display:"inline",verticalAlign:"-2px"}}/> Tus cambios fueron enviados. El admin está revisando la edición final.
            </div>
          )}

          {/* Vendedor: cambios rechazados en revisión final */}
          {!isAdmin && es==="cambios rechazados" && (
            <div style={{background:"#fdecea",border:"1.5px solid #f1948a",borderRadius:10,padding:"10px 14px",marginBottom:14}}>
              <div style={{fontWeight:800,fontSize:13,color:"#c0392b",marginBottom:3,display:"flex",alignItems:"center",gap:6}}><XCircle size={13} strokeWidth={2.4}/>Cambios rechazados por el admin</div>
              <div style={{fontSize:12,color:"#922b21"}}>Motivo: "{o.editRejectReason}"</div>
              <div style={{fontSize:11,color:"#888",marginTop:4}}>El pedido quedó con los datos originales.</div>
            </div>
          )}

          {/* ITEMS LIST */}
          {o.items.map((it,i)=>{const p=getP(it.pid);return <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #f9f9f9",fontSize:13}}><span style={{color:"#444"}}>{p?.name||it.name} x {it.qty}</span><span style={{fontWeight:600}}>{fARS(it.price*it.qty)}</span></div>;})}
          <div style={{display:"flex",justifyContent:"flex-end",fontWeight:800,fontSize:16,color:RED,margin:"8px 0 12px"}}>{fARS(o.total)}</div>
          {o.notes&&<div style={{background:"#f9f9f9",borderRadius:8,padding:"8px 12px",fontSize:13,color:"#555",marginBottom:12}}> {o.notes}</div>}

          {/* NOTA INTERNA */}
          <div style={{background:"#fffbf0",border:"1.5px solid #f0d080",borderRadius:10,padding:"10px 14px",marginBottom:14}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6}}>
              <span style={{display:"inline-flex",alignItems:"center",gap:4,fontSize:12,fontWeight:700,color:"#b7770d"}}><Pencil size={11} strokeWidth={2.4}/>Nota interna (solo admin)</span>
              {!editNote&&<button onClick={(e)=>{e.stopPropagation();setEditNote(true);setNoteVal(o.internalNote||"");}}
                style={{padding:"3px 10px",borderRadius:6,border:"1px solid #f0d080",background:"#fff",color:"#b7770d",fontSize:11,fontWeight:700,cursor:"pointer"}}>
                {o.internalNote?<><Pencil size={11} strokeWidth={2.4}/> Editar</>:"+ Agregar"}
              </button>}
            </div>
            {editNote ? (
              <div onClick={e=>e.stopPropagation()}>
                <textarea value={noteVal} onChange={e=>setNoteVal(e.target.value)} placeholder="Escribí una nota interna..."
                  style={{width:"100%",padding:"8px 10px",borderRadius:8,border:"1.5px solid #f0d080",fontSize:13,resize:"vertical",minHeight:64,outline:"none",boxSizing:"border-box",background:"#fff"}}/>
                <div style={{display:"flex",gap:8,marginTop:6}}>
                  <button onClick={async(e)=>{e.stopPropagation();await onSaveNote(o.id,noteVal);setEditNote(false);}}
                    style={{padding:"6px 14px",borderRadius:7,border:"none",background:"#b7770d",color:"#fff",fontWeight:700,fontSize:12,cursor:"pointer"}}> Guardar</button>
                  <button onClick={(e)=>{e.stopPropagation();setEditNote(false);setNoteVal(o.internalNote||"");}}
                    style={{padding:"6px 12px",borderRadius:7,border:"1px solid #e5e5e5",background:"#fff",color:"#666",fontSize:12,cursor:"pointer"}}>Cancelar</button>
                </div>
              </div>
            ) : (
              <div style={{fontSize:13,color:o.internalNote?"#5d4037":"#aaa",fontStyle:o.internalNote?"normal":"italic"}}>{o.internalNote||"Sin nota interna"}</div>
            )}
          </div>

          {/* ── PAGO (comprobante digital + efectivo, ambos con aval del admin) ── */}
          <div style={{background:"#f8fffe",border:"1.5px solid #a9dfbf",borderRadius:10,padding:"10px 14px",marginBottom:14}} onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:12,fontWeight:700,color:"#1e8449",marginBottom:8,display:"flex",alignItems:"center",gap:5}}><CreditCard size={13} strokeWidth={2.4}/>Pago</div>

            {/* ── Sin pago registrado: dos opciones ── */}
            {!pt && (
              <div style={{display:"flex",flexDirection:"column",gap:7}}>
                <label style={{display:"flex",alignItems:"center",gap:8,padding:"11px 14px",borderRadius:9,border:"1px solid #a9dfbf",background:"#fff",color:"#1e8449",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                  <input ref={compFileRef} type="file" accept="image/*" onChange={handleCompFile} style={{display:"none"}}/>
                  {uploadingComp
                    ? <><div style={{width:14,height:14,borderRadius:"50%",border:"2px solid #a9dfbf",borderTop:"2px solid #1e8449",animation:"lm-spin 0.8s linear infinite",flexShrink:0}}/><style>{`@keyframes lm-spin{to{transform:rotate(360deg)}}`}</style>Subiendo...</>
                    : <><Camera size={12} strokeWidth={2.4}/> Foto del comprobante</>
                  }
                </label>
                <label style={{display:"flex",alignItems:"center",gap:8,padding:"11px 14px",borderRadius:9,border:"1px solid #a9dfbf",background:"#fff",color:"#1e8449",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                  <input ref={compFileRefDoc} type="file" accept="application/pdf,.pdf" onChange={handleCompFile} style={{display:"none"}}/>
                  <FileText size={12} strokeWidth={2.4}/> Subir PDF
                </label>
                <button disabled={savingEfectivo}
                  onClick={async()=>{setSavingEfectivo(true);try{await onMarcarEfectivo(o.id);}catch(e){console.warn(e);toast.error("No se pudo registrar. Probá de nuevo.");}finally{setSavingEfectivo(false);}}}
                  style={{display:"flex",alignItems:"center",gap:8,padding:"11px 14px",borderRadius:9,border:"1px solid #f0d080",background:"#fff",color:"#b7770d",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                  <Banknote size={12} strokeWidth={2.4}/> Pago en efectivo
                </button>
              </div>
            )}

            {/* ── Comprobante pendiente de revisión ── */}
            {pt==="comprobante_pendiente" && (
              <div>
                <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
                  {/\.(jpe?g|png|gif|webp)$/i.test(o.comprobanteUrl)
                    ? <img src={o.comprobanteUrl} alt="comprobante" style={{width:48,height:48,borderRadius:8,objectFit:"cover",border:"1.5px solid #aed6f1",flexShrink:0}}/>
                    : <div style={{width:48,height:48,borderRadius:8,border:"1.5px solid #aed6f1",background:"#fff",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><FileText size={20} color="#1a5276" strokeWidth={2}/></div>}
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:12,color:"#1a5276",fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{o.comprobanteNombre||"Comprobante"}</div>
                    <div style={{fontSize:11,color:"#888"}}>{o.comprobanteFecha&&`Subido el ${o.comprobanteFecha}`} · monto del pedido: <strong>{fARS(o.total)}</strong></div>
                  </div>
                  <a href={o.comprobanteUrl} target="_blank" rel="noopener noreferrer"
                    style={{padding:"5px 10px",borderRadius:6,border:"1px solid #aed6f1",background:"#fff",color:"#1a5276",fontSize:11,fontWeight:700,textDecoration:"none",flexShrink:0}}>
                    <Eye size={11} strokeWidth={2.4}/> Ver
                  </a>
                </div>

                {isAdmin ? (!showRejectComp ? (
                  <div style={{display:"flex",flexDirection:"column",gap:7}}>
                    <button disabled={savingEfectivo}
                      onClick={async()=>{setSavingEfectivo(true);try{await onConfirmarComprobante(o.id);}catch(e){console.warn(e);toast.error("No se pudo confirmar. Probá de nuevo.");}finally{setSavingEfectivo(false);}}}
                      style={{padding:"11px 14px",borderRadius:9,border:"none",background:"#1e8449",color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                      {savingEfectivo?"Guardando...":<><CheckCircle size={12} strokeWidth={2.4}/> Confirmar — el pago es correcto</>}
                    </button>
                    <button onClick={()=>setShowRejectComp(true)}
                      style={{padding:"11px 14px",borderRadius:9,border:"1.5px solid #f1948a",background:"#fff",color:"#c0392b",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                      <XCircle size={12} strokeWidth={2.4}/> Rechazar
                    </button>
                    <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
                      <label style={{display:"flex",alignItems:"center",gap:5,padding:"7px 12px",borderRadius:7,border:"1px solid #e5e5e5",background:"#fff",color:"#666",fontSize:11,fontWeight:600,cursor:"pointer"}}>
                        <input ref={compFileRef} type="file" accept="image/*" onChange={handleCompFile} style={{display:"none"}}/>
                        <RefreshCw size={10} strokeWidth={2.5}/> Reemplazar (foto)
                      </label>
                      <label style={{display:"flex",alignItems:"center",gap:5,padding:"7px 12px",borderRadius:7,border:"1px solid #e5e5e5",background:"#fff",color:"#666",fontSize:11,fontWeight:600,cursor:"pointer"}}>
                        <input ref={compFileRefDoc} type="file" accept="application/pdf,.pdf" onChange={handleCompFile} style={{display:"none"}}/>
                        <RefreshCw size={10} strokeWidth={2.5}/> Reemplazar (PDF)
                      </label>
                    </div>
                  </div>
                ) : (
                  <div style={{padding:"10px 12px",background:"#fdecea",borderRadius:8,border:"1px solid #f1948a"}}>
                    <div style={{fontSize:11.5,fontWeight:700,color:"#c0392b",marginBottom:6}}>Motivo del rechazo (el vendedor lo va a ver)</div>
                    <textarea value={rejectCompMotivo} onChange={e=>setRejectCompMotivo(e.target.value)} rows={2}
                      placeholder="Ej: el monto no coincide, la imagen no se lee..."
                      style={{width:"100%",padding:8,borderRadius:6,border:"1px solid #e5b4b0",fontSize:12.5,fontFamily:"inherit",resize:"vertical",boxSizing:"border-box"}}/>
                    <div style={{display:"flex",gap:8,marginTop:8}}>
                      <button disabled={!rejectCompMotivo.trim()||savingEfectivo}
                        onClick={async()=>{setSavingEfectivo(true);try{await onRechazarComprobante(o.id,rejectCompMotivo);setShowRejectComp(false);setRejectCompMotivo("");}catch(e){console.warn(e);toast.error("No se pudo rechazar. Probá de nuevo.");}finally{setSavingEfectivo(false);}}}
                        style={{padding:"6px 14px",borderRadius:7,border:"none",background:rejectCompMotivo.trim()?"#c0392b":"#e5b4b0",color:"#fff",fontWeight:700,fontSize:12,cursor:rejectCompMotivo.trim()?"pointer":"not-allowed"}}>
                        Confirmar rechazo
                      </button>
                      <button onClick={()=>{setShowRejectComp(false);setRejectCompMotivo("");}}
                        style={{padding:"6px 14px",borderRadius:7,border:"1px solid #ccc",background:"#fff",color:"#666",fontWeight:600,fontSize:12,cursor:"pointer"}}>
                        Cancelar
                      </button>
                    </div>
                  </div>
                )) : (
                  <div style={{display:"flex",alignItems:"center",gap:8,padding:"8px 10px",background:"#eaf2f8",border:"1px solid #aed6f1",borderRadius:8}}>
                    <Paperclip size={17} strokeWidth={2.2}/>
                    <div>
                      <div style={{fontSize:12,fontWeight:700,color:"#1a5276"}}>Comprobante en revisión</div>
                      <div style={{fontSize:11,color:"#888"}}>Esperando confirmación del admin</div>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* ── Efectivo pendiente: vendedor espera, admin decide ── */}
            {pt==="efectivo_pendiente" && (
              <div>
                {isAdmin ? (
                  <div>
                    <div style={{fontSize:13,color:"#b7770d",marginBottom:10}}>
                      <Banknote size={12} strokeWidth={2.4} style={{display:"inline",verticalAlign:"-2px"}}/> <strong>{o.vendedor}</strong> registró pago en efectivo{o.pagoEfectivoFecha&&` el ${o.pagoEfectivoFecha}`} por <strong>{fARS(o.total)}</strong>. ¿Confirmás?
                    </div>
                    <div style={{display:"flex",flexDirection:"column",gap:7}}>
                      <button disabled={savingEfectivo}
                        onClick={async()=>{setSavingEfectivo(true);try{await onConfirmarEfectivo(o.id);}catch(e){console.warn(e);toast.error("No se pudo confirmar. Probá de nuevo.");}finally{setSavingEfectivo(false);}}}
                        style={{padding:"11px 14px",borderRadius:9,border:"none",background:"#1e8449",color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                        {savingEfectivo?"Guardando...":<><CheckCircle size={12} strokeWidth={2.4}/> Confirmar efectivo</>}
                      </button>
                      <button disabled={savingEfectivo}
                        onClick={async()=>{setSavingEfectivo(true);try{await onRechazarEfectivo(o.id);}catch(e){console.warn(e);toast.error("No se pudo rechazar. Probá de nuevo.");}finally{setSavingEfectivo(false);}}}
                        style={{padding:"11px 14px",borderRadius:9,border:"1.5px solid #f1948a",background:"#fff",color:"#c0392b",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                        <XCircle size={12} strokeWidth={2.4}/> No confirmar
                      </button>
                    </div>
                  </div>
                ) : (
                  <div style={{display:"flex",alignItems:"center",gap:8,padding:"9px 10px",background:"#fef9e7",border:"1px solid #f0d080",borderRadius:8}}>
                    <Banknote size={17} strokeWidth={2.2}/>
                    <div>
                      <div style={{fontSize:12,fontWeight:700,color:"#b7770d"}}>Pago en efectivo registrado</div>
                      <div style={{fontSize:11,color:"#888"}}>Esperando confirmación del admin{o.pagoEfectivoFecha&&` · ${o.pagoEfectivoFecha}`}</div>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* ── Pagado (comprobante o efectivo, ya confirmado) — leyenda unificada ── */}
            {pagoConfirmado && (
              <div style={{display:"flex",alignItems:"center",gap:8,padding:"9px 10px",background:"#eafaf1",border:"1px solid #a9dfbf",borderRadius:8,flexWrap:"wrap"}}>
                {pt==="comprobante_confirmado" && /\.(jpe?g|png|gif|webp)$/i.test(o.comprobanteUrl)
                  ? <img src={o.comprobanteUrl} alt="comprobante" style={{width:36,height:36,borderRadius:7,objectFit:"cover",border:"1.5px solid #a9dfbf",flexShrink:0}}/>
                  : <CheckCircle size={17} strokeWidth={2.2}/>}
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontSize:12.5,fontWeight:800,color:"#1e8449",display:"flex",alignItems:"center",gap:4}}>Pagado · {pt==="comprobante_confirmado"?<><Paperclip size={11} strokeWidth={2.4}/>Comprobante</>:<><Banknote size={11} strokeWidth={2.4}/>Efectivo</>}</div>
                  <div style={{fontSize:11,color:"#888"}}>
                    {pt==="comprobante_confirmado"
                      ? (o.comprobanteFecha && `Confirmado · subido el ${o.comprobanteFecha}`)
                      : (o.pagoEfectivoFecha && `Confirmado · registrado el ${o.pagoEfectivoFecha}`)}
                  </div>
                </div>
                {pt==="comprobante_confirmado" && o.comprobanteUrl && (
                  <a href={o.comprobanteUrl} target="_blank" rel="noopener noreferrer"
                    style={{padding:"5px 10px",borderRadius:6,border:"1px solid #a9dfbf",background:"#fff",color:"#1e8449",fontSize:11,fontWeight:700,textDecoration:"none",flexShrink:0}}>
                    <Eye size={11} strokeWidth={2.4}/> Ver
                  </a>
                )}
                {isAdmin && (
                  <button disabled={savingEfectivo}
                    onClick={async()=>{
                      setSavingEfectivo(true);
                      try{ pt==="comprobante_confirmado" ? await onRechazarComprobante(o.id,"Anulado por el admin") : await onRechazarEfectivo(o.id); }
                      catch(e){console.warn(e);}
                      finally{setSavingEfectivo(false);}
                    }}
                    style={{padding:"5px 9px",borderRadius:6,border:"1px solid #f1948a",background:"#fff",color:"#c0392b",fontSize:10.5,fontWeight:700,cursor:"pointer",flexShrink:0}}>
                    Anular
                  </button>
                )}
              </div>
            )}
          </div>

          {/* ACTIONS */}
          <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
            {next && !es && entregaBloqueada ? (
              <div style={{width:"100%"}}>
                <button disabled style={{padding:"8px 14px",borderRadius:8,border:"none",cursor:"not-allowed",background:"#e8e8e8",color:"#aaa",fontWeight:700,fontSize:13}}>
                  <Lock size={12} strokeWidth={2.4}/> {SCFG[next].label} (bloqueado)
                </button>
                <div style={{fontSize:11.5,color:"#b7770d",marginTop:5}}>
                  <AlertTriangle size={11} strokeWidth={2.4} style={{display:"inline",verticalAlign:"-2px"}}/> Esperando confirmación de pago para habilitar la entrega
                </div>
                {isAdmin && onGoToPagos && (
                  <button onClick={onGoToPagos} style={{marginTop:7,display:"flex",alignItems:"center",gap:5,background:"none",border:"none",color:"#1a5276",fontWeight:700,fontSize:12,cursor:"pointer",padding:0}}>
                    Ver en Pagos <ArrowRightIcon size={12} strokeWidth={2.3}/>
                  </button>
                )}
              </div>
            ) : (
              next&&!es&&<button onClick={()=>onStage(o.id,next)} style={{padding:"8px 14px",borderRadius:8,border:"none",cursor:"pointer",background:SCFG[next].color,color:"#fff",fontWeight:700,fontSize:13}}>{SCFG[next].icon} Pasar a {SCFG[next].label}</button>
            )}
            {idx>0&&o.stage!=="entregado"&&!es&&<button onClick={()=>onStage(o.id,STAGES[idx-1])} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",cursor:"pointer",background:"#fff",color:"#666",fontWeight:600,fontSize:13,display:"inline-flex",alignItems:"center",gap:5}}><ArrowLeftIcon size={12} strokeWidth={2.4}/> Retroceder</button>}
            <button onClick={()=>printDoc(o, o.stage==="reserva"?"reserva":"confirmado")} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #d6eaf8",cursor:"pointer",background:"#fff",color:"#1a5276",fontWeight:600,fontSize:13}}>
              <Printer size={12} strokeWidth={2.4} style={{display:"inline",verticalAlign:"-2px"}}/> {o.stage==="reserva"?(o.docNum||"Imprimir"):(o.compNum||"Imprimir")}
            </button>
            <button onClick={()=>window.open(`https://wa.me/?text=${encodeURIComponent(buildWAOrder(o))}`, "_blank")}
              style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #d5f5e3",cursor:"pointer",background:"#25D366",color:"#fff",fontWeight:700,fontSize:13,display:"flex",alignItems:"center",gap:5}}>
              <MessageCircle size={14} strokeWidth={2.3}/> Enviar por WhatsApp
            </button>
            {/* Solicitar edición — solo vendedor, solo si no hay edición en curso */}
            {!isAdmin && !es && o.stage!=="entregado" && (
              !showReqForm
                ? <button onClick={(e)=>{e.stopPropagation();setShowReqForm(true);}}
                    style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #aed6f1",background:"#fff",color:"#1a5276",fontWeight:600,fontSize:13,cursor:"pointer"}}>
                    <Pencil size={12} strokeWidth={2.4}/> Solicitar edición
                  </button>
                : <div style={{width:"100%",marginTop:8}} onClick={e=>e.stopPropagation()}>
                    <textarea value={reqReason} onChange={e=>setReqReason(e.target.value)}
                      placeholder="Motivo de la edición..."
                      style={{width:"100%",padding:"8px 10px",borderRadius:8,border:"1.5px solid #aed6f1",fontSize:13,resize:"vertical",minHeight:60,outline:"none",boxSizing:"border-box"}}/>
                    <div style={{display:"flex",gap:8,marginTop:6}}>
                      <button onClick={async()=>{if(!reqReason.trim())return;setSaving(true);try{await onRequestEdit(o.id,reqReason);setShowReqForm(false);setReqReason("");}catch(e){console.warn(e);toast.error("No se pudo enviar la solicitud. Probá de nuevo.");}finally{setSaving(false);}}}
                        disabled={!reqReason.trim()||saving}
                        style={{padding:"6px 14px",borderRadius:7,border:"none",background:reqReason.trim()?"#1a5276":"#e5e5e5",color:reqReason.trim()?"#fff":"#aaa",fontWeight:700,fontSize:12,cursor:reqReason.trim()?"pointer":"not-allowed"}}>
                        {saving?"Enviando...":<><Send size={12} strokeWidth={2.4}/> Enviar solicitud</>}
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
    if(!form.name.trim()||!form.phone.trim()) { toast.error("Nombre y teléfono son obligatorios"); return; }
    setSaving(true);
    const newClient = {id:genId(), name:form.name.trim(), phone:form.phone.trim(), email:"", cuit:"", address:"", notes:"", deleteRequested:false, deleteReason:"", createdBy:currentUser.name, createdAt:today()};
    await onSaveClient(newClient);
    onSelect(newClient);
    setSaving(false);
  };

  return (
    <div>
      <div style={{position:"relative",marginBottom:8}}>
        <Search size={13} color="#aaa" strokeWidth={2.3} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)"}}/>
        <input value={search} onChange={e=>setSearch(e.target.value)}
          placeholder="Buscar cliente por nombre o teléfono..."
          style={{width:"100%",padding:"10px 12px 10px 32px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",boxSizing:"border-box"}}/>
      </div>

      {/* Results */}
      <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:10}}>
        {found.map(c=>(
          <div key={c.id} onClick={()=>onSelect(c)}
            style={{background:"#f9f9f9",borderRadius:10,padding:"11px 14px",cursor:"pointer",border:"1.5px solid #f0f0f0",display:"flex",alignItems:"center",gap:10}}
            onMouseEnter={e=>e.currentTarget.style.borderColor="#c0392b"}
            onMouseLeave={e=>e.currentTarget.style.borderColor="#f0f0f0"}>
            <div style={{width:34,height:34,borderRadius:9,background:"#f4f6f9",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><Users size={16} color="#888" strokeWidth={2.1}/></div>
            <div style={{flex:1}}>
              <div style={{fontWeight:700,fontSize:13}}>{c.name}</div>
              <div style={{fontSize:11,color:"#888",display:"flex",alignItems:"center",gap:4}}><Phone size={10} strokeWidth={2.4}/>{c.phone}{c.email&&<><span style={{margin:"0 2px"}}>·</span><Mail size={10} strokeWidth={2.4}/>{c.email}</>}</div>
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
            style={{width:"100%",padding:"10px",borderRadius:10,border:"1.5px dashed #aed6f1",background:"#f0f8ff",color:"#1a5276",fontWeight:700,fontSize:13,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
            <Plus size={14} strokeWidth={2.4}/> Cliente nuevo
          </button>
        : <div style={{background:"#f0f8ff",border:"1.5px solid #aed6f1",borderRadius:10,padding:"12px 14px"}}>
            <div style={{fontWeight:700,fontSize:12,color:"#1a5276",marginBottom:10,display:"flex",alignItems:"center",gap:5}}><Plus size={12} strokeWidth={2.4}/>Nuevo cliente rápido</div>
            <Field label="Nombre / Razón social *">
              <input value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} placeholder="Ej: Librería Nueva" style={inputStyle}/>
            </Field>
            <Field label="Teléfono *">
              <input value={form.phone} onChange={e=>setForm(f=>({...f,phone:e.target.value}))} placeholder="+54 11 1234-5678" style={inputStyle}/>
            </Field>
            <div style={{fontSize:11,color:"#888",marginBottom:10}}>Los datos opcionales (CUIT, email, dirección) se completan desde la sección Clientes.</div>
            <div style={{display:"flex",gap:8}}>
              <button onClick={createAndSelect} disabled={saving||!form.name.trim()||!form.phone.trim()}
                style={{flex:1,padding:"9px",borderRadius:8,border:"none",background:(!form.name.trim()||!form.phone.trim())?"#e5e5e5":"#1a5276",color:(!form.name.trim()||!form.phone.trim())?"#aaa":"#fff",fontWeight:700,fontSize:13,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
                {saving?"Creando...":<><CheckCircle size={13} strokeWidth={2.4}/> Crear y usar</>}
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
        <ArrowLeftIcon size={12} strokeWidth={2.4}/> Volver a clientes
      </button>

      {/* Header cliente */}
      <div style={{background:"#fff",borderRadius:12,padding:"16px 18px",marginBottom:12,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:16}}>{client.name}</div>
        <div style={{fontSize:12,color:"#888",marginTop:3,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
          {client.phone&&<span style={{display:"inline-flex",alignItems:"center",gap:4}}><Phone size={11} strokeWidth={2.4}/>{client.phone}</span>}
          {client.email&&<span style={{display:"inline-flex",alignItems:"center",gap:4}}><Mail size={11} strokeWidth={2.4}/>{client.email}</span>}
          {client.cuit&&<span style={{display:"inline-flex",alignItems:"center",gap:4}}><IdCard size={11} strokeWidth={2.4}/>{client.cuit}</span>}
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
            <div style={{marginBottom:8,display:"flex",justifyContent:"center"}}><ClipboardList size={32} color="#ddd" strokeWidth={1.8}/></div>
            <div style={{fontWeight:600}}>Sin pedidos en los últimos {period==="30"?"30 días":period==="180"?"6 meses":"año"}</div>
          </div>
        : clientOrders.map(o=>(
          <div key={o.id} style={{background:"#fff",borderRadius:12,padding:"13px 16px",marginBottom:8,boxShadow:"0 1px 4px #0001"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:4}}>
              <div>
                <div style={{fontWeight:700,fontSize:13}}>{o.docNum||o.compNum||o.id.slice(0,8)}</div>
                <div style={{fontSize:11,color:"#aaa"}}>{o.date} · {o.vendedor}</div>
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

function ClientesPanel({clients, onSave, onDelete, onRequestDelete, onRejectDelete, currentUser, isMobile, orders=[]}) {
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
    if(!form.name.trim()||!form.phone.trim()){toast.error("Nombre y teléfono son obligatorios");return;}
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
        <button onClick={()=>{cancelEdit();setView("lista");}} style={{flex:1,padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="lista"?`linear-gradient(135deg,#922b21,#c0392b)`:"transparent",color:view==="lista"?"#fff":"#555",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
          <Users size={13} strokeWidth={2.3}/> Clientes ({clients.length})
        </button>
        <button onClick={()=>setView("form")} style={{flex:1,padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="form"?`linear-gradient(135deg,#922b21,#c0392b)`:"transparent",color:view==="form"?"#fff":"#555",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
          {editing?<><Pencil size={12} strokeWidth={2.4}/> Editando</>:<><Plus size={13} strokeWidth={2.4}/> Nuevo cliente</>}
        </button>
      </div>

      {/* HISTORIAL */}
      {view==="historial" && historialClient && (
        <ClientHistorial client={historialClient} orders={orders} period={historialPeriod} setPeriod={setHistorialPeriod} onBack={()=>setView("lista")}/>
      )}

      {/* LISTA */}
      {view==="lista" && <>
        <div style={{position:"relative",marginBottom:12}}>
          <Search size={13} color="#aaa" strokeWidth={2.3} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)"}}/>
          <input value={search} onChange={e=>setSearch(e.target.value)}
            placeholder="Buscar por nombre, teléfono o CUIT..."
            style={{width:"100%",padding:"10px 12px 10px 32px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",boxSizing:"border-box"}}/>
        </div>
        {filtered.length===0
          ? <div style={{textAlign:"center",padding:40,color:"#aaa",background:"#fff",borderRadius:12}}>
              <div style={{marginBottom:8,display:"flex",justifyContent:"center"}}><Users size={36} color="#ddd" strokeWidth={1.8}/></div>
              <div>{search?"No se encontraron clientes":"No hay clientes aún. ¡Agregá el primero!"}</div>
            </div>
          : filtered.map(c=>(
            <div key={c.id} style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 4px #0001",marginBottom:8,border:"1.5px solid #f0f0f0"}}>
              <div style={{display:"flex",alignItems:"flex-start",gap:10}}>
                <div style={{width:40,height:40,borderRadius:11,background:"#f4f6f9",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><Users size={19} color="#888" strokeWidth={2}/></div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:700,fontSize:14}}>{c.name}</div>
                  <div style={{fontSize:12,color:"#555",marginTop:3,display:"flex",alignItems:"center",gap:4}}><Phone size={11} strokeWidth={2.4}/>{c.phone}</div>
                  {c.email&&<div style={{fontSize:12,color:"#888",display:"flex",alignItems:"center",gap:4}}><Mail size={11} strokeWidth={2.4}/>{c.email}</div>}
                  {c.cuit&&<div style={{fontSize:12,color:"#888",display:"flex",alignItems:"center",gap:4}}><IdCard size={11} strokeWidth={2.4}/>CUIT: {c.cuit}</div>}
                  {c.address&&<div style={{fontSize:12,color:"#888",display:"flex",alignItems:"center",gap:4}}><MapPin size={11} strokeWidth={2.4}/>{c.address}</div>}
                  {c.notes&&<div style={{fontSize:11,color:"#aaa",marginTop:4,fontStyle:"italic"}}>"{c.notes}"</div>}
                  {c.deleteRequested&&<div style={{background:"#fef9e7",border:"1px solid #f0d080",borderRadius:6,padding:"4px 8px",fontSize:11,color:"#b7770d",fontWeight:600,marginTop:6,display:"flex",alignItems:"center",gap:5,flexWrap:"wrap"}}>
                    <Clock size={11} strokeWidth={2.4}/>Solicitud de baja pendiente — "{c.deleteReason}"
                    {isAdmin&&<>
                      <button onClick={()=>onDelete(c.id)} style={{marginLeft:8,padding:"2px 8px",borderRadius:4,border:"none",background:"#c0392b",color:"#fff",fontSize:10,fontWeight:700,cursor:"pointer"}}>Aprobar baja</button>
                      <button onClick={()=>onRejectDelete(c.id)} style={{marginLeft:6,padding:"2px 8px",borderRadius:4,border:"1px solid #ccc",background:"#fff",color:"#666",fontSize:10,fontWeight:700,cursor:"pointer"}}>Rechazar</button>
                    </>}
                  </div>}
                </div>
                <div style={{display:"flex",flexDirection:"column",gap:5,flexShrink:0}}>
                  <button onClick={()=>{setHistorialClient(c);setView("historial");}} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #eaf4fc",background:"#fff",color:"#1a5276",cursor:"pointer",fontSize:11,fontWeight:600,display:"flex",alignItems:"center",justifyContent:"center"}}><ClipboardList size={13} strokeWidth={2.3}/></button>
            <button onClick={()=>startEdit(c)} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:11,fontWeight:600,display:"flex",alignItems:"center",justifyContent:"center"}}><Pencil size={12} strokeWidth={2.3}/></button>
                  {isAdmin
                    ? <button onClick={async()=>{if(await confirmDialog("¿Eliminar cliente?",`Vas a eliminar a "${c.name}" definitivamente.`,true))onDelete(c.id);}} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #fcc",background:"#fff",color:"#c0392b",cursor:"pointer",fontSize:11,display:"flex",alignItems:"center",justifyContent:"center"}}><Trash size={12} strokeWidth={2.3}/></button>
                    : !c.deleteRequested&&<button onClick={()=>setDelForm({id:c.id,reason:""})} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #fcc",background:"#fff",color:"#c0392b",cursor:"pointer",fontSize:10,fontWeight:600}}>Solicitar baja</button>
                  }
                </div>
              </div>
              {delForm.id===c.id&&<div style={{marginTop:10,background:"#fef9e7",border:"1.5px solid #f0d080",borderRadius:8,padding:"10px 12px"}} onClick={e=>e.stopPropagation()}>
                <div style={{fontSize:12,fontWeight:700,color:"#b7770d",marginBottom:6}}>Motivo de la solicitud de baja</div>
                <input value={delForm.reason} onChange={e=>setDelForm(f=>({...f,reason:e.target.value}))} placeholder="Ej: Cliente duplicado..." style={{...inputStyle,fontSize:12,marginBottom:8}}/>
                <div style={{display:"flex",gap:6}}>
                  <button onClick={async()=>{if(!delForm.reason.trim()){toast.error("Escribí un motivo");return;}await onRequestDelete(delForm.id,delForm.reason);setDelForm({id:null,reason:""}); }}
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
          <div style={{fontWeight:800,fontSize:15,marginBottom:14,display:"flex",alignItems:"center",gap:7}}>{editing?<><Pencil size={14} strokeWidth={2.4}/>Editar cliente</>:<><Plus size={15} strokeWidth={2.4}/>Nuevo cliente</>}</div>
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
            <button onClick={save} disabled={saving} style={{flex:1,padding:"11px",borderRadius:10,border:"none",background:"#c0392b",color:"#fff",fontWeight:800,cursor:"pointer",fontSize:14,display:"flex",alignItems:"center",justifyContent:"center",gap:7}}>{saving?"Guardando...":editing?<><Save size={14} strokeWidth={2.3}/>Guardar cambios</>:<><CheckCircle size={14} strokeWidth={2.3}/>Crear cliente</>}</button>
            <button onClick={cancelEdit} style={{padding:"11px 16px",borderRadius:10,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontWeight:600,color:"#666"}}>Cancelar</button>
          </div>
        </div>
      )}
    </div>
  );
}

function Precios({products,canScan}) {
  const [search,setSearch]=useState("");
  const [sortBy,setSortBy]=useState("name");
  const [showScanner,setShowScanner]=useState(false);
  const shown=useMemo(()=>{
    const q=norm(search);
    return products
      .filter(p=>!q||norm(p.name).includes(q)||normSKU(p.id).includes(normSKU(search)))
      .sort((a,b)=>sortBy==="name"?a.name.localeCompare(b.name):b.salePrice-a.salePrice)
      .slice(0,200);
  },[products,search,sortBy]);
  const handleBarcode = (code) => {
    setShowScanner(false);
    const p = products.find(p=>(p.barcode&&p.barcode===code) || normSKU(p.id)===normSKU(code));
    if(p) { setSearch(p.id); toast.success(`${p.name} — ${fARS(p.salePrice)}`); }
    else toast.error(`Código "${code}" no encontrado en el catálogo`);
  };
  return (
    <div>
      {showScanner && <BarcodeScanner onDetected={handleBarcode} onClose={()=>setShowScanner(false)}/>}
      <div style={{background:"#fff",borderRadius:12,padding:16,marginBottom:14,boxShadow:"0 1px 4px #0001"}}>
        <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
          {canScan && <button onClick={()=>setShowScanner(true)}
            style={{padding:"8px 14px",borderRadius:8,border:"none",background:"linear-gradient(135deg,#1a5276,#2980b9)",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer",display:"flex",alignItems:"center",gap:6,whiteSpace:"nowrap"}}>
            <Camera size={13} strokeWidth={2.3}/> Escanear
          </button>}
          <div style={{position:"relative",flex:1,minWidth:160}}>
            <Search size={13} color="#aaa" strokeWidth={2.3} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)"}}/>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar producto o código..."
              style={{width:"100%",padding:"8px 12px 8px 32px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",boxSizing:"border-box"}}/>
          </div>
          <div style={{display:"flex",gap:6}}>
            {[{v:"name",l:"A-Z"},{v:"price",l:"Precio"}].map(opt=>(
              <button key={opt.v} onClick={()=>setSortBy(opt.v)} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid",cursor:"pointer",fontSize:12,fontWeight:600,borderColor:sortBy===opt.v?RED:"#e5e5e5",background:sortBy===opt.v?"#fdecea":"#fff",color:sortBy===opt.v?RED:"#666"}}>{opt.l}</button>
            ))}
          </div>
          <div style={{fontSize:12,color:"#aaa",whiteSpace:"nowrap"}}>{shown.length} producto{shown.length!==1?"s":""}{shown.length===200?" (máx 200)":""}</div>
        </div>
      </div>
      {shown.length===0
        ? <div style={{textAlign:"center",padding:60,color:"#aaa",background:"#fff",borderRadius:12}}><div style={{marginBottom:8,display:"flex",justifyContent:"center"}}><Search size={40} color="#ddd" strokeWidth={1.7}/></div><div>No se encontraron productos</div></div>
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
function ProductSelector({products,cart,setCart,isMobile,promos=[],loteMode=false}) {
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

  const combos = useMemo(()=>getVigentCombos(promos),[promos]);
  const shownCombos = search.trim() ? combos.filter(c=>norm(c.nombre).includes(norm(search))) : combos;

  // ── Productos regulares (no-combo) ──
  // En loteMode (Solicitud de Compra), el proveedor exige pedir en multiplos de caja (columna UNIDAD del Excel)
  const step = p => loteMode ? (p.multiploCompra||1) : 1;
  const addC=p=>setCart(c=>{
    const promo = getProductPromo(promos,p.id);
    const ex = c.find(i=>i.cartKey===p.id);
    if(ex) {
      const newQty = ex.qty+step(p);
      const {disc,label} = computeAutoDisc(promo,newQty);
      return c.map(i=>i.cartKey===p.id?{...i,qty:newQty,disc,promoLabel:label}:i);
    }
    const qtyInicial = step(p);
    const {disc,label} = computeAutoDisc(promo,qtyInicial);
    return [...c,{pid:p.id,id:p.id,cartKey:p.id,qty:qtyInicial,price:p.salePrice,name:p.name,disc,promoLabel:label}];
  });
  const setQ=(cartKey,qty)=>{
    if(qty<=0){setCart(c=>c.filter(i=>i.cartKey!==cartKey));return;}
    setCart(c=>c.map(i=>{
      if(i.cartKey!==cartKey) return i;
      if(i.comboId) return {...i,qty}; // los combos ajustan su cantidad como set completo, ver setComboQty
      const promo = getProductPromo(promos,i.pid);
      const {disc,label} = computeAutoDisc(promo,qty);
      return {...i,qty,disc,promoLabel:label};
    }));
  };

  // ── Combos ──
  const comboQtyOf = (promo) => {
    const comp0 = promo.data?.componentes?.[0];
    if(!comp0) return 0;
    const item = cart.find(i=>i.comboId===promo.id && i.pid===comp0.pid);
    return item ? Math.round(item.qty/comp0.qty) : 0;
  };
  const setComboQty = (promo, newQty) => {
    setCart(c=>{
      const without = c.filter(i=>i.comboId!==promo.id);
      if(newQty<=0) return without;
      const componentes = promo.data?.componentes||[];
      const precioNormal = componentes.reduce((s,comp)=>{const pr=products.find(x=>x.id===comp.pid);return s+(pr?pr.salePrice*comp.qty:0);},0);
      const pct = precioNormal>0 ? Math.max(0,(1-(promo.data?.precioFijo||0)/precioNormal)*100) : 0;
      const nuevos = componentes.map(comp=>{
        const pr = products.find(x=>x.id===comp.pid);
        return {pid:comp.pid,cartKey:promo.id+":"+comp.pid,qty:comp.qty*newQty,price:pr?pr.salePrice:0,name:pr?pr.name:comp.pid,
          disc:{type:"%",value:Math.round(pct*10)/10},promoLabel:`🎁 ${promo.nombre}`,comboId:promo.id,comboNombre:promo.nombre};
      });
      return [...without,...nuevos];
    });
  };

  return (
    <div>
      <div style={{background:"#fff",borderRadius:12,padding:16,marginBottom:12,boxShadow:"0 1px 4px #0001"}}>
        <div style={{position:"relative",marginBottom:10}}>
          <Search size={14} color="#aaa" strokeWidth={2.3} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)"}}/>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar por nombre o código..." style={{width:"100%",padding:"8px 12px 8px 32px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",boxSizing:"border-box"}}/>
        </div>
        <div style={{position:"relative"}}>
          <button onClick={()=>setCatOpen(o=>!o)} style={{display:"flex",alignItems:"center",justifyContent:"space-between",width:"100%",padding:"8px 12px",borderRadius:8,border:`1.5px solid ${catOpen?RED:"#e5e5e5"}`,background:cat!=="todos"?"#fdecea":"#fff",color:cat!=="todos"?RED:"#666",cursor:"pointer",fontSize:13,fontWeight:600}}>
            <span style={{display:"flex",alignItems:"center",gap:6}}><Tag size={13} strokeWidth={2.2}/> {cat==="todos"?"Todas las categorías":cat}</span><span style={{fontSize:10,marginLeft:6}}>{catOpen?"▲":"▼"}</span>
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
          <Package size={13} strokeWidth={2.2}/> {soloStock?"Solo con stock":"Mostrar solo con stock"}{soloStock&&<CheckCircle size={12} strokeWidth={2.4}/>}
        </button>
      </div>
      {isMobile
        ? <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {shownCombos.map(promo=>{const cq=comboQtyOf(promo); return (
              <div key={promo.id} style={{background:"#fffdf8",borderRadius:10,padding:"12px 14px",display:"flex",alignItems:"center",gap:10,boxShadow:"0 1px 4px #0001",border:cq>0?`2px solid ${RED}`:"2px solid #c9a96a"}}>
                <div style={{flex:1,minWidth:0}}>
                  <span style={{background:"#7b1a1a",color:"#fff",fontSize:10,fontWeight:800,borderRadius:6,padding:"2px 7px",display:"inline-flex",alignItems:"center",gap:4}}><Gift size={10} strokeWidth={2.6}/> COMBO</span>
                  <div style={{fontWeight:700,fontSize:12,color:"#1a1a1a",lineHeight:1.3,margin:"4px 0 2px"}}>{promo.nombre}</div>
                  <div style={{fontSize:11,color:"#666"}}>{(promo.data?.componentes||[]).map(c=>products.find(x=>x.id===c.pid)?.name).filter(Boolean).join(" + ")}</div>
                  <span style={{fontSize:15,fontWeight:800,color:RED,marginTop:4,display:"block"}}>{fARS(promo.data?.precioFijo||0)}</span>
                </div>
                {cq>0
                  ? <div style={{display:"flex",alignItems:"center",gap:6,flexShrink:0}}>
                      <button onClick={()=>setComboQty(promo,cq-1)} style={{width:30,height:30,borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:16,fontWeight:700}}>−</button>
                      <span style={{minWidth:24,textAlign:"center",fontWeight:800,fontSize:14}}>{cq}</span>
                      <button onClick={()=>setComboQty(promo,cq+1)} style={{width:30,height:30,borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:16,fontWeight:700}}>+</button>
                    </div>
                  : <button onClick={()=>setComboQty(promo,1)} style={{padding:"7px 12px",borderRadius:7,border:"none",cursor:"pointer",background:"#7b1a1a",color:"#fff",fontWeight:700,fontSize:12,flexShrink:0}}>+ Agregar</button>
                }
              </div>
            );})}
            {shown.map(p=>{
              const ic=cart.find(i=>i.cartKey===p.id);
              const promo=getProductPromo(promos,p.id);
              const isDesc=promo?.tipo==="descuento";
              const finalPrice=isDesc?Math.max(0,promo.data?.tipoValor==="%"?p.salePrice*(1-(promo.data?.valor||0)/100):p.salePrice-(promo.data?.valor||0)):p.salePrice;
              return (
              <div key={p.id} style={{background:"#fff",borderRadius:10,padding:"12px 14px",display:"flex",alignItems:"center",gap:10,boxShadow:"0 1px 4px #0001",border:ic?`2px solid ${RED}`:promo?"2px solid #f3d98a":"2px solid transparent"}}>
                <div style={{width:46,height:46,borderRadius:8,background:p.imageUrl?"transparent":"#f4f6f9",overflow:"hidden",flexShrink:0,border:"1px solid #f0f0f0",display:"flex",alignItems:"center",justifyContent:"center"}}>
                  {p.imageUrl ? <img src={p.imageUrl} alt="" style={{width:"100%",height:"100%",objectFit:"cover"}}/> : <ImageIcon size={18} color="#ddd" strokeWidth={1.6}/>}
                </div>
                <div style={{flex:1,minWidth:0}}>
                  {promo && <span style={{fontSize:10,fontWeight:800,borderRadius:6,padding:"1px 6px",marginBottom:3,display:"inline-block",background:promo.tipo==="3x2"?"#fef9e7":"#fdecea",color:promo.tipo==="3x2"?"#b7770d":"#c0392b",border:promo.tipo==="3x2"?"1px solid #f3d98a":"1px solid #f5b7b1"}}><span style={{display:"inline-flex",alignItems:"center",gap:3}}>{promo.tipo==="3x2"?<Tag size={10} strokeWidth={2.6}/>:<TrendDown size={10} strokeWidth={2.6}/>} {promo.tipo==="3x2"?`${promo.data.comprar}×${promo.data.pagar}`:`-${promo.data.valor}${promo.data.tipoValor==="%"?"%":""}`}</span></span>}
                  <div style={{fontWeight:700,fontSize:12,color:"#1a1a1a",lineHeight:1.3,marginBottom:2}}>{p.name}</div>
                  <div style={{fontSize:11,color:"#666"}}>{p.id} · {p.category}</div>
                  {loteMode && p.multiploCompra>1 && <span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#fef9e7",color:"#b7770d",border:"1px solid #f0d080",borderRadius:20,padding:"1.5px 8px",fontSize:10,fontWeight:700,marginTop:3}}><BoxIcon size={10} strokeWidth={2.4}/>Caja de {p.multiploCompra}</span>}
                  <div style={{display:"flex",alignItems:"center",gap:8,marginTop:4}}>
                    {isDesc&&<span style={{fontSize:11,color:"#aaa",textDecoration:"line-through"}}>{fARS(p.salePrice)}</span>}
                    <span style={{fontSize:15,fontWeight:800,color:RED}}>{fARS(finalPrice)}</span>
                    <SPill n={p.stock}/>
                  </div>
                </div>
                {ic
                  ? <div style={{display:"flex",alignItems:"center",gap:6,flexShrink:0}}>
                      <button onClick={()=>setQ(p.id,ic.qty-step(p))} style={{width:30,height:30,borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:16,fontWeight:700}}>−</button>
                      <span style={{minWidth:24,textAlign:"center",fontWeight:800,fontSize:14}}>{ic.qty}</span>
                      <button onClick={()=>setQ(p.id,ic.qty+step(p))} style={{width:30,height:30,borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:16,fontWeight:700}}>+</button>
                    </div>
                  : <button onClick={()=>addC(p)} style={{padding:"7px 12px",borderRadius:7,border:"none",cursor:"pointer",background:RED,color:"#fff",fontWeight:700,fontSize:12,flexShrink:0}}>+ Agregar</button>
                }
              </div>
            );})}
          </div>
        : <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(195px,1fr))",gap:10}}>
            {shownCombos.map(promo=>{const cq=comboQtyOf(promo); return (
              <div key={promo.id} style={{background:"#fffdf8",borderRadius:10,padding:14,border:cq>0?`2px solid ${RED}`:"2px solid #c9a96a",boxShadow:"0 1px 4px #0001"}}>
                <div style={{marginBottom:6}}><span style={{background:"#7b1a1a",color:"#fff",fontSize:10,fontWeight:800,borderRadius:6,padding:"2px 7px",display:"inline-flex",alignItems:"center",gap:4}}><Gift size={10} strokeWidth={2.6}/> COMBO</span></div>
                <div style={{fontWeight:700,fontSize:12,color:"#1a1a1a",marginBottom:3,lineHeight:1.3}}>{promo.nombre}</div>
                <div style={{fontSize:11,color:"#666",marginBottom:7}}>{(promo.data?.componentes||[]).map(c=>products.find(x=>x.id===c.pid)?.name).filter(Boolean).join(" + ")}</div>
                <div style={{marginBottom:10}}><span style={{fontSize:17,fontWeight:800,color:RED}}>{fARS(promo.data?.precioFijo||0)}</span></div>
                {cq>0?<div style={{display:"flex",alignItems:"center",gap:5}}>
                  <button onClick={()=>setComboQty(promo,cq-1)} style={{width:27,height:27,borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:15,fontWeight:700}}>−</button>
                  <span style={{flex:1,textAlign:"center",fontWeight:800,fontSize:13}}>{cq}</span>
                  <button onClick={()=>setComboQty(promo,cq+1)} style={{width:27,height:27,borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:15,fontWeight:700}}>+</button>
                  <span style={{color:"#1e8449",fontSize:12,fontWeight:700}}>✓</span>
                </div>:<button onClick={()=>setComboQty(promo,1)} style={{width:"100%",padding:"7px",borderRadius:7,border:"none",cursor:"pointer",background:"#7b1a1a",color:"#fff",fontWeight:700,fontSize:12}}>+ Agregar combo</button>}
              </div>
            );})}
            {shown.map(p=>{
              const ic=cart.find(i=>i.cartKey===p.id);
              const promo=getProductPromo(promos,p.id);
              const isDesc=promo?.tipo==="descuento";
              const finalPrice=isDesc?Math.max(0,promo.data?.tipoValor==="%"?p.salePrice*(1-(promo.data?.valor||0)/100):p.salePrice-(promo.data?.valor||0)):p.salePrice;
              return <div key={p.id} style={{background:"#fff",borderRadius:10,padding:14,border:ic?`2px solid ${RED}`:promo?"2px solid #f3d98a":"2px solid transparent",boxShadow:"0 1px 4px #0001"}}>
              <div style={{width:"100%",height:110,borderRadius:8,background:p.imageUrl?"transparent":"#f4f6f9",overflow:"hidden",marginBottom:9,display:"flex",alignItems:"center",justifyContent:"center"}}>
                {p.imageUrl ? <img src={p.imageUrl} alt="" style={{width:"100%",height:"100%",objectFit:"cover"}}/> : <ImageIcon size={26} color="#ddd" strokeWidth={1.5}/>}
              </div>
              {promo && <div style={{marginBottom:5}}><span style={{fontSize:10,fontWeight:800,borderRadius:6,padding:"2px 7px",display:"inline-block",background:promo.tipo==="3x2"?"#fef9e7":"#fdecea",color:promo.tipo==="3x2"?"#b7770d":"#c0392b",border:promo.tipo==="3x2"?"1px solid #f3d98a":"1px solid #f5b7b1"}}><span style={{display:"inline-flex",alignItems:"center",gap:3}}>{promo.tipo==="3x2"?<Tag size={10} strokeWidth={2.6}/>:<TrendDown size={10} strokeWidth={2.6}/>} {promo.tipo==="3x2"?`${promo.data.comprar}×${promo.data.pagar}`:`-${promo.data.valor}${promo.data.tipoValor==="%"?"%":""}`}</span></span></div>}
              <div style={{fontWeight:700,fontSize:12,color:"#1a1a1a",marginBottom:3,lineHeight:1.3}}>{p.name}</div>
              <div style={{fontSize:12,color:"#666",marginBottom:7,fontWeight:500}}>{p.id} · {p.category}</div>
              {loteMode && p.multiploCompra>1 && <div style={{marginBottom:7}}><span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#fef9e7",color:"#b7770d",border:"1px solid #f0d080",borderRadius:20,padding:"1.5px 8px",fontSize:10,fontWeight:700}}><BoxIcon size={10} strokeWidth={2.4}/>Caja de {p.multiploCompra}</span></div>}
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                <span>{isDesc&&<span style={{fontSize:11,color:"#aaa",textDecoration:"line-through",marginRight:5}}>{fARS(p.salePrice)}</span>}<span style={{fontSize:17,fontWeight:800,color:RED}}>{fARS(finalPrice)}</span></span><SPill n={p.stock}/>
              </div>
              {ic?<div style={{display:"flex",alignItems:"center",gap:5}}>
                <button onClick={()=>setQ(p.id,ic.qty-step(p))} style={{width:27,height:27,borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:15,fontWeight:700}}>−</button>
                <input type="number" value={ic.qty} onChange={e=>{
                  const val = +e.target.value||0;
                  if(loteMode && p.multiploCompra>1 && val>0 && val%p.multiploCompra!==0) {
                    const snapped = Math.max(p.multiploCompra, Math.round(val/p.multiploCompra)*p.multiploCompra);
                    toast.info(`${p.name} se vende por caja de ${p.multiploCompra} — ajustado a ${snapped}`);
                    setQ(p.id, snapped);
                  } else setQ(p.id, val);
                }} style={{width:40,textAlign:"center",padding:3,borderRadius:6,border:`1.5px solid ${RED}`,fontWeight:700,fontSize:13,outline:"none"}}/>
                <button onClick={()=>setQ(p.id,ic.qty+step(p))} style={{width:27,height:27,borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:15,fontWeight:700}}>+</button>
                <span style={{color:"#1e8449",fontSize:12,fontWeight:700}}>✓</span>
              </div>:<button onClick={()=>addC(p)} style={{width:"100%",padding:"7px",borderRadius:7,border:"none",cursor:"pointer",background:RED,color:"#fff",fontWeight:700,fontSize:12}}>+ Agregar</button>}
            </div>;})}
          </div>
      }
    </div>
  );
}


// Muestra las líneas del carrito agrupando los componentes de combos bajo su propio encabezado
function CartSummaryLines({cart}) {
  const comboIds = [...new Set(cart.filter(i=>i.comboId).map(i=>i.comboId))];
  const regulares = cart.filter(i=>!i.comboId);
  return (
    <>
      {comboIds.map(cid=>{
        const items = cart.filter(i=>i.comboId===cid);
        const tot = items.reduce((s,i)=>s+applyItemDiscount(i.price,i.qty,i.disc),0);
        return (
          <div key={cid} style={{background:"#fff8ec",border:"1px dashed #c9a96a",borderRadius:8,padding:"8px 10px",margin:"6px 0"}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:12,fontWeight:800,color:"#7b1a1a",marginBottom:3}}>
              <span style={{display:"inline-flex",alignItems:"center",gap:5}}><Gift size={11} strokeWidth={2.5}/>{items[0].comboNombre}</span><span>{fARS(tot)}</span>
            </div>
            {items.map(i=><div key={i.cartKey} style={{fontSize:11,color:"#888",padding:"1px 0 1px 10px"}}>↳ {i.name} × {i.qty}</div>)}
          </div>
        );
      })}
      {regulares.map(i=>{
        const lineTotal=applyItemDiscount(i.price,i.qty,i.disc);
        const hasD=parseFloat(i.disc?.value)>0;
        return <div key={i.cartKey} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"4px 0",borderBottom:"1px solid #f9f9f9",color:"#555",gap:6}}>
          <span style={{flex:1,lineHeight:1.3}}>{i.name} × {i.qty}{i.promoLabel&&<div style={{fontSize:10,color:"#b7770d",fontWeight:700,marginTop:1}}>{i.promoLabel}</div>}</span>
          <div style={{textAlign:"right",whiteSpace:"nowrap"}}>
            {hasD&&<div style={{fontSize:10,color:"#aaa",textDecoration:"line-through"}}>{fARS(i.price*i.qty)}</div>}
            <span style={{fontWeight:600,color:hasD?"#1e8449":undefined}}>{fARS(lineTotal)}</span>
          </div>
        </div>;
      })}
    </>
  );
}

function Nuevo({products,vendors,onAdd,onDone,currentUser,isMobile,clients,onSaveClient,promos}) {
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
    if(!selectedClient){toast.error("Seleccioná un cliente");return;}
    if(!vendedor&&currentUser.role==="admin"){toast.error("Seleccioná un vendedor");return;}
    if(!cart.length){toast.error("Agregá productos");return;}
    setSaving(true);
    await onAdd({id:genId(),client:selectedClient.name,clientId:selectedClient.id,notes,vendedor:vendedor||currentUser.vendedor||currentUser.name,items:cart,total,subtotal,globalDisc,stage:"reserva",date:today()});
    setSaving(false);
    setOk(true); setTimeout(()=>onDone(),1400);
  };

  if(saving) return <SaveSpinner label="Registrando pedido..." color={RED}/>;
  if(ok) return <div style={{textAlign:"center",padding:80}}><div style={{display:"flex",justifyContent:"center"}}><CheckCircle size={52} color="#1e8449" strokeWidth={1.8}/></div><div style={{fontWeight:800,color:"#1e8449",fontSize:20,marginTop:12}}>¡Pedido registrado!</div></div>;

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
            <div style={{fontWeight:800,fontSize:14,marginBottom:12,display:"flex",alignItems:"center",gap:6}}><Users size={14} strokeWidth={2.3}/>Seleccioná el cliente</div>
            {selectedClient
              ? <div style={{background:"#fdecea",border:"1.5px solid #f5b7b1",borderRadius:10,padding:"12px 14px",marginBottom:12,display:"flex",alignItems:"center",gap:10}}>
                  <CheckCircle size={22} color="#1e8449" strokeWidth={2}/>
                  <div style={{flex:1}}>
                    <div style={{fontWeight:700,fontSize:14,color:"#922b21"}}>{selectedClient.name}</div>
                    <div style={{fontSize:12,color:"#c0392b",display:"flex",alignItems:"center",gap:4}}><Phone size={10} strokeWidth={2.4}/>{selectedClient.phone}</div>
                  </div>
                  <button onClick={()=>setSelectedClient(null)} style={{padding:"4px 10px",borderRadius:7,border:"1px solid #f5b7b1",background:"#fff",color:"#c0392b",fontSize:11,fontWeight:600,cursor:"pointer"}}>Cambiar</button>
                </div>
              : <ClientSelector clients={clients} onSelect={c=>{setSelectedClient(c);}} onSaveClient={onSaveClient} currentUser={currentUser}/>
            }
            {currentUser.role==="admin"&&<Field label="Vendedor *"><select value={vendedor} onChange={e=>setVendedor(e.target.value)} style={{...inputStyle,cursor:"pointer",color:vendedor?"#1a1a1a":"#aaa",marginBottom:8}}><option value="">— Seleccioná vendedor —</option>{vendors.map(v=><option key={v} value={v}>{v}</option>)}</select></Field>}
            <Field label="Notas"><textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Observaciones..." style={{...inputStyle,resize:"vertical",minHeight:55,fontSize:12}}/></Field>
            <button onClick={()=>{if(!selectedClient){toast.error("Seleccioná un cliente");return;}if(!vendedor&&currentUser.role==="admin"){toast.error("Seleccioná un vendedor");return;}setMStep(2);}}
              style={{width:"100%",padding:"12px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:selectedClient?`linear-gradient(135deg,${REDD},${RED})`:"#e5e5e5",color:selectedClient?"#fff":"#aaa",marginTop:8,display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
              Siguiente Productos <ArrowRightIcon size={13} strokeWidth={2.4}/>
            </button>
          </div>
        )}

        {mStep===2 && (
          <div style={{flex:1,overflow:"auto",paddingBottom:cart.length>0?72:0}}>
            <ProductSelector products={products} cart={cart} setCart={setCart} isMobile={true} promos={promos}/>
          </div>
        )}
        {mStep===2 && cart.length>0&&(
          <div style={{position:"fixed",bottom:0,left:0,right:0,background:`linear-gradient(135deg,${REDD},${RED})`,color:"#fff",padding:"13px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",zIndex:200,boxShadow:"0 -3px 16px #0003"}}>
            <div>
              <div style={{fontWeight:800,fontSize:15}}>{fARS(subtotal)}</div>
              <div style={{fontSize:11,opacity:.85}}>{cart.length} producto{cart.length!==1?"s":""} seleccionado{cart.length!==1?"s":""}</div>
            </div>
            <button onClick={()=>setMStep(3)} style={{padding:"10px 20px",borderRadius:10,border:"none",background:"#fff",color:RED,fontWeight:800,fontSize:14,cursor:"pointer",boxShadow:"0 2px 8px #0002",display:"flex",alignItems:"center",gap:6}}>
              Ver resumen <ArrowRightIcon size={13} strokeWidth={2.4}/>
            </button>
          </div>
        )}

        {mStep===3 && (
          <div style={{background:"#fff",borderRadius:12,padding:16,boxShadow:"0 1px 4px #0001"}}>
            <div style={{fontWeight:800,fontSize:15,marginBottom:14,color:"#1a1a1a",display:"flex",alignItems:"center",gap:7}}><ClipboardList size={15} strokeWidth={2.3}/>Confirmar pedido</div>
            <div style={{fontSize:13,color:"#555",marginBottom:4,display:"flex",alignItems:"center",gap:5}}><Users size={12} strokeWidth={2.4}/><strong>{selectedClient?.name}</strong> · {vendedor}</div>
            <div style={{borderTop:"1px solid #f5f5f5",margin:"8px 0",paddingTop:8}}>
              <CartSummaryLines cart={cart}/>
            </div>
            <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:17,color:RED,padding:"8px 0",borderTop:"2px solid #f5f5f5",marginBottom:14}}><span>Total</span><span>{fARS(total)}</span></div>
            <button onClick={submit} style={{width:"100%",padding:"12px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:`linear-gradient(135deg,${REDD},${RED})`,color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",gap:7}}>
              <CheckCircle size={14} strokeWidth={2.3}/> Registrar como Reserva
            </button>
            <button onClick={()=>setMStep(2)} style={{width:"100%",padding:"10px",borderRadius:10,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontWeight:600,fontSize:13,cursor:"pointer",marginTop:8,display:"flex",alignItems:"center",justifyContent:"center",gap:5}}><ArrowLeftIcon size={11} strokeWidth={2.4}/> Volver a productos</button>
          </div>
        )}
      </div>
    );
  }

  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 330px",gap:18,alignItems:"start"}}>
      <div>
        <div style={{fontWeight:800,fontSize:15,marginBottom:12,display:"flex",alignItems:"center",gap:7}}><ShoppingCart size={15} strokeWidth={2.3}/>Nuevo Pedido — Seleccioná productos</div>
        <ProductSelector products={products} cart={cart} setCart={setCart} promos={promos}/>
      </div>
      <div style={{position:"sticky",top:16}}>
        <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 2px 12px #0002"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14,display:"flex",alignItems:"center",gap:7}}><ClipboardList size={15} strokeWidth={2.3}/>Resumen del Pedido</div>
          {/* Client selector */}
          {selectedClient
            ? <div style={{background:"#fdecea",border:"1.5px solid #f5b7b1",borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",gap:10}}>
                <CheckCircle size={18} color="#1e8449" strokeWidth={2}/>
                <div style={{flex:1}}>
                  <div style={{fontWeight:700,fontSize:13,color:"#922b21"}}>{selectedClient.name}</div>
                  <div style={{fontSize:11,color:"#c0392b",display:"flex",alignItems:"center",gap:4}}><Phone size={10} strokeWidth={2.4}/>{selectedClient.phone}</div>
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
            :<CartSummaryLines cart={cart}/>}
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
          <button onClick={submit} disabled={!cart.length||!selectedClient||(!vendedor&&currentUser.role==="admin")} style={{width:"100%",padding:"11px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:(!cart.length||!selectedClient||(!vendedor&&currentUser.role==="admin"))?"#e5e5e5":`linear-gradient(135deg,${REDD},${RED})`,color:(!cart.length||!selectedClient||(!vendedor&&currentUser.role==="admin"))?"#aaa":"#fff",display:"flex",alignItems:"center",justifyContent:"center",gap:7}}>
            <CheckCircle size={14} strokeWidth={2.3}/> Registrar como Reserva
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
        <button onClick={()=>setView("lista")} style={{flex:1,padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="lista"?`linear-gradient(135deg,${REDD},${RED})`:"transparent",color:view==="lista"?"#fff":"#555",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}><FileText size={13} strokeWidth={2.3}/>Lista de Cotizaciones ({quotes.filter(q=>!q.convertida).length})</button>
        <button onClick={()=>setView("nueva")} style={{flex:1,padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="nueva"?`linear-gradient(135deg,${REDD},${RED})`:"transparent",color:view==="nueva"?"#fff":"#555",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}><Plus size={13} strokeWidth={2.4}/>Nueva Cotización</button>
      </div>
      {view==="nueva" && <NuevaCotizacion products={products} vendors={vendors} onAdd={async(q)=>{await onAdd(q);setView("lista");}} currentUser={currentUser} isMobile={isMobile} clients={clients} onSaveClient={onSaveClient}/>}
      {view==="lista" && (quotes.length===0
        ? <div style={{textAlign:"center",padding:60,color:"#aaa"}}><div style={{display:"flex",justifyContent:"center"}}><FileText size={42} color="#ddd" strokeWidth={1.7}/></div><div style={{marginTop:8}}>No hay cotizaciones aún</div></div>
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
    if(q.convertida) return <span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#d5f5e3",color:"#1e8449",border:"1px solid #1e844944",borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}><CheckCircle size={10} strokeWidth={2.5}/>Convertida</span>;
    if(status==="vencida")   return <span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#fdecea",color:"#c0392b",border:"1px solid #f1948a",borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}><Clock size={10} strokeWidth={2.5}/>Vencida</span>;
    if(status==="extendida") return <span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#fef9e7",color:"#b7770d",border:"1px solid #f0d080",borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}><RefreshCw size={10} strokeWidth={2.5}/>Extendida</span>;
    return <span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#d5f5e3",color:"#1e8449",border:"1px solid #a9dfbf",borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}><CheckCircle size={10} strokeWidth={2.5}/>Vigente</span>;
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
            {q.vendedor&&<span>· {q.vendedor}</span>}
            <StatusBdg/>
            {hoursLeft>0&&<span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#eaf4fc",color:"#1a5276",border:"1px solid #aed6f1",borderRadius:20,padding:"2px 8px",fontSize:10,fontWeight:700}}><Clock size={9} strokeWidth={2.6}/>{hoursLeft}hs</span>}
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
              <div style={{fontWeight:800,fontSize:13,color:"#c0392b",marginBottom:4,display:"flex",alignItems:"center",gap:6}}><Clock size={13} strokeWidth={2.4}/>Cotización vencida</div>
              <div style={{fontSize:12,color:"#922b21",marginBottom:10}}>Las cotizaciones vencen a las 48 hs. Podés extenderla 48 hs más por única vez.</div>
              <button onClick={(e)=>{e.stopPropagation();setShowExtForm(true);}}
                style={{padding:"8px 16px",borderRadius:8,border:"none",cursor:"pointer",background:"linear-gradient(135deg,#e67e22,#d35400)",color:"#fff",fontWeight:700,fontSize:13,display:"flex",alignItems:"center",gap:6}}>
                <RefreshCw size={12} strokeWidth={2.4}/> Extender 48 horas
              </button>
            </div>
          )}

          {/* FORMULARIO EXTENSIÓN */}
          {showExtForm && (
            <div style={{background:"#fffbf0",border:"1.5px solid #f0d080",borderRadius:10,padding:"12px 14px",marginBottom:14}} onClick={e=>e.stopPropagation()}>
              <div style={{fontSize:12,fontWeight:700,color:"#b7770d",marginBottom:6,display:"flex",alignItems:"center",gap:5}}><Pencil size={11} strokeWidth={2.4}/>Razón de la extensión (obligatorio)</div>
              <textarea value={extReason} onChange={e=>setExtReason(e.target.value)}
                placeholder="Ej: Cliente solicitó más tiempo para confirmar con su jefe..."
                style={{width:"100%",padding:"8px 10px",borderRadius:8,border:`1.5px solid ${extReason.trim()?"#f0d080":"#f1948a"}`,fontSize:13,resize:"vertical",minHeight:70,outline:"none",boxSizing:"border-box",fontFamily:"inherit"}}/>
              <div style={{display:"flex",gap:8,marginTop:8}}>
                <button onClick={handleExtend} disabled={!extReason.trim()||saving}
                  style={{padding:"7px 14px",borderRadius:7,border:"none",background:extReason.trim()?"#b7770d":"#e5e5e5",color:extReason.trim()?"#fff":"#aaa",fontWeight:700,fontSize:12,cursor:extReason.trim()?"pointer":"not-allowed",display:"flex",alignItems:"center",gap:5}}>
                  {saving?"Guardando...":<><CheckCircle size={11} strokeWidth={2.5}/>Confirmar extensión</>}
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
              <span style={{display:"inline-flex",alignItems:"center",gap:5}}><RefreshCw size={11} strokeWidth={2.4}/><strong>Extensión aplicada</strong></span> {q.extendDate&&`(${q.extendDate})`}<br/>
              Motivo: "{q.extendReason}"<br/>
              <span style={{color:"#aaa",fontSize:11}}>Extensión ya usada — no se puede extender de nuevo</span>
            </div>
          )}

          {/* VIGENTE */}
          {status==="vigente" && !q.convertida && hoursLeft>0 && (
            <div style={{background:"#eafaf1",border:"1.5px solid #a9dfbf",borderRadius:10,padding:"9px 14px",marginBottom:14,fontSize:12,color:"#1e8449",fontWeight:600,display:"flex",alignItems:"center",gap:6}}>
              <CheckCircle size={12} strokeWidth={2.4}/> Vigente — {hoursLeft} hs restantes
            </div>
          )}

          {q.validity&&<div style={{background:"#fef9e7",borderRadius:8,padding:"7px 12px",fontSize:12,color:"#7d6608",marginBottom:12,display:"flex",alignItems:"center",gap:5}}><Clock size={11} strokeWidth={2.4}/>Validez: <strong>{q.validity}</strong></div>}
          {q.items.map((it,i)=>{const p=getP(it.pid);return <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #f9f9f9",fontSize:13}}><span style={{color:"#444"}}>{p?.name||it.name} x {it.qty}</span><span style={{fontWeight:600}}>{fARS(it.price*it.qty)}</span></div>;})}
          <div style={{display:"flex",justifyContent:"flex-end",fontWeight:800,fontSize:16,color:PURPLE,margin:"8px 0 12px"}}>{fARS(q.total)}</div>
          {q.notes&&<div style={{background:"#f9f9f9",borderRadius:8,padding:"8px 12px",fontSize:13,color:"#555",marginBottom:12}}>{q.notes}</div>}
          <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
            <button onClick={()=>printDoc(q,"cotizacion")} style={{padding:"8px 12px",borderRadius:8,border:`1.5px solid ${PURPLEBG}`,cursor:"pointer",background:"#fff",color:PURPLE,fontWeight:600,fontSize:13,display:"flex",alignItems:"center",gap:6}}><Printer size={13} strokeWidth={2.3}/>Imprimir</button>
            <button onClick={()=>window.open(`https://wa.me/?text=${encodeURIComponent(buildWAQuote(q))}`, "_blank")}
              style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #25D366",cursor:"pointer",background:"#25D366",color:"#fff",fontWeight:700,fontSize:13,display:"flex",alignItems:"center",gap:5}}>
              <MessageCircle size={14} strokeWidth={2.3}/> Enviar por WhatsApp
            </button>
            {!q.convertida
              ? <button onClick={()=>onConvert(q)} disabled={isVencida}
                  style={{padding:"8px 14px",borderRadius:8,border:"none",cursor:isVencida?"not-allowed":"pointer",background:isVencida?"#e5e5e5":"linear-gradient(135deg,#922b21,#c0392b)",color:isVencida?"#aaa":"#fff",fontWeight:700,fontSize:13,display:"flex",alignItems:"center",gap:6}}>
                  <ShoppingCart size={13} strokeWidth={2.3}/> Pasar a Reserva
                </button>
              : <span style={{display:"inline-flex",alignItems:"center",gap:5,background:"#d5f5e3",color:"#1e8449",borderRadius:8,padding:"6px 12px",fontSize:12,fontWeight:700}}><CheckCircle size={11} strokeWidth={2.5}/>Convertida a Reserva</span>
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
    if(!selectedClient){toast.error("Seleccioná un cliente");return;}
    if(!cart.length){toast.error("Agregá productos");return;}
    setSaving(true);
    const q={id:genId(),client:selectedClient.name,clientId:selectedClient.id,notes,vendedor,validity,items:cart,total,subtotal,globalDisc,date:today()};
    await onAdd(q);
    setSaving(false);
    setOk(true);
  };

  if(saving) return <SaveSpinner label="Guardando cotización..." color={PURPLE}/>;
  if(ok) return <div style={{textAlign:"center",padding:80}}><div style={{display:"flex",justifyContent:"center"}}><FileText size={52} color={PURPLE} strokeWidth={1.8}/></div><div style={{fontWeight:800,color:PURPLE,fontSize:20,marginTop:12}}>¡Cotización guardada!</div></div>;

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
          <div style={{background:"#f5eef8",borderRadius:8,padding:"8px 12px",fontSize:12,color:PURPLE,marginBottom:12,display:"flex",alignItems:"flex-start",gap:6}}>
            <Info size={13} strokeWidth={2.2} style={{flexShrink:0,marginTop:1}}/> Las cotizaciones <strong>no descuentan stock</strong>. Son solo presupuestos.
          </div>
          <div style={{fontWeight:800,fontSize:14,marginBottom:12,display:"flex",alignItems:"center",gap:6}}><Users size={14} strokeWidth={2.3}/>Seleccioná el cliente</div>
          {selectedClient
            ? <div style={{background:"#f5eef8",border:"1.5px solid #d7bde2",borderRadius:10,padding:"12px 14px",marginBottom:12,display:"flex",alignItems:"center",gap:10}}>
                <CheckCircle size={22} color={PURPLE} strokeWidth={2}/>
                <div style={{flex:1}}>
                  <div style={{fontWeight:700,fontSize:14,color:PURPLE}}>{selectedClient.name}</div>
                  <div style={{fontSize:12,color:"#888",display:"flex",alignItems:"center",gap:4}}><Phone size={10} strokeWidth={2.4}/>{selectedClient.phone}</div>
                </div>
                <button onClick={()=>setSelectedClient(null)} style={{padding:"4px 10px",borderRadius:7,border:"1px solid #d7bde2",background:"#fff",color:PURPLE,fontSize:11,fontWeight:600,cursor:"pointer"}}>Cambiar</button>
              </div>
            : <ClientSelector clients={clients||[]} onSelect={setSelectedClient} onSaveClient={onSaveClient} currentUser={currentUser}/>
          }
          {currentUser?.role==="admin"&&<Field label="Vendedor"><select value={vendedor} onChange={e=>setVendedor(e.target.value)} style={{...inputStyle,cursor:"pointer"}}><option value="">- Sin asignar -</option>{vendors.map(v=><option key={v} value={v}>{v}</option>)}</select></Field>}
          <Field label="Válida hasta"><input value={validity} onChange={e=>setValidity(e.target.value)} placeholder="Ej: 48 horas" style={inputStyle}/></Field>
          <Field label="Notas"><textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Observaciones, condiciones..." style={{...inputStyle,resize:"vertical",minHeight:55,fontSize:12}}/></Field>
          <button onClick={()=>{if(!selectedClient){toast.error("Seleccioná un cliente");return;}setMStep(2);}}
            style={{width:"100%",padding:"12px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:selectedClient?PURPLEG:"#e5e5e5",color:selectedClient?"#fff":"#aaa",marginTop:8,display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
            Siguiente Productos <ArrowRightIcon size={13} strokeWidth={2.4}/>
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
          <button onClick={()=>setMStep(3)} style={{padding:"10px 20px",borderRadius:10,border:"none",background:"#fff",color:PURPLE,fontWeight:800,fontSize:14,cursor:"pointer",boxShadow:"0 2px 8px #0002",display:"flex",alignItems:"center",gap:6}}>
            Ver resumen <ArrowRightIcon size={13} strokeWidth={2.4}/>
          </button>
        </div>
      )}

      {/* Paso 3 — Confirmar */}
      {mStep===3 && (
        <div style={{background:"#fff",borderRadius:12,padding:16,boxShadow:"0 1px 4px #0001"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14,color:PURPLE,display:"flex",alignItems:"center",gap:7}}><FileText size={15} strokeWidth={2.3}/>Confirmar cotización</div>
          <div style={{fontSize:13,color:"#555",marginBottom:8,display:"flex",alignItems:"center",gap:5}}><Users size={12} strokeWidth={2.4}/><strong>{selectedClient?.name}</strong> · {vendedor} · {validity}</div>
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
          <button onClick={submit} style={{width:"100%",padding:"12px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:PURPLEG,color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",gap:7}}>
            <FileText size={14} strokeWidth={2.3}/> Guardar Cotización
          </button>
          <button onClick={()=>setMStep(2)} style={{width:"100%",padding:"10px",borderRadius:10,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontWeight:600,fontSize:13,cursor:"pointer",marginTop:8,display:"flex",alignItems:"center",justifyContent:"center",gap:5}}>
            <ArrowLeftIcon size={11} strokeWidth={2.4}/> Volver a productos
          </button>
        </div>
      )}
    </div>
  );

  // ── DESKTOP ──
  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 330px",gap:18,alignItems:"start"}}>
      <div>
        <div style={{fontWeight:800,fontSize:15,marginBottom:12,display:"flex",alignItems:"center",gap:7}}><FileText size={15} strokeWidth={2.3}/>Nueva Cotización - Seleccioná productos</div>
        <ProductSelector products={products} cart={cart} setCart={setCart}/>
      </div>
      <div style={{position:"sticky",top:16}}>
        <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 2px 12px #0002",border:"2px solid #e8daef"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14,color:PURPLE,display:"flex",alignItems:"center",gap:7}}><FileText size={15} strokeWidth={2.3}/>Resumen de Cotización</div>
          <div style={{background:"#e8daef",borderRadius:8,padding:"7px 12px",fontSize:12,color:PURPLE,marginBottom:14,display:"flex",alignItems:"flex-start",gap:6}}>
            <Info size={13} strokeWidth={2.2} style={{flexShrink:0,marginTop:1}}/> Las cotizaciones <strong>no descuentan stock</strong>. Son solo presupuestos para el cliente.
          </div>
          {selectedClient
            ? <div style={{background:"#f5eef8",border:"1.5px solid #d7bde2",borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",gap:10}}>
                <CheckCircle size={18} color={PURPLE} strokeWidth={2}/>
                <div style={{flex:1}}><div style={{fontWeight:700,fontSize:13,color:PURPLE}}>{selectedClient.name}</div><div style={{fontSize:11,color:"#888",display:"flex",alignItems:"center",gap:4}}><Phone size={10} strokeWidth={2.4}/>{selectedClient.phone}</div></div>
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
          <button onClick={submit} disabled={!cart.length||!selectedClient||saving} style={{width:"100%",padding:"11px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:(!cart.length||!selectedClient)?"#e5e5e5":PURPLEG,color:(!cart.length||!selectedClient)?"#aaa":"#fff",display:"flex",alignItems:"center",justifyContent:"center",gap:7}}>
            {saving?"Guardando...":<><FileText size={14} strokeWidth={2.3}/> Guardar Cotización</>}
          </button>
        </div>
      </div>
    </div>
  );
}


// ─── STOCK ────────────────────────────────────────────────────────────────────
function StockAlert({low, onPedirReposicion}) {
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
            <div key={p.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:"1px solid #f1c40f22",gap:8,flexWrap:"wrap"}}>
              <div>
                <span style={{fontWeight:600,fontSize:13,color:"#555"}}>{p.name}</span>
                <span style={{fontSize:11,color:"#aaa",marginLeft:8}}>{p.id}</span>
                {p.multiploCompra>1&&<span style={{fontSize:10,color:"#9a7d0a",marginLeft:6}}>(caja de {p.multiploCompra})</span>}
              </div>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <span style={{background:p.stock===0?"#fdecea":"#fff3cd",color:p.stock===0?RED:"#856404",borderRadius:8,padding:"2px 8px",fontSize:12,fontWeight:700,border:`1px solid ${p.stock===0?"#f5c6cb":"#ffc107"}`}}>
                  {p.stock===0?"Sin stock":`🎉 ${p.stock} u.`}
                </span>
                {onPedirReposicion && (
                  <button onClick={(e)=>{e.stopPropagation();onPedirReposicion(p.id,p.multiploCompra||1);}}
                    style={{display:"flex",alignItems:"center",gap:4,padding:"3px 9px",borderRadius:7,border:"1px solid #6c348355",background:"#f5eef8",color:"#6c3483",fontSize:11,fontWeight:700,cursor:"pointer",whiteSpace:"nowrap"}}>
                    📦 Pedir reposición
                  </button>
                )}
              </div>
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
    if(!reason.trim()){toast.error("El motivo es obligatorio para registrar el movimiento");return;}
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
      if(qty===0){toast.error("El ajuste no puede ser 0");return;}
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

function Stock({products,onUpd,onDel,onAdjust,isAdmin,addLog,stockLog,setStockLog,isMobile,onCrearOferta,onPedirReposicion}) {
  const [search,setSearch]=useState("");
  const [cat,setCat]=useState("todos");
  const [editing,setEditing]=useState(null);
  const [stockTab,setStockTab]=useState("lista");
  const [page,setPage]=useState(0);
  const PAGE_SIZE = 100;

  const [soloConStock, setSoloConStock] = useState(false);
  const [soloSinRotacion, setSoloSinRotacion] = useState(false);  // ← nuevo filtro admin
  const CATS=useMemo(()=>["todos",...new Set(products.map(p=>p.category))].sort(),[products]);
  const q=search.toLowerCase();

  // ── Último movimiento por producto (solo admin) ───────────────────────────
  const lastMovByPid = useMemo(()=>{
    const map={};
    stockLog.forEach(e=>{
      const d = parseFechaLog(e.fecha);
      if(!d) return;
      if(!map[e.productoId] || d > map[e.productoId]) map[e.productoId] = d;
    });
    return map;
  }, [stockLog]);

  const diasSinRot = (pid) => {
    const last = lastMovByPid[pid];
    if(!last) return null;
    return Math.floor((Date.now() - last.getTime()) / 86400000);
  };

  // Reset page when filters change
  useEffect(()=>setPage(0),[search,cat,soloConStock,soloSinRotacion]);

  const filtered=useMemo(()=>products.filter(p=>{
    if(soloConStock && p.stock<=0) return false;
    if(cat!=="todos"&&p.category!==cat) return false;
    if(soloSinRotacion && isAdmin){ const d=diasSinRot(p.id); if(!d || d<180) return false; }
    if(q) return norm(p.name).includes(norm(q))||normSKU(p.id).includes(normSKU(q));
    return true;
  }),[products,cat,q,soloConStock,soloSinRotacion,lastMovByPid]);

  const totalPages = Math.ceil(filtered.length / PAGE_SIZE);
  const shown = filtered.slice(page*PAGE_SIZE, (page+1)*PAGE_SIZE);
  const low=products.filter(p=>p.stock>0&&p.stock<=5);

  // Contadores de productos sin rotación (solo relevantes para admin)
  const stale180 = isAdmin ? products.filter(p=>{ const d=diasSinRot(p.id); return d!==null&&d>=180&&d<360; }).length : 0;
  const stale360 = isAdmin ? products.filter(p=>{ const d=diasSinRot(p.id); return d!==null&&d>=360; }).length : 0;

  return (
    <div>
      {isAdmin && (
        <div style={{background:"#fff",borderRadius:12,padding:4,marginBottom:14,display:"flex",gap:4,boxShadow:"0 1px 4px #0001"}}>
          {[{k:"lista",l:"Lista de Stock",Icon:ClipboardList},{k:"ajuste",l:"Ajuste / Baja",Icon:Scale},{k:"log",l:"Movimientos",Icon:ListChecks}].map(t=>(
            <button key={t.k} onClick={()=>setStockTab(t.k)} style={{flex:1,padding:"9px 14px",borderRadius:9,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:stockTab===t.k?`linear-gradient(135deg,${REDD},${RED})`:"transparent",color:stockTab===t.k?"#fff":"#555",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
              <t.Icon size={13} strokeWidth={2.3}/> {t.l}
            </button>
          ))}
        </div>
      )}
      {isAdmin && stockTab==="ajuste" && <StockAdjust products={products} onDel={onDel} onAdjust={onAdjust} addLog={addLog}/>}
      {isAdmin && stockTab==="log" && <StockLog log={stockLog} onClear={async()=>{setStockLog([]);await db.clearStockLog();}}/>}
      {stockTab==="lista" && <>
        {low.length>0&&<StockAlert low={low} onPedirReposicion={onPedirReposicion}/>}

        {/* ── Alerta de stock sin rotación (solo admin) ── */}
        {isAdmin && (stale180>0||stale360>0) && (
          <div style={{background:"#fff",borderRadius:12,padding:"12px 16px",marginBottom:10,boxShadow:"0 1px 4px #0001",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
            <span style={{fontWeight:700,fontSize:13,color:"#1a1a1a",display:"inline-flex",alignItems:"center",gap:5}}><Package size={13} strokeWidth={2.3}/>Stock inmovilizado</span>
            {stale360>0&&<span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#fdecea",color:"#c0392b",border:"1px solid #f1948a",borderRadius:20,padding:"3px 10px",fontSize:12,fontWeight:700}}><span style={{width:7,height:7,borderRadius:"50%",background:"#c0392b",flexShrink:0}}/>{stale360} producto{stale360!==1?"s":""} sin rotación +360 días</span>}
            {stale180>0&&<span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#fef9e7",color:"#b7770d",border:"1px solid #f0d080",borderRadius:20,padding:"3px 10px",fontSize:12,fontWeight:700}}><span style={{width:7,height:7,borderRadius:"50%",background:"#f1c40f",flexShrink:0}}/>{stale180} producto{stale180!==1?"s":""} sin rotación 180–360 días</span>}
            <button onClick={()=>setSoloSinRotacion(s=>!s)}
              style={{marginLeft:"auto",padding:"4px 12px",borderRadius:8,border:`1.5px solid ${soloSinRotacion?"#c0392b":"#e5e5e5"}`,background:soloSinRotacion?"#fdecea":"#fff",color:soloSinRotacion?"#c0392b":"#666",fontWeight:700,fontSize:11,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:4}}>
              {soloSinRotacion?<><XIcon size={10} strokeWidth={2.6}/>Ver todos</>:"Ver solo estos"}
            </button>
          </div>
        )}
        <div style={{background:"#fff",borderRadius:12,padding:14,marginBottom:14,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center",boxShadow:"0 1px 4px #0001"}}>
          <div style={{fontWeight:800,fontSize:15,flex:1,display:"flex",alignItems:"center",gap:6}}>
            <Package size={15} strokeWidth={2.2}/>Stock — <span style={{color:RED}}>{filtered.length.toLocaleString("es-AR")}</span> productos
            {q||cat!=="todos" ? <span style={{fontSize:12,color:"#888",fontWeight:400,marginLeft:6}}>de {products.length.toLocaleString("es-AR")} total</span> : null}
          </div>
          <div style={{position:"relative"}}>
            <Search size={13} color="#aaa" strokeWidth={2.3} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)"}}/>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar por nombre o código..."
              style={{padding:"7px 12px 7px 30px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:12,outline:"none",width:220}} autoFocus={false}/>
          </div>
          <button onClick={()=>setSoloConStock(s=>!s)}
            style={{padding:"7px 14px",borderRadius:8,border:`1.5px solid ${soloConStock?"#1e8449":"#e5e5e5"}`,background:soloConStock?"#d5f5e3":"#fff",color:soloConStock?"#1e8449":"#666",fontWeight:700,fontSize:12,cursor:"pointer",whiteSpace:"nowrap",display:"inline-flex",alignItems:"center",gap:5}}>
            {soloConStock ? <CheckCircle size={12} strokeWidth={2.4}/> : <Package size={12} strokeWidth={2.4}/>} Con stock
          </button>
          {isAdmin && (stale180+stale360>0) && (
            <button onClick={()=>setSoloSinRotacion(s=>!s)}
              style={{padding:"7px 14px",borderRadius:8,border:`1.5px solid ${soloSinRotacion?"#c0392b":"#e5e5e5"}`,background:soloSinRotacion?"#fdecea":"#fff",color:soloSinRotacion?"#c0392b":"#666",fontWeight:700,fontSize:12,cursor:"pointer",whiteSpace:"nowrap",display:"inline-flex",alignItems:"center",gap:5}}>
              {soloSinRotacion?<><XIcon size={11} strokeWidth={2.6}/>Sin rotación</>:<><span style={{width:8,height:8,borderRadius:"50%",background:"#c0392b",display:"inline-block"}}/>Sin rotación</>}
            </button>
          )}
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
                    const dias = isAdmin ? diasSinRot(p.id) : null;
                    const esRojo   = dias!==null && dias>=360;
                    const esAmarillo = dias!==null && dias>=180 && dias<360;
                    const staleBadge = isAdmin && (esRojo||esAmarillo) ? (
                      <span style={{display:"inline-flex",alignItems:"center",gap:4,background:esRojo?"#fdecea":"#fef9e7",color:esRojo?"#c0392b":"#b7770d",border:`1px solid ${esRojo?"#f1948a":"#f0d080"}`,borderRadius:20,padding:"1px 7px",fontSize:10,fontWeight:700,whiteSpace:"nowrap",marginLeft:4}}>
                        <span style={{width:6,height:6,borderRadius:"50%",background:esRojo?"#c0392b":"#f1c40f",flexShrink:0}}/>{dias}d sin mov.
                      </span>
                    ) : null;
                    const crearOfertaBtn = isAdmin && (esRojo||esAmarillo) && onCrearOferta ? (
                      <button onClick={()=>onCrearOferta(p.id)} style={{display:"inline-flex",alignItems:"center",gap:4,marginLeft:4,padding:"1px 8px",borderRadius:20,border:"1px solid #6c348355",background:"#f5eef8",color:"#6c3483",fontSize:10,fontWeight:700,cursor:"pointer",whiteSpace:"nowrap"}}>
                        <Gift size={10} strokeWidth={2.5}/> Crear oferta
                      </button>
                    ) : null;
                    return isMobile ? (
                      <tr key={p.id} style={{borderTop:"1px solid #f5f5f5"}}>
                        <td style={{padding:"10px 12px"}}>
                          <div style={{fontWeight:600,fontSize:13}}>{p.name}{staleBadge}</div>
                          <div style={{fontSize:10,color:"#aaa",marginTop:2}}>{p.id} · {p.category}</div>
                          {isAdmin&&<button onClick={()=>setEditing(p)} style={{padding:"3px 8px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:11,fontWeight:600,marginTop:4,display:"inline-flex",alignItems:"center",gap:4}}><Pencil size={10} strokeWidth={2.4}/>Editar</button>}
                          {crearOfertaBtn && <span style={{marginTop:4,display:"inline-block"}}>{crearOfertaBtn}</span>}
                        </td>
                        <td style={{padding:"10px 12px",textAlign:"right"}}><SPill n={p.stock}/></td>
                        <td style={{padding:"10px 12px",fontWeight:700,color:RED,textAlign:"right",whiteSpace:"nowrap"}}>{fARS(p.salePrice)}</td>
                      </tr>
                    ) : (
                    <tr key={p.id} style={{borderTop:"1px solid #f5f5f5",background:esRojo?"#fff8f8":esAmarillo?"#fffef5":"#fff"}}>
                      <td style={{padding:"9px 12px",color:"#aaa",fontSize:11}}>{p.id}</td>
                      <td style={{padding:"9px 12px",fontWeight:600,color:"#1a1a1a",maxWidth:260}}>
                        <div style={{display:"flex",alignItems:"center",gap:4,flexWrap:"wrap"}}>
                          <span>{p.name}</span>
                          {staleBadge}
                          {crearOfertaBtn}
                        </div>
                      </td>
                      <td style={{padding:"9px 12px",color:"#aaa",fontSize:11}}>{p.category}</td>
                      <td style={{padding:"9px 12px"}}><SPill n={p.stock}/></td>
                      <td style={{padding:"9px 12px",fontWeight:700,color:RED}}>{fARS(p.salePrice)}</td>
                      {isAdmin&&<td style={{padding:"9px 12px",color:"#666"}}>{fARS(p.costPrice)}</td>}
                      {isAdmin&&<td style={{padding:"9px 12px",fontWeight:700,color:+m>=40?"#1e8449":"#e67e22"}}>{m}%</td>}
                      <td style={{padding:"9px 12px"}}>{isAdmin&&<button onClick={()=>setEditing(p)} style={{padding:"4px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:11,fontWeight:600,display:"inline-flex",alignItems:"center"}}><Pencil size={11} strokeWidth={2.4}/></button>}</td>
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
                style={{padding:"5px 12px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:page===0?"not-allowed":"pointer",color:page===0?"#ccc":"#555",fontSize:12,fontWeight:600,display:"inline-flex",alignItems:"center",gap:5}}>
                <ArrowLeftIcon size={11} strokeWidth={2.4}/> Anterior
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
                style={{padding:"5px 12px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:page===totalPages-1?"not-allowed":"pointer",color:page===totalPages-1?"#ccc":"#555",fontSize:12,fontWeight:600,display:"inline-flex",alignItems:"center",gap:5}}>
                Siguiente <ArrowRightIcon size={11} strokeWidth={2.4}/>
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
  const [imageUrl,setImageUrl]=useState(p.imageUrl||"");
  const [uploadingImg,setUploadingImg]=useState(false);
  const m=cost>0?((sale-cost)/cost*100).toFixed(1):"-";

  const handleImage = async (e) => {
    const file = e.target.files[0];
    if(!file) return;
    setUploadingImg(true);
    try {
      const url = await uploadProductImage(p.id, file);
      setImageUrl(url);
    } catch(err) {
      console.warn(err);
      toast.error("No se pudo subir la foto. Probá de nuevo.");
    } finally {
      setUploadingImg(false);
      e.target.value = "";
    }
  };

  return (
    <div style={{position:"fixed",inset:0,background:"#0007",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999,padding:16}}>
      <div style={{background:"#fff",borderRadius:16,padding:24,width:"100%",maxWidth:440,boxShadow:"0 20px 60px #0003"}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:16}}>
          <div style={{fontWeight:800,fontSize:14,color:"#1a1a1a",flex:1,marginRight:8,lineHeight:1.3}}>{p.name}</div>
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:22,color:"#aaa"}}>x</button>
        </div>

        <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:18}}>
          <div style={{width:64,height:64,borderRadius:12,background:imageUrl?"transparent":"#f4f6f9",overflow:"hidden",flexShrink:0,border:"1.5px solid #f0f0f0",display:"flex",alignItems:"center",justifyContent:"center"}}>
            {uploadingImg
              ? <RefreshCw size={20} color="#aaa" strokeWidth={2}/>
              : imageUrl ? <img src={imageUrl} alt="" style={{width:"100%",height:"100%",objectFit:"cover"}}/> : <ImageIcon size={24} color="#ccc" strokeWidth={1.6}/>}
          </div>
          <div>
            <label style={{display:"inline-flex",alignItems:"center",gap:6,padding:"7px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",fontSize:11.5,fontWeight:600,cursor:uploadingImg?"default":"pointer",color:"#555",opacity:uploadingImg?0.6:1}}>
              <Camera size={12} strokeWidth={2.3}/> {imageUrl?"Cambiar foto":"Subir foto"}
              <input type="file" accept="image/*" onChange={handleImage} disabled={uploadingImg} style={{display:"none"}}/>
            </label>
            {imageUrl&&!uploadingImg&&<button onClick={()=>setImageUrl("")} style={{marginLeft:6,padding:"7px 10px",borderRadius:8,border:"none",background:"#fdecea",color:RED,fontSize:11,cursor:"pointer",fontWeight:600,display:"inline-flex",alignItems:"center",gap:4}}><XIcon size={9} strokeWidth={2.6}/>Quitar</button>}
            <div style={{fontSize:10,color:"#aaa",marginTop:5}}>Opcional. Se ve al elegir el producto.</div>
          </div>
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
          <button onClick={()=>onSave({...p,costPrice:cost,salePrice:sale,stock,imageUrl})} style={{padding:"8px 14px",borderRadius:8,border:"none",background:RED,color:"#fff",cursor:"pointer",fontWeight:700}}>Guardar</button>
        </div>
      </div>
    </div>
  );
}

// ─── INGRESAR STOCK DESDE SOLICITUD ──────────────────────────────────────────
function IngresarDesdeSolicitud({po, products, onStock, onDone, onArrived}) {
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
    if(!toIngresar.length) { toast.error("Seleccioná al menos un producto para ingresar"); return; }
    setSaving(true);
    try {
      for(const it of toIngresar) {
        await onStock(it.pid, +it.qtyRecibida, +it.cost);
      }
      if(onArrived) { try { await onArrived(); } catch(e) { console.warn(e); } }
      setOk(true);
      setTimeout(()=>onDone(), 1500);
    } catch(e) {
      console.warn("Error al ingresar mercadería:", e);
      toast.error("Hubo un problema de conexión. Revisá el Stock — puede que ya se haya actualizado igual.");
    } finally {
      setSaving(false);
    }
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
  abierta:   {label:"Abierta",    color:"#1a5276", bg:"#d6eaf8",  Icon:ClipboardList},
  revisando: {label:"Revisando",  color:"#e67e22", bg:"#fef9e7",  Icon:Search},
  cerrada:   {label:"Cerrada",    color:"#1e8449", bg:"#d5f5e3",  Icon:CheckCircle},
};

function printSolicitudPDF(po, logoSrc) {
  const rows = po.items.map(it=>`
    <tr>
      <td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;font-size:14px;">${it.name}</td>
      <td style="padding:9px 12px;border-bottom:1px solid #f0f0f0;font-size:14px;">${it.id||it.pid||""}</td>
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
  .header-img{width:100%;height:auto;display:block;}
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
  <img class="header-img" src="${logoSrc}" alt="Libreria Madrid" onerror="this.style.display='none'"/>
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
    ...po.items.map(it=>[it.name, it.id||it.pid||"", it.qty, it.notas||""]),
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

function SolicitudCompra({products,currentUser,isAdmin,purchaseOrders,setPurchaseOrders,isMobile,onStockExternal,addLog,onCreated,autoOpenId,onConsumedAutoOpen,prefillItem,onConsumedPrefillItem}) {
  const [view, setView] = useState("lista"); // lista | nueva | detalle
  const [selected, setSelected] = useState(null);
  const [showAddProduct, setShowAddProduct] = useState(false);
  const [addSearch, setAddSearch] = useState("");
  const [cart, setCart] = useState([]);
  const [notas, setNotas] = useState("");
  const [saving, setSaving] = useState(false);
  const [itemNotas, setItemNotas] = useState({});

  // Deep-link desde la consola de aprobaciones: abre directo el detalle de una solicitud puntual
  useEffect(() => {
    if(!autoOpenId) return;
    const po = purchaseOrders.find(p=>p.id===autoOpenId);
    if(po) { setSelected(po); setView("detalle"); }
    onConsumedAutoOpen && onConsumedAutoOpen();
  }, [autoOpenId, purchaseOrders]);

  // Deep-link desde "Pedir reposicion" (stock bajo) — abre Nueva Solicitud con el producto
  // ya cargado, en la cantidad del multiplo del proveedor (el admin la puede modificar despues)
  useEffect(() => {
    if(!prefillItem) return;
    const p = products.find(pr=>pr.id===prefillItem.pid);
    if(p) {
      setCart([{pid:p.id, id:p.id, cartKey:p.id, qty:prefillItem.qty||1, price:p.salePrice, name:p.name}]);
      setView("nueva");
    }
    onConsumedPrefillItem && onConsumedPrefillItem();
  }, [prefillItem]);

  const myOrders = isAdmin ? purchaseOrders : purchaseOrders.filter(po=>po.vendedor===currentUser.vendedor||po.vendedor===currentUser.name);

  const savePO = async (po) => {
    setPurchaseOrders(prev=>prev.find(x=>x.id===po.id)?prev.map(x=>x.id===po.id?po:x):[po,...prev]);
    await db.savePurchaseOrder(po);
  };

  const createNew = async () => {
    if(!cart.length){toast.error("Agregá al menos un producto");return;}
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
    if(!await confirmDialog("¿Eliminar esta solicitud?","Esta acción no se puede deshacer.",true)) return;
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
            <div style={{marginBottom:8,display:"flex",justifyContent:"center"}}><ClipboardList size={42} color="#ccc" strokeWidth={1.6}/></div>
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
                    <div style={{fontSize:11,color:"#888",marginTop:2}}>{po.fecha} · {po.vendedor}</div>
                  </div>
                  <div style={{textAlign:"right",flexShrink:0}}>
                    <span style={{display:"inline-flex",alignItems:"center",gap:5,background:cfg.bg,color:cfg.color,borderRadius:8,padding:"3px 10px",fontSize:11,fontWeight:700}}>{cfg.Icon&&<cfg.Icon size={11} strokeWidth={2.4}/>}{cfg.label}</span>
                    <div style={{fontSize:11,color:"#888",marginTop:4}}>{po.items.length} productos · {po.items.reduce((s,i)=>s+i.qty,0)} uds.</div>
                  </div>
                </div>
                <div style={{display:"flex",gap:6,padding:"8px 12px",background:"#fafafa",borderTop:"1px solid #f0f0f0"}}>
                  <button onClick={()=>openDetail(po)} style={{flex:1,padding:"8px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",fontSize:11,fontWeight:700,cursor:"pointer"}}>Ver detalle</button>
                  {isAdmin && po.estado==="abierta" && <button onClick={()=>changeEstado(po,"revisando")} style={{flex:1,padding:"8px",borderRadius:8,border:"none",background:"#e67e22",color:"#fff",fontSize:11,fontWeight:700,cursor:"pointer"}}><><Search size={11} strokeWidth={2.6}/> Revisar</></button>}
                  {isAdmin && po.estado==="revisando" && <button onClick={()=>changeEstado(po,"cerrada")} style={{flex:1,padding:"8px",borderRadius:8,border:"none",background:"#1e8449",color:"#fff",fontSize:11,fontWeight:700,cursor:"pointer"}}><><CheckCircle size={11} strokeWidth={2.6}/> Cerrar</></button>}
                  {isAdmin && <button onClick={()=>deletePO(po.id)} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #fcc",background:"#fff",color:"#c0392b",fontSize:11,fontWeight:700,cursor:"pointer"}}><Trash size={12} strokeWidth={2.4}/></button>}
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
          <button onClick={()=>setView("lista")} style={{padding:"7px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",fontWeight:600,fontSize:12,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:5}}><ArrowLeftIcon size={11} strokeWidth={2.4}/> Volver</button>
          <div style={{fontWeight:800,fontSize:15}}>Solicitud #{po.id.slice(-6).toUpperCase()}</div>
          <span style={{display:"inline-flex",alignItems:"center",gap:5,background:cfg.bg,color:cfg.color,borderRadius:8,padding:"3px 10px",fontSize:11,fontWeight:700,marginLeft:"auto"}}>{cfg.Icon&&<cfg.Icon size={11} strokeWidth={2.4}/>}{cfg.label}</span>
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
              <span style={{display:"inline-flex",alignItems:"center",gap:5}}><Pencil size={12} strokeWidth={2.4}/>Modo edición — modificá cantidades o agregá productos</span>
              <button onClick={()=>setShowAddProduct(s=>!s)}
                style={{marginLeft:"auto",padding:"5px 12px",borderRadius:7,border:"none",background:showAddProduct?"#1a5276":"#e67e22",color:"#fff",fontWeight:700,fontSize:11,cursor:"pointer",whiteSpace:"nowrap"}}>
                {showAddProduct?<><XIcon size={11} strokeWidth={2.6}/> Cerrar buscador</>:"+ Agregar producto"}
              </button>
            </div>
          )}
          {isAdmin && po.estado!=="cerrada" && showAddProduct && (
            <div style={{padding:"12px 14px",borderBottom:"1px solid #f0f0f0",background:"#f9f9fb"}}>
              <input value={addSearch} onChange={e=>setAddSearch(e.target.value)} autoFocus
                placeholder="Buscar producto para agregar..."
                style={{width:"100%",padding:"8px 12px",borderRadius:8,border:"1.5px solid #1a5276",fontSize:13,outline:"none",marginBottom:8}}/>
              {addSearch.length>1 && (
                <div style={{maxHeight:220,overflowY:"auto",display:"flex",flexDirection:"column",gap:5}}>
                  {products.filter(p=>norm(p.name).includes(norm(addSearch))||normSKU(p.id).includes(normSKU(addSearch))).slice(0,20).map(p=>{
                    const already = po.items.find(i=>(i.pid||i.id)===p.id);
                    return (
                      <div key={p.id} style={{display:"flex",alignItems:"center",gap:10,background:"#fff",borderRadius:8,padding:"8px 12px",border:"1.5px solid #e5e5e5"}}>
                        <div style={{flex:1,minWidth:0}}>
                          <div style={{fontWeight:600,fontSize:12,lineHeight:1.3}}>{p.name}</div>
                          <div style={{fontSize:10,color:"#aaa"}}>{p.id}</div>
                        </div>
                        {already
                          ? <span style={{display:"inline-flex",alignItems:"center",gap:4,fontSize:11,color:"#1e8449",fontWeight:700,background:"#d5f5e3",borderRadius:6,padding:"3px 8px"}}><CheckCircle size={11} strokeWidth={2.6}/>Ya está (×{already.qty})</span>
                          : <button onClick={()=>{
                              const newItems=[...po.items,{pid:p.id,id:p.id,name:p.name,qty:p.multiploCompra||1,notas:""}];
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
              {isAdmin && <th style={{padding:"10px 12px",textAlign:"right",fontSize:11,color:"#888",fontWeight:700,textTransform:"uppercase"}}>Costo</th>}
              <th style={{padding:"10px 12px",textAlign:"left",fontSize:11,color:"#888",fontWeight:700,textTransform:"uppercase"}}>Obs.</th>
              {isAdmin && po.estado!=="cerrada" && <th style={{padding:"10px 12px",textAlign:"center",fontSize:11,color:"#888",fontWeight:700,textTransform:"uppercase"}}>Quitar</th>}
            </tr></thead>
            <tbody>
              {po.items.map((it,i)=>(
                <tr key={i} style={{borderTop:"1px solid #f5f5f5",background:i%2===0?"#fff":"#fafafa"}}>
                  <td style={{padding:"9px 12px",fontWeight:600}}>{it.name}</td>
                  <td style={{padding:"9px 12px",color:"#aaa",fontSize:11}}>{it.id||it.pid||""}</td>
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
                  {isAdmin && (()=>{
                    const prod = products.find(p=>p.id===(it.id||it.pid));
                    const costo = prod?.costPrice||0;
                    return <td style={{padding:"9px 12px",textAlign:"right",fontSize:12,color:"#666",whiteSpace:"nowrap"}}>{costo>0?fARS(costo*it.qty):"—"}</td>;
                  })()}
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
          <div style={{padding:"10px 14px",borderTop:"1px solid #f0f0f0",display:"flex",justifyContent:"flex-end",gap:20,fontSize:12,color:"#888",flexWrap:"wrap"}}>
            <span>Total productos: <strong style={{color:"#1a5276",fontSize:14}}>{po.items.length}</strong></span>
            <span>Total unidades: <strong style={{color:"#1a5276",fontSize:14}}>{po.items.reduce((s,i)=>s+i.qty,0)}</strong></span>
            {isAdmin && <span>Costo total: <strong style={{color:RED,fontFamily:SERIF,fontSize:14}}>{fARS(po.items.reduce((s,i)=>{const prod=products.find(p=>p.id===(i.id||i.pid));return s+(prod?.costPrice||0)*i.qty;},0))}</strong></span>}
          </div>
        </div>

        {/* Actions */}
        <div style={{display:"flex",gap:8,flexWrap:"wrap",padding:isMobile?"0 12px":"0"}}>
          {isAdmin && po.estado==="abierta" && <button onClick={()=>changeEstado(po,"revisando")} style={{flex:1,padding:"11px",borderRadius:10,border:"none",background:"#e67e22",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}>🔍 Marcar como Revisando</button>}
          {isAdmin && po.estado==="revisando" && <button onClick={()=>changeEstado(po,"cerrada")} style={{flex:1,padding:"11px",borderRadius:10,border:"none",background:"#1e8449",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}><><CheckCircle size={12} strokeWidth={2.4}/> Cerrar</> solicitud</button>}
          <button onClick={()=>printSolicitudPDF(po, PDF_LOGO_BANNER)} style={{flex:1,padding:"11px",borderRadius:10,border:"1.5px solid #1a5276",background:"#fff",color:"#1a5276",fontWeight:700,fontSize:13,cursor:"pointer",display:"inline-flex",alignItems:"center",justifyContent:"center",gap:6}}><Printer size={13} strokeWidth={2.4}/> PDF</button>
          <button onClick={()=>exportSolicitudXLSX(po)} style={{flex:1,padding:"11px",borderRadius:10,border:"1.5px solid #1e8449",background:"#fff",color:"#1e8449",fontWeight:700,fontSize:13,cursor:"pointer",display:"inline-flex",alignItems:"center",justifyContent:"center",gap:6}}><BarChart size={13} strokeWidth={2.4}/> Excel</button>
          {isAdmin&&<button onClick={()=>deletePO(po.id)} style={{padding:"11px 16px",borderRadius:10,border:"1.5px solid #fcc",background:"#fff",color:"#c0392b",fontWeight:700,fontSize:13,cursor:"pointer",display:"inline-flex",alignItems:"center"}}><Trash size={13} strokeWidth={2.4}/></button>}
        </div>

        {/* Ingresar mercadería desde solicitud */}
        {isAdmin && po.estado==="cerrada" && (
          <div style={{background:"#d5f5e3",border:"2px solid #1e8449",borderRadius:12,padding:16,marginTop:12,margin:isMobile?"12px 12px 0":"12px 0 0"}}>
            <div style={{fontWeight:800,fontSize:14,color:"#1e8449",marginBottom:8,display:"flex",alignItems:"center",gap:6}}><Package size={14} strokeWidth={2.3}/>¿Llegó la mercadería?</div>
            <div style={{fontSize:12,color:"#1e8449",marginBottom:12,lineHeight:1.5}}>
              Podés ingresar el stock recibido directamente desde esta solicitud. Ajustá las cantidades si recibiste diferente a lo pedido.
            </div>
            <IngresarDesdeSolicitud po={po} products={products} onStock={onStockExternal} onDone={()=>setView("lista")}
              onArrived={async()=>{
                const updated = {...po, fechaRecibido: today()};
                setPurchaseOrders(prev=>prev.map(x=>x.id===po.id?updated:x));
                setSelected(updated);
                await db.savePurchaseOrder(updated);
              }}/>
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
          <div style={{fontWeight:800,fontSize:15,marginBottom:12,color:"#1a5276",display:"flex",alignItems:"center",gap:7}}><ClipboardList size={15} strokeWidth={2.3}/>Nueva Solicitud de Compra</div>
          <input onChange={e=>{
              const q=e.target.value.toLowerCase();
              // handled by ProductSelector below
            }} placeholder="Usá el buscador de productos abajo" style={{...inputStyle,background:"#f9f9f9",color:"#aaa"}} readOnly/>
        </div>
        <ProductSelector products={products} cart={cart} setCart={setCart} isMobile={isMobile} loteMode={true}/>
      </div>

      <div style={{position:isMobile?"static":"sticky",top:16,margin:isMobile?"12px":"0"}}>
        <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 2px 12px #0002",border:"2px solid #d6eaf8"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14,color:"#1a5276",display:"flex",alignItems:"center",gap:7}}><ClipboardList size={15} strokeWidth={2.3}/>Resumen de Solicitud</div>
          <div style={{background:"#d6eaf8",borderRadius:8,padding:"7px 12px",fontSize:12,color:"#1a5276",marginBottom:14}}>
            ℹ️ Esta solicitud <strong>no modifica el stock</strong>. Es solo para pedir al proveedor.
          </div>

          {/* Items con notas por item */}
          {cart.length===0
            ? <div style={{textAlign:"center",color:"#aaa",fontSize:12,padding:"16px 0"}}>Agregá productos del catálogo</div>
            : cart.map(i=>{
                const prod = products.find(p=>p.id===i.pid);
                const costo = prod?.costPrice||0;
                return (
                <div key={i.pid} style={{background:"#f9f9f9",borderRadius:8,padding:"10px 12px",marginBottom:8}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                    <span style={{fontWeight:600,fontSize:12,flex:1,lineHeight:1.3,marginRight:8}}>{i.name}</span>
                    <div style={{display:"flex",alignItems:"center",gap:4}}>
                      <button onClick={()=>setQty(i.pid,i.qty-1)} style={{width:24,height:24,borderRadius:5,border:"1.5px solid #ccc",background:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>-</button>
                      <input type="number" value={i.qty} onChange={e=>setQty(i.pid,+e.target.value||0)} style={{width:40,textAlign:"center",padding:"2px",borderRadius:5,border:"1.5px solid #1a5276",fontWeight:700,fontSize:13,outline:"none"}}/>
                      <button onClick={()=>setQty(i.pid,i.qty+1)} style={{width:24,height:24,borderRadius:5,border:"1.5px solid #ccc",background:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>+</button>
                    </div>
                  </div>
                  {isAdmin && costo>0 && (
                    <div style={{fontSize:10.5,color:"#999",marginBottom:6}}>{fARS(costo)} c/u · <strong style={{color:"#555"}}>{fARS(costo*i.qty)}</strong> subtotal</div>
                  )}
                  <input value={itemNotas[i.pid]||""} onChange={e=>setItemNotas(n=>({...n,[i.pid]:e.target.value}))}
                    placeholder="Obs. (color, medida, marca...)" style={{...inputStyle,fontSize:11,padding:"4px 8px",background:"#fff"}}/>
                </div>
                );
              })
          }
          {isAdmin && cart.length>0 && (
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 12px",background:"#fdecea",borderRadius:8,marginBottom:14}}>
              <span style={{fontWeight:700,fontSize:12.5,color:"#1a1a1a"}}>Costo total estimado</span>
              <span style={{fontFamily:SERIF,fontWeight:700,fontSize:16,color:RED}}>{fARS(cart.reduce((s,i)=>s+(products.find(p=>p.id===i.pid)?.costPrice||0)*i.qty,0))}</span>
            </div>
          )}

          <Field label="Notas generales"><textarea value={notas} onChange={e=>setNotas(e.target.value)} placeholder="Urgencia, proveedor sugerido..." style={{...inputStyle,resize:"vertical",minHeight:52,fontSize:12}}/></Field>

          <div style={{display:"flex",gap:8,marginTop:8}}>
            <button onClick={createNew} disabled={!cart.length||saving}
              style={{flex:1,padding:"11px",borderRadius:10,border:"none",fontWeight:800,fontSize:14,cursor:"pointer",
                background:(!cart.length||saving)?"#e5e5e5":"linear-gradient(135deg,#1a5e20,#1e8449)",
                color:(!cart.length||saving)?"#aaa":"#fff"}}>
              {saving?"Guardando...":<><ClipboardList size={13} strokeWidth={2.4}/> Crear Solicitud</>}
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
  const [saving,setSaving]=useState(false);
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
        setScanMsg({tipo:"ok", texto:`${p.name} agregado`});
      } else {
        setScanMsg({tipo:"info", texto:`${p.name} ya está en la lista`});
      }
    } else {
      setScanMsg({tipo:"warn", texto:`Código "${code}" no encontrado en el catálogo`});
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

  const submit=async()=>{
    if(saving) return;
    setSaving(true);
    try {
      for(const i of items) { await onStock(i.pid,+i.qty,+i.cost); }
      setItems([]); setSearch(""); setMStep(1);
      setOk(true); setTimeout(()=>setOk(false),2000);
    } catch(e) {
      console.warn("Error al ingresar al stock:", e);
      toast.error("Hubo un problema de conexión. Revisá el Stock — puede que ya se haya actualizado igual.");
    } finally {
      setSaving(false);
    }
  };

  if(saving) return <SaveSpinner label="Ingresando al stock..." color="#1e8449"/>;
  if(ok) return <div style={{textAlign:"center",padding:80}}><div style={{display:"flex",justifyContent:"center"}}><Package size={52} color="#1e8449" strokeWidth={1.8}/></div><div style={{fontWeight:800,color:"#1e8449",fontSize:20,marginTop:12}}>¡Stock actualizado!</div></div>;

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
      {scanMsg && <div style={{display:"flex",alignItems:"center",gap:7,background:scanMsg.tipo==="ok"?"#eafaf1":scanMsg.tipo==="info"?"#eaf4fc":"#fef9e7",border:"1.5px solid",borderColor:scanMsg.tipo==="ok"?"#a9dfbf":scanMsg.tipo==="info"?"#aed6f1":"#f0d080",borderRadius:10,padding:"10px 14px",marginBottom:10,fontSize:13,fontWeight:600,color:scanMsg.tipo==="ok"?"#1e8449":scanMsg.tipo==="info"?"#1a5276":"#b7770d"}}>
        {scanMsg.tipo==="ok"?<CheckCircle size={14} strokeWidth={2.4}/>:scanMsg.tipo==="info"?<Info size={14} strokeWidth={2.2}/>:<AlertTriangle size={14} strokeWidth={2.3}/>}
        {scanMsg.texto}
      </div>}

      {/* Paso 1 — Buscar */}
      {mStep===1 && <>
        <div style={{background:"#fff",borderRadius:12,padding:16,marginBottom:12,boxShadow:"0 1px 4px #0001"}}>
          <div style={{fontWeight:800,fontSize:14,marginBottom:10,color:"#1a1a1a",display:"flex",alignItems:"center",gap:6}}><Package size={14} strokeWidth={2.3}/>Buscá los productos recibidos</div>

          {/* Scanner button — only for admin/enabled users */}
          {canScan && <button onClick={()=>setShowScanner(true)}
            style={{width:"100%",padding:"12px",borderRadius:10,border:"none",background:"linear-gradient(135deg,#1a5276,#2980b9)",color:"#fff",fontWeight:800,fontSize:14,cursor:"pointer",marginBottom:10,display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>
            <Camera size={15} strokeWidth={2.3}/> Escanear código de barras
          </button>}

          <div style={{position:"relative"}}>
            <Search size={14} color="#aaa" strokeWidth={2.3} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)"}}/>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="O buscá por nombre o código SKU..."
              style={{width:"100%",padding:"10px 12px 10px 32px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:14,outline:"none",boxSizing:"border-box"}}/>
          </div>
          {search&&<div style={{fontSize:11,color:"#aaa",marginTop:6}}>{found.length} resultados</div>}

          {/* Producto manual */}
          <div style={{marginTop:12,borderTop:"1px solid #f0f0f0",paddingTop:12}}>
            <button onClick={()=>{setShowManual(s=>!s);setManualError("");}}
              style={{width:"100%",padding:"9px",borderRadius:8,border:"1.5px solid #e67e22",background:showManual?"#fef9e7":"#fff",color:"#e67e22",fontWeight:700,fontSize:13,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
              {showManual?<><XIcon size={12} strokeWidth={2.6}/> Cancelar</>:<><Plus size={13} strokeWidth={2.4}/> Producto no está en el catálogo</>}
            </button>
            {showManual && (
              <div style={{background:"#fef9e7",border:"1.5px solid #e67e22",borderRadius:10,padding:14,marginTop:10}}>
                <div style={{fontWeight:700,fontSize:13,color:"#e67e22",marginBottom:6,display:"flex",alignItems:"center",gap:5}}><AlertTriangle size={13} strokeWidth={2.3}/>Producto nuevo (excepcional)</div>
                <Field label="SKU *"><input value={manualForm.sku} onChange={e=>setManualForm(f=>({...f,sku:e.target.value.toUpperCase()}))} placeholder="Ej: PROD-001" style={inputStyle}/></Field>
                <Field label="Detalle *"><input value={manualForm.name} onChange={e=>setManualForm(f=>({...f,name:e.target.value}))} placeholder="Nombre completo del producto" style={inputStyle}/></Field>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                  <Field label="Cantidad"><input type="number" min={1} value={manualForm.qty} onChange={e=>setManualForm(f=>({...f,qty:e.target.value}))} style={inputStyle}/></Field>
                  <Field label="Costo ($)"><input type="number" min={0} value={manualForm.cost} onChange={e=>setManualForm(f=>({...f,cost:e.target.value}))} style={inputStyle}/></Field>
                </div>
                {manualError && <div style={{display:"flex",alignItems:"center",gap:5,color:"#c0392b",fontSize:12,margin:"6px 0",fontWeight:600}}><AlertTriangle size={11} strokeWidth={2.4}/>{manualError}</div>}
                <button onClick={addManual} style={{width:"100%",padding:"10px",borderRadius:8,border:"none",background:"#e67e22",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer",marginTop:4,display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
                  <Plus size={13} strokeWidth={2.4}/> Agregar a la compra
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
                    cursor:inL?"not-allowed":"pointer",flexShrink:0,whiteSpace:"nowrap",display:"flex",alignItems:"center",gap:5}}>
                  {inL?<><CheckCircle size={12} strokeWidth={2.5}/> Agregado</>:"+ Agregar"}
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
            <button onClick={()=>setMStep(2)} style={{padding:"10px 20px",borderRadius:10,border:"none",background:"#fff",color:"#1e8449",fontWeight:800,fontSize:14,cursor:"pointer",display:"flex",alignItems:"center",gap:6}}>
              Ver detalle <ArrowRightIcon size={13} strokeWidth={2.4}/>
            </button>
          </div>
        )}
      </>}

      {/* Paso 2 — Detalle y confirmar */}
      {mStep===2 && (
        <div style={{paddingBottom:80}}>
          <div style={{fontWeight:800,fontSize:14,marginBottom:12,color:"#1a1a1a",display:"flex",alignItems:"center",gap:6}}><FileText size={14} strokeWidth={2.3}/>Detalle de Compra</div>
          {items.map(it=>(
            <div key={it.pid} style={{background:"#fff",borderRadius:12,padding:14,marginBottom:10,boxShadow:"0 1px 4px #0001",border:it.esNuevo?"1.5px solid #e67e22":"none"}}>
              {it.esNuevo && <div style={{fontSize:10,color:"#e67e22",fontWeight:700,marginBottom:4,display:"flex",alignItems:"center",gap:4}}><AlertTriangle size={9} strokeWidth={2.6}/>NUEVO en catálogo</div>}
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

          <button onClick={()=>setMStep(1)} style={{width:"100%",padding:"10px",borderRadius:10,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontWeight:600,fontSize:13,cursor:"pointer",marginBottom:8,display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
            <ArrowLeftIcon size={11} strokeWidth={2.4}/> Seguir agregando productos
          </button>

          {/* Barra flotante confirmar */}
          <div style={{position:"fixed",bottom:0,left:0,right:0,background:"linear-gradient(135deg,#1a5e20,#1e8449)",color:"#fff",padding:"13px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",zIndex:200,boxShadow:"0 -3px 16px #0003"}}>
            <div>
              <div style={{fontWeight:800,fontSize:15}}>{fARS(totalCost)}</div>
              <div style={{fontSize:11,opacity:.85}}>{items.length} producto{items.length!==1?"s":""}</div>
            </div>
            <button onClick={submit} style={{padding:"10px 20px",borderRadius:10,border:"none",background:"#fff",color:"#1e8449",fontWeight:800,fontSize:14,cursor:"pointer",display:"flex",alignItems:"center",gap:6}}>
              <Package size={13} strokeWidth={2.4}/> Ingresar al Stock
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
          <div style={{fontWeight:800,fontSize:15,marginBottom:12,display:"flex",alignItems:"center",gap:7}}><Package size={15} strokeWidth={2.3}/>Ingresar al Stock</div>
          {scanMsg && <div style={{display:"flex",alignItems:"center",gap:7,background:scanMsg.tipo==="ok"?"#eafaf1":"#fef9e7",borderRadius:8,padding:"8px 12px",fontSize:13,fontWeight:600,marginBottom:10,color:scanMsg.tipo==="ok"?"#1e8449":"#b7770d"}}>{scanMsg.tipo==="ok"?<CheckCircle size={14} strokeWidth={2.4}/>:<AlertTriangle size={14} strokeWidth={2.3}/>}{scanMsg.texto}</div>}
          <div style={{display:"flex",gap:8,marginBottom:10}}>
            <div style={{position:"relative",flex:1}}>
              <Search size={13} color="#aaa" strokeWidth={2.3} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)"}}/>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscá los productos que recibiste..."
                style={{width:"100%",padding:"8px 12px 8px 32px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",boxSizing:"border-box"}}/>
            </div>
            {canScan && <button onClick={()=>setShowScanner(true)}
              style={{padding:"8px 14px",borderRadius:8,border:"none",background:"linear-gradient(135deg,#1a5276,#2980b9)",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer",whiteSpace:"nowrap",display:"flex",alignItems:"center",gap:6}}>
              <Camera size={13} strokeWidth={2.3}/> Escanear
            </button>}
          </div>
          {search&&<div style={{fontSize:11,color:"#aaa",marginBottom:6}}>{found.length} resultados</div>}
          <div style={{marginTop:12,borderTop:"1px solid #f0f0f0",paddingTop:12}}>
            <button onClick={()=>{setShowManual(s=>!s);setManualError("");}}
              style={{padding:"6px 14px",borderRadius:8,border:"1.5px solid #e67e22",background:showManual?"#fef9e7":"#fff",color:"#e67e22",fontWeight:700,fontSize:12,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6}}>
              {showManual?<><XIcon size={11} strokeWidth={2.6}/> Cancelar</>:<><Plus size={12} strokeWidth={2.4}/> Producto no está en el catálogo</>}
            </button>
            {showManual && (
              <div style={{background:"#fef9e7",border:"1.5px solid #e67e22",borderRadius:10,padding:14,marginTop:10}}>
                <div style={{fontWeight:700,fontSize:13,color:"#e67e22",marginBottom:10,display:"flex",alignItems:"center",gap:6}}><AlertTriangle size={13} strokeWidth={2.3}/>Producto nuevo (excepcional)</div>
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
                {manualError && <div style={{display:"flex",alignItems:"center",gap:5,color:"#c0392b",fontSize:12,marginBottom:8,fontWeight:600}}><AlertTriangle size={11} strokeWidth={2.4}/>{manualError}</div>}
                <button onClick={addManual} style={{width:"100%",padding:"8px",borderRadius:8,border:"none",background:"#e67e22",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}><Plus size={12} strokeWidth={2.4}/> Agregar a la compra</button>
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
                  <button onClick={()=>addI(p)} disabled={!!inL} style={{width:"100%",padding:"7px",borderRadius:7,border:"none",fontSize:12,fontWeight:700,background:inL?"#d5f5e3":"#1e8449",color:inL?"#1a5276":"#fff",cursor:inL?"not-allowed":"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:5}}>{inL?<><CheckCircle size={11} strokeWidth={2.6}/> Agregado</>:"+ Agregar"}</button>
                </div>;
              })
            : <div style={{padding:20,color:"#aaa",fontSize:13}}>{search?"Sin resultados.":"Escribí el nombre del producto a ingresar."}</div>
          }
        </div>
      </div>
      <div style={{position:"sticky",top:16}}>
        <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 2px 12px #0002"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14,display:"flex",alignItems:"center",gap:7}}><FileText size={15} strokeWidth={2.3}/>Detalle de Compra</div>
          {items.length===0
            ? <div style={{textAlign:"center",color:"#aaa",fontSize:12,padding:"16px 0"}}>Seleccioná productos</div>
            : items.map(it=>(
                <div key={it.pid} style={{background:"#f9f9f9",borderRadius:8,padding:12,marginBottom:10,border:it.esNuevo?"1.5px solid #e67e22":"none"}}>
                  {it.esNuevo && <div style={{fontSize:10,color:"#e67e22",fontWeight:700,marginBottom:4,display:"flex",alignItems:"center",gap:4}}><AlertTriangle size={9} strokeWidth={2.6}/>NUEVO en catálogo</div>}
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
          <button onClick={submit} disabled={!items.length||saving} style={{width:"100%",padding:"11px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:!items.length?"#e5e5e5":"linear-gradient(135deg,#1a5e20,#1e8449)",color:!items.length?"#aaa":"#fff",display:"flex",alignItems:"center",justifyContent:"center",gap:7}}><Package size={13} strokeWidth={2.4}/> Ingresar al Stock</button>
        </div>
      </div>
    </div>
  );
}

// ─── SANDBOX STOCK MANAGER ────────────────────────────────────────────────────
function SandboxStockManager({products, sandboxStock, updateSandboxStock}) {
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
  const applyChanges = () => { updateSandboxStock(prev=>({...prev,...edited})); setEdited({}); };
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
// ─── PAGOS — cola centralizada de revisión de comprobantes y efectivo ─────────
function PagosPanel({orders, onConfirmarComprobante, onRechazarComprobante, onConfirmarEfectivo, onRechazarEfectivo, settings, onToggleExigirPago}) {
  const [tab, setTab] = useState("pendientes");
  const [busyId, setBusyId] = useState(null);
  const [rejectId, setRejectId] = useState(null);
  const [rejectMotivo, setRejectMotivo] = useState("");

  const pendientes = orders.filter(o=>o.pagoTipo==="comprobante_pendiente"||o.pagoTipo==="efectivo_pendiente")
    .sort((a,b)=>(b.date||"").localeCompare(a.date||""));
  const confirmados = orders.filter(o=>normPagoTipo(o.pagoTipo)==="comprobante_confirmado"||normPagoTipo(o.pagoTipo)==="efectivo_confirmado")
    .sort((a,b)=>(b.date||"").localeCompare(a.date||""));

  const TABS = [
    {k:"pendientes",  label:"Pendientes",  n:pendientes.length,  color:"#b7770d"},
    {k:"confirmados", label:"Confirmados", n:confirmados.length, color:"#1e8449"},
  ];

  const handleConfirm = async (o) => {
    setBusyId(o.id);
    try {
      if(o.pagoTipo==="comprobante_pendiente") await onConfirmarComprobante(o.id);
      else await onConfirmarEfectivo(o.id);
    } catch(e) { console.warn(e); toast.error("No se pudo confirmar. Probá de nuevo."); }
    finally { setBusyId(null); }
  };

  const handleReject = (o) => {
    if(o.pagoTipo==="efectivo_pendiente") { rejectEfectivo(o); return; }
    setRejectId(o.id);
  };

  const rejectEfectivo = async (o) => {
    setBusyId(o.id);
    try { await onRechazarEfectivo(o.id); } catch(e) { console.warn(e); toast.error("No se pudo rechazar. Probá de nuevo."); }
    finally { setBusyId(null); }
  };

  const confirmRejectComprobante = async (o) => {
    if(!rejectMotivo.trim()) return;
    setBusyId(o.id);
    try { await onRechazarComprobante(o.id, rejectMotivo); setRejectId(null); setRejectMotivo(""); }
    catch(e) { console.warn(e); toast.error("No se pudo rechazar. Probá de nuevo."); }
    finally { setBusyId(null); }
  };

  return (
    <div>
      {/* Header + setting */}
      <div style={{background:"#fff",borderRadius:12,padding:"16px 18px",marginBottom:16,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:16,color:"#1a1a1a",marginBottom:4,display:"flex",alignItems:"center",gap:7}}><CreditCard size={16} strokeWidth={2.2}/>Control de Pagos</div>
        <div style={{fontSize:12.5,color:"#888",marginBottom:14}}>Revisá y confirmá los comprobantes y pagos en efectivo declarados por los vendedores.</div>

        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 14px",background:settings?.exigirPagoConfirmado?"#fdecea":"#f8f9fa",borderRadius:9,border:`1px solid ${settings?.exigirPagoConfirmado?"#f1948a":"#e5e5e5"}`,flexWrap:"wrap",gap:10}}>
          <div style={{flex:1,minWidth:240}}>
            <div style={{fontWeight:700,fontSize:13,color:settings?.exigirPagoConfirmado?"#c0392b":"#333",display:"flex",alignItems:"center",gap:6}}><Lock size={13} strokeWidth={2.3}/>Exigir pago confirmado para habilitar "Entregar"</div>
            <div style={{fontSize:11.5,color:"#888",marginTop:2}}>
              {settings?.exigirPagoConfirmado
                ? "Activado: los pedidos no se pueden marcar como Entregado hasta confirmar el comprobante o el efectivo."
                : "Desactivado: la entrega funciona como siempre, sin esperar la confirmación de pago."}
            </div>
          </div>
          <button onClick={onToggleExigirPago}
            style={{position:"relative",width:50,height:28,borderRadius:20,border:"none",cursor:"pointer",background:settings?.exigirPagoConfirmado?"#c0392b":"#ccc",flexShrink:0}}>
            <div style={{position:"absolute",top:3,left:settings?.exigirPagoConfirmado?25:3,width:22,height:22,borderRadius:"50%",background:"#fff",boxShadow:"0 1px 3px #0003",transition:"left .2s"}}/>
          </button>
        </div>
      </div>

      {/* Tabs */}
      <div style={{display:"flex",gap:6,marginBottom:14,flexWrap:"wrap"}}>
        {TABS.map(t=>(
          <button key={t.k} onClick={()=>setTab(t.k)}
            style={{padding:"7px 16px",borderRadius:20,border:`1.5px solid ${tab===t.k?t.color:"#e5e5e5"}`,background:tab===t.k?t.color:"#fff",color:tab===t.k?"#fff":"#666",fontWeight:700,fontSize:12.5,cursor:"pointer"}}>
            {t.label} ({t.n})
          </button>
        ))}
      </div>

      {/* List */}
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        {tab==="pendientes" && (pendientes.length ? pendientes.map(o=>(
          <div key={o.id} style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 4px #0001",border:`1.5px solid ${o.pagoTipo==="comprobante_pendiente"?"#aed6f1":"#f0d080"}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:12,flexWrap:"wrap"}}>
              <div style={{flex:1,minWidth:200}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4,flexWrap:"wrap"}}>
                  <span style={{fontWeight:700,fontSize:14,color:"#1a1a1a"}}>{o.client}</span>
                  {o.pagoTipo==="comprobante_pendiente"
                    ? <span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#eaf2f8",color:"#1a5276",border:"1px solid #aed6f1",borderRadius:20,padding:"2px 8px",fontSize:11,fontWeight:700}}><Paperclip size={10} strokeWidth={2.5}/>Comprobante</span>
                    : <span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#fef9e7",color:"#b7770d",border:"1px solid #f0d080",borderRadius:20,padding:"2px 8px",fontSize:11,fontWeight:700}}><Banknote size={10} strokeWidth={2.5}/>Efectivo</span>}
                </div>
                <div style={{fontSize:11.5,color:"#888"}}>{o.compNum||o.docNum} · {o.date} · {o.vendedor}</div>
              </div>
              <div style={{fontSize:18,fontWeight:800,color:RED,whiteSpace:"nowrap"}}>{fARS(o.total)}</div>
            </div>

            {o.pagoTipo==="comprobante_pendiente" && (
              <div style={{display:"flex",alignItems:"center",gap:10,marginTop:10,padding:"8px 10px",background:"#fafbfc",borderRadius:8}}>
                {/\.(jpe?g|png|gif|webp)$/i.test(o.comprobanteUrl)
                  ? <img src={o.comprobanteUrl} alt="comprobante" style={{width:44,height:44,borderRadius:7,objectFit:"cover",border:"1.5px solid #aed6f1",flexShrink:0}}/>
                  : <div style={{width:44,height:44,borderRadius:7,background:"#e8eef3",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><FileText size={18} color="#1a5276" strokeWidth={2}/></div>}
                <div style={{fontSize:12,color:"#555",flex:1,minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{o.comprobanteNombre||"Comprobante"}</div>
                <a href={o.comprobanteUrl} target="_blank" rel="noopener noreferrer"
                  style={{fontSize:11.5,fontWeight:700,color:"#1a5276",textDecoration:"none",padding:"4px 10px",border:"1px solid #aed6f155",borderRadius:6,flexShrink:0,display:"inline-flex",alignItems:"center",gap:4}}><Eye size={11} strokeWidth={2.4}/>Ver</a>
              </div>
            )}
            {o.pagoTipo==="efectivo_pendiente" && (
              <div style={{marginTop:10,padding:"8px 10px",background:"#fffdf6",borderRadius:8,fontSize:12,color:"#b7770d",fontWeight:600,display:"flex",alignItems:"flex-start",gap:6}}>
                <Banknote size={13} strokeWidth={2.3} style={{flexShrink:0,marginTop:1}}/>El vendedor declaró haber cobrado en efectivo{o.pagoEfectivoFecha&&` el ${o.pagoEfectivoFecha}`}.
              </div>
            )}

            {rejectId===o.id ? (
              <div style={{marginTop:12,padding:"10px 12px",background:"#fdecea",borderRadius:8,border:"1px solid #f1948a"}}>
                <div style={{fontSize:11.5,fontWeight:700,color:"#c0392b",marginBottom:6}}>Motivo del rechazo (el vendedor lo va a ver)</div>
                <textarea value={rejectMotivo} onChange={e=>setRejectMotivo(e.target.value)} rows={2}
                  placeholder="Ej: el monto no coincide, la imagen no se lee..."
                  style={{width:"100%",padding:8,borderRadius:6,border:"1px solid #e5b4b0",fontSize:12.5,fontFamily:"inherit",resize:"vertical",boxSizing:"border-box"}}/>
                <div style={{display:"flex",gap:8,marginTop:8}}>
                  <button disabled={!rejectMotivo.trim()||busyId===o.id}
                    onClick={()=>confirmRejectComprobante(o)}
                    style={{padding:"6px 14px",borderRadius:7,border:"none",background:rejectMotivo.trim()?"#c0392b":"#e5b4b0",color:"#fff",fontWeight:700,fontSize:12,cursor:rejectMotivo.trim()?"pointer":"not-allowed"}}>
                    Confirmar rechazo
                  </button>
                  <button onClick={()=>{setRejectId(null);setRejectMotivo("");}}
                    style={{padding:"6px 14px",borderRadius:7,border:"1px solid #ccc",background:"#fff",color:"#666",fontWeight:600,fontSize:12,cursor:"pointer"}}>
                    Cancelar
                  </button>
                </div>
              </div>
            ) : (
              <div style={{display:"flex",gap:8,marginTop:12}}>
                <button disabled={busyId===o.id} onClick={()=>handleConfirm(o)}
                  style={{flex:1,padding:"8px 14px",borderRadius:8,border:"none",background:"#1e8449",color:"#fff",fontWeight:700,fontSize:12.5,cursor:"pointer",display:"inline-flex",alignItems:"center",justifyContent:"center",gap:6}}>
                  {busyId===o.id?"Guardando...":<><CheckCircle size={13} strokeWidth={2.4}/> Confirmar — el pago es correcto</>}
                </button>
                <button onClick={()=>handleReject(o)}
                  style={{padding:"8px 14px",borderRadius:8,border:"1.5px solid #f1948a",background:"#fff",color:"#c0392b",fontWeight:700,fontSize:12.5,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6}}>
                  <XCircle size={13} strokeWidth={2.4}/> Rechazar
                </button>
              </div>
            )}
          </div>
        )) : (
          <div style={{textAlign:"center",padding:40,background:"#fff",borderRadius:12,color:"#aaa"}}>
            <div style={{marginBottom:6,display:"flex",justifyContent:"center"}}><PartyPopper size={32} color="#a9dfbf" strokeWidth={1.8}/></div>
            <div style={{fontWeight:600,fontSize:13}}>No hay pagos pendientes de revisión</div>
          </div>
        ))}

        {tab==="confirmados" && (confirmados.length ? confirmados.map(o=>(
          <div key={o.id} style={{background:"#fff",borderRadius:12,padding:"12px 16px",boxShadow:"0 1px 4px #0001",border:"1px solid #e5e5e5"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:12,flexWrap:"wrap"}}>
              <div>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:3,flexWrap:"wrap"}}>
                  <span style={{fontWeight:700,fontSize:13.5,color:"#1a1a1a"}}>{o.client}</span>
                  {normPagoTipo(o.pagoTipo)==="comprobante_confirmado"
                    ? <span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#eaf2f8",color:"#1a5276",border:"1px solid #aed6f1",borderRadius:20,padding:"2px 8px",fontSize:11,fontWeight:700}}><Paperclip size={10} strokeWidth={2.5}/>Comprobante</span>
                    : <span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#fef9e7",color:"#b7770d",border:"1px solid #f0d080",borderRadius:20,padding:"2px 8px",fontSize:11,fontWeight:700}}><Banknote size={10} strokeWidth={2.5}/>Efectivo</span>}
                  <span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#eafaf1",color:"#1e8449",border:"1px solid #a9dfbf",borderRadius:20,padding:"2px 8px",fontSize:11,fontWeight:700}}><CheckCircle size={10} strokeWidth={2.5}/>Confirmado</span>
                </div>
                <div style={{fontSize:11,color:"#888"}}>{o.compNum||o.docNum} · {o.vendedor}</div>
              </div>
              <div style={{fontSize:15,fontWeight:800,color:"#444"}}>{fARS(o.total)}</div>
            </div>
          </div>
        )) : (
          <div style={{textAlign:"center",padding:40,background:"#fff",borderRadius:12,color:"#aaa"}}>
            <div style={{marginBottom:6,display:"flex",justifyContent:"center"}}><ClipboardList size={32} color="#ddd" strokeWidth={1.8}/></div>
            <div style={{fontWeight:600,fontSize:13}}>Todavía no confirmaste ningún pago</div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── ANÁLISIS DE COMPRAS Y COSTOS — cambios de costo del proveedor + trazabilidad ──
function AnalisisComprasPanel({products, purchaseOrders}) {
  const [tab, setTab] = useState("costos");

  const cambios = useMemo(()=>products
    .filter(p=>p.costPriceAnterior>0 && p.costPriceAnterior!==p.costPrice)
    .map(p=>{
      const pct = (p.costPrice - p.costPriceAnterior) / p.costPriceAnterior * 100;
      return {...p, pct, negativo: p.costPrice >= p.salePrice};
    })
    .sort((a,b)=>b.pct-a.pct), [products]);
  const negativos = cambios.filter(c=>c.negativo);
  const subieronFuerte = cambios.filter(c=>!c.negativo && c.pct>=15);

  const entregas = useMemo(()=>(purchaseOrders||[])
    .filter(po=>po.fechaCierre && po.fechaRecibido)
    .map(po=>{
      const d1 = parseFechaLog(po.fechaCierre), d2 = parseFechaLog(po.fechaRecibido);
      const dias = (d1&&d2) ? Math.round((d2-d1)/86400000) : null;
      return {...po, dias};
    })
    .filter(po=>po.dias!==null)
    .sort((a,b)=>parseFechaLog(b.fechaRecibido)-parseFechaLog(a.fechaRecibido)), [purchaseOrders]);
  const promedio = entregas.length ? (entregas.reduce((s,e)=>s+e.dias,0)/entregas.length).toFixed(1) : null;
  const masLenta = entregas.length ? [...entregas].sort((a,b)=>b.dias-a.dias)[0] : null;

  return (
    <div>
      <div style={{display:"flex",gap:8,marginBottom:18,maxWidth:640}}>
        <button onClick={()=>setTab("costos")} style={{flex:1,padding:"9px",borderRadius:9,border:"none",cursor:"pointer",fontWeight:700,fontSize:12.5,background:tab==="costos"?"#1a1a1a":"#fff",color:tab==="costos"?"#fff":"#666",boxShadow:"0 1px 4px #0001"}}>
          🔺 Cambios de Costo {cambios.length>0&&`(${cambios.length})`}
        </button>
        <button onClick={()=>setTab("entregas")} style={{flex:1,padding:"9px",borderRadius:9,border:"none",cursor:"pointer",fontWeight:700,fontSize:12.5,background:tab==="entregas"?"#1a1a1a":"#fff",color:tab==="entregas"?"#fff":"#666",boxShadow:"0 1px 4px #0001"}}>
          🚚 Trazabilidad de Entregas
        </button>
      </div>

      <div style={{maxWidth:640}}>
        {tab==="costos" ? (
          cambios.length===0 ? (
            <div style={{textAlign:"center",padding:40,background:"#fff",borderRadius:12,color:"#aaa"}}>
              <div style={{fontSize:36,marginBottom:6}}>📭</div>
              <div style={{fontWeight:600,fontSize:13}}>Todavía no hay cambios de costo registrados</div>
              <div style={{fontSize:11.5,marginTop:4}}>Aparecen acá la próxima vez que reimportes la lista de precios y algún costo cambie.</div>
            </div>
          ) : (
            <>
              <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap"}}>
                <div style={{flex:1,minWidth:140,background:"#fff",borderRadius:12,padding:"14px 16px",borderLeft:"3px solid #c0392b",boxShadow:"0 1px 4px #0001"}}>
                  <div style={{fontFamily:SERIF,fontSize:24,fontWeight:700,color:"#c0392b"}}>{negativos.length}</div>
                  <div style={{fontSize:11.5,color:"#888",marginTop:2,fontWeight:600}}>Margen negativo</div>
                </div>
                <div style={{flex:1,minWidth:140,background:"#fff",borderRadius:12,padding:"14px 16px",borderLeft:"3px solid #b7770d",boxShadow:"0 1px 4px #0001"}}>
                  <div style={{fontFamily:SERIF,fontSize:24,fontWeight:700,color:"#b7770d"}}>{subieronFuerte.length}</div>
                  <div style={{fontSize:11.5,color:"#888",marginTop:2,fontWeight:600}}>Subieron +15%</div>
                </div>
                <div style={{flex:1,minWidth:140,background:"#fff",borderRadius:12,padding:"14px 16px",borderLeft:"3px solid #1a5276",boxShadow:"0 1px 4px #0001"}}>
                  <div style={{fontFamily:SERIF,fontSize:24,fontWeight:700,color:"#1a5276"}}>{cambios.length}</div>
                  <div style={{fontSize:11.5,color:"#888",marginTop:2,fontWeight:600}}>Productos con cambio</div>
                </div>
              </div>

              <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px #0001",overflow:"hidden"}}>
                <div style={{padding:"12px 16px",borderBottom:"1px solid #f0f0f0",fontWeight:800,fontSize:13.5}}>Cambios del último import</div>
                {cambios.map(c=>(
                  <div key={c.id} style={{padding:"11px 16px",borderBottom:"1px solid #f5f5f5",display:"flex",alignItems:"center",gap:10,background:c.negativo?"#fff8f8":"transparent",flexWrap:"wrap"}}>
                    <div style={{flex:1,minWidth:160}}>
                      <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                        <span style={{fontWeight:700,fontSize:12.5,color:"#1a1a1a"}}>{c.name}</span>
                        {c.negativo && <span style={{background:"#fdecea",color:"#c0392b",borderRadius:20,padding:"2px 9px",fontSize:10.5,fontWeight:700}}>🔴 Margen negativo</span>}
                      </div>
                      <div style={{fontSize:11,color:"#999",marginTop:2}}>{c.id} · venta actual {fARS(c.salePrice)}</div>
                    </div>
                    <div style={{textAlign:"right",flexShrink:0}}>
                      <div style={{fontSize:11,color:"#aaa"}}>{fARS(c.costPriceAnterior)} → <strong style={{color:"#555"}}>{fARS(c.costPrice)}</strong></div>
                      <div style={{display:"flex",alignItems:"center",gap:3,justifyContent:"flex-end",marginTop:2}}>
                        {c.pct>0?<TrendUp size={11} color={c.pct>=15?"#c0392b":"#b7770d"} strokeWidth={2.5}/>:<TrendDown size={11} color="#1e8449" strokeWidth={2.5}/>}
                        <span style={{fontSize:12,fontWeight:800,color:c.pct>=15?"#c0392b":c.pct>0?"#b7770d":"#1e8449"}}>{c.pct>0?"+":""}{c.pct.toFixed(1)}%</span>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </>
          )
        ) : (
          entregas.length===0 ? (
            <div style={{textAlign:"center",padding:40,background:"#fff",borderRadius:12,color:"#aaa"}}>
              <div style={{fontSize:36,marginBottom:6}}>📭</div>
              <div style={{fontWeight:600,fontSize:13}}>Todavía no hay entregas registradas</div>
              <div style={{fontSize:11.5,marginTop:4}}>Aparecen acá cada vez que confirmes "Ingresar mercadería" en una solicitud cerrada.</div>
            </div>
          ) : (
            <>
              <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap"}}>
                <div style={{flex:1,minWidth:160,background:"#fff",borderRadius:12,padding:"14px 16px",borderLeft:"3px solid #6c3483",boxShadow:"0 1px 4px #0001"}}>
                  <div style={{fontFamily:SERIF,fontSize:24,fontWeight:700,color:"#6c3483"}}>{promedio}<span style={{fontSize:13}}> días</span></div>
                  <div style={{fontSize:11.5,color:"#888",marginTop:2,fontWeight:600}}>Tiempo de entrega promedio</div>
                </div>
                {masLenta && <div style={{flex:1,minWidth:160,background:"#fff",borderRadius:12,padding:"14px 16px",borderLeft:"3px solid #b7770d",boxShadow:"0 1px 4px #0001"}}>
                  <div style={{fontFamily:SERIF,fontSize:24,fontWeight:700,color:"#b7770d"}}>{masLenta.dias}<span style={{fontSize:13}}> días</span></div>
                  <div style={{fontSize:11.5,color:"#888",marginTop:2,fontWeight:600}}>Entrega más lenta</div>
                </div>}
              </div>

              <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px #0001",overflow:"hidden"}}>
                <div style={{padding:"12px 16px",borderBottom:"1px solid #f0f0f0",fontWeight:800,fontSize:13.5}}>Últimas entregas recibidas</div>
                {entregas.map(e=>(
                  <div key={e.id} style={{padding:"11px 16px",borderBottom:"1px solid #f5f5f5",display:"flex",alignItems:"center",gap:10}}>
                    <div style={{width:30,height:30,borderRadius:9,background:"#f5eef8",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                      <Truck size={14} color="#6c3483" strokeWidth={2.2}/>
                    </div>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontWeight:700,fontSize:12.5,color:"#1a1a1a"}}>{e.items?.length||0} productos</div>
                      <div style={{fontSize:11,color:"#999"}}>👤 {e.vendedor} · cerrada {e.fechaCierre} → recibida {e.fechaRecibido}</div>
                    </div>
                    <div style={{display:"flex",alignItems:"center",gap:5,flexShrink:0}}>
                      <Clock size={12} color={e.dias>7?"#c0392b":e.dias>4?"#b7770d":"#1e8449"} strokeWidth={2.3}/>
                      <span style={{fontWeight:800,fontSize:14,color:e.dias>7?"#c0392b":e.dias>4?"#b7770d":"#1e8449"}}>{e.dias}d</span>
                    </div>
                  </div>
                ))}
              </div>
            </>
          )
        )}
      </div>
    </div>
  );
}

function AdminPanel({users,setUsers,vendors,setVendors,products,setProducts,stockLog,setStockLog,notifs,setNotifs,activity,setActivity,orders,purchaseOrders,priceLists,setPriceLists,isMobile,sandboxStock,setSandboxStock,updateSandboxStock,promos,setPromos,settings,onToggleExigirPago,onConfirmarComprobante,onRechazarComprobante,onConfirmarEfectivo,onRechazarEfectivo,autoSection,onConsumedAutoSection,prefillProductId,onConsumedPrefill}) {
  const [section, setSection] = useState("home");
  // Deep-link desde otras pantallas (ej: "Ver en Pagos" desde un pedido bloqueado)
  useEffect(() => {
    if(!autoSection) return;
    setSection(autoSection);
    onConsumedAutoSection && onConsumedAutoSection();
  }, [autoSection]);

  const pagosPendientes = orders.filter(o=>o.pagoTipo==="comprobante_pendiente"||o.pagoTipo==="efectivo_pendiente").length;
  const negativosCount = products.filter(p=>p.costPrice>0 && p.costPrice>=p.salePrice).length;

  const SECTIONS = [
    {k:"home",        label:"Administración",   Icon:Home,        grp:"home"},
    {k:"ventas",      label:"Ventas",           Icon:TrendUp,      grp:"negocio"},
    {k:"pagos",       label:"Pagos",            Icon:CreditCard,   grp:"negocio", badge:pagosPendientes, urgentColor:"#b7770d"},
    {k:"ofertas",     label:"Ofertas y Combos", Icon:Gift,         grp:"negocio"},
    {k:"pricelists",  label:"Listas de Precio", Icon:CircleDollarSign, grp:"negocio"},
    {k:"compras",     label:"Análisis Compras", Icon:TrendDown,    grp:"negocio", badge:negativosCount, urgentColor:"#c0392b"},
    {k:"vendors",     label:"Vendedores",       Icon:Users,        grp:"equipo"},
    {k:"users",       label:"Usuarios",         Icon:Lock,         grp:"equipo"},
    {k:"sandbox",     label:"Demo Sandbox",     Icon:Beaker,       grp:"config"},
    {k:"sandboxstock",label:"Stock Sandbox",     Icon:Package,      grp:"config"},
    {k:"activity",    label:"Actividad",        Icon:Pencil,       grp:"config"},
    {k:"excel",       label:"Importar Precios", Icon:BarChart,     grp:"config"},
    {k:"notifcfg",    label:"Notificaciones",   Icon:Bell,         grp:"config"},
  ];
  const GRP_LABELS = {negocio:"📈 Negocio", equipo:"👥 Equipo", config:"⚙️ Configuración"};

  const sandboxOrders = orders.filter(o=>o.isSandbox);
  const realOrders    = orders.filter(o=>!o.isSandbox);
  const [ventasView, setVentasView] = useState("real");

  // Card reutilizable para las dos grillas (entrada mobile + dashboard "home")
  const AdminCard = ({s}) => {
    const urgent = s.urgentColor && s.badge>0;
    return (
      <div onClick={()=>setSection(s.k)} style={{
        background: urgent ? `${s.urgentColor}0a` : "#fff",
        borderRadius:14, padding:"16px 14px", textAlign:"center", cursor:"pointer", position:"relative",
        boxShadow: urgent ? `0 1px 4px #0001, 0 0 0 1.5px ${s.urgentColor}44` : "0 1px 4px #0001",
        border: urgent ? `1.5px solid ${s.urgentColor}33` : "1px solid transparent",
      }}>
        {urgent && (
          <span style={{position:"absolute",top:8,right:8,background:s.urgentColor,color:"#fff",borderRadius:20,fontSize:10,fontWeight:800,padding:"2px 7px",minWidth:18}}>
            {s.badge>999?"999+":s.badge}
          </span>
        )}
        <div style={{width:42,height:42,borderRadius:12,background:urgent?`${s.urgentColor}18`:"#f4f6f9",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 9px"}}>
          <s.Icon size={20} color={urgent?s.urgentColor:"#555"} strokeWidth={2.1}/>
        </div>
        <div style={{fontSize:12,fontWeight:700,color:"#1a1a1a"}}>{s.label}</div>
      </div>
    );
  };
  const AdminGrid = ({cols}) => {
    const grouped = {};
    SECTIONS.filter(s=>s.k!=="home").forEach(s=>{ (grouped[s.grp]=grouped[s.grp]||[]).push(s); });
    return Object.entries(grouped).map(([grp,items])=>(
      <div key={grp} style={{marginBottom:18}}>
        <div style={{fontSize:11.5,fontWeight:700,color:"#999",textTransform:"uppercase",letterSpacing:0.5,marginBottom:8}}>{GRP_LABELS[grp]}</div>
        <div style={{display:"grid",gridTemplateColumns:cols,gap:10}}>
          {items.map(s=><AdminCard key={s.k} s={s}/>)}
        </div>
      </div>
    ));
  };

  // Mobile: show icon grid when no section selected, back button when inside
  if(isMobile && !section) return (
    <div style={{padding:12}}>
      <div style={{fontWeight:800,fontSize:16,marginBottom:12,display:"flex",alignItems:"center",gap:7}}><Star size={16} color="#1a1a1a" strokeWidth={2.2}/> Administración</div>
      <AdminGrid cols="1fr 1fr"/>
    </div>
  );

  return (
    <div>
      {isMobile && section && (
        <div style={{display:"flex",alignItems:"center",gap:10,padding:"12px 12px 0",marginBottom:8}}>
          <button onClick={()=>setSection(null)} style={{padding:"7px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",fontWeight:600,fontSize:12,cursor:"pointer",color:"#555",display:"inline-flex",alignItems:"center",gap:5}}><ArrowLeftIcon size={11} strokeWidth={2.4}/> Volver</button>
          <div style={{fontWeight:800,fontSize:15,display:"flex",alignItems:"center",gap:6}}>
            {(() => { const cur = SECTIONS.find(s=>s.k===section); return cur ? <><cur.Icon size={15} color="#1a1a1a" strokeWidth={2.2}/> {cur.label}</> : null; })()}
          </div>
        </div>
      )}
      {!isMobile && <div style={{background:"#fff",borderRadius:12,padding:4,marginBottom:16,display:"flex",gap:4,boxShadow:"0 1px 4px #0001",flexWrap:"wrap"}}>
        {SECTIONS.map(s=>(
          <button key={s.k} onClick={()=>setSection(s.k)} style={{flex:1,minWidth:120,padding:"10px 16px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:section===s.k?`linear-gradient(135deg,${REDD},${RED})`:"transparent",color:section===s.k?"#fff":"#555",transition:"all .15s",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
            <s.Icon size={14} strokeWidth={2.2}/> {s.label}{s.urgentColor&&s.badge>0?` (${s.badge>999?"999+":s.badge})`:""}
          </button>
        ))}
      </div>}
      {section==="home" && (
        <div>
          <div style={{fontWeight:800,fontSize:16,marginBottom:16,color:"#1a1a1a",display:"flex",alignItems:"center",gap:7}}><Home size={17} color="#1a1a1a" strokeWidth={2.2}/> Central de Administración</div>
          <AdminGrid cols={isMobile?"1fr 1fr":"repeat(auto-fill,minmax(160px,1fr))"}/>
        </div>
      )}
      {section==="ventas"      && (
        <div>
          {sandboxOrders.length>0 && (
            <div style={{display:"flex",gap:8,marginBottom:14}}>
              <button onClick={()=>setVentasView("real")}
                style={{padding:"6px 14px",borderRadius:8,border:"1.5px solid #e5e5e5",cursor:"pointer",fontWeight:700,fontSize:12,
                  background:ventasView==="real"?RED:"#fff",color:ventasView==="real"?"#fff":"#666"}}>
                📊 Datos reales
              </button>
              <button onClick={()=>setVentasView("sandbox")}
                style={{padding:"6px 14px",borderRadius:8,border:"1.5px solid #9b59b6",cursor:"pointer",fontWeight:700,fontSize:12,
                  background:ventasView==="sandbox"?"#9b59b6":"#fff",color:ventasView==="sandbox"?"#fff":"#6c3483"}}>
                🧪 Datos de prueba (sandbox)
              </button>
            </div>
          )}
          <VentasPanel orders={ventasView==="sandbox" ? sandboxOrders : realOrders}/>
        </div>
      )}
      {section==="pagos"       && (
        <PagosPanel orders={orders} onConfirmarComprobante={onConfirmarComprobante} onRechazarComprobante={onRechazarComprobante}
          onConfirmarEfectivo={onConfirmarEfectivo} onRechazarEfectivo={onRechazarEfectivo}
          settings={settings} onToggleExigirPago={onToggleExigirPago}/>
      )}
      {section==="ofertas"     && <OfertasPanel   promos={promos} setPromos={setPromos} products={products} isMobile={isMobile} prefillProductId={prefillProductId} onConsumedPrefill={onConsumedPrefill}/>}
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
          <SandboxStockManager products={products} sandboxStock={sandboxStock} updateSandboxStock={updateSandboxStock}/>
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
      {section==="compras"     && <AnalisisComprasPanel products={products} purchaseOrders={purchaseOrders}/>}
      {section==="notifcfg"    && <NotifConfig    users={users} setUsers={setUsers} notifs={notifs} setNotifs={setNotifs}/>}
    </div>
  );
}

// ── Panel de Ventas ───────────────────────────────────────────────────────────
function VentasPanel({orders}) {
  const vendidas = orders.filter(o=>o.stage==="entregado");
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

// ── Panel de Ofertas y Combos ──────────────────────────────────────────────
function TypeBadge({tipo}) {
  const CFG = {
    combo:     {label:"🎁 COMBO",  bg:"#7b1a1a", color:"#fff",    border:"none"},
    "3x2":     {label:"🏷️ OFERTA", bg:"#fef9e7", color:"#b7770d", border:"1px solid #f3d98a"},
    descuento: {label:"🔻 DESC.",  bg:"#fdecea", color:"#c0392b", border:"1px solid #f5b7b1"},
  };
  const c = CFG[tipo] || CFG.descuento;
  return <span style={{fontSize:10,fontWeight:800,borderRadius:6,padding:"3px 8px",whiteSpace:"nowrap",background:c.bg,color:c.color,border:c.border,display:"inline-block"}}>{c.label}</span>;
}

function nextPromoSku(promos) {
  const used = new Set();
  (promos||[]).forEach(p => {
    const m = /^999-(\d{3})$/.exec(p.id||"");
    if(m) used.add(parseInt(m[1],10));
  });
  let n = 1;
  while(used.has(n)) n++;
  return `999-${String(n).padStart(3,"0")}`;
}

function promoDisplay(promo, products) {
  const d = promo.data || {};
  let titulo = promo.nombre, sub = "";
  if(promo.tipo==="combo") {
    sub = `${promo.id} · ${(d.componentes||[]).length} productos · ${fARS(d.precioFijo||0)}`;
  } else if(promo.tipo==="3x2") {
    const prod = products.find(p=>p.id===d.productoId);
    titulo = prod?.name || promo.nombre;
    sub = `${promo.id} · Comprá ${d.comprar||3}, pagá ${d.pagar||2}`;
  } else if(promo.tipo==="descuento") {
    const prod = products.find(p=>p.id===d.productoId);
    titulo = prod?.name || promo.nombre;
    const finalP = prod ? (d.tipoValor==="%" ? prod.salePrice*(1-(d.valor||0)/100) : Math.max(0,prod.salePrice-(d.valor||0))) : null;
    sub = `${promo.id}${prod?` · ${fARS(prod.salePrice)} → ${fARS(finalP)}`:""}`;
  }
  return {titulo, sub};
}

function PromoEstado({promo, vigente, dias}) {
  let bg="#d5f5e3", color="#1e8449", label="● Vigente";
  if(!promo.activa) { bg="#f0f0f0"; color="#888"; label="⏸ Pausada"; }
  else if(!vigente && dias!==null && dias<0) { bg="#fdecea"; color="#c0392b"; label="Vencida"; }
  else if(!vigente) { bg="#fef9e7"; color="#b7770d"; label="Aún no empieza"; }
  else if(dias!==null && dias<=5) { bg="#fef9e7"; color="#b7770d"; label=`⏳ Vence en ${dias} día${dias===1?"":"s"}`; }
  return <span style={{fontSize:11,fontWeight:700,borderRadius:10,padding:"3px 9px",background:bg,color,whiteSpace:"nowrap"}}>{label}</span>;
}

function OfertasPanel({promos, setPromos, products, isMobile, prefillProductId, onConsumedPrefill}) {
  const [view, setView] = useState("vigentes"); // vigentes | historial | elegir | combo | 3x2 | descuento
  const [editing, setEditing] = useState(null);

  // Deep-link desde la alerta de stock sin rotacion: "Crear oferta con este producto"
  useEffect(() => {
    if(!prefillProductId) return;
    setEditing(null);
    setView("descuento");
    onConsumedPrefill && onConsumedPrefill();
  }, [prefillProductId]);

  const fDate = (iso) => { if(!iso) return ""; const [y,m,d]=iso.split("-"); return `${d}/${m}/${y}`; };
  const diasRestantes = (iso) => { if(!iso) return null; return Math.ceil((new Date(iso)-new Date(todayISO()))/86400000); };

  const vigentes  = promos.filter(isPromoVigente);
  const historial = promos.filter(p=>!isPromoVigente(p));

  const savePromo = async (promoIn) => {
    const { _replacesId, ...promo } = promoIn;
    await db.savePromo(promo);
    if(_replacesId) { try { await db.deletePromo(_replacesId); } catch(e) { console.warn(e); } }
    setPromos(prev => {
      const base = _replacesId ? prev.filter(p=>p.id!==_replacesId) : prev;
      return base.find(p=>p.id===promo.id) ? base.map(p=>p.id===promo.id?promo:p) : [promo,...base];
    });
    setEditing(null); setView("vigentes");
  };
  const togglePausa = async (promo) => {
    const updated = {...promo, activa:!promo.activa};
    await db.savePromo(updated);
    setPromos(prev=>prev.map(p=>p.id===promo.id?updated:p));
  };
  const delPromo = async (promo) => {
    if(!await confirmDialog("¿Eliminar oferta?",`Vas a eliminar "${promo.nombre}". Esta acción no se puede deshacer.`,true)) return;
    await db.deletePromo(promo.id);
    setPromos(prev => prev.filter(p=>p.id!==promo.id));
  };
  const editPromo = (promo) => { setEditing(promo); setView(promo.tipo); };
  const cancelForm = () => { setEditing(null); setView("vigentes"); };

  if(view==="combo")     return <ComboForm products={products} editing={editing} nextSku={nextPromoSku(promos)} onSave={savePromo} onCancel={cancelForm}/>;
  if(view==="3x2")       return <TresPorDosForm products={products} editing={editing} nextSku={nextPromoSku(promos)} onSave={savePromo} onCancel={cancelForm}/>;
  if(view==="descuento") return <DescuentoForm products={products} editing={editing} prefillProductId={prefillProductId} onSave={savePromo} onCancel={cancelForm}/>;

  if(view==="elegir") {
    const TIPOS = [
      {k:"combo",     Icon:Gift,    title:"Combo",            desc:"Varios productos distintos agrupados a un precio fijo"},
      {k:"3x2",       Icon:Tag,     title:"Oferta 3×N",        desc:"Comprando X unidades de un producto, paga menos"},
      {k:"descuento", Icon:TrendDown, title:"Descuento simple", desc:"Un producto puntual con % o $ de descuento"},
    ];
    return (
      <div>
        <div style={{textAlign:"center",marginBottom:20}}>
          <div style={{fontWeight:800,fontSize:16,marginBottom:4}}>¿Qué querés crear?</div>
          <div style={{fontSize:13,color:"#888"}}>Elegí el tipo de promoción</div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr 1fr",gap:14,maxWidth:760,margin:"0 auto"}}>
          {TIPOS.map(t=>(
            <div key={t.k} onClick={()=>setView(t.k)}
              style={{background:"#fff",borderRadius:14,padding:"24px 18px",textAlign:"center",boxShadow:"0 1px 6px #0001",cursor:"pointer",border:"2px solid transparent",transition:"border .15s"}}
              onMouseEnter={e=>e.currentTarget.style.border=`2px solid ${RED}`}
              onMouseLeave={e=>e.currentTarget.style.border="2px solid transparent"}>
              <div style={{marginBottom:10,display:"flex",justifyContent:"center"}}><t.Icon size={30} color="#555" strokeWidth={1.8}/></div>
              <div style={{fontWeight:800,fontSize:14,marginBottom:6}}>{t.title}</div>
              <div style={{fontSize:12,color:"#888",lineHeight:1.4}}>{t.desc}</div>
            </div>
          ))}
        </div>
        <div style={{textAlign:"center",marginTop:18}}>
          <button onClick={()=>setView("vigentes")} style={{padding:"8px 16px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontWeight:600,fontSize:12,color:"#666"}}>← Cancelar</button>
        </div>
      </div>
    );
  }

  // vigentes | historial
  const list = view==="vigentes" ? vigentes : historial;
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10,marginBottom:16}}>
        <div style={{display:"flex",gap:6}}>
          <button onClick={()=>setView("vigentes")} style={{padding:"8px 16px",borderRadius:8,border:"1.5px solid",cursor:"pointer",fontWeight:700,fontSize:12,borderColor:view==="vigentes"?RED:"#e5e5e5",background:view==="vigentes"?"#fdecea":"#fff",color:view==="vigentes"?RED:"#666"}}>Vigentes ({vigentes.length})</button>
          <button onClick={()=>setView("historial")} style={{padding:"8px 16px",borderRadius:8,border:"1.5px solid",cursor:"pointer",fontWeight:700,fontSize:12,borderColor:view==="historial"?RED:"#e5e5e5",background:view==="historial"?"#fdecea":"#fff",color:view==="historial"?RED:"#666"}}>Historial ({historial.length})</button>
        </div>
        <button onClick={()=>{setEditing(null);setView("elegir");}} style={{padding:"9px 16px",borderRadius:8,border:"none",cursor:"pointer",fontWeight:800,fontSize:12,background:`linear-gradient(135deg,${REDD},${RED})`,color:"#fff"}}>+ Nueva</button>
      </div>

      {list.length===0
        ? <div style={{textAlign:"center",color:"#aaa",padding:"40px 0"}}>{view==="vigentes"?"No hay ofertas ni combos vigentes":"Todavía no hay historial"}</div>
        : isMobile
          ? <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {list.map(p=>{
                const {titulo,sub} = promoDisplay(p,products);
                const vig = !p.vigenciaDesde && !p.vigenciaHasta ? "Sin vencimiento" : `${p.vigenciaDesde?fDate(p.vigenciaDesde):"…"} al ${p.vigenciaHasta?fDate(p.vigenciaHasta):"…"}`;
                return (
                  <div key={p.id} style={{background:"#fff",borderRadius:10,padding:14,boxShadow:"0 1px 4px #0001"}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                      <TypeBadge tipo={p.tipo}/>
                      <div>
                        <button onClick={()=>editPromo(p)} style={{background:"none",border:"none",cursor:"pointer",padding:"2px 6px",display:"flex",alignItems:"center"}}><Pencil size={13} strokeWidth={2.3} color="#888"/></button>
                        <button onClick={()=>togglePausa(p)} style={{background:"none",border:"none",cursor:"pointer",fontSize:14,padding:"2px 6px"}}>{p.activa?"⏸️":"▶️"}</button>
                        <button onClick={()=>delPromo(p)} style={{background:"none",border:"none",cursor:"pointer",padding:"2px 6px",display:"flex",alignItems:"center"}}><Trash size={13} strokeWidth={2.3} color="#888"/></button>
                      </div>
                    </div>
                    <div style={{fontWeight:700,fontSize:13}}>{titulo}</div>
                    <div style={{fontSize:11,color:"#888",marginTop:2}}>{sub}</div>
                    <div style={{marginTop:6,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <span style={{fontSize:11,color:"#aaa"}}>{vig}</span>
                      <PromoEstado promo={p} vigente={isPromoVigente(p)} dias={p.vigenciaHasta?diasRestantes(p.vigenciaHasta):null}/>
                    </div>
                  </div>
                );
              })}
            </div>
          : <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 6px #0001",overflow:"hidden"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead><tr style={{background:"#f9f9f9"}}>
                  <th style={{textAlign:"left",padding:"10px 14px",fontSize:11,color:"#888",textTransform:"uppercase",letterSpacing:.5}}>Tipo</th>
                  <th style={{textAlign:"left",padding:"10px 14px",fontSize:11,color:"#888",textTransform:"uppercase",letterSpacing:.5}}>Nombre / SKU</th>
                  <th style={{textAlign:"left",padding:"10px 14px",fontSize:11,color:"#888",textTransform:"uppercase",letterSpacing:.5}}>Vigencia</th>
                  <th style={{textAlign:"left",padding:"10px 14px",fontSize:11,color:"#888",textTransform:"uppercase",letterSpacing:.5}}>Estado</th>
                  <th></th>
                </tr></thead>
                <tbody>
                  {list.map(p=>{
                    const {titulo,sub} = promoDisplay(p,products);
                    const vig = !p.vigenciaDesde && !p.vigenciaHasta ? "Sin vencimiento" : `${p.vigenciaDesde?fDate(p.vigenciaDesde):"…"} al ${p.vigenciaHasta?fDate(p.vigenciaHasta):"…"}`;
                    return (
                      <tr key={p.id} style={{borderBottom:"1px solid #f5f5f5"}}>
                        <td style={{padding:"11px 14px"}}><TypeBadge tipo={p.tipo}/></td>
                        <td style={{padding:"11px 14px"}}>
                          <div style={{fontWeight:700,color:"#1a1a1a",fontSize:13}}>{titulo}</div>
                          <div style={{fontSize:11,color:"#888",marginTop:2}}>{sub}</div>
                        </td>
                        <td style={{padding:"11px 14px",fontSize:12,color:"#555"}}>{vig}</td>
                        <td style={{padding:"11px 14px"}}><PromoEstado promo={p} vigente={isPromoVigente(p)} dias={p.vigenciaHasta?diasRestantes(p.vigenciaHasta):null}/></td>
                        <td style={{padding:"11px 14px",textAlign:"right",whiteSpace:"nowrap"}}>
                          <button onClick={()=>editPromo(p)} title="Editar" style={{background:"none",border:"none",cursor:"pointer",padding:"4px 6px",color:"#888",display:"flex",alignItems:"center"}}><Pencil size={13} strokeWidth={2.3}/></button>
                          <button onClick={()=>togglePausa(p)} title={p.activa?"Pausar":"Reactivar"} style={{background:"none",border:"none",cursor:"pointer",fontSize:14,padding:"4px 6px",color:"#888"}}>{p.activa?"⏸️":"▶️"}</button>
                          <button onClick={()=>delPromo(p)} title="Eliminar" style={{background:"none",border:"none",cursor:"pointer",padding:"4px 6px",color:"#888",display:"flex",alignItems:"center"}}><Trash size={13} strokeWidth={2.3}/></button>
                        </td>
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

function ComboForm({products, editing, nextSku, onSave, onCancel}) {
  const isEdit = !!editing;
  const [sku] = useState(editing?.id || nextSku);
  const [nombre, setNombre] = useState(editing?.nombre || "");
  const [search, setSearch] = useState("");
  const [componentes, setComponentes] = useState(editing?.data?.componentes || []);
  const [precioFijo, setPrecioFijo] = useState(editing?.data?.precioFijo!=null ? String(editing.data.precioFijo) : "");
  const [vigDesde, setVigDesde] = useState(editing?.vigenciaDesde || "");
  const [vigHasta, setVigHasta] = useState(editing?.vigenciaHasta || "");
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);

  const found = useMemo(()=>{
    if(!search.trim()) return [];
    const q = norm(search);
    return products.filter(p=>!componentes.find(c=>c.pid===p.id) && (norm(p.name).includes(q)||normSKU(p.id).includes(normSKU(search)))).slice(0,8);
  },[search,products,componentes]);

  const addComp  = (p) => { setComponentes(c=>[...c,{pid:p.id,qty:1}]); setSearch(""); };
  const setQty   = (pid,qty) => setComponentes(c=>c.map(i=>i.pid===pid?{...i,qty:Math.max(1,qty)}:i));
  const delComp  = (pid) => setComponentes(c=>c.filter(i=>i.pid!==pid));

  const precioNormal = componentes.reduce((s,c)=>{ const p=products.find(pr=>pr.id===c.pid); return s+(p?p.salePrice*c.qty:0); },0);
  const ahorro = precioNormal - (parseFloat(precioFijo)||0);
  const ahorroPct = precioNormal>0 ? (ahorro/precioNormal*100) : 0;

  const submit = async () => {
    if(!sku.trim())               { setError("El SKU es obligatorio"); return; }
    if(!nombre.trim())            { setError("El nombre es obligatorio"); return; }
    if(componentes.length===0)    { setError("Agregá al menos un producto"); return; }
    if(!precioFijo || +precioFijo<=0) { setError("Ingresá el precio fijo del combo"); return; }
    setError(""); setSaving(true);
    try {
      await onSave({
        id:sku.trim(), tipo:"combo", nombre:nombre.trim(), activa:editing?.activa!==false,
        vigenciaDesde:vigDesde, vigenciaHasta:vigHasta,
        data:{precioFijo:+precioFijo, componentes}, createdAt:editing?.createdAt,
      });
    } finally { setSaving(false); }
  };

  return (
    <div style={{maxWidth:680}}>
      <div style={{fontWeight:800,fontSize:15,marginBottom:14,display:"flex",alignItems:"center",gap:7}}><Gift size={15} strokeWidth={2.3}/>{isEdit?"Editar":"Nuevo"} Combo</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <Field label="SKU del combo (automático)"><input value={sku} disabled style={{...inputStyle,background:"#f5f5f5",color:"#888",fontWeight:700}}/></Field>
        <Field label="Nombre del combo *"><input value={nombre} onChange={e=>setNombre(e.target.value)} style={inputStyle}/></Field>
      </div>

      <Field label="Productos incluidos">
        <div style={{position:"relative",marginBottom:8}}><Search size={13} color="#aaa" strokeWidth={2.3} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)"}}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar producto para agregar al combo..." style={{...inputStyle,paddingLeft:32}}/></div>
        {found.length>0 && <div style={{border:"1.5px solid #f0f0f0",borderRadius:8,marginBottom:8,overflow:"hidden"}}>
          {found.map(p=><div key={p.id} onClick={()=>addComp(p)} style={{padding:"8px 12px",cursor:"pointer",fontSize:12,borderBottom:"1px solid #f5f5f5"}}>{p.name} <span style={{color:"#aaa"}}>· {fARS(p.salePrice)}</span></div>)}
        </div>}
        {componentes.map(c=>{
          const p = products.find(pr=>pr.id===c.pid);
          return (
            <div key={c.pid} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 10px",background:"#f9f9f9",borderRadius:8,marginBottom:6,border:"1.5px solid #f0f0f0"}}>
              <input type="number" min="1" value={c.qty} onChange={e=>setQty(c.pid,+e.target.value||1)} style={{width:46,padding:"4px 6px",borderRadius:6,border:"1.5px solid #e5e5e5",fontWeight:800,fontSize:12,textAlign:"center"}}/>
              <span style={{flex:1,fontSize:13,fontWeight:600}}>{p?.name||c.pid}</span>
              <span style={{fontSize:12,color:"#888"}}>{fARS(p?.salePrice||0)} c/u</span>
              <span onClick={()=>delComp(c.pid)} style={{color:RED,cursor:"pointer",padding:"0 4px"}}>✕</span>
            </div>
          );
        })}
      </Field>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,alignItems:"end"}}>
        <Field label="Precio normal (suma)"><div style={{padding:"8px 12px",fontSize:14,color:"#888",textDecoration:"line-through"}}>{fARS(precioNormal)}</div></Field>
        <Field label="Precio fijo del combo *"><input type="number" value={precioFijo} onChange={e=>setPrecioFijo(e.target.value)} style={{...inputStyle,fontWeight:800,color:RED,fontSize:15}}/></Field>
      </div>

      {precioFijo && precioNormal>0 && (
        <div style={{background:"#eafaf1",border:"1.5px solid #a9dfbf",borderRadius:8,padding:"12px 14px",margin:"14px 0",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{display:"inline-flex",alignItems:"center",gap:5,fontSize:13,fontWeight:700,color:"#1e8449"}}><Banknote size={12} strokeWidth={2.4}/>Ahorro para el cliente</span>
          <span style={{fontSize:15,fontWeight:800,color:ahorro>=0?"#1e8449":RED}}>{fARS(ahorro)} ({ahorroPct.toFixed(1)}%)</span>
        </div>
      )}

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginTop:precioFijo&&precioNormal>0?0:14}}>
        <Field label="Vigencia desde (opcional)"><input type="date" value={vigDesde} onChange={e=>setVigDesde(e.target.value)} style={inputStyle}/></Field>
        <Field label="Vigencia hasta (opcional)"><input type="date" value={vigHasta} onChange={e=>setVigHasta(e.target.value)} style={inputStyle}/></Field>
      </div>

      {error && <div style={{color:RED,fontSize:12,fontWeight:600,marginBottom:10}}>{error}</div>}
      <div style={{display:"flex",gap:10,marginTop:6}}>
        <button onClick={onCancel} style={{flex:1,padding:"11px",borderRadius:10,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontWeight:600,fontSize:13,cursor:"pointer"}}>Cancelar</button>
        <button onClick={submit} disabled={saving} style={{flex:2,padding:"11px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:`linear-gradient(135deg,${REDD},${RED})`,color:"#fff",opacity:saving?.7:1,display:"flex",alignItems:"center",justifyContent:"center",gap:7}}>{saving?"Guardando...":<><CheckCircle size={13} strokeWidth={2.4}/> Guardar Combo</>}</button>
      </div>
    </div>
  );
}

function TresPorDosForm({products, editing, nextSku, onSave, onCancel}) {
  const isEdit = !!editing;
  const [sku] = useState(editing?.id || nextSku);
  const [search, setSearch] = useState("");
  const [producto, setProducto] = useState(()=> editing?.data?.productoId ? (products.find(p=>p.id===editing.data.productoId)||null) : null);
  const [tipoSel, setTipoSel] = useState(()=>{
    const c=editing?.data?.comprar, p=editing?.data?.pagar;
    if(c===3&&p===2) return "3x2"; if(c===2&&p===1) return "2x1"; return "3x2";
  });
  const [comprar, setComprar] = useState(editing?.data?.comprar ?? 3);
  const [pagar, setPagar]     = useState(editing?.data?.pagar ?? 2);
  const [vigDesde, setVigDesde] = useState(editing?.vigenciaDesde || "");
  const [vigHasta, setVigHasta] = useState(editing?.vigenciaHasta || "");
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);

  const found = useMemo(()=>{
    if(!search.trim()) return [];
    const q = norm(search);
    return products.filter(p=>norm(p.name).includes(q)||normSKU(p.id).includes(normSKU(search))).slice(0,8);
  },[search,products]);

  const selectTipo = (k) => { setTipoSel(k); if(k==="3x2"){setComprar(3);setPagar(2);} else if(k==="2x1"){setComprar(2);setPagar(1);} };
  const pct = comprar>0 ? (1 - pagar/comprar)*100 : 0;

  const submit = async () => {
    if(!sku.trim())  { setError("El SKU es obligatorio"); return; }
    if(!producto)    { setError("Elegí el producto al que aplica"); return; }
    if(!comprar || !pagar || +pagar>=+comprar) { setError("Revisá los números: tiene que pagar menos unidades de las que compra"); return; }
    setError(""); setSaving(true);
    try {
      await onSave({
        id:sku.trim(), tipo:"3x2", nombre:producto.name, activa:editing?.activa!==false,
        vigenciaDesde:vigDesde, vigenciaHasta:vigHasta,
        data:{productoId:producto.id, comprar:+comprar, pagar:+pagar}, createdAt:editing?.createdAt,
      });
    } finally { setSaving(false); }
  };

  return (
    <div style={{maxWidth:680}}>
      <div style={{fontWeight:800,fontSize:15,marginBottom:14,display:"flex",alignItems:"center",gap:7}}><Tag size={15} strokeWidth={2.3}/>{isEdit?"Editar":"Nueva"} Oferta</div>
      <Field label="SKU de la oferta (automático)"><input value={sku} disabled style={{...inputStyle,maxWidth:300,background:"#f5f5f5",color:"#888",fontWeight:700}}/></Field>

      <Field label="Producto al que aplica *">
        {producto
          ? <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",background:"#f9f9f9",borderRadius:8,border:"1.5px solid #f0f0f0"}}>
              <span style={{flex:1,fontWeight:700,fontSize:13}}>{producto.name}</span>
              <span style={{fontSize:12,color:"#888"}}>{fARS(producto.salePrice)} c/u</span>
              <span onClick={()=>setProducto(null)} style={{color:RED,cursor:"pointer",display:"inline-flex"}}><XIcon size={13} strokeWidth={2.6}/></span>
            </div>
          : <>
              <div style={{position:"relative"}}><Search size={13} color="#aaa" strokeWidth={2.3} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)"}}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar producto..." style={{...inputStyle,paddingLeft:32}}/></div>
              {found.length>0 && <div style={{border:"1.5px solid #f0f0f0",borderRadius:8,marginTop:6,overflow:"hidden"}}>
                {found.map(p=><div key={p.id} onClick={()=>{setProducto(p);setSearch("");}} style={{padding:"8px 12px",cursor:"pointer",fontSize:12,borderBottom:"1px solid #f5f5f5"}}>{p.name} <span style={{color:"#aaa"}}>· {fARS(p.salePrice)}</span></div>)}
              </div>}
            </>
        }
      </Field>

      <Field label="Tipo de oferta">
        <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
          {[["3x2","3 × 2"],["2x1","2 × 1"],["custom","Personalizado"]].map(([k,label])=>(
            <div key={k} onClick={()=>selectTipo(k)} style={{padding:"8px 16px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:13,fontWeight:700,borderColor:tipoSel===k?"#b7770d":"#e5e5e5",background:tipoSel===k?"#fef9e7":"#fff",color:tipoSel===k?"#b7770d":"#666"}}>{label}</div>
          ))}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10,background:"#f9f9f9",borderRadius:10,padding:14,flexWrap:"wrap"}}>
          <span style={{fontSize:13,fontWeight:600,color:"#555"}}>Comprando</span>
          <input type="number" min="2" value={comprar} onChange={e=>{setComprar(+e.target.value||0);setTipoSel("custom");}} style={{width:60,padding:8,textAlign:"center",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:15,fontWeight:800}}/>
          <span style={{fontSize:13,fontWeight:600,color:"#555"}}>unidades, el cliente paga</span>
          <input type="number" min="1" value={pagar} onChange={e=>{setPagar(+e.target.value||0);setTipoSel("custom");}} style={{width:60,padding:8,textAlign:"center",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:15,fontWeight:800}}/>
          <span style={{fontSize:13,fontWeight:600,color:"#555"}}>unidades</span>
        </div>
      </Field>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <Field label="Vigencia desde (opcional)"><input type="date" value={vigDesde} onChange={e=>setVigDesde(e.target.value)} style={inputStyle}/></Field>
        <Field label="Vigencia hasta (opcional)"><input type="date" value={vigHasta} onChange={e=>setVigHasta(e.target.value)} style={inputStyle}/></Field>
      </div>

      {comprar>0 && pagar>0 && pagar<comprar && (
        <div style={{background:"#fef9e7",border:"1.5px solid #f8e9b0",borderRadius:8,padding:"12px 14px",margin:"14px 0",fontSize:13,color:"#7d6608",fontWeight:600}}>
          <Lightbulb size={12} strokeWidth={2.3} style={{display:"inline",verticalAlign:"-2px",marginRight:4}}/>Equivale a un <b>{pct.toFixed(1)}% de descuento</b> sobre la compra de este producto, aplicado automáticamente a partir de {comprar} unidades ({comprar*2}, {comprar*3}...). No afecta compras de menos de {comprar} unidades.
        </div>
      )}

      {error && <div style={{color:RED,fontSize:12,fontWeight:600,marginBottom:10}}>{error}</div>}
      <div style={{display:"flex",gap:10,marginTop:6}}>
        <button onClick={onCancel} style={{flex:1,padding:"11px",borderRadius:10,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontWeight:600,fontSize:13,cursor:"pointer"}}>Cancelar</button>
        <button onClick={submit} disabled={saving} style={{flex:2,padding:"11px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:`linear-gradient(135deg,${REDD},${RED})`,color:"#fff",opacity:saving?.7:1,display:"flex",alignItems:"center",justifyContent:"center",gap:7}}>{saving?"Guardando...":<><CheckCircle size={13} strokeWidth={2.4}/> Guardar Oferta</>}</button>
      </div>
    </div>
  );
}

function DescuentoForm({products, editing, prefillProductId, onSave, onCancel}) {
  const isEdit = !!editing;
  const [search, setSearch] = useState("");
  const [producto, setProducto] = useState(()=> editing?.data?.productoId ? (products.find(p=>p.id===editing.data.productoId)||null) : (prefillProductId ? (products.find(p=>p.id===prefillProductId)||null) : null));
  const [tipoValor, setTipoValor] = useState(editing?.data?.tipoValor || "%");
  const [valor, setValor] = useState(editing?.data?.valor!=null ? String(editing.data.valor) : "");
  const [vigDesde, setVigDesde] = useState(editing?.vigenciaDesde || "");
  const [vigHasta, setVigHasta] = useState(editing?.vigenciaHasta || "");
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);

  const found = useMemo(()=>{
    if(!search.trim()) return [];
    const q = norm(search);
    return products.filter(p=>norm(p.name).includes(q)||normSKU(p.id).includes(normSKU(search))).slice(0,8);
  },[search,products]);

  const precioFinal = producto ? Math.max(0, tipoValor==="%" ? producto.salePrice*(1-(parseFloat(valor)||0)/100) : producto.salePrice-(parseFloat(valor)||0)) : null;

  const submit = async () => {
    if(!producto)       { setError("Elegí el producto al que aplica"); return; }
    if(!valor || +valor<=0) { setError("Ingresá el valor del descuento"); return; }
    setError(""); setSaving(true);
    try {
      await onSave({
        id:producto.id, tipo:"descuento", nombre:producto.name, activa:editing?.activa!==false,
        vigenciaDesde:vigDesde, vigenciaHasta:vigHasta,
        data:{productoId:producto.id, tipoValor, valor:+valor}, createdAt:editing?.createdAt,
        // Si veníamos de un promo editado con un id viejo (ej. autogenerado de una versión anterior), lo limpiamos
        _replacesId: (editing && editing.id!==producto.id) ? editing.id : undefined,
      });
    } finally { setSaving(false); }
  };

  return (
    <div style={{maxWidth:680}}>
      <div style={{fontWeight:800,fontSize:15,marginBottom:14,display:"flex",alignItems:"center",gap:7}}><TrendDown size={15} strokeWidth={2.3}/>{isEdit?"Editar":"Nuevo"} Descuento</div>

      <Field label="Producto al que aplica *">
        {producto
          ? <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",background:"#f9f9f9",borderRadius:8,border:"1.5px solid #f0f0f0"}}>
              <span style={{flex:1,fontWeight:700,fontSize:13}}>{producto.name}</span>
              <span style={{fontSize:12,color:"#888"}}>{fARS(producto.salePrice)} c/u</span>
              <span onClick={()=>setProducto(null)} style={{color:RED,cursor:"pointer",display:"inline-flex"}}><XIcon size={13} strokeWidth={2.6}/></span>
            </div>
          : <>
              <div style={{position:"relative"}}><Search size={13} color="#aaa" strokeWidth={2.3} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)"}}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar producto..." style={{...inputStyle,paddingLeft:32}}/></div>
              {found.length>0 && <div style={{border:"1.5px solid #f0f0f0",borderRadius:8,marginTop:6,overflow:"hidden"}}>
                {found.map(p=><div key={p.id} onClick={()=>{setProducto(p);setSearch("");}} style={{padding:"8px 12px",cursor:"pointer",fontSize:12,borderBottom:"1px solid #f5f5f5"}}>{p.name} <span style={{color:"#aaa"}}>· {fARS(p.salePrice)}</span></div>)}
              </div>}
            </>
        }
      </Field>

      <Field label="Tipo de descuento">
        <div style={{display:"flex",gap:8,marginBottom:14}}>
          <div onClick={()=>setTipoValor("%")} style={{padding:"8px 16px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:13,fontWeight:700,borderColor:tipoValor==="%"?RED:"#e5e5e5",background:tipoValor==="%"?"#fdecea":"#fff",color:tipoValor==="%"?RED:"#666"}}>% Porcentaje</div>
          <div onClick={()=>setTipoValor("$")} style={{padding:"8px 16px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:13,fontWeight:700,borderColor:tipoValor==="$"?RED:"#e5e5e5",background:tipoValor==="$"?"#fdecea":"#fff",color:tipoValor==="$"?RED:"#666"}}>$ Monto fijo</div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <input type="number" value={valor} onChange={e=>setValor(e.target.value)} style={{width:100,padding:"10px 12px",borderRadius:8,border:`1.5px solid ${RED}`,fontSize:18,fontWeight:800,color:RED,textAlign:"center",outline:"none"}}/>
          <span style={{fontSize:14,color:"#555",fontWeight:600}}>{tipoValor==="%" ? "% de descuento sobre el precio de venta" : "$ de descuento sobre el precio de venta"}</span>
        </div>
      </Field>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <Field label="Vigencia desde"><input type="date" value={vigDesde} onChange={e=>setVigDesde(e.target.value)} style={inputStyle}/></Field>
        <Field label="Vigencia hasta"><input type="date" value={vigHasta} onChange={e=>setVigHasta(e.target.value)} style={inputStyle}/></Field>
      </div>

      {producto && valor && (
        <div style={{background:"#eafaf1",border:"1.5px solid #a9dfbf",borderRadius:8,padding:"12px 14px",margin:"14px 0",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{fontSize:13,fontWeight:700,color:"#1e8449"}}>Precio con descuento</span>
          <span style={{fontSize:15,fontWeight:800}}><span style={{color:"#aaa",textDecoration:"line-through",fontSize:13,marginRight:8}}>{fARS(producto.salePrice)}</span><span style={{color:"#1e8449"}}>{fARS(precioFinal)}</span></span>
        </div>
      )}

      {error && <div style={{color:RED,fontSize:12,fontWeight:600,marginBottom:10}}>{error}</div>}
      <div style={{display:"flex",gap:10,marginTop:6}}>
        <button onClick={onCancel} style={{flex:1,padding:"11px",borderRadius:10,border:"1.5px solid #e5e5e5",background:"#fff",color:"#666",fontWeight:600,fontSize:13,cursor:"pointer"}}>Cancelar</button>
        <button onClick={submit} disabled={saving} style={{flex:2,padding:"11px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:`linear-gradient(135deg,${REDD},${RED})`,color:"#fff",opacity:saving?.7:1,display:"flex",alignItems:"center",justifyContent:"center",gap:7}}>{saving?"Guardando...":<><CheckCircle size={13} strokeWidth={2.4}/> Guardar Descuento</>}</button>
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
    "Nuevo pedido":                {color:"#1a5276", bg:"#d6eaf8",  Icon:ShoppingCart},
    "Cambio estado: Confirmado":   {color:"#1a5276", bg:"#d6eaf8",  Icon:CheckCircle},
    "Cambio estado: En Armado":    {color:"#6c3483", bg:"#e8daef",  Icon:Package},
    "Cambio estado: Entregado":    {color:"#1e8449", bg:"#d5f5e3",  Icon:PartyPopper},
    "Cambio estado: Reserva":      {color:"#c0392b", bg:"#fdecea",  Icon:Clock},
    "Nueva cotización":            {color:"#6c3483", bg:"#e8daef",  Icon:FileText},
    "Cotización convertida a reserva":{color:"#e67e22",bg:"#fef9e7",Icon:RefreshCw},
    "Pedido eliminado":            {color:"#c0392b", bg:"#fdecea",  Icon:Trash},
    "Precio/stock editado":        {color:"#e67e22", bg:"#fef9e7",  Icon:Pencil},
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
          <div style={{fontSize:11,color:"#888",fontWeight:600,display:"flex",alignItems:"center",gap:4}}><Pencil size={10} strokeWidth={2.4}/>Total acciones</div>
          <div style={{fontSize:24,fontWeight:900,color:RED}}>{activity.length}</div>
        </div>
        {byUser.slice(0,3).map(([u,n])=>(
          <div key={u} style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 4px #0001",borderLeft:"4px solid #1a5276"}}>
            <div style={{fontSize:11,color:"#888",fontWeight:600,display:"flex",alignItems:"center",gap:4}}><Users size={10} strokeWidth={2.4}/>{u}</div>
            <div style={{fontSize:24,fontWeight:900,color:"#1a5276"}}>{n}</div>
          </div>
        ))}
      </div>

      {/* Filtros */}
      <div style={{background:"#fff",borderRadius:12,padding:12,marginBottom:12,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center",boxShadow:"0 1px 4px #0001"}}>
        <div style={{position:"relative",flex:1,minWidth:200}}>
          <Search size={13} color="#aaa" strokeWidth={2.3} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)"}}/>
          <input value={search} onChange={e=>setSearch(e.target.value)}
            placeholder="Buscar por usuario, acción o detalle..."
            style={{width:"100%",padding:"7px 12px 7px 32px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",boxSizing:"border-box"}}/>
        </div>
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
            : <button onClick={()=>setConfirmClear(true)} style={{padding:"6px 12px",borderRadius:8,border:"1.5px solid #fcc",background:"#fff",color:RED,cursor:"pointer",fontSize:11,fontWeight:600,whiteSpace:"nowrap",display:"flex",alignItems:"center",gap:5}}><Trash size={11} strokeWidth={2.4}/>Limpiar</button>
        )}
      </div>

      {/* Tabla */}
      {filtered.length===0
        ? <div style={{textAlign:"center",padding:50,color:"#aaa",background:"#fff",borderRadius:12}}>
            <div style={{marginBottom:8,display:"flex",justifyContent:"center"}}><Pencil size={34} color="#ddd" strokeWidth={1.8}/></div>
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
                  const cfg = TIPOS[a.accion]||{color:"#666",bg:"#f5f5f5",Icon:Info};
                  return (
                    <tr key={a.id||i} style={{borderTop:"1px solid #f5f5f5"}}>
                      <td style={{padding:"9px 12px",fontSize:11,color:"#888",whiteSpace:"nowrap"}}>{a.fecha}</td>
                      <td style={{padding:"9px 12px",whiteSpace:"nowrap"}}>
                        <span style={{fontWeight:700,fontSize:12}}>{a.usuario}</span>
                        <span style={{fontSize:10,color:a.rol==="admin"?RED:"#1a5276",marginLeft:4,fontWeight:600}}>({a.rol})</span>
                      </td>
                      <td style={{padding:"9px 12px"}}>
                        <span style={{display:"inline-flex",alignItems:"center",gap:5,background:cfg.bg,color:cfg.color,borderRadius:8,padding:"3px 9px",fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>
                          <cfg.Icon size={11} strokeWidth={2.4}/> {a.accion}
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
    if(!form.name.trim()) { toast.error("Ingresá un nombre"); return; }
    const disc = parseFloat(form.discount)||0;
    if(disc < 0 || disc >= 100) { toast.error("El descuento debe ser entre 0 y 99"); return; }
    setSaving(true);
    const pl = {id: editing || genId(), name: form.name.trim(), discount: disc};
    setPriceLists(list => editing ? list.map(x=>x.id===editing?pl:x) : [...list, pl]);
    await db.savePriceList(pl);
    setForm({name:"", discount:""}); setEditing(null); setSaving(false);
  };

  const del = async (id) => {
    if(id==="default") { toast.error("No se puede eliminar la lista Normal"); return; }
    if(!await confirmDialog("¿Eliminar esta lista?","Esta acción no se puede deshacer.",true)) return;
    setPriceLists(list=>list.filter(x=>x.id!==id));
    await db.deletePriceList(id);
  };

  const startEdit = (pl) => { setEditing(pl.id); setForm({name:pl.name, discount:pl.discount}); };
  const cancel = () => { setEditing(null); setForm({name:"", discount:""}); };

  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,alignItems:"start"}}>
      {/* Form */}
      <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:15,marginBottom:16,display:"flex",alignItems:"center",gap:7}}>
          {editing ? <><Pencil size={14} strokeWidth={2.4}/> Editar lista</> : <><Plus size={15} strokeWidth={2.4}/> Nueva lista de precios</>}
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
                <button onClick={()=>startEdit(pl)} style={{padding:"5px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12,display:"flex",alignItems:"center"}}><Pencil size={12} strokeWidth={2.3}/></button>
                <button onClick={()=>del(pl.id)} style={{padding:"5px 10px",borderRadius:6,border:"1.5px solid #fcc",background:"#fff",cursor:"pointer",fontSize:12,color:RED,display:"flex",alignItems:"center"}}><Trash size={12} strokeWidth={2.3}/></button>
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
    if(vendors.includes(n)){toast.error("Ya existe ese vendedor");return;}
    setLoading(true);
    try {
      await db.addVendor(n);
      setVendors(v=>[...v,n]);
      setNewName("");
    } catch(e) {
      toast.error("Error al guardar el vendedor: " + e.message);
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
      toast.error("Error al eliminar el vendedor: " + e.message);
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
      toast.error("Error al editar el vendedor: " + e.message);
    }
  };

  return (
    <div style={{background:"#fff",borderRadius:12,padding:24,boxShadow:"0 1px 4px #0001",maxWidth:520}}>
      <div style={{fontWeight:800,fontSize:16,marginBottom:16,display:"flex",alignItems:"center",gap:7}}><Users size={16} strokeWidth={2.2}/>Gestión de Vendedores</div>
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
            <div style={{width:30,height:30,borderRadius:9,background:"#f4f6f9",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><Users size={14} color="#888" strokeWidth={2.1}/></div>
            {editing===v
              ? <>
                  <input value={editVal} onChange={e=>setEditVal(e.target.value)}
                    onKeyDown={e=>{if(e.key==="Enter")saveEdit(v);if(e.key==="Escape")setEditing(null);}}
                    style={{...inputStyle,flex:1,fontSize:13}} autoFocus/>
                  <button onClick={()=>saveEdit(v)} style={{padding:"5px 12px",borderRadius:7,border:"none",background:"#1e8449",color:"#fff",fontWeight:700,cursor:"pointer",fontSize:12,display:"flex",alignItems:"center"}}><CheckCircle size={13} strokeWidth={2.4}/></button>
                  <button onClick={()=>setEditing(null)} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12,display:"flex",alignItems:"center"}}><XIcon size={12} strokeWidth={2.4}/></button>
                </>
              : <>
                  <span style={{flex:1,fontWeight:600,fontSize:14}}>{v}</span>
                  <button onClick={()=>{setEditing(v);setEditVal(v);}} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12,display:"flex",alignItems:"center",gap:5}}><Pencil size={11} strokeWidth={2.4}/>Editar</button>
                  {confirmDel===v
                    ? <div style={{display:"flex",alignItems:"center",gap:5,background:"#fdecea",borderRadius:8,padding:"4px 8px",border:`1px solid ${RED}44`}}>
                        <span style={{fontSize:11,color:RED,fontWeight:600,whiteSpace:"nowrap"}}>?Eliminar?</span>
                        <button onClick={doRemove} style={{padding:"3px 9px",borderRadius:6,border:"none",background:RED,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:11}}>Sí</button>
                        <button onClick={()=>setConfirmDel(null)} style={{padding:"3px 8px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:11}}>No</button>
                      </div>
                    : <button onClick={()=>setConfirmDel(v)} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #fcc",background:"#fff",color:RED,cursor:"pointer",fontSize:12,display:"flex",alignItems:"center"}}><Trash size={12} strokeWidth={2.3}/></button>
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
    if(!form.username.trim()||!form.password.trim()||!form.name.trim()){toast.error("Completá nombre, usuario y contraseña");return;}
    try {
      if(editing) {
        const updated = {...users.find(u=>u.id===editing), ...form, priceList:form.priceList||"default", canSeeAll:form.canSeeAll!==false};
        setUsers(us=>us.map(u=>u.id===editing?updated:u));
        await db.saveUser(updated);
      } else {
        if(users.find(u=>u.username===form.username.trim())){toast.error("Ese usuario ya existe");return;}
        const newUser = {id:genId(),username:form.username.trim(),password:form.password,name:form.name.trim(),role:form.role||"vendedor",email:form.email||"",phone:form.phone||"",cargo:form.cargo||"",vendedor:form.vendedor||"",priceList:form.priceList||"default",canSeeAll:form.canSeeAll!==false,avatar:form.avatar||"",barcodeEnabled:form.barcodeEnabled||false};
        setUsers(us=>[...us,newUser]);
        await db.saveUser(newUser);
      }
      cancelEdit();
    } catch(e) {
      toast.error("Error al guardar: " + (e.message||JSON.stringify(e)));
    }
  };

  const remove = async (id) => {
    if(users.filter(u=>u.role==="admin").length===1&&users.find(u=>u.id===id)?.role==="admin"){toast.error("Debe haber al menos un administrador");return;}
    setUsers(us=>us.filter(u=>u.id!==id));
    await db.deleteUser(id);
  };

  const Toggle = ({label,sub,val,onChange}) => (
    <div style={{background:"#f9f9f9",borderRadius:8,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",gap:12}}>
      <div>
        <div style={{fontWeight:700,fontSize:13}}>{label}</div>
        {sub&&<div style={{fontSize:11,color:"#888",marginTop:2}}>{sub}</div>}
      </div>
      <button onClick={onChange} style={{padding:"6px 14px",borderRadius:20,border:"none",cursor:"pointer",fontWeight:700,fontSize:12,background:val?"#1e8449":"#e5e5e5",color:val?"#fff":"#888",whiteSpace:"nowrap",display:"flex",alignItems:"center",gap:5}}>
        {val?<><CheckCircle size={11} strokeWidth={2.5}/>Sí</>:<><Lock size={11} strokeWidth={2.5}/>No</>}
      </button>
    </div>
  );

  const isMobilePanel = useIsMobile();
  const [view, setView] = useState("lista");

  return (
    <div>
      <div style={{background:"#fff",borderRadius:12,padding:4,marginBottom:16,display:"flex",gap:4,boxShadow:"0 1px 4px #0001"}}>
        <button onClick={()=>{setView("lista");cancelEdit();}} style={{flex:1,padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="lista"?`linear-gradient(135deg,#922b21,${RED})`:"transparent",color:view==="lista"?"#fff":"#555",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
          <Lock size={13} strokeWidth={2.3}/> Usuarios ({users.length})
        </button>
        <button onClick={()=>setView("form")} style={{flex:1,padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:view==="form"?`linear-gradient(135deg,#922b21,${RED})`:"transparent",color:view==="form"?"#fff":"#555",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
          {editing?<><Pencil size={12} strokeWidth={2.4}/> Editando</>:<><Plus size={13} strokeWidth={2.4}/> Nuevo usuario</>}
        </button>
      </div>

      {view==="lista" && (
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {users.map(u=>(
            <div key={u.id} style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 4px #0001",border:`1.5px solid ${editing===u.id?"#c0392b":"#f0f0f0"}`}}>
              <div style={{display:"flex",alignItems:"center",gap:12}}>
                <div style={{width:48,height:48,borderRadius:"50%",background:u.avatar?"transparent":"linear-gradient(135deg,#7b1a1a,#c0392b)",overflow:"hidden",flexShrink:0,border:"2px solid #f0f0f0",display:"flex",alignItems:"center",justifyContent:"center"}}>
                  {u.avatar ? <img src={u.avatar} alt={u.name} style={{width:"100%",height:"100%",objectFit:"cover"}}/> : (u.role==="admin" ? <Crown size={20} color="#fff" strokeWidth={2}/> : <Users size={19} color="#fff" strokeWidth={2}/>)}
                </div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:700,fontSize:14}}>{u.name}{u.cargo&&<span style={{fontWeight:400,color:"#888",fontSize:12}}> · {u.cargo}</span>}</div>
                  <div style={{display:"flex",gap:5,flexWrap:"wrap",marginTop:4,alignItems:"center"}}>
                    <span style={{background:u.role==="admin"?"#fdecea":"#eaf4fc",color:u.role==="admin"?RED:"#1a5276",borderRadius:6,padding:"1px 8px",fontWeight:700,fontSize:10}}>{u.role==="admin"?"Admin":"Vendedor"}</span>
                    {u.vendedor&&<span style={{display:"inline-flex",alignItems:"center",gap:3,background:"#f5eef8",color:"#6c3483",borderRadius:6,padding:"1px 8px",fontWeight:700,fontSize:10}}><Users size={9} strokeWidth={2.5}/>{u.vendedor}</span>}
                    {u.barcodeEnabled&&<span style={{display:"inline-flex",alignItems:"center",gap:3,background:"#eaf4fc",color:"#1a5276",borderRadius:6,padding:"1px 8px",fontWeight:700,fontSize:10}}><Camera size={9} strokeWidth={2.5}/>Lector</span>}
                    {u.canSeeAll===false&&<span style={{display:"inline-flex",alignItems:"center",gap:3,background:"#fef9e7",color:"#b7770d",borderRadius:6,padding:"1px 8px",fontWeight:700,fontSize:10}}><Lock size={9} strokeWidth={2.5}/>Solo suyos</span>}
                  </div>
                  {(u.phone||u.email)&&<div style={{fontSize:11,color:"#aaa",marginTop:4,display:"flex",alignItems:"center",gap:5,flexWrap:"wrap"}}>{u.phone&&<span style={{display:"inline-flex",alignItems:"center",gap:3}}><Phone size={9} strokeWidth={2.4}/>{u.phone}</span>}{u.phone&&u.email&&"·"}{u.email&&<span style={{display:"inline-flex",alignItems:"center",gap:3}}><Mail size={9} strokeWidth={2.4}/>{u.email}</span>}</div>}
                </div>
                <div style={{display:"flex",gap:6,flexShrink:0}}>
                  <button onClick={()=>{startEdit(u);setView("form");}} style={{padding:"6px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:5}}><Pencil size={11} strokeWidth={2.4}/>Editar</button>
                  <button onClick={()=>remove(u.id)} style={{padding:"6px 10px",borderRadius:8,border:"1.5px solid #fcc",background:"#fff",color:RED,cursor:"pointer",fontSize:12,display:"flex",alignItems:"center"}}><Trash size={12} strokeWidth={2.3}/></button>
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
          <div style={{fontWeight:800,fontSize:15,marginBottom:16,display:"flex",alignItems:"center",gap:7}}>{editing?<><Pencil size={14} strokeWidth={2.4}/>Editando usuario</>:<><Plus size={15} strokeWidth={2.4}/>Nuevo usuario</>}</div>
          <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:16,padding:"14px 16px",background:"#f9f9f9",borderRadius:12}}>
            <div style={{width:72,height:72,borderRadius:"50%",background:avatarPreview?"transparent":"linear-gradient(135deg,#7b1a1a,#c0392b)",overflow:"hidden",flexShrink:0,border:"3px solid #fff",boxShadow:"0 2px 8px #0002",display:"flex",alignItems:"center",justifyContent:"center"}}>
              {avatarPreview ? <img src={avatarPreview} alt="avatar" style={{width:"100%",height:"100%",objectFit:"cover"}}/> : <Users size={30} color="#fff" strokeWidth={2}/>}
            </div>
            <div style={{flex:1}}>
              <div style={{fontWeight:700,fontSize:13,marginBottom:6}}>Foto de perfil</div>
              <label style={{display:"inline-flex",alignItems:"center",gap:6,padding:"8px 14px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",fontSize:12,fontWeight:600,cursor:"pointer",color:"#555"}}>
                <Camera size={13} strokeWidth={2.3}/> Subir foto
                <input type="file" accept="image/*" onChange={handleAvatar} style={{display:"none"}}/>
              </label>
              {avatarPreview&&<button onClick={()=>{setAvatarPreview("");setForm(f=>({...f,avatar:""}));}} style={{marginLeft:8,padding:"8px 12px",borderRadius:8,border:"none",background:"#fdecea",color:RED,fontSize:12,cursor:"pointer",fontWeight:600,display:"inline-flex",alignItems:"center",gap:4}}><XIcon size={10} strokeWidth={2.6}/>Quitar</button>}
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
                <button onClick={()=>setShowPass(s=>({...s,form:!s.form}))} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:"#aaa",display:"flex",alignItems:"center"}}>{showPass.form?<EyeOff size={15} strokeWidth={2.2}/>:<Eye size={15} strokeWidth={2.2}/>}</button>
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
          <Toggle label={<><Camera size={12} strokeWidth={2.4} style={{display:"inline",verticalAlign:"-2px",marginRight:5}}/>Lector de código de barras</>} sub={form.barcodeEnabled?"Puede usar el lector":"Sin acceso al lector"} val={form.barcodeEnabled} onChange={()=>setForm(f=>({...f,barcodeEnabled:!f.barcodeEnabled}))}/>
          <div style={{display:"flex",gap:8,marginTop:8}}>
            <button onClick={save} style={{flex:1,padding:"11px",borderRadius:10,border:"none",background:RED,color:"#fff",fontWeight:800,cursor:"pointer",fontSize:14,display:"flex",alignItems:"center",justifyContent:"center",gap:7}}>{editing?<><Save size={14} strokeWidth={2.3}/>Guardar cambios</>:<><CheckCircle size={14} strokeWidth={2.3}/>Crear usuario</>}</button>
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
  const [fixing, setFixing] = useState(false);

  // Productos con el precio de venta roto (igual o menor al costo) por un import anterior
  const rotos = useMemo(()=>products.filter(p=>p.costPrice>0 && p.salePrice<=p.costPrice), [products]);

  const corregirPrecios = async () => {
    const ok = await confirmDialog(
      "¿Corregir precios?",
      `Se va a recalcular el precio de venta de ${rotos.length} producto${rotos.length!==1?"s":""} como costo × 1,5 (tu fórmula: ×1.25 y ×1.20). Esta acción no se puede deshacer.`,
      true
    );
    if(!ok) return;
    setFixing(true);
    try {
      const corregidos = rotos.map(p=>({...p, salePrice: Math.round(p.costPrice*1.5*100)/100}));
      await db.upsertProducts(corregidos);
      setProducts(prev=>prev.map(p=>{
        const fix = corregidos.find(c=>c.id===p.id);
        return fix ? fix : p;
      }));
      toast.success(`${corregidos.length} producto${corregidos.length!==1?"s":""} corregido${corregidos.length!==1?"s":""}`);
    } catch(e) {
      console.warn(e);
      toast.error("No se pudieron corregir los precios. Probá de nuevo.");
    } finally {
      setFixing(false);
    }
  };

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

        // FIX 1: rango de busqueda del encabezado mas amplio (algunas listas de proveedores
        // traen varias filas de texto legal/contacto antes del encabezado real) + corte
        // correcto del loop externo apenas se encuentra (antes seguia pisando headerRow).
        let headerRow = 0;
        let headerFound = false;
        for(let r = 0; r <= Math.min(40, totalRows) && !headerFound; r++) {
          for(let c = 0; c <= 20; c++) {
            const v = cellStr(r, c).toUpperCase()
              .normalize("NFD").replace(/[̀-ͯ]/g,"");
            if(v === "CODIGO" || v === "COD") { headerRow = r; headerFound = true; break; }
          }
        }

        const norm = s => String(s).toUpperCase().trim()
          .normalize("NFD").replace(/[̀-ͯ]/g,"");

        // FIX 2: cada columna se asigna solo si todavia no tiene match ("COL.x<0 &&"), asi la
        // PRIMERA columna que matchea gana siempre. Antes "CODIGO DE BARRAS" pisaba a "CODIGO"
        // porque ambas contienen la palabra CODIGO y el codigo se quedaba con la ultima.
        const COL = { codigo:-1, descripcion:-1, precioIVA:-1, precioOferta:-1, precioFinal:-1, fecha:-1, unidad:-1, barras:-1 };
        for(let c = range.s.c; c <= range.e.c; c++) {
          const h = norm(cellStr(headerRow, c));
          if(COL.codigo<0 && (h==="CODIGO" || h==="COD" || h==="ID")) COL.codigo = c;
          else if(COL.descripcion<0 && (h.includes("DESCRIP") || h.includes("NOMBRE"))) COL.descripcion = c;
          else if(COL.precioIVA<0 && (h.includes("CON IVA") || (h.includes("IVA") && !h.includes("OFERTA") && !h.includes("FINAL")))) COL.precioIVA = c;
          else if(COL.precioOferta<0 && h.includes("OFERTA")) COL.precioOferta = c;
          else if(COL.precioFinal<0 && h.includes("FINAL")) COL.precioFinal = c;
          else if(COL.fecha<0 && (h.includes("FECHA") || h.includes("ULTIMA") || h.includes("ACT"))) COL.fecha = c;
          else if(COL.unidad<0 && h==="UNIDAD") COL.unidad = c;
          else if(COL.barras<0 && h.includes("BARRAS")) COL.barras = c;
        }

        if(COL.codigo      < 0) COL.codigo      = 0;
        if(COL.descripcion < 0) COL.descripcion = 1;
        if(COL.precioIVA   < 0) COL.precioIVA   = 2;
        if(COL.precioOferta< 0) COL.precioOferta = 3;
        if(COL.fecha       < 0) COL.fecha        = 4;
        // FIX 3: "Precio Final" NO tiene un fallback adivinado. Si la lista no trae esa columna
        // (como la de Papelera Bariloche), se deja sin mapear: mas abajo se ignora en vez de
        // leer por accidente la columna de fecha y guardar un numero de serie como si fuera precio.
        // "Unidad" (multiplo obligatorio de compra al proveedor) y "Codigo de barras" tampoco
        // tienen fallback adivinado: si la columna no existe, simplemente no se completan.

        const colLetter = c => c < 0 ? "-" : String.fromCharCode(65 + c);
        const detectedCols = {
          "Código":        colLetter(COL.codigo),
          "Descripción":   colLetter(COL.descripcion),
          "Precio c/IVA":  colLetter(COL.precioIVA),
          "Precio Oferta": colLetter(COL.precioOferta),
          "Fecha":         colLetter(COL.fecha),
          "Precio Final":  colLetter(COL.precioFinal),
          "Múltiplo (Unidad)": colLetter(COL.unidad),
          "Código de Barras": colLetter(COL.barras),
        };

        const parsed = [];
        for(let r = headerRow + 1; r <= totalRows; r++) {
          const id = cellStr(r, COL.codigo);
          if(!id) continue;
          const pIVA    = cellNum(r, COL.precioIVA);
          const pOferta = cellNum(r, COL.precioOferta);
          const pFinal  = COL.precioFinal >= 0 ? cellNum(r, COL.precioFinal) : null;
          const unidadRaw = COL.unidad >= 0 ? cellNum(r, COL.unidad) : null;
          const multiplo = (unidadRaw && unidadRaw > 0) ? Math.round(unidadRaw) : 1;
          const barcode = COL.barras >= 0 ? cellStr(r, COL.barras) : "";
          parsed.push({
            id,
            name:         cellStr(r, COL.descripcion),
            precioIVA:    pIVA,
            precioOferta: pOferta,
            multiplo,
            barcode,
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

    // Formula real de venta de Sebastian: costo x 1.25 x 1.20 (= costo x 1.5)
    // Se aplica SOLO cuando el archivo no trae una columna "Precio Final" ya calculada
    // (ej: su Excel propio con la formula puesta) — si la trae, esa tiene prioridad.
    const MARKUP_VENTA = 1.25 * 1.20;

    // El costo real es el de oferta si el proveedor tiene uno activo (es un costo temporal
    // mas bajo, NO un precio de venta), sino el normal con IVA.
    const resolveCosto = (row) => {
      const oferta = row.precioOferta;
      const iva    = row.precioIVA;
      if(oferta !== null && oferta !== undefined && !isNaN(oferta) && oferta > 0) return oferta;
      return iva;
    };

    preview.all.forEach(row=>{
      const idx = newProds.findIndex(p=>p.id===row.id);
      const costo = resolveCosto(row);  // oferta si esta activa, sino IVA
      if(idx>=0){
        if(row.precioFinal!==null && row.precioFinal>0)
                          newProds[idx].salePrice = row.precioFinal;
        else if(costo!==null) newProds[idx].salePrice = Math.round(costo*MARKUP_VENTA*100)/100;
        if(costo!==null) {
          // Guarda el costo previo ANTES de pisarlo, para poder mostrar el cambio en el dashboard
          if(costo !== newProds[idx].costPrice) newProds[idx].costPriceAnterior = newProds[idx].costPrice;
          newProds[idx].costPrice = costo;
        }
        newProds[idx].multiploCompra = row.multiplo||1;
        if(row.barcode) newProds[idx].barcode = row.barcode;
        if(row.name) newProds[idx].name = row.name;
        updated.push(row.id);
      } else {
        notFound.push(row.id);
      }
    });

    if(mode==="full") {
      preview.all.forEach(row=>{
        if(!newProds.find(p=>p.id===row.id)) {
          const costo = resolveCosto(row);
          newProds.push({
            id:row.id, name:row.name||row.id,
            costPrice:costo||0,
            salePrice:(row.precioFinal&&row.precioFinal>0)?row.precioFinal:Math.round((costo||0)*MARKUP_VENTA*100)/100,
            category:"Importado", stock:0, multiploCompra:row.multiplo||1, barcode:row.barcode||""
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
    setStatus({type:"success",msg:`${updated.length} productos actualizados.${notFound.length>0?` ${notFound.length} códigos no encontrados en el catálogo.`:""}`});
    setPreview(null);
    if(fileRef.current) fileRef.current.value="";
  };

  return (
    <div>
      {rotos.length>0 && (
        <div style={{background:"#fdecea",border:"1px solid #f1948a",borderRadius:12,padding:"14px 18px",marginBottom:16,display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
          <div style={{flex:1,minWidth:220}}>
            <div style={{fontWeight:800,fontSize:14,color:RED,display:"flex",alignItems:"center",gap:6}}><AlertTriangle size={14} strokeWidth={2.3}/>{rotos.length} producto{rotos.length!==1?"s":""} con precio de venta roto</div>
            <div style={{fontSize:12,color:"#a33",marginTop:2}}>Quedaron con precio de venta igual o menor al costo, probablemente por un import anterior. Esto los recalcula como costo × 1,5 (tu fórmula habitual).</div>
          </div>
          <button onClick={corregirPrecios} disabled={fixing}
            style={{padding:"9px 16px",borderRadius:9,border:"none",background:RED,color:"#fff",fontWeight:700,fontSize:13,cursor:fixing?"default":"pointer",opacity:fixing?0.7:1,whiteSpace:"nowrap",display:"flex",alignItems:"center",gap:6}}>
            {fixing?"Corrigiendo...":<><Wrench size={12} strokeWidth={2.4}/> Corregir {rotos.length} producto{rotos.length!==1?"s":""}</>}
          </button>
        </div>
      )}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,alignItems:"start"}}>
      <div style={{background:"#fff",borderRadius:12,padding:24,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:16,marginBottom:6,display:"flex",alignItems:"center",gap:7}}><BarChart size={16} strokeWidth={2.2}/>Importar Lista de Precios</div>
        <div style={{fontSize:13,color:"#666",marginBottom:16,lineHeight:1.6}}>
          El sistema lee automáticamente las columnas:<br/>
          <strong>CÓDIGO . DESCRIPCIÓN . PRECIO CON IVA . PRECIO OFERTA . FECHA ULTIMA ACTUALIZACIÓN . PRECIO FINAL . UNIDAD (múltiplo de compra) . CÓDIGO DE BARRAS</strong>
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
          <div style={{marginBottom:8,display:"flex",justifyContent:"center"}}><Folder size={32} color="#ccc" strokeWidth={1.7}/></div>
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
          <button onClick={applyImport} disabled={status?.type==="progress"} style={{width:"100%",padding:"11px",borderRadius:10,border:"none",background:status?.type==="progress"?"#aaa":`linear-gradient(135deg,${REDD},${RED})`,color:"#fff",fontWeight:800,fontSize:14,cursor:status?.type==="progress"?"not-allowed":"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:7}}>
            {status?.type==="progress"?<><Clock size={13} strokeWidth={2.4}/> Importando, no cierres esta página...</>:<><Save size={13} strokeWidth={2.4}/> Aplicar importación ({preview.total} productos)</>}
          </button>
        )}
      </div>
      <div style={{background:"#fff",borderRadius:12,padding:24,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:15,marginBottom:12,display:"flex",alignItems:"center",gap:6}}>
          {preview ? <><Eye size={14} strokeWidth={2.3}/>Vista previa ({preview.total} filas)</> : <><ClipboardList size={14} strokeWidth={2.3}/>Instrucciones</>}
        </div>
        {preview?.detectedCols && (
          <div style={{background:"#f0fdf4",border:"1px solid #a7f3d0",borderRadius:8,padding:"10px 12px",marginBottom:12}}>
            <div style={{fontWeight:700,fontSize:11,color:"#065f46",marginBottom:6,display:"flex",alignItems:"center",gap:5}}><CheckCircle size={11} strokeWidth={2.5}/>Columnas detectadas:</div>
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
                <div><strong>Col G:</strong> UNIDAD (múltiplo de compra, opcional)</div>
              </div>
              <div style={{marginTop:12,fontSize:12,color:"#aaa"}}>El sistema detecta automáticamente las columnas por nombre. Las columnas no tienen que estar en un orden específico.</div>
            </div>
          : <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                <thead><tr style={{background:"#f9f9f9"}}>
                  {["Código","Descripción","P. IVA","P. Oferta","P. Final","Fecha","Múltiplo","Cód. Barras"].map(h=><th key={h} style={{padding:"7px 8px",textAlign:"left",fontWeight:700,color:"#888",whiteSpace:"nowrap"}}>{h}</th>)}
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
                      <td style={{padding:"6px 8px",color:r.multiplo>1?"#b7770d":"#ccc",fontWeight:r.multiplo>1?700:400}}>{r.multiplo>1?`×${r.multiplo}`:"—"}</td>
                      <td style={{padding:"6px 8px",color:"#aaa",fontSize:10,whiteSpace:"nowrap"}}>{r.barcode||"—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {preview.total>10&&<div style={{fontSize:11,color:"#aaa",marginTop:8,textAlign:"center"}}>... y {preview.total-10} filas más</div>}
            </div>
        }
      </div>
    </div>
    </div>
  );
}
