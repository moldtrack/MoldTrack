import { useState, useMemo, useRef, useEffect, useCallback } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";

const SUPABASE_URL = "https://jqyehnohnwlvjonciyiw.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpxeWVobm9obndsdmpvbmNpeWl3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAxNjQ1NzcsImV4cCI6MjA5NTc0MDU3N30.CBl41ugxddq44hqC2lKc0LUoR0PgMwSoZbQPaGtvI48";
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

const PROBLEMS = [
  { id:"MOR_ATAS_DALAM",         label:"MOR Atas Ke Dalam",       group:"MOR" },
  { id:"MOR_ATAS_KELUAR",        label:"MOR Atas Ke Luar",        group:"MOR" },
  { id:"MOR_BAWAH_DALAM",        label:"MOR Bawah Ke Dalam",      group:"MOR" },
  { id:"MOR_BAWAH_KELUAR",       label:"MOR Bawah Ke Luar",       group:"MOR" },
  { id:"OOR",                    label:"OOR (Out of Register)",   group:"OOR" },
  { id:"OVERFLOW_SHOULDER_ATAS", label:"Overflow Shoulder Atas",  group:"Overflow" },
  { id:"OVERFLOW_SHOULDER_BAWAH",label:"Overflow Shoulder Bawah", group:"Overflow" },
  { id:"OS",                     label:"OS (Open Segment)",       group:"OS" },
];
const SECTORS  = ["A","B","C","D","E","F","G","H"];
const PLANTS   = ["D","K"];
const LINES    = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N"];
const MACHINES_D = Array.from({length:15},(_,i)=>String(i+1).padStart(2,"0"));   // 01-15
const MACHINES_K = Array.from({length:10},(_,i)=>String(i+21).padStart(2,"0"));  // 21-30
const getMachines = (plant) => plant==="K" ? MACHINES_K : MACHINES_D;
const MAKER_CONTAINER = ["Greatoo","Seahwa","Himille","Sumhing","Herbert","AZ German","Tyangyang"];
const TYPE_CONTAINER  = ["S-TYPE","AZIII/2","AZIII","AZIV","L46","L48","AZ V","AW 200"];
const PRESSES  = ["L","R"];

const ROLES    = ["teknisi","persiapan","qcgate","analyst","adh","dh","admin"];

// ── EXPORT TO EXCEL ───────────────────────────────────────────────────────────
const exportToExcel = (records) => {
  const getProblemLabel = (pid) => PROBLEMS.find(p=>p.id===pid)?.label || pid;

  const calcDuration = (start, end) => {
    if (!start || !end) return "";
    const [h1,m1] = start.split(":").map(Number);
    const [h2,m2] = end.split(":").map(Number);
    const diff = (h2*60+m2) - (h1*60+m1);
    return diff > 0 ? diff : "";
  };

  const getTindakan = (problems, problemDetails) => {
    const parts = [];
    (problems||[]).forEach(pid => {
      const det = (problemDetails||{})[pid];
      if (!det) return;
      const lbl = getProblemLabel(pid);
      const isMOR = pid.startsWith("MOR");
      const isOverflow = pid.startsWith("OVERFLOW");
      const isOS = pid === "OS";
      const isOOR = pid === "OOR";
      if ((isMOR||isOverflow||isOS) && det.shimJumlah) {
        const lokasi = isMOR && det.shimSectors?.length>0
          ? `Sektor ${det.shimSectors.join(",")}`
          : det.shimLokasi || "";
        parts.push(`${lbl}: Shim ${det.shimAction} ${det.shimJumlah} lbr${lokasi?` (${lokasi})`:""}` );
      }
      if (isOOR && det.oorSectors?.length>0) {
        const tipe = det.oorType === "segmented" ? "Segmented" : "Two Piece";
        parts.push(`${lbl}: ${tipe} - Sektor ${det.oorSectors.join(",")}`);
      }
    });
    return parts.join(" | ");
  };

  const rows = records.map((r, i) => ({
    "No":               i + 1,
    "Tanggal":          r.date || "",
    "Size Mold":        r.mold_size || "",
    "Maker Container L":r.maker_container_l || "",
    "Maker Container R":r.maker_container_r || "",
    "Type Container L": r.type_container_l || "",
    "Type Container R": r.type_container_r || "",
    "No Container L":   r.container_no_l || "",
    "No Container R":   r.container_no_r || "",
    "No Mold L":        r.mold_no_l || "",
    "No Mold R":        r.mold_no_r || "",
    "Kode Mesin":       r.machine_code || "",
    "Plant":            r.plant || "",
    "Line":             r.line || "",
    "Mesin":            r.machine || "",
    "Press":            Array.isArray(r.press) ? r.press.join(" & ") : (r.press || ""),
    "Jenis Problem":    (r.problems||[]).map(getProblemLabel).join(" | "),
    "Detail Tindakan":  getTindakan(r.problems, r.problem_details),
    "Teknisi":          r.technician || "",
    "Jam Mulai":        r.jam_mulai || "",
    "Jam Selesai":      r.jam_selesai || "",
    "Durasi (menit)":   calcDuration(r.jam_mulai, r.jam_selesai),
    "Catatan":          r.notes || "",
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);

  // Set column widths
  ws["!cols"] = [
    {wch:5},   // No
    {wch:12},  // Tanggal
    {wch:14},  // Size Mold
    {wch:16},  // Maker Container L
    {wch:16},  // Maker Container R
    {wch:14},  // Type Container L
    {wch:14},  // Type Container R
    {wch:14},  // No Container L
    {wch:14},  // No Container R
    {wch:12},  // No Mold L
    {wch:12},  // No Mold R
    {wch:12},  // Kode Mesin
    {wch:8},   // Plant
    {wch:8},   // Line
    {wch:8},   // Mesin
    {wch:8},   // Press
    {wch:30},  // Jenis Problem
    {wch:40},  // Detail Tindakan
    {wch:16},  // Teknisi
    {wch:10},  // Jam Mulai
    {wch:10},  // Jam Selesai
    {wch:14},  // Durasi
    {wch:40},  // Catatan
  ];

  const lastRow = rows.length + 1;
  const lastCol = "W";
  ws["!autofilter"] = { ref: `A1:${lastCol}${lastRow}` };

  XLSX.utils.book_append_sheet(wb, ws, "Data Perbaikan");

  const date = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `MoldTrack_${date}.xlsx`);
};

// role permissions
const canDashboard   = (r) => ["analyst","adh","dh","admin"].includes(r);
const canDatabase    = (r) => ["analyst","adh","dh","admin","teknisi"].includes(r);
const canPersiapan   = (r) => ["persiapan","analyst","adh","dh","admin"].includes(r);
const canQCGate      = (r) => ["qcgate","analyst","adh","dh","admin"].includes(r);
const canNaik        = (r) => ["naik","analyst","adh","dh","admin"].includes(r);
const canEntry       = (r) => ["teknisi","analyst","adh","dh","admin"].includes(r);
const canManageUsers = (r) => r === "admin";
const getGrup = (username) => {
  if (!username) return null;
  const match = username.match(/-(a|b|c|d)$/i);
  return match ? match[1].toUpperCase() : null;
};
const canDeleteEdit  = (r, recordUserId, currentUserId) =>
  r === "admin" || recordUserId === currentUserId;

const emptyProblemDetail = (pid) => {
  if (pid.startsWith("MOR"))      return { shimAction:"tambah", shimJumlah:"", shimSectors:[] };
  if (pid.startsWith("OVERFLOW")) return { shimAction:"tambah", shimJumlah:"", shimLokasi:"" };
  if (pid === "OS")               return { shimAction:"tambah", shimJumlah:"", shimLokasi:"" };
  if (pid === "OOR")              return { oorType:"segmented", oorSectors:[] };
  return {};
};
const emptyForm = () => ({
  moldSize:"",
  makerContainerL:"", makerContainerR:"", typeContainerL:"", typeContainerR:"",
  containerNoL:"", containerNoR:"", moldNoL:"", moldNoR:"",
  date:new Date().toISOString().slice(0,10),
  jamMulai:"", jamSelesai:"",
  technician:"",
  plant:"", line:"", machine:"", press:[],
  problems:[], problemDetails:{}, notes:"",
});

const CSS = `
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:system-ui,-apple-system,sans-serif;background:#f5f5f5;}
  .app{display:flex;flex-direction:column;min-height:100vh;}
  .sidebar{display:none;}
  .bottom-nav{position:fixed;bottom:0;left:0;right:0;background:#fff;border-top:1px solid #e5e5e5;display:flex;z-index:50;padding-bottom:env(safe-area-inset-bottom);}
  .bottom-nav-item{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:8px 4px;cursor:pointer;font-size:10px;color:#999;gap:3px;border:none;background:none;}
  .bottom-nav-item.active{color:#1D9E75;}
  .bottom-nav-item i{font-size:20px;}
  .topbar{background:#fff;border-bottom:1px solid #e5e5e5;padding:12px 16px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:40;}
  .topbar-logo{display:flex;align-items:center;gap:8px;}
  .logo-dot{width:28px;height:28px;background:#1D9E75;border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0;}
  .topbar-title{font-size:14px;font-weight:600;color:#111;}
  .topbar-sub{font-size:10px;color:#999;}
  .content{flex:1;overflow-y:auto;padding:16px;padding-bottom:80px;}
  .card{background:#fff;border-radius:12px;padding:16px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.06);}
  .card-title{font-size:13px;font-weight:600;color:#111;margin-bottom:12px;}
  .stat-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px;}
  .stat-card{background:#f8f8f8;border-radius:10px;padding:12px;}
  .stat-label{font-size:10px;color:#999;margin-bottom:4px;}
  .stat-val{font-size:22px;font-weight:700;color:#111;}
  .stat-sub{font-size:10px;color:#999;margin-top:2px;}
  .pbadge{font-size:11px;padding:2px 8px;border-radius:4px;display:inline-block;margin-right:3px;margin-bottom:3px;font-weight:500;}
  .form-label{font-size:11px;color:#666;font-weight:500;margin-bottom:5px;display:block;}
  .form-input{width:100%;padding:10px 12px;border:1px solid #e0e0e0;border-radius:8px;font-size:14px;color:#111;background:#fff;outline:none;}
  .form-input:focus{border-color:#1D9E75;}
  .form-group{margin-bottom:14px;}
  .form-row{display:grid;grid-template-columns:1fr 1fr;gap:10px;}
  .type-btn-group{display:grid;grid-template-columns:1fr 1fr;gap:8px;}
  .type-btn{padding:10px 8px;border-radius:8px;border:1.5px solid #e0e0e0;background:#fff;font-size:12px;color:#666;cursor:pointer;text-align:center;font-weight:400;}
  .type-btn.active{border-color:#1D9E75;background:#E1F5EE;color:#085041;font-weight:500;}
  .prob-card{border:1.5px solid #e0e0e0;border-radius:10px;margin-bottom:8px;overflow:hidden;background:#f9f9f9;}
  .prob-card.active{border-color:#1D9E75;background:#f0faf5;}
  .prob-card-header{display:flex;align-items:center;gap:10px;padding:12px;cursor:pointer;}
  .prob-check{width:18px;height:18px;border-radius:5px;border:1.5px solid #ccc;background:#fff;display:flex;align-items:center;justify-content:center;flex-shrink:0;}
  .prob-check.active{background:#1D9E75;border-color:#1D9E75;}
  .prob-card-detail{padding:0 12px 14px;border-top:1px solid #c8eedd;}
  .shim-toggle-group{display:flex;gap:8px;margin-bottom:10px;}
  .shim-toggle{flex:1;padding:9px;border-radius:8px;border:1.5px solid #e0e0e0;background:#f9f9f9;font-size:12px;color:#666;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:6px;}
  .shim-toggle.active{border-color:#1D9E75;background:#E1F5EE;color:#085041;font-weight:500;}
  .radio-dot{width:12px;height:12px;border-radius:50%;background:#ccc;flex-shrink:0;}
  .radio-dot.active{background:#1D9E75;}
  .sector-grid{display:flex;flex-wrap:wrap;gap:8px;}
  .sector-btn{width:40px;height:40px;border-radius:8px;border:1.5px solid #e0e0e0;background:#f9f9f9;font-size:14px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;color:#444;}
  .sector-btn.active{background:#1D9E75;border-color:#1D9E75;color:#fff;}
  .btn-primary{background:#1D9E75;color:#fff;border:none;padding:12px 20px;border-radius:10px;font-size:14px;font-weight:600;cursor:pointer;width:100%;margin-bottom:8px;}
  .btn-secondary{background:#fff;color:#444;border:1.5px solid #e0e0e0;padding:11px 20px;border-radius:10px;font-size:14px;cursor:pointer;width:100%;margin-bottom:8px;}
  .btn-danger{background:#fff;color:#E24B4A;border:1.5px solid #E24B4A;padding:11px 20px;border-radius:10px;font-size:14px;cursor:pointer;width:100%;}
  .btn-sm{padding:6px 12px;border-radius:6px;border:1px solid #e0e0e0;background:#fff;font-size:12px;cursor:pointer;color:#444;}
  .btn-sm-danger{padding:6px 12px;border-radius:6px;border:1px solid #E24B4A;background:#fff;font-size:12px;cursor:pointer;color:#E24B4A;}
  .record-card{background:#fff;border-radius:10px;padding:14px;margin-bottom:10px;box-shadow:0 1px 3px rgba(0,0,0,0.06);}
  .record-card-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px;}
  .record-size{font-size:16px;font-weight:700;color:#111;}
  .record-type{font-size:10px;color:#999;margin-top:1px;}
  .record-meta{font-size:11px;color:#999;margin-top:6px;}
  .record-actions{display:flex;gap:6px;margin-top:10px;padding-top:10px;border-top:1px solid #f0f0f0;}
  .detail-section{background:#f8f8f8;border-radius:8px;padding:12px;margin-bottom:10px;}
  .toast{position:fixed;top:16px;left:50%;transform:translateX(-50%);z-index:9999;padding:10px 20px;border-radius:10px;font-size:13px;font-weight:500;white-space:nowrap;box-shadow:0 4px 12px rgba(0,0,0,0.15);}
  .toast.success{background:#E1F5EE;color:#085041;border:1px solid #9FE1CB;}
  .toast.error{background:#FCEBEB;color:#791F1F;border:1px solid #F09595;}
  .bar-row{margin-bottom:10px;}
  .bar-row-header{display:flex;justify-content:space-between;font-size:11px;color:#666;margin-bottom:4px;}
  .bar-track{height:6px;background:#f0f0f0;border-radius:3px;}
  .bar-fill{height:6px;background:#1D9E75;border-radius:3px;}
  .search-wrap{position:relative;margin-bottom:10px;}
  .search-icon{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:#999;font-size:16px;}
  .search-input{width:100%;padding:10px 12px 10px 36px;border:1px solid #e0e0e0;border-radius:8px;font-size:14px;background:#fff;outline:none;color:#111;}
  .filter-row{display:flex;gap:8px;margin-bottom:12px;overflow-x:auto;padding-bottom:2px;}
  .filter-select{padding:8px 10px;border:1px solid #e0e0e0;border-radius:8px;font-size:12px;background:#fff;white-space:nowrap;flex-shrink:0;color:#111;}
  .autocomplete-wrap{position:relative;}
  .autocomplete-dropdown{position:absolute;top:100%;left:0;right:0;background:#fff;border:1px solid #e0e0e0;border-radius:8px;z-index:100;max-height:160px;overflow-y:auto;margin-top:2px;box-shadow:0 4px 12px rgba(0,0,0,0.1);}
  .autocomplete-item{padding:10px 12px;font-size:13px;cursor:pointer;color:#111;}
  .autocomplete-item:hover{background:#f5f5f5;}
  .status-dot{width:8px;height:8px;border-radius:50%;background:#1D9E75;display:inline-block;margin-right:5px;}
  .group-label{font-size:10px;color:#999;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;margin-bottom:6px;margin-top:4px;}
  .role-badge{display:inline-block;padding:2px 10px;border-radius:20px;font-size:11px;font-weight:600;text-transform:uppercase;}
  .role-admin{background:#FEE2E2;color:#991B1B;}
  .role-dh{background:#FEF3C7;color:#92400E;}
  .role-adh{background:#E0F2FE;color:#075985;}
  .role-analyst{background:#F3E8FF;color:#6B21A8;}
  .role-teknisi{background:#E1F5EE;color:#085041;}
  /* LOGIN PAGE */
  .login-page{min-height:100vh;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#f0faf5 0%,#e8f5f0 100%);padding:20px;}
  .login-card{background:#fff;border-radius:20px;padding:32px 28px;width:100%;max-width:380px;box-shadow:0 8px 32px rgba(0,0,0,0.08);}
  .login-logo{display:flex;align-items:center;gap:12px;margin-bottom:28px;}
  .login-logo-dot{width:44px;height:44px;background:#1D9E75;border-radius:12px;display:flex;align-items:center;justify-content:center;}
  .login-title{font-size:20px;font-weight:700;color:#111;}
  .login-sub{font-size:12px;color:#999;margin-top:2px;}
  .login-input-wrap{position:relative;margin-bottom:14px;}
  .login-input{width:100%;padding:12px 14px 12px 42px;border:1.5px solid #e0e0e0;border-radius:10px;font-size:14px;color:#111;outline:none;background:#fff;}
  .login-input:focus{border-color:#1D9E75;}
  .login-input-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#999;font-size:17px;}
  .login-btn{width:100%;padding:13px;background:#1D9E75;color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;margin-top:4px;}
  .login-btn:disabled{opacity:0.6;cursor:not-allowed;}
  .login-error{background:#FCEBEB;color:#791F1F;border:1px solid #F09595;border-radius:8px;padding:10px 14px;font-size:13px;margin-bottom:14px;}
  @media (min-width:768px) {
    .app{flex-direction:row;}
    .bottom-nav{display:none;}
    .sidebar{width:210px;height:100vh;position:sticky;top:0;background:#fff;border-right:1px solid #e5e5e5;display:flex;flex-direction:column;flex-shrink:0;}
    .sidebar-logo{padding:18px 16px;border-bottom:1px solid #e5e5e5;}
    .sidebar-nav{padding:10px 0;flex:1;}
    .sidebar-nav-item{display:flex;align-items:center;gap:10px;padding:10px 16px;cursor:pointer;font-size:13px;color:#666;border-left:2px solid transparent;}
    .sidebar-nav-item.active{color:#1D9E75;background:#f0faf5;border-left-color:#1D9E75;}
    .sidebar-nav-item i{font-size:17px;flex-shrink:0;}
    .sidebar-footer{padding:12px 16px;border-top:1px solid #e5e5e5;}
    .main{flex:1;min-width:0;display:flex;flex-direction:column;}
    .content{padding:20px;padding-bottom:20px;}
    .stat-grid{grid-template-columns:repeat(4,1fr);}
    .dashboard-charts{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:12px;}
    .topbar{padding:12px 20px;}
  }
`;

// ── SIMPLE PASSWORD HASH (base64 - cukup untuk internal app) ──────────────────
const hashPass = (p) => btoa(unescape(encodeURIComponent(p)));
const checkPass = (plain, hashed) => hashPass(plain) === hashed;

function AutocompleteInput({ value, onChange, suggestions, placeholder, label }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  const filtered = suggestions.filter(s => s.toLowerCase().includes(value.toLowerCase()) && s.toLowerCase() !== value.toLowerCase());
  useEffect(() => {
    const h = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, []);
  return (
    <div className="form-group autocomplete-wrap" ref={ref}>
      {label && <label className="form-label">{label}</label>}
      <input className="form-input" value={value} onChange={e=>{onChange(e.target.value);setOpen(true);}} onFocus={()=>setOpen(true)} placeholder={placeholder} autoComplete="off" />
      {open && filtered.length > 0 && (
        <div className="autocomplete-dropdown">
          {filtered.map(s=><div key={s} className="autocomplete-item" onMouseDown={()=>{onChange(s);setOpen(false);}}>{s}</div>)}
        </div>
      )}
    </div>
  );
}

const ProblemBadge = ({ pid }) => {
  const p = PROBLEMS.find(x=>x.id===pid);
  const colors = { MOR:["#EAF3DE","#27500A"], OOR:["#E6F1FB","#0C447C"], Overflow:["#FAEEDA","#633806"], OS:["#FAECE7","#712B13"] };
  const [bg,fg] = colors[p?.group]||["#F1EFE8","#444441"];
  return <span className="pbadge" style={{ background:bg, color:fg }}>{p?.label}</span>;
};

const RoleBadge = ({ role }) => (
  <span className={`role-badge role-${role}`}>{role}</span>
);

function ProblemDetailForm({ pid, detail, onChange }) {
  const isMOR=pid.startsWith("MOR"), isOverflow=pid.startsWith("OVERFLOW"), isOS=pid==="OS", isOOR=pid==="OOR";
  const toggleArr = (key, s) => {
    const cur = detail[key]||[];
    onChange({...detail,[key]:cur.includes(s)?cur.filter(x=>x!==s):[...cur,s]});
  };
  if (isMOR) return (
    <div style={{ paddingTop:10 }}>
      <div className="shim-toggle-group">
        {["tambah","kurangi"].map(a=>(
          <div key={a} className={`shim-toggle${detail.shimAction===a?" active":""}`} onClick={()=>onChange({...detail,shimAction:a})}>
            <div className={`radio-dot${detail.shimAction===a?" active":""}`}></div>
            {a==="tambah"?"Tambah shim":"Kurangi shim"}
          </div>
        ))}
      </div>
      <div className="form-group">
        <label className="form-label">Jumlah shim (lembar)</label>
        <input type="number" min="1" className="form-input" style={{ width:140 }} placeholder="cth: 2" value={detail.shimJumlah||""} onChange={e=>onChange({...detail,shimJumlah:e.target.value})} />
      </div>
      <div className="form-group">
        <label className="form-label">Pilih sektor</label>
        <div className="sector-grid">
          {SECTORS.map(s=><div key={s} className={`sector-btn${(detail.shimSectors||[]).includes(s)?" active":""}`} onClick={()=>toggleArr("shimSectors",s)}>{s}</div>)}
        </div>
        {(detail.shimSectors||[]).length>0&&<div style={{ marginTop:8,fontSize:11,color:"#666" }}>Dipilih: <strong>{detail.shimSectors.join(", ")}</strong></div>}
      </div>
    </div>
  );
  if (isOverflow||isOS) {
    const opts = isOverflow?["Shoulder Atas","Shoulder Bawah","Atas & Bawah"]:["Atas","Bawah","Atas & Bawah"];
    return (
      <div style={{ paddingTop:10 }}>
        <div className="shim-toggle-group">
          {["tambah","kurangi"].map(a=>(
            <div key={a} className={`shim-toggle${detail.shimAction===a?" active":""}`} onClick={()=>onChange({...detail,shimAction:a})}>
              <div className={`radio-dot${detail.shimAction===a?" active":""}`}></div>
              {a==="tambah"?"Tambah shim":"Kurangi shim"}
            </div>
          ))}
        </div>
        <div className="form-row">
          <div className="form-group">
            <label className="form-label">Jumlah shim (lembar)</label>
            <input type="number" min="1" className="form-input" placeholder="cth: 2" value={detail.shimJumlah||""} onChange={e=>onChange({...detail,shimJumlah:e.target.value})} />
          </div>
          <div className="form-group">
            <label className="form-label">Lokasi shim</label>
            <select className="form-input" value={detail.shimLokasi||""} onChange={e=>onChange({...detail,shimLokasi:e.target.value})}>
              <option value="">— Pilih —</option>
              {opts.map(o=><option key={o}>{o}</option>)}
            </select>
          </div>
        </div>
      </div>
    );
  }
  if (isOOR) return (
    <div style={{ paddingTop:10 }}>
      <div className="shim-toggle-group">
        {[["segmented","Segmented (A–H)"],["two_piece","Two Piece"]].map(([val,lbl])=>(
          <div key={val} className={`shim-toggle${detail.oorType===val?" active":""}`} onClick={()=>onChange({...detail,oorType:val,oorSectors:[]})}>
            <div className={`radio-dot${detail.oorType===val?" active":""}`}></div>{lbl}
          </div>
        ))}
      </div>
      {detail.oorType==="segmented"&&(
        <div className="form-group">
          <label className="form-label">Pilih sektor yang diperbaiki</label>
          <div className="sector-grid">
            {SECTORS.map(s=><div key={s} className={`sector-btn${(detail.oorSectors||[]).includes(s)?" active":""}`} onClick={()=>toggleArr("oorSectors",s)}>{s}</div>)}
          </div>
          {(detail.oorSectors||[]).length>0&&<div style={{ marginTop:8,fontSize:11,color:"#666" }}>Dipilih: <strong>{detail.oorSectors.join(", ")}</strong></div>}
        </div>
      )}
      {detail.oorType==="two_piece"&&(
        <div className="form-group">
          <label className="form-label">Pilih bagian yang diperbaiki</label>
          <div className="shim-toggle-group">
            {["Mold Atas","Mold Bawah"].map(s=>(
              <div key={s} className={`shim-toggle${(detail.oorSectors||[]).includes(s)?" active":""}`} onClick={()=>toggleArr("oorSectors",s)}>
                <div className={`radio-dot${(detail.oorSectors||[]).includes(s)?" active":""}`}></div>{s}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
  return null;
}

const DetailSummary = ({ rec }) => {
  const d = rec.problemDetails||{};
  return (
    <div>
      {(rec.problems||[]).map(pid=>{
        const det=d[pid]; if(!det) return null;
        const isMOR=pid.startsWith("MOR"),isOverflow=pid.startsWith("OVERFLOW"),isOS=pid==="OS",isOOR=pid==="OOR";
        return (
          <div key={pid} style={{ marginBottom:4 }}>
            <ProblemBadge pid={pid}/>
            {(isMOR||isOverflow||isOS)&&det.shimJumlah&&(
              <span style={{ fontSize:11,color:"#666",display:"block",marginTop:2,paddingLeft:4 }}>
                Shim {det.shimAction} {det.shimJumlah} lbr
                {isMOR&&det.shimSectors?.length>0?` — Sektor: ${det.shimSectors.join(", ")}`:det.shimLokasi?` (${det.shimLokasi})`:""}
              </span>
            )}
            {isOOR&&det.oorSectors?.length>0&&(
              <span style={{ fontSize:11,color:"#666",display:"block",marginTop:2,paddingLeft:4 }}>Sektor: {det.oorSectors.join(", ")}</span>
            )}
          </div>
        );
      })}
    </div>
  );
};

// ── LOGIN PAGE ────────────────────────────────────────────────────────────────
function LoginPage({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError]       = useState("");
  const [loading, setLoading]   = useState(false);
  const [showPass, setShowPass] = useState(false);

  const handleLogin = async (e) => {
    e.preventDefault();
    if (!username.trim() || !password.trim()) { setError("Username dan password wajib diisi."); return; }
    setLoading(true); setError("");
    const { data, error: err } = await supabase
      .from("users")
      .select("*")
      .eq("username", username.trim().toLowerCase())
      .eq("is_active", true)
      .single();
    if (err || !data) { setError("Username tidak ditemukan atau akun nonaktif."); setLoading(false); return; }
    if (!checkPass(password, data.password_hash)) { setError("Password salah."); setLoading(false); return; }
    onLogin(data);
    setLoading(false);
  };

  return (
    <div className="login-page">
      <div className="login-card">
        <div className="login-logo">
          <div className="login-logo-dot">
            <svg width="22" height="22" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="6" stroke="white" strokeWidth="1.5"/><circle cx="8" cy="8" r="2" fill="white"/></svg>
          </div>
          <div>
            <div className="login-title">MoldTrack</div>
            <div className="login-sub">© by DMM</div>
          </div>
        </div>
        {error && <div className="login-error">{error}</div>}
        <form onSubmit={handleLogin}>
          <div className="login-input-wrap">
            <i className="ti ti-user login-input-icon"></i>
            <input className="login-input" type="text" placeholder="Username" value={username} onChange={e=>setUsername(e.target.value)} autoComplete="username" />
          </div>
          <div className="login-input-wrap" style={{ position:"relative" }}>
            <i className="ti ti-lock login-input-icon"></i>
            <input className="login-input" type={showPass?"text":"password"} placeholder="Password" value={password} onChange={e=>setPassword(e.target.value)} autoComplete="current-password" style={{ paddingRight:44 }} />
            <button type="button" onClick={()=>setShowPass(s=>!s)} style={{ position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:"#999",fontSize:17 }}>
              <i className={`ti ${showPass?"ti-eye-off":"ti-eye"}`}></i>
            </button>
          </div>
          <button className="login-btn" type="submit" disabled={loading}>
            {loading ? "Masuk..." : "Masuk"}
          </button>
        </form>
        <div style={{ textAlign:"center",marginTop:16,fontSize:11,color:"#999" }}>Hubungi Admin untuk mendapatkan akun</div>
      </div>
    </div>
  );
}

// ── USER MANAGEMENT PAGE (Admin only) ────────────────────────────────────────
function UserManagementPage({ showToast }) {
  const [users, setUsers]       = useState([]);
  const [loading, setLoading]   = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [editUser, setEditUser] = useState(null);
  const [form, setForm]         = useState({ username:"", full_name:"", password:"", role:"teknisi", is_active:true });

  const loadUsers = async () => {
    setLoading(true);
    const { data } = await supabase.from("users").select("id,username,full_name,role,is_active,created_at").order("created_at",{ascending:false});
    setUsers(data||[]);
    setLoading(false);
  };
  useEffect(()=>{ loadUsers(); },[]);

  const saveUser = async () => {
    if (!form.username.trim()||!form.full_name.trim()||(!editUser&&!form.password.trim())) {
      showToast("Lengkapi username, nama lengkap, dan password.","error"); return;
    }
    const payload = {
      username: form.username.trim().toLowerCase(),
      full_name: form.full_name.trim(),
      role: form.role,
      is_active: form.is_active,
    };
    if (form.password.trim()) payload.password_hash = hashPass(form.password);
    if (editUser) {
      const { error } = await supabase.from("users").update(payload).eq("id", editUser.id);
      if (error) { showToast("Gagal update: "+error.message,"error"); return; }
      showToast("User berhasil diupdate.");
    } else {
      const { error } = await supabase.from("users").insert(payload);
      if (error) { showToast("Gagal buat user: "+error.message,"error"); return; }
      showToast("User berhasil dibuat.");
    }
    setShowForm(false); setEditUser(null); setForm({ username:"",full_name:"",password:"",role:"teknisi",is_active:true });
    loadUsers();
  };

  const toggleActive = async (user) => {
    await supabase.from("users").update({ is_active: !user.is_active }).eq("id", user.id);
    loadUsers();
  };

  return (
    <div>
      <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12 }}>
        <div style={{ fontSize:14,fontWeight:600,color:"#111" }}>Kelola User</div>
        <button className="btn-primary" style={{ width:"auto",padding:"8px 16px",margin:0 }} onClick={()=>{ setShowForm(true); setEditUser(null); setForm({ username:"",full_name:"",password:"",role:"teknisi",is_active:true }); }}>
          <i className="ti ti-plus" style={{ marginRight:4 }}></i>Tambah User
        </button>
      </div>

      {showForm && (
        <div className="card" style={{ border:"1.5px solid #1D9E75" }}>
          <div className="card-title">{editUser?"Edit User":"Tambah User Baru"}</div>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">Username *</label>
              <input className="form-input" placeholder="cth: budi.s" value={form.username} onChange={e=>setForm(f=>({...f,username:e.target.value}))} disabled={!!editUser} />
            </div>
            <div className="form-group">
              <label className="form-label">Nama lengkap *</label>
              <input className="form-input" placeholder="cth: Budi Santoso" value={form.full_name} onChange={e=>setForm(f=>({...f,full_name:e.target.value}))} />
            </div>
          </div>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">Password {editUser?"(kosongkan jika tidak ganti)":""} *</label>
              <input className="form-input" type="password" placeholder="Password" value={form.password} onChange={e=>setForm(f=>({...f,password:e.target.value}))} />
            </div>
            <div className="form-group">
              <label className="form-label">Role *</label>
              <select className="form-input" value={form.role} onChange={e=>setForm(f=>({...f,role:e.target.value}))}>
                {ROLES.map(r=><option key={r} value={r}>{r.toUpperCase()}</option>)}
              </select>
            </div>
          </div>
          <div className="form-group" style={{ display:"flex",alignItems:"center",gap:10 }}>
            <input type="checkbox" id="is_active" checked={form.is_active} onChange={e=>setForm(f=>({...f,is_active:e.target.checked}))} style={{ width:16,height:16 }} />
            <label htmlFor="is_active" style={{ fontSize:13,color:"#333",cursor:"pointer" }}>Akun aktif</label>
          </div>
          <div style={{ display:"flex",gap:8 }}>
            <button className="btn-primary" style={{ margin:0 }} onClick={saveUser}>{editUser?"Update":"Simpan"}</button>
            <button className="btn-secondary" style={{ margin:0 }} onClick={()=>{setShowForm(false);setEditUser(null);}}>Batal</button>
          </div>
        </div>
      )}

      {loading && <div className="card" style={{ textAlign:"center",color:"#999",fontSize:13 }}>Memuat...</div>}
      {users.map(u=>(
        <div key={u.id} className="record-card">
          <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6 }}>
            <div>
              <div style={{ fontSize:15,fontWeight:700,color:"#111" }}>{u.full_name}</div>
              <div style={{ fontSize:12,color:"#666",marginTop:2 }}>@{u.username}</div>
            </div>
            <RoleBadge role={u.role} />
          </div>
          <div style={{ fontSize:11,color:u.is_active?"#1D9E75":"#E24B4A",marginBottom:8,fontWeight:500 }}>
            {u.is_active?"● Aktif":"● Nonaktif"}
          </div>
          <div className="record-actions">
            <button className="btn-sm" onClick={()=>{ setEditUser(u); setForm({ username:u.username,full_name:u.full_name,password:"",role:u.role,is_active:u.is_active }); setShowForm(true); }}>Edit</button>
            <button className={u.is_active?"btn-sm-danger":"btn-sm"} onClick={()=>toggleActive(u)}>
              {u.is_active?"Nonaktifkan":"Aktifkan"}
            </button>
          </div>
        </div>
      ))}
    </div>
  );
}

// ── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [page, setPage]       = useState("dashboard");
  const [records, setRecords] = useState([]);
  const [loading, setLoading] = useState(true);
  const [form, setForm]       = useState(emptyForm());
  const [editId, setEditId]   = useState(null);
  const [detailId, setDetailId] = useState(null);
  const [search, setSearch]   = useState("");
  const [filterProblem, setFilterProblem] = useState("");
  const [filterTech, setFilterTech]       = useState("");
  const [toast, setToast]     = useState(null);
  const [filterMonth, setFilterMonth]     = useState(new Date().toISOString().slice(0,7));
  const [naikTab, setNaikTab]             = useState("form");

  // persiapan state
  const [prepRecords, setPrepRecords] = useState([]);
  const [prepForm, setPrepForm]       = useState({ moldSize:"", moldSerial:"", slotLocation:"", operator:"", date:new Date().toISOString().slice(0,10), notes:"" });
  const [prepMode, setPrepMode]       = useState("out");
  const [prepSearch, setPrepSearch]   = useState("");

  // qc gate state
  const [qcRecords, setQcRecords]   = useState([]);
  const [qcMode, setQcMode]         = useState("entry");
  const [qcSearch, setQcSearch]     = useState("");
  const emptyQcForm = () => ({ moldSize:"", moldSerial:"", cavityCondition:"", defects:[], status:"ok", repairNotes:"", checker:"", jamMulai:"", jamSelesai:"", date:new Date().toISOString().slice(0,10) });
  const [qcForm, setQcForm]         = useState(emptyQcForm());

  const emptyNaikForm = () => ({ moldSizeNaik:"", moldSizeTurun:"", plant:"", line:"", machine:"", containerNoL:"", containerNoR:"", moldNoL:"", moldNoR:"", turunContainerNoL:"", turunContainerNoR:"", turunMoldNoL:"", turunMoldNoR:"", jamMulai:"", jamSelesai:"", operator:"", notes:"", date:new Date().toISOString().slice(0,10) });
  const [naikForm, setNaikForm]     = useState(emptyNaikForm());
  const [naikRecords, setNaikRecords] = useState([]);

  const role = currentUser?.role || "";
  const showToast = (msg, type="success") => { setToast({msg,type}); setTimeout(()=>setToast(null),2800); };

  // session persist
  useEffect(() => {
    const saved = sessionStorage.getItem("moldtrack_user");
    if (saved) { setCurrentUser(JSON.parse(saved)); }
  }, []);

  const handleLogin = (user) => {
    setCurrentUser(user);
    sessionStorage.setItem("moldtrack_user", JSON.stringify(user));
    if (user.role === "teknisi")        setPage("entry");
    else if (user.role === "persiapan") setPage("persiapan");
    else if (user.role === "qcgate")    setPage("qcgate");
    else if (user.role === "naik")      setPage("naik");
    else setPage("dashboard");
  };

  const handleLogout = () => {
    setCurrentUser(null);
    sessionStorage.removeItem("moldtrack_user");
    setPage("dashboard");
  };

  const loadRecords = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase.from("repair_records").select("*").order("created_at",{ascending:false});
    if (error) showToast("Gagal memuat data: "+error.message,"error");
    else setRecords(data||[]);
    setLoading(false);
  },[]);

  const loadPrepRecords = useCallback(async () => {
    const { data } = await supabase.from("preparation_records").select("*").order("created_at",{ascending:false});
    setPrepRecords(data||[]);
  },[]);

  const loadQcRecords = useCallback(async () => {
    const { data } = await supabase.from("qc_records").select("*").order("created_at",{ascending:false});
    setQcRecords(data||[]);
  },[]);

  const loadNaikRecords = useCallback(async () => {
    const { data } = await supabase.from("naik_mold_records").select("*").order("created_at",{ascending:false});
    setNaikRecords(data||[]);
  },[]);

  useEffect(()=>{ if(currentUser){ loadRecords(); loadPrepRecords(); loadQcRecords(); loadNaikRecords(); } },[loadRecords, loadPrepRecords, loadQcRecords, loadNaikRecords, currentUser]);

  useEffect(()=>{
    if (!currentUser) return;
    const ch1 = supabase.channel("rr").on("postgres_changes",{event:"*",schema:"public",table:"repair_records"},()=>loadRecords()).subscribe();
    const ch2 = supabase.channel("pr").on("postgres_changes",{event:"*",schema:"public",table:"preparation_records"},()=>loadPrepRecords()).subscribe();
    const ch3 = supabase.channel("qr").on("postgres_changes",{event:"*",schema:"public",table:"qc_records"},()=>loadQcRecords()).subscribe();
    const ch4 = supabase.channel("nr").on("postgres_changes",{event:"*",schema:"public",table:"naik_mold_records"},()=>loadNaikRecords()).subscribe();
    return ()=>{ supabase.removeChannel(ch1); supabase.removeChannel(ch2); supabase.removeChannel(ch3); supabase.removeChannel(ch4); };
  },[loadRecords, loadPrepRecords, loadQcRecords, loadNaikRecords, currentUser]);

  const submitNaik = async () => {
    if (!naikForm.moldSizeNaik.trim()||!naikForm.plant||!naikForm.line||!naikForm.machine||!naikForm.operator.trim()) {
      showToast("Lengkapi: size mold naik, mesin, dan operator.","error"); return;
    }
    const payload = {
      mold_size_naik:   naikForm.moldSizeNaik.trim().toUpperCase(),
      mold_size_turun:      naikForm.moldSizeTurun.trim().toUpperCase()||null,
      turun_container_no_l: naikForm.turunContainerNoL.trim().toUpperCase()||null,
      turun_container_no_r: naikForm.turunContainerNoR.trim().toUpperCase()||null,
      turun_mold_no_l:      naikForm.turunMoldNoL.trim().toUpperCase()||null,
      turun_mold_no_r:      naikForm.turunMoldNoR.trim().toUpperCase()||null,
      plant:            naikForm.plant,
      line:             naikForm.line,
      machine:          naikForm.machine,
      machine_code:     `${naikForm.plant}-${naikForm.line}${naikForm.machine}`,
      container_no_l:   naikForm.containerNoL.trim().toUpperCase()||null,
      container_no_r:   naikForm.containerNoR.trim().toUpperCase()||null,
      mold_no_l:        naikForm.moldNoL.trim().toUpperCase()||null,
      mold_no_r:        naikForm.moldNoR.trim().toUpperCase()||null,
      jam_mulai:        naikForm.jamMulai||null,
      jam_selesai:      naikForm.jamSelesai||null,
      operator:         naikForm.operator.trim(),
      date:             naikForm.date,
      notes:            naikForm.notes||null,
      created_by:       currentUser?.id,
      grup:             getGrup(currentUser?.username),
    };
    const { error } = await supabase.from("naik_mold_records").insert(payload);
    if (error) { showToast("Gagal simpan: "+error.message,"error"); return; }
    showToast("Naik mold berhasil dicatat! ✅");
    setNaikForm(emptyNaikForm());
  };

  const submitPrepOut = async () => {
    if (!prepForm.moldSize.trim()||!prepForm.moldSerial.trim()||!prepForm.slotLocation.trim()||!prepForm.operator.trim()) {
      showToast("Lengkapi: size mold, nomor seri, lokasi slot, dan operator.","error"); return;
    }
    const payload = {
      mold_size: prepForm.moldSize.trim().toUpperCase(),
      mold_serial: prepForm.moldSerial.trim().toUpperCase(),
      slot_location: prepForm.slotLocation.trim().toUpperCase(),
      operator: prepForm.operator.trim(),
      date: prepForm.date,
      notes: prepForm.notes,
      status: "keluar",
      created_by: currentUser?.id,
      grup: getGrup(currentUser?.username),
    };
    const { error } = await supabase.from("preparation_records").insert(payload);
    if (error) { showToast("Gagal simpan: "+error.message,"error"); return; }
    showToast("Mold keluar gudang berhasil dicatat.");
    setPrepForm({ moldSize:"", moldSerial:"", slotLocation:"", operator:"", date:new Date().toISOString().slice(0,10), notes:"" });
  };

  const submitPrepIn = async (id) => {
    const { error } = await supabase.from("preparation_records").update({ status:"kembali", returned_at: new Date().toISOString() }).eq("id", id);
    if (error) { showToast("Gagal update: "+error.message,"error"); return; }
    showToast("Mold berhasil dicatat kembali ke gudang.");
    loadPrepRecords();
  };

  const submitQc = async () => {
    if (!qcForm.moldSize.trim()||!qcForm.moldSerial.trim()||!qcForm.checker.trim()) {
      showToast("Lengkapi: size mold, nomor seri, dan nama checker.","error"); return;
    }
    const payload = {
      mold_size:        qcForm.moldSize.trim().toUpperCase(),
      mold_serial:      qcForm.moldSerial.trim().toUpperCase(),
      cavity_condition: qcForm.cavityCondition.trim(),
      defects:          qcForm.defects,
      status:           qcForm.status,
      repair_notes:     qcForm.repairNotes,
      checker:          qcForm.checker.trim(),
      jam_mulai:        qcForm.jamMulai,
      jam_selesai:      qcForm.jamSelesai,
      date:             qcForm.date,
      created_by:       currentUser?.id,
      grup: getGrup(currentUser?.username),
    };
    const { error } = await supabase.from("qc_records").insert(payload);
    if (error) { showToast("Gagal simpan: "+error.message,"error"); return; }
    showToast("Hasil QC berhasil disimpan.");
    setQcForm(emptyQcForm());
  };

  const knownSizes = useMemo(()=>[...new Set(records.map(r=>r.mold_size))].sort(),[records]);
  const knownTechs = useMemo(()=>[...new Set(records.map(r=>r.technician))].sort(),[records]);
  const machineCode = (f) => {
    if (!f.plant||!f.line||!f.machine||!f.press?.length) return "";
    const pressStr = f.press.sort().join("&");
    return `${f.plant}-${f.line}${f.machine}${pressStr}`;
  };

  const toggleProblem = (pid) => {
    setForm(f=>{
      const active=f.problems.includes(pid);
      const problems=active?f.problems.filter(x=>x!==pid):[...f.problems,pid];
      const pd={...f.problemDetails};
      if(active) delete pd[pid]; else pd[pid]=emptyProblemDetail(pid);
      return{...f,problems,problemDetails:pd};
    });
  };
  const updateDetail = (pid,detail) => setForm(f=>({...f,problemDetails:{...f.problemDetails,[pid]:detail}}));

  const submitForm = async () => {
    if(!form.moldSize.trim()||!form.technician.trim()||form.problems.length===0){
      showToast("Lengkapi: size mold, nama teknisi, dan minimal 1 problem.","error"); return;
    }
    if(!form.plant||!form.line||!form.machine||!form.press?.length){
      showToast("Lengkapi: plant, line, mesin, dan press.","error"); return;
    }
    const payload={
      mold_size:form.moldSize.trim().toUpperCase(),
      maker_container_l:form.makerContainerL, maker_container_r:form.makerContainerR,
      type_container_l:form.typeContainerL,   type_container_r:form.typeContainerR,
      container_no_l:form.containerNoL.trim().toUpperCase(),
      container_no_r:form.containerNoR.trim().toUpperCase(),
      mold_no_l:form.moldNoL.trim().toUpperCase(),
      mold_no_r:form.moldNoR.trim().toUpperCase(),
      date:form.date, jam_mulai:form.jamMulai, jam_selesai:form.jamSelesai,
      technician:form.technician.trim(),
      machine_code:machineCode(form), plant:form.plant, line:form.line, machine:form.machine, press:form.press,
      problems:form.problems, problem_details:form.problemDetails, notes:form.notes,
      created_by: currentUser?.id,
      grup: getGrup(currentUser?.username),
    };
    if(editId){
      const{error}=await supabase.from("repair_records").update(payload).eq("id",editId);
      if(error){showToast("Gagal update: "+error.message,"error");return;}
      showToast("Record berhasil diupdate.");
    }else{
      const{error}=await supabase.from("repair_records").insert(payload);
      if(error){showToast("Gagal simpan: "+error.message,"error");return;}
      showToast("Record berhasil disimpan.");
    }
    setForm(emptyForm());setEditId(null);setPage("database");
  };

  const startEdit = (rec) => {
    const pressVal = Array.isArray(rec.press) ? rec.press : (rec.press ? [rec.press] : []);
    setForm({moldSize:rec.mold_size,makerContainerL:rec.maker_container_l||"",makerContainerR:rec.maker_container_r||"",typeContainerL:rec.type_container_l||"",typeContainerR:rec.type_container_r||"",containerNoL:rec.container_no_l||"",containerNoR:rec.container_no_r||"",moldNoL:rec.mold_no_l||"",moldNoR:rec.mold_no_r||"",date:rec.date,jamMulai:rec.jam_mulai||"",jamSelesai:rec.jam_selesai||"",technician:rec.technician,plant:rec.plant||"",line:rec.line||"",machine:rec.machine||"",press:pressVal,problems:rec.problems||[],problemDetails:rec.problem_details||{},notes:rec.notes||""});
    setEditId(rec.id);setPage("entry");
  };

  const deleteRecord = async (id) => {
    if(!window.confirm("Hapus record ini?")) return;
    const{error}=await supabase.from("repair_records").delete().eq("id",id);
    if(error){showToast("Gagal hapus: "+error.message,"error");return;}
    showToast("Record dihapus.");
    if(detailId===id){setDetailId(null);setPage("database");}
  };

  const canEditDelete = (rec) => canDeleteEdit(role, rec.created_by, currentUser?.id);

  const filtered = useMemo(()=>records.filter(r=>{
    if(search&&!r.mold_size?.toLowerCase().includes(search.toLowerCase())&&!r.technician?.toLowerCase().includes(search.toLowerCase())&&!r.machine_code?.toLowerCase().includes(search.toLowerCase())) return false;
    if(filterProblem&&!(r.problems||[]).includes(filterProblem)) return false;
    if(filterTech&&r.technician!==filterTech) return false;
    return true;
  }),[records,search,filterProblem,filterTech]);

  const today=new Date().toISOString().slice(0,10);
  const todayRecs=records.filter(r=>r.date===today);
  const problemCounts=PROBLEMS.map(p=>({...p,count:records.filter(r=>(r.problems||[]).includes(p.id)).length})).sort((a,b)=>b.count-a.count);
  const detailRec=records.find(r=>r.id===detailId);
  const sameSize=detailRec?records.filter(r=>r.mold_size===detailRec.mold_size&&r.id!==detailRec.id).slice(0,5):[];

  const navTo = (p) => {
    setPage(p);
    if(p==="entry"){setForm(emptyForm());setEditId(null);}
  };

  // build nav items based on role
  const navItems = [
    ...(canDashboard(role)   ? [["dashboard","ti-layout-dashboard","Dashboard"]] : []),
    ...(canEntry(role)       ? [["entry","ti-plus","Action Problem"]] : []),
    ...(canDatabase(role)    ? [["database","ti-database","Database"]] : []),
    ...(canPersiapan(role)   ? [["persiapan","ti-package","Persiapan"]] : []),
    ...(canQCGate(role)      ? [["qcgate","ti-clipboard-check","QC Gate"]] : []),
    ...(canDashboard(role)||canQCGate(role)||canPersiapan(role) ? [
      ["rakit","ti-tools","Rakit Mold"],
    ] : []),
    ...(canNaik(role)        ? [["naik","ti-arrow-up","Naik Mold"]] : []),
    ...(canManageUsers(role) ? [["users","ti-users","Users"]] : []),
  ];

  if (!currentUser) return (
    <>
      <style>{CSS}</style>
      <LoginPage onLogin={handleLogin} />
    </>
  );

  if (loading) return (
    <>
      <style>{CSS}</style>
      <div style={{ display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",flexDirection:"column",gap:12,background:"#f5f5f5" }}>
        <div style={{ width:36,height:36,border:"3px solid #E1F5EE",borderTop:"3px solid #1D9E75",borderRadius:"50%",animation:"spin 0.8s linear infinite" }}></div>
        <div style={{ fontSize:13,color:"#999" }}>Memuat data...</div>
        <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      </div>
    </>
  );

  return (
    <>
      <style>{CSS}</style>
      {toast&&<div className={`toast ${toast.type}`}>{toast.msg}</div>}
      <div className="app">
        {/* SIDEBAR desktop */}
        <div className="sidebar">
          <div className="sidebar-logo">
            <div style={{ display:"flex",alignItems:"center",gap:8 }}>
              <div className="logo-dot"><svg width="14" height="14" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="6" stroke="white" strokeWidth="1.5"/><circle cx="8" cy="8" r="2" fill="white"/></svg></div>
              <div><div style={{ fontSize:13,fontWeight:600,color:"#111" }}>MoldTrack</div><div style={{ fontSize:10,color:"#999" }}>© by DMM</div></div>
            </div>
          </div>
          <div className="sidebar-nav">
            {navItems.map(([id,icon,lbl])=>(
              <div key={id} className={`sidebar-nav-item${page===id?" active":""}`} onClick={()=>navTo(id)}>
                <i className={`ti ${icon}`}></i><span>{lbl}</span>
              </div>
            ))}
          </div>
          <div className="sidebar-footer">
            <div style={{ marginBottom:6 }}>
              <div style={{ fontSize:12,fontWeight:500,color:"#111" }}>{currentUser.full_name}</div>
              <div style={{ marginTop:2 }}><RoleBadge role={role} /></div>
            </div>
            <button style={{ background:"none",border:"none",cursor:"pointer",color:"#E24B4A",fontSize:12,padding:0,display:"flex",alignItems:"center",gap:4 }} onClick={handleLogout}>
              <i className="ti ti-logout" style={{ fontSize:14 }}></i>Keluar
            </button>
          </div>
        </div>

        {/* MAIN */}
        <div className="main">
          <div className="topbar">
            <div className="topbar-logo">
              <div className="logo-dot"><svg width="14" height="14" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="6" stroke="white" strokeWidth="1.5"/><circle cx="8" cy="8" r="2" fill="white"/></svg></div>
              <div style={{ marginLeft:8 }}>
                <div className="topbar-title">
                  {page==="dashboard"&&"Dashboard"}
                  {page==="entry"&&(editId?"Edit Record":"Action Problem")}
                  {page==="database"&&"Database Record"}
                  {page==="detail"&&"Detail Record"}
                  {page==="users"&&"Kelola User"}
                  {page==="persiapan"&&"Persiapan Mold"}
                  {page==="qcgate"&&"QC Gate"}
                  {page==="rakit"&&"Rakit Mold"}
                  {page==="naik"&&"Naik Mold"}
                </div>
                <div className="topbar-sub">{currentUser.full_name} · <RoleBadge role={role} /></div>
              </div>
            </div>
            <div style={{ display:"flex",alignItems:"center",gap:8 }}>
              <button className="btn-sm" onClick={loadRecords} style={{ fontSize:11 }}>↻ Refresh</button>
              <button className="btn-sm" onClick={handleLogout} style={{ color:"#E24B4A",borderColor:"#E24B4A",fontSize:11 }}>Keluar</button>
            </div>
          </div>

          <div className="content">

            {/* DASHBOARD */}
            {page==="dashboard"&&canDashboard(role)&&(
              <div>
                {/* FILTER PERIODE */}
                {(()=>{
                  const now = new Date();
                  const thisMonth = now.toISOString().slice(0,7);
                  const filtered = filterMonth==="all" ? records : records.filter(r=>r.date?.startsWith(filterMonth));
                  const todayStr = now.toISOString().slice(0,10);
                  const todayFiltered = filtered.filter(r=>r.date===todayStr);

                  // Problem counts
                  const probCounts = PROBLEMS.map(p=>({
                    label:p.label, id:p.id,
                    count:filtered.filter(r=>(r.problems||[]).includes(p.id)).length
                  })).filter(p=>p.count>0).sort((a,b)=>b.count-a.count);

                  // Teknisi counts
                  const techCounts = [...new Set(filtered.map(r=>r.technician).filter(Boolean))]
                    .map(t=>({name:t, count:filtered.filter(r=>r.technician===t).length}))
                    .sort((a,b)=>b.count-a.count);

                  // Maker container counts
                  const makerCounts = {};
                  filtered.forEach(r=>{
                    if(r.maker_container_l) makerCounts[r.maker_container_l]=(makerCounts[r.maker_container_l]||0)+1;
                    if(r.maker_container_r) makerCounts[r.maker_container_r]=(makerCounts[r.maker_container_r]||0)+1;
                  });
                  const makerArr = Object.entries(makerCounts).map(([name,count])=>({name,count})).sort((a,b)=>b.count-a.count);

                  // Plant counts
                  const plantD = filtered.filter(r=>r.plant==="D").length;
                  const plantK = filtered.filter(r=>r.plant==="K").length;

                  // Tren 7 hari terakhir
                  const last7 = Array.from({length:7},(_,i)=>{
                    const d = new Date(now); d.setDate(d.getDate()-i);
                    const ds = d.toISOString().slice(0,10);
                    return {date:ds.slice(5), count:records.filter(r=>r.date===ds).length};
                  }).reverse();
                  const maxTren = Math.max(...last7.map(d=>d.count), 1);

                  // Avg durasi
                  const durations = filtered.map(r=>{
                    if(!r.jam_mulai||!r.jam_selesai) return null;
                    const [h1,m1]=r.jam_mulai.split(":").map(Number);
                    const [h2,m2]=r.jam_selesai.split(":").map(Number);
                    const d=(h2*60+m2)-(h1*60+m1);
                    return d>0?d:null;
                  }).filter(Boolean);
                  const avgDur = durations.length ? Math.round(durations.reduce((a,b)=>a+b,0)/durations.length) : 0;

                  return (
                    <div>
                      {/* FILTER BAR */}
                      <div style={{ display:"flex",gap:8,marginBottom:16,alignItems:"center",flexWrap:"wrap" }}>
                        <div style={{ fontSize:13,fontWeight:600,color:"#111",marginRight:4 }}>Periode:</div>
                        {[
                          [thisMonth,"Bulan ini"],
                          [new Date(now.getFullYear(),now.getMonth()-1).toISOString().slice(0,7),"Bulan lalu"],
                          ["all","Semua data"]
                        ].map(([val,lbl])=>(
                          <button key={val} onClick={()=>setFilterMonth(val)}
                            style={{ padding:"6px 14px",borderRadius:20,border:"1.5px solid",fontSize:12,cursor:"pointer",
                              borderColor:filterMonth===val?"#1D9E75":"#e0e0e0",
                              background:filterMonth===val?"#1D9E75":"#fff",
                              color:filterMonth===val?"#fff":"#666",fontWeight:filterMonth===val?600:400 }}>
                            {lbl}
                          </button>
                        ))}
                        <input type="month" value={filterMonth==="all"?"":filterMonth}
                          onChange={e=>setFilterMonth(e.target.value||"all")}
                          style={{ padding:"5px 10px",borderRadius:8,border:"1.5px solid #e0e0e0",fontSize:12,color:"#111" }}/>
                      </div>

                      {/* STAT CARDS */}
                      <div className="stat-grid" style={{ gridTemplateColumns:"repeat(4,1fr)",marginBottom:16 }}>
                        {[
                          {label:"Total Perbaikan",val:filtered.length,sub:filterMonth==="all"?"semua waktu":filterMonth,icon:"ti-tools"},
                          {label:"Hari ini",val:todayFiltered.length,sub:todayStr,icon:"ti-calendar-today"},
                          {label:"Rata-rata durasi",val:avgDur?`${avgDur} mnt`:"-",sub:"per perbaikan",icon:"ti-clock"},
                          {label:"Teknisi aktif",val:techCounts.length,sub:"orang",icon:"ti-users"},
                        ].map((c,i)=>(
                          <div key={i} className="stat-card" style={{ textAlign:"center" }}>
                            <i className={`ti ${c.icon}`} style={{ fontSize:22,color:"#1D9E75",marginBottom:6,display:"block" }}></i>
                            <div className="stat-label">{c.label}</div>
                            <div className="stat-val">{c.val}</div>
                            <div className="stat-sub">{c.sub}</div>
                          </div>
                        ))}
                      </div>

                      {/* TREN 7 HARI */}
                      <div className="card" style={{ marginBottom:16 }}>
                        <div className="card-title">📈 Tren Perbaikan 7 Hari Terakhir</div>
                        <div style={{ display:"flex",alignItems:"flex-end",gap:8,height:80,marginTop:8 }}>
                          {last7.map((d,i)=>(
                            <div key={i} style={{ flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:4 }}>
                              <div style={{ fontSize:11,fontWeight:600,color:"#1D9E75" }}>{d.count||""}</div>
                              <div style={{ width:"100%",background:d.count?"#1D9E75":"#e0e0e0",borderRadius:"4px 4px 0 0",
                                height:`${Math.round((d.count/maxTren)*60)+4}px`,minHeight:4,transition:"height 0.3s" }}></div>
                              <div style={{ fontSize:10,color:"#999" }}>{d.date}</div>
                            </div>
                          ))}
                        </div>
                      </div>

                      {/* CHARTS ROW 1 */}
                      <div className="dashboard-charts" style={{ marginBottom:16 }}>
                        {/* Problem distribution */}
                        <div className="card">
                          <div className="card-title">🔧 Distribusi Problem</div>
                          {probCounts.length===0&&<div style={{ fontSize:12,color:"#999" }}>Belum ada data.</div>}
                          {probCounts.map(p=>(
                            <div key={p.id} className="bar-row">
                              <div className="bar-row-header"><span style={{ fontSize:11 }}>{p.label}</span><span style={{ fontWeight:600,color:"#111" }}>{p.count}</span></div>
                              <div className="bar-track"><div className="bar-fill" style={{ width:`${Math.round((p.count/filtered.length)*100)}%` }}></div></div>
                            </div>
                          ))}
                        </div>

                        {/* Aktivitas teknisi */}
                        <div className="card">
                          <div className="card-title">👷 Aktivitas Teknisi</div>
                          {techCounts.length===0&&<div style={{ fontSize:12,color:"#999" }}>Belum ada data.</div>}
                          {techCounts.map(t=>(
                            <div key={t.name} className="bar-row">
                              <div className="bar-row-header"><span>{t.name}</span><span style={{ fontWeight:600,color:"#111" }}>{t.count}</span></div>
                              <div className="bar-track"><div className="bar-fill" style={{ width:`${Math.round((t.count/filtered.length)*100)}%` }}></div></div>
                            </div>
                          ))}
                        </div>
                      </div>

                      {/* CHARTS ROW 2 */}
                      <div className="dashboard-charts" style={{ marginBottom:16 }}>
                        {/* Maker container */}
                        <div className="card">
                          <div className="card-title">🏭 Maker Container</div>
                          {makerArr.length===0&&<div style={{ fontSize:12,color:"#999" }}>Belum ada data.</div>}
                          {makerArr.map(m=>(
                            <div key={m.name} className="bar-row">
                              <div className="bar-row-header"><span>{m.name}</span><span style={{ fontWeight:600,color:"#111" }}>{m.count}</span></div>
                              <div className="bar-track"><div className="bar-fill" style={{ width:`${Math.round((m.count/Math.max(...makerArr.map(x=>x.count)))*100)}%`,background:"#3B82F6" }}></div></div>
                            </div>
                          ))}
                        </div>

                        {/* Plant distribution */}
                        <div className="card">
                          <div className="card-title">🏗️ Distribusi Plant</div>
                          {filtered.length===0&&<div style={{ fontSize:12,color:"#999" }}>Belum ada data.</div>}
                          {[{name:"Plant D",count:plantD},{name:"Plant K",count:plantK}].map(p=>(
                            <div key={p.name} className="bar-row">
                              <div className="bar-row-header"><span>{p.name}</span><span style={{ fontWeight:600,color:"#111" }}>{p.count}</span></div>
                              <div className="bar-track"><div className="bar-fill" style={{ width:filtered.length?`${Math.round((p.count/filtered.length)*100)}%`:"0%",background:"#F59E0B" }}></div></div>
                            </div>
                          ))}
                          {/* Per line breakdown */}
                          <div style={{ marginTop:12,fontSize:11,color:"#999",fontWeight:600 }}>PER LINE</div>
                          {LINES.map(l=>{
                            const cnt=filtered.filter(r=>r.line===l).length;
                            if(!cnt) return null;
                            return(
                              <div key={l} className="bar-row">
                                <div className="bar-row-header"><span>Line {l}</span><span style={{ fontWeight:600,color:"#111" }}>{cnt}</span></div>
                                <div className="bar-track"><div className="bar-fill" style={{ width:`${Math.round((cnt/filtered.length)*100)}%`,background:"#F59E0B" }}></div></div>
                              </div>
                            );
                          })}
                        </div>
                      </div>

                      {/* PERBANDINGAN GRUP PER PROSES */}
                      {(()=>{
                        const grups = ["A","B","C","D"];
                        const colors = ["#1D9E75","#3B82F6","#F59E0B","#EF4444"];
                        const proses = [
                          {label:"Action Problem", data:filtered},
                          {label:"Persiapan Mold", data:prepRecords},
                          {label:"QC Gate",        data:qcRecords},
                          {label:"Naik Mold",      data:naikRecords},
                        ];
                        const grupTotals = grups.map(g=>({
                          name:`Grup ${g}`,
                          color:colors[grups.indexOf(g)],
                          counts:proses.map(p=>p.data.filter(r=>r.grup===g).length)
                        }));
                        const maxVal = Math.max(...grupTotals.flatMap(g=>g.counts),1);
                        return(
                          <div className="card" style={{ marginBottom:16 }}>
                            <div className="card-title">👥 Perbandingan Grup per Proses</div>
                            {/* Summary cards */}
                            <div style={{ display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:16 }}>
                              {grupTotals.map(g=>(
                                <div key={g.name} style={{ textAlign:"center",padding:"10px 6px",borderRadius:8,background:"#f8f8f8",border:`2px solid ${g.color}30` }}>
                                  <div style={{ fontSize:20,fontWeight:700,color:g.color }}>{g.counts.reduce((a,b)=>a+b,0)}</div>
                                  <div style={{ fontSize:11,color:"#666",marginTop:2 }}>{g.name}</div>
                                  <div style={{ fontSize:10,color:"#999",marginTop:4 }}>total semua proses</div>
                                </div>
                              ))}
                            </div>
                            {/* Per proses breakdown */}
                            {proses.map((p,pi)=>(
                              <div key={p.label} style={{ marginBottom:12 }}>
                                <div style={{ fontSize:11,fontWeight:600,color:"#999",marginBottom:6 }}>{p.label.toUpperCase()}</div>
                                {grupTotals.every(g=>g.counts[pi]===0)
                                  ? <div style={{ fontSize:11,color:"#ccc",fontStyle:"italic",paddingLeft:4 }}>Belum ada data</div>
                                  : grupTotals.map(g=>(
                                    <div key={g.name} className="bar-row">
                                      <div className="bar-row-header"><span style={{ fontSize:11 }}>{g.name}</span><span style={{ fontWeight:600,color:"#111" }}>{g.counts[pi]}</span></div>
                                      <div className="bar-track"><div className="bar-fill" style={{ width:`${Math.round((g.counts[pi]/maxVal)*100)}%`,background:g.color }}></div></div>
                                    </div>
                                  ))
                                }
                              </div>
                            ))}
                          </div>
                        );
                      })()}

                      {/* RECORD TERBARU */}
                      <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10 }}>
                        <div style={{ fontSize:13,fontWeight:600,color:"#111" }}>Record Terbaru</div>
                        <button className="btn-primary" style={{ width:"auto",padding:"8px 16px",margin:0 }} onClick={()=>navTo("entry")}>
                          <i className="ti ti-plus" style={{ marginRight:4 }}></i>Entry baru
                        </button>
                      </div>
                      {filtered.length===0&&<div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>Belum ada record.</div>}
                      {filtered.slice(0,5).map(r=>(
                        <div key={r.id} className="record-card" onClick={()=>{setDetailId(r.id);setPage("detail");}}>
                          <div className="record-card-header">
                            <div><div className="record-size">{r.mold_size}</div><div className="record-type">{r.maker_container_l||r.maker_container_r||"-"}</div></div>
                            <div style={{ fontSize:11,color:"#999",textAlign:"right" }}><div>{r.date}</div><div>{r.jam_mulai&&r.jam_selesai?`${r.jam_mulai}–${r.jam_selesai}`:""}</div></div>
                          </div>
                          {r.machine_code&&<div style={{ fontSize:12,fontWeight:600,color:"#1D9E75",marginBottom:6 }}><i className="ti ti-robot" style={{ fontSize:13,marginRight:4 }}></i>{r.machine_code}</div>}
                          <DetailSummary rec={{...r,problemDetails:r.problem_details,problems:r.problems||[]}}/>
                          <div className="record-meta"><i className="ti ti-user" style={{ fontSize:12,marginRight:4 }}></i>{r.technician}</div>
                        </div>
                      ))}
                    </div>
                  );
                })()}
              </div>
            )}

            {/* ENTRY FORM */}
            {page==="entry"&&(
              <div>
                <div className="card">
                  <div className="card-title">Identitas perbaikan</div>
                  <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:0 }}>
                    <AutocompleteInput label="Size mold *" value={form.moldSize} onChange={v=>setForm(f=>({...f,moldSize:v}))} suggestions={knownSizes} placeholder="cth: 205/65R15"/>
                    <div className="form-group">
                      <label className="form-label">Tanggal *</label>
                      <input type="date" className="form-input" value={form.date} onChange={e=>setForm(f=>({...f,date:e.target.value}))}/>
                    </div>
                  </div>
                  <AutocompleteInput label="Nama teknisi *" value={form.technician} onChange={v=>setForm(f=>({...f,technician:v}))} suggestions={knownTechs} placeholder="Ketik nama kamu..."/>
                  <div className="form-row">
                    <div className="form-group">
                      <label className="form-label">Maker Container L</label>
                      <select className="form-input" value={form.makerContainerL} onChange={e=>setForm(f=>({...f,makerContainerL:e.target.value}))}>
                        <option value="">— Pilih —</option>
                        {MAKER_CONTAINER.map(m=><option key={m} value={m}>{m}</option>)}
                      </select>
                    </div>
                    <div className="form-group">
                      <label className="form-label">Maker Container R</label>
                      <select className="form-input" value={form.makerContainerR} onChange={e=>setForm(f=>({...f,makerContainerR:e.target.value}))}>
                        <option value="">— Pilih —</option>
                        {MAKER_CONTAINER.map(m=><option key={m} value={m}>{m}</option>)}
                      </select>
                    </div>
                  </div>
                  <div className="form-row">
                    <div className="form-group">
                      <label className="form-label">Type Container L</label>
                      <select className="form-input" value={form.typeContainerL} onChange={e=>setForm(f=>({...f,typeContainerL:e.target.value}))}>
                        <option value="">— Pilih —</option>
                        {TYPE_CONTAINER.map(t=><option key={t} value={t}>{t}</option>)}
                      </select>
                    </div>
                    <div className="form-group">
                      <label className="form-label">Type Container R</label>
                      <select className="form-input" value={form.typeContainerR} onChange={e=>setForm(f=>({...f,typeContainerR:e.target.value}))}>
                        <option value="">— Pilih —</option>
                        {TYPE_CONTAINER.map(t=><option key={t} value={t}>{t}</option>)}
                      </select>
                    </div>
                  </div>
                  <div className="form-row">
                    <div className="form-group">
                      <label className="form-label">Nomor Container L</label>
                      <input className="form-input" placeholder="cth: C-001L" value={form.containerNoL} onChange={e=>setForm(f=>({...f,containerNoL:e.target.value}))} />
                    </div>
                    <div className="form-group">
                      <label className="form-label">Nomor Container R</label>
                      <input className="form-input" placeholder="cth: C-001R" value={form.containerNoR} onChange={e=>setForm(f=>({...f,containerNoR:e.target.value}))} />
                    </div>
                  </div>
                  <div className="form-row">
                    <div className="form-group">
                      <label className="form-label">Nomor Mold L</label>
                      <input className="form-input" placeholder="cth: ML-001" value={form.moldNoL} onChange={e=>setForm(f=>({...f,moldNoL:e.target.value}))} />
                    </div>
                    <div className="form-group">
                      <label className="form-label">Nomor Mold R</label>
                      <input className="form-input" placeholder="cth: MR-001" value={form.moldNoR} onChange={e=>setForm(f=>({...f,moldNoR:e.target.value}))} />
                    </div>
                  </div>
                  <div className="form-group">
                    <label className="form-label">Jam pengerjaan</label>
                    <div style={{ display:"flex",alignItems:"center",gap:8 }}>
                      <input type="time" className="form-input" style={{ flex:1 }} value={form.jamMulai} onChange={e=>setForm(f=>({...f,jamMulai:e.target.value}))}/>
                      <span style={{ fontSize:12,color:"#999",flexShrink:0 }}>–</span>
                      <input type="time" className="form-input" style={{ flex:1 }} value={form.jamSelesai} onChange={e=>setForm(f=>({...f,jamSelesai:e.target.value}))}/>
                    </div>
                    {form.jamMulai&&form.jamSelesai&&(()=>{
                      const[h1,m1]=form.jamMulai.split(":").map(Number);
                      const[h2,m2]=form.jamSelesai.split(":").map(Number);
                      const diff=(h2*60+m2)-(h1*60+m1);
                      if(diff>0) return <div style={{ fontSize:11,color:"#1D9E75",marginTop:4,fontWeight:500 }}>{Math.floor(diff/60)>0?`${Math.floor(diff/60)} jam `:""}{diff%60>0?`${diff%60} menit`:""}</div>;
                    })()}
                  </div>
                </div>

                <div className="card">
                  <div className="card-title">Mesin yang dikerjakan *</div>
                  <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10 }}>
                    <div className="form-group">
                      <label className="form-label">Plant *</label>
                      <select className="form-input" value={form.plant} onChange={e=>setForm(f=>({...f,plant:e.target.value,line:"",machine:"",press:[]}))}>
                        <option value="">— Pilih —</option>
                        {PLANTS.map(p=><option key={p} value={p}>Plant {p}</option>)}
                      </select>
                    </div>
                    <div className="form-group">
                      <label className="form-label">Line *</label>
                      <select className="form-input" value={form.line} onChange={e=>setForm(f=>({...f,line:e.target.value,machine:"",press:""}))} disabled={!form.plant}>
                        <option value="">— Pilih —</option>
                        {LINES.map(l=><option key={l} value={l}>Line {l}</option>)}
                      </select>
                    </div>
                    <div className="form-group">
                      <label className="form-label">Mesin *</label>
                      <select className="form-input" value={form.machine} onChange={e=>setForm(f=>({...f,machine:e.target.value,press:""}))} disabled={!form.line}>
                        <option value="">— Pilih —</option>
                        {getMachines(form.plant).map(m=><option key={m} value={m}>Mesin {m}</option>)}
                      </select>
                    </div>
                    <div className="form-group">
                      <label className="form-label">Press * <span style={{ fontSize:10,color:"#999",fontWeight:400 }}>(bisa pilih keduanya)</span></label>
                      <div style={{ display:"flex",gap:8 }}>
                        {PRESSES.map(p=>{
                          const selected = (form.press||[]).includes(p);
                          const togglePress = () => {
                            const cur = form.press||[];
                            setForm(f=>({...f, press: cur.includes(p) ? cur.filter(x=>x!==p) : [...cur,p]}));
                          };
                          return (
                            <button key={p} className={`type-btn${selected?" active":""}`} style={{ flex:1 }} onClick={togglePress} disabled={!form.machine}>
                              {p==="L"?"◀ Left (L)":"Right (R) ▶"}
                            </button>
                          );
                        })}
                      </div>
                      {(form.press||[]).length>0&&(
                        <div style={{ marginTop:6,fontSize:11,color:"#1D9E75",fontWeight:500 }}>
                          Dipilih: {form.press.join(" & ")}
                        </div>
                      )}
                    </div>
                  </div>
                  {machineCode(form)&&(
                    <div style={{ background:"#E1F5EE",borderRadius:8,padding:"10px 14px",display:"flex",alignItems:"center",gap:8 }}>
                      <i className="ti ti-robot" style={{ fontSize:18,color:"#1D9E75" }}></i>
                      <div>
                        <div style={{ fontSize:10,color:"#666",marginBottom:2 }}>Kode mesin</div>
                        <div style={{ fontSize:18,fontWeight:700,color:"#085041",letterSpacing:1 }}>{machineCode(form)}</div>
                      </div>
                    </div>
                  )}
                </div>

                <div className="card">
                  <div className="card-title">Jenis problem <span style={{ fontSize:11,color:"#999",fontWeight:400 }}>(bisa lebih dari 1)</span></div>
                  {["MOR","OOR","Overflow","OS"].map(grp=>(
                    <div key={grp}>
                      <div className="group-label">{grp}</div>
                      {PROBLEMS.filter(p=>p.group===grp).map(p=>{
                        const active=form.problems.includes(p.id);
                        const det=form.problemDetails[p.id];
                        return(
                          <div key={p.id} className={`prob-card${active?" active":""}`}>
                            <div className="prob-card-header" onClick={()=>toggleProblem(p.id)}>
                              <div className={`prob-check${active?" active":""}`}>
                                {active&&<i className="ti ti-check" style={{ fontSize:12,color:"#fff" }}></i>}
                              </div>
                              <span style={{ fontSize:13,color:active?"#085041":"#333",fontWeight:active?500:400 }}>{p.label}</span>
                            </div>
                            {active&&det&&(
                              <div className="prob-card-detail">
                                <ProblemDetailForm pid={p.id} detail={det} onChange={d=>updateDetail(p.id,d)}/>
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  ))}
                </div>

                <div className="card">
                  <div className="card-title">Catatan perbaikan</div>
                  <textarea className="form-input" rows={4} style={{ resize:"vertical" }} placeholder="Deskripsikan langkah perbaikan..." value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))}/>
                </div>

                <button className="btn-primary" onClick={submitForm}>{editId?"Update record":"Simpan record"}</button>
                <button className="btn-secondary" onClick={()=>{setForm(emptyForm());setEditId(null);}}>Reset</button>
                <button className="btn-secondary" onClick={()=>setPage(canDashboard(role)?"dashboard":"database")}>Batal</button>
              </div>
            )}

            {/* DATABASE */}
            {page==="database"&&canDatabase(role)&&(
              <div>
                <div className="search-wrap">
                  <i className="ti ti-search search-icon"></i>
                  <input className="search-input" placeholder="Cari size, teknisi, kode mesin..." value={search} onChange={e=>setSearch(e.target.value)}/>
                </div>
                <div className="filter-row">
                  <select className="filter-select" value={filterProblem} onChange={e=>setFilterProblem(e.target.value)}>
                    <option value="">Semua problem</option>
                    {PROBLEMS.map(p=><option key={p.id} value={p.id}>{p.label}</option>)}
                  </select>
                  <select className="filter-select" value={filterTech} onChange={e=>setFilterTech(e.target.value)}>
                    <option value="">Semua teknisi</option>
                    {knownTechs.map(t=><option key={t}>{t}</option>)}
                  </select>
                  <button className="btn-primary" style={{ width:"auto",padding:"8px 14px",margin:0,flexShrink:0 }} onClick={()=>navTo("entry")}>
                    + Entry
                  </button>
                  <button style={{ padding:"8px 14px",borderRadius:8,border:"1.5px solid #1D9E75",background:"#E1F5EE",color:"#085041",fontSize:12,fontWeight:600,cursor:"pointer",flexShrink:0,whiteSpace:"nowrap" }} onClick={()=>exportToExcel(records)}>
                    ⬇ Excel
                  </button>
                </div>
                <div style={{ fontSize:12,color:"#999",marginBottom:10 }}>{filtered.length} record ditemukan</div>
                {filtered.length===0&&<div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>Tidak ada record yang sesuai filter.</div>}
                {filtered.map(r=>(
                  <div key={r.id} className="record-card">
                    <div className="record-card-header" onClick={()=>{setDetailId(r.id);setPage("detail");}}>
                      <div><div className="record-size">{r.mold_size}</div><div className="record-type">{r.mold_type==="segmented"?"Segmented":"Two Piece"}</div></div>
                      <div style={{ fontSize:11,color:"#999",textAlign:"right" }}><div>{r.date}</div><div>{r.jam_mulai&&r.jam_selesai?`${r.jam_mulai}–${r.jam_selesai}`:""}</div></div>
                    </div>
                    {r.machine_code&&<div style={{ fontSize:12,fontWeight:600,color:"#1D9E75",marginBottom:6 }}><i className="ti ti-robot" style={{ fontSize:13,marginRight:4 }}></i>{r.machine_code}</div>}
                    <div style={{ marginBottom:6 }}><DetailSummary rec={{...r,problemDetails:r.problem_details,problems:r.problems||[]}}/></div>
                    <div className="record-meta" style={{ marginBottom:8 }}><i className="ti ti-user" style={{ fontSize:12,marginRight:4 }}></i>{r.technician}</div>
                    <div className="record-actions">
                      <button className="btn-sm" onClick={()=>{setDetailId(r.id);setPage("detail");}}>Detail</button>
                      {canEditDelete(r)&&<button className="btn-sm" onClick={()=>startEdit(r)}>Edit</button>}
                      {canEditDelete(r)&&<button className="btn-sm-danger" onClick={()=>deleteRecord(r.id)}>Hapus</button>}
                    </div>
                  </div>
                ))}
              </div>
            )}

            {/* DETAIL */}
            {page==="detail"&&detailRec&&(
              <div>
                <button className="btn-secondary" style={{ marginBottom:12 }} onClick={()=>setPage("database")}>
                  <i className="ti ti-arrow-left" style={{ marginRight:6 }}></i>Kembali
                </button>
                <div className="card">
                  <div style={{ marginBottom:12 }}>
                    <div style={{ fontSize:20,fontWeight:700,color:"#111" }}>{detailRec.mold_size}</div>
                    <div style={{ fontSize:11,color:"#999",marginTop:2 }}>{detailRec.mold_type==="segmented"?"Segmented (A–H)":"Two Piece"}</div>
                  </div>
                  <div className="stat-grid" style={{ marginBottom:12 }}>
                    <div className="stat-card"><div className="stat-label">Teknisi</div><div style={{ fontSize:13,fontWeight:600,color:"#111" }}>{detailRec.technician}</div></div>
                    <div className="stat-card"><div className="stat-label">Tanggal</div><div style={{ fontSize:13,fontWeight:600,color:"#111" }}>{detailRec.date}</div></div>
                    <div className="stat-card"><div className="stat-label">Jam mulai</div><div style={{ fontSize:13,fontWeight:600,color:"#111" }}>{detailRec.jam_mulai||"-"}</div></div>
                    <div className="stat-card"><div className="stat-label">Jam selesai</div><div style={{ fontSize:13,fontWeight:600,color:"#111" }}>{detailRec.jam_selesai||"-"}</div></div>
                  </div>
                  {detailRec.machine_code&&(
                    <div style={{ background:"#E1F5EE",borderRadius:8,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",gap:10 }}>
                      <i className="ti ti-robot" style={{ fontSize:22,color:"#1D9E75" }}></i>
                      <div>
                        <div style={{ fontSize:10,color:"#666",marginBottom:2 }}>Kode mesin</div>
                        <div style={{ fontSize:20,fontWeight:700,color:"#085041",letterSpacing:1 }}>{detailRec.machine_code}</div>
                        <div style={{ fontSize:11,color:"#666" }}>Plant {detailRec.plant} · Line {detailRec.line} · Mesin {detailRec.machine} · Press {detailRec.press}</div>
                      </div>
                    </div>
                  )}
                  <div style={{ fontSize:11,color:"#999",fontWeight:600,marginBottom:8 }}>DETAIL TINDAKAN</div>
                  {(detailRec.problems||[]).map(pid=>{
                    const det=(detailRec.problem_details||{})[pid];
                    const isMOR=pid.startsWith("MOR"),isOverflow=pid.startsWith("OVERFLOW"),isOS=pid==="OS",isOOR=pid==="OOR";
                    return(
                      <div key={pid} className="detail-section">
                        <div style={{ marginBottom:8 }}><ProblemBadge pid={pid}/></div>
                        {det&&(isMOR||isOverflow||isOS)&&(
                          <div style={{ fontSize:13 }}>
                            <div>Shim <strong>{det.shimAction}</strong> {det.shimJumlah||"—"} lembar{!isMOR&&det.shimLokasi?` — ${det.shimLokasi}`:""}</div>
                            {isMOR&&det.shimSectors?.length>0&&<div style={{ marginTop:4 }}>Sektor: {det.shimSectors.map(s=><span key={s} style={{ background:"#EAF3DE",color:"#27500A",fontSize:11,padding:"1px 8px",borderRadius:4,marginRight:4,fontWeight:500 }}>{s}</span>)}</div>}
                          </div>
                        )}
                        {det&&isOOR&&(
                          <div style={{ fontSize:13 }}>
                            <div>Tipe: <strong>{det.oorType==="segmented"?"Segmented":"Two Piece"}</strong></div>
                            {det.oorSectors?.length>0&&<div style={{ marginTop:4 }}>Sektor: {det.oorSectors.map(s=><span key={s} style={{ background:"#E1F5EE",color:"#085041",fontSize:11,padding:"1px 8px",borderRadius:4,marginRight:4,fontWeight:500 }}>{s}</span>)}</div>}
                          </div>
                        )}
                        {!det&&<div style={{ fontSize:12,color:"#999" }}>Tidak ada detail tindakan.</div>}
                      </div>
                    );
                  })}
                  {detailRec.notes&&(
                    <div style={{ marginTop:12 }}>
                      <div style={{ fontSize:11,color:"#999",fontWeight:600,marginBottom:6 }}>CATATAN</div>
                      <div style={{ fontSize:13,lineHeight:1.7,color:"#333" }}>{detailRec.notes}</div>
                    </div>
                  )}
                  {canEditDelete(detailRec)&&(
                    <div style={{ marginTop:16,display:"flex",gap:8 }}>
                      <button className="btn-primary" style={{ margin:0 }} onClick={()=>startEdit(detailRec)}>Edit</button>
                      <button className="btn-danger" onClick={()=>deleteRecord(detailRec.id)}>Hapus</button>
                    </div>
                  )}
                </div>
                {sameSize.length>0&&(
                  <div className="card">
                    <div className="card-title"><i className="ti ti-history" style={{ marginRight:6,color:"#999" }}></i>Riwayat size {detailRec.mold_size}</div>
                    {sameSize.map(r=>(
                      <div key={r.id} style={{ paddingBottom:10,marginBottom:10,borderBottom:"1px solid #f0f0f0",cursor:"pointer" }} onClick={()=>setDetailId(r.id)}>
                        <div style={{ fontSize:11,color:"#999",marginBottom:6 }}>{r.date}{r.jam_mulai?` · ${r.jam_mulai}–${r.jam_selesai}`:""}{r.machine_code?` · ${r.machine_code}`:""}</div>
                        <DetailSummary rec={{...r,problemDetails:r.problem_details,problems:r.problems||[]}}/>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            {/* PERSIAPAN */}
            {page==="persiapan"&&canPersiapan(role)&&(
              <div>
                {/* TAB: Keluar / Kembali / History */}
                <div style={{ display:"flex",gap:8,marginBottom:16 }}>
                  {[["out","📤 Catat Keluar"],["in","📥 Catat Kembali"],["history","📋 History"]].map(([m,lbl])=>(
                    <button key={m} onClick={()=>setPrepMode(m)}
                      style={{ flex:1,padding:"10px 8px",borderRadius:10,border:"1.5px solid",borderColor:prepMode===m?"#1D9E75":"#e0e0e0",background:prepMode===m?"#E1F5EE":"#fff",color:prepMode===m?"#085041":"#666",fontSize:12,fontWeight:prepMode===m?600:400,cursor:"pointer" }}>
                      {lbl}
                    </button>
                  ))}
                </div>

                {/* CATAT KELUAR */}
                {prepMode==="out"&&(
                  <div className="card">
                    <div className="card-title">📤 Catat Mold Keluar Gudang</div>
                    <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
                      <div className="form-group">
                        <label className="form-label">Size mold *</label>
                        <input className="form-input" placeholder="cth: 205/65R15" value={prepForm.moldSize} onChange={e=>setPrepForm(f=>({...f,moldSize:e.target.value}))} />
                      </div>
                      <div className="form-group">
                        <label className="form-label">Nomor seri / ID mold *</label>
                        <input className="form-input" placeholder="cth: M-20241001" value={prepForm.moldSerial} onChange={e=>setPrepForm(f=>({...f,moldSerial:e.target.value}))} />
                      </div>
                      <div className="form-group">
                        <label className="form-label">Line *</label>
                        <input className="form-input" placeholder="cth: A-3" value={prepForm.slotLocation} onChange={e=>setPrepForm(f=>({...f,slotLocation:e.target.value}))} />
                      </div>
                      <div className="form-group">
                        <label className="form-label">Tanggal *</label>
                        <input type="date" className="form-input" value={prepForm.date} onChange={e=>setPrepForm(f=>({...f,date:e.target.value}))} />
                      </div>
                    </div>
                    <div className="form-group">
                      <label className="form-label">Nama operator *</label>
                      <input className="form-input" placeholder="Nama operator persiapan" value={prepForm.operator} onChange={e=>setPrepForm(f=>({...f,operator:e.target.value}))} />
                    </div>
                    <div className="form-group">
                      <label className="form-label">Catatan (opsional)</label>
                      <textarea className="form-input" rows={2} style={{ resize:"vertical" }} placeholder="Catatan tambahan..." value={prepForm.notes} onChange={e=>setPrepForm(f=>({...f,notes:e.target.value}))} />
                    </div>
                    <button className="btn-primary" onClick={submitPrepOut}>Simpan — Mold Keluar Gudang</button>
                    <button className="btn-secondary" onClick={()=>setPrepForm({ moldSize:"",moldSerial:"",slotLocation:"",operator:"",date:new Date().toISOString().slice(0,10),notes:"" })}>Reset</button>
                  </div>
                )}

                {/* CATAT KEMBALI */}
                {prepMode==="in"&&(
                  <div>
                    <div style={{ position:"relative", marginBottom:10 }}>
                      <span style={{ position:"absolute", left:12, top:"50%", transform:"translateY(-50%)", color:"#999", fontSize:15 }}>🔍</span>
                      <input style={{ width:"100%", padding:"10px 12px 10px 36px", border:"1px solid #e0e0e0", borderRadius:8, fontSize:14, outline:"none", background:"#fff" }}
                        placeholder="Cari size mold atau nomor seri..."
                        value={prepSearch}
                        onChange={e=>setPrepSearch(e.target.value)}
                      />
                    </div>
                    {(()=>{
                      const keluarList = prepRecords.filter(r=>r.status==="keluar").filter(r=>
                        !prepSearch || r.mold_size?.toLowerCase().includes(prepSearch.toLowerCase()) || r.mold_serial?.toLowerCase().includes(prepSearch.toLowerCase())
                      );
                      return (
                        <>
                          <div style={{ fontSize:12,color:"#999",marginBottom:10 }}>{keluarList.length} mold sedang keluar gudang</div>
                          {keluarList.length===0&&(
                            <div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>
                              {prepSearch?"Tidak ada mold yang sesuai pencarian.":"Tidak ada mold yang sedang keluar gudang."}
                            </div>
                          )}
                          {keluarList.map(r=>(
                            <div key={r.id} className="record-card">
                              <div className="record-card-header">
                                <div>
                                  <div className="record-size">{r.mold_size}</div>
                                  <div className="record-type">SN: {r.mold_serial}</div>
                                </div>
                                <div style={{ textAlign:"right",fontSize:11,color:"#999" }}>
                                  <div>{r.date}</div>
                                  <div style={{ color:"#E8A020",fontWeight:600,marginTop:2 }}>● Keluar</div>
                                </div>
                              </div>
                              <div style={{ fontSize:11,color:"#666",marginBottom:8 }}>
                                <div>Line: <strong>{r.slot_location}</strong></div>
                                <div>Operator: {r.operator}</div>
                              </div>
                              <button className="btn-primary" style={{ margin:0 }} onClick={()=>submitPrepIn(r.id)}>
                                📥 Catat Kembali ke Gudang
                              </button>
                            </div>
                          ))}
                        </>
                      );
                    })()}
                  </div>
                )}

                {/* HISTORY */}
                {prepMode==="history"&&(
                  <div>
                    <div style={{ fontSize:12,color:"#999",marginBottom:10 }}>{prepRecords.length} total record persiapan</div>
                    {prepRecords.length===0&&(
                      <div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>Belum ada record persiapan.</div>
                    )}
                    {prepRecords.map(r=>(
                      <div key={r.id} className="record-card">
                        <div className="record-card-header">
                          <div>
                            <div className="record-size">{r.mold_size}</div>
                            <div className="record-type">SN: {r.mold_serial}</div>
                          </div>
                          <div style={{ textAlign:"right" }}>
                            <div style={{ fontSize:11,color:"#999" }}>{r.date}</div>
                            <div style={{ marginTop:4 }}>
                              <span style={{ fontSize:11,padding:"2px 10px",borderRadius:20,fontWeight:600, background:r.status==="keluar"?"#FEF3C7":"#E1F5EE", color:r.status==="keluar"?"#92400E":"#085041" }}>
                                {r.status==="keluar"?"● Keluar":"✓ Kembali"}
                              </span>
                            </div>
                          </div>
                        </div>
                        <div style={{ fontSize:11,color:"#666" }}>
                          <div>Line: <strong>{r.slot_location}</strong></div>
                          <div>Operator: {r.operator}</div>
                          {r.returned_at&&<div style={{ color:"#1D9E75",marginTop:2 }}>Kembali: {new Date(r.returned_at).toLocaleDateString("id-ID")}</div>}
                          {r.notes&&<div style={{ marginTop:4,color:"#999" }}>📝 {r.notes}</div>}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            {/* QC GATE */}
            {page==="qcgate"&&canQCGate(role)&&(
              <div>
                {/* TABS */}
                <div style={{ display:"flex",gap:8,marginBottom:16 }}>
                  {[["entry","🔍 Entry QC"],["history","📋 History QC"]].map(([m,lbl])=>(
                    <button key={m} onClick={()=>setQcMode(m)}
                      style={{ flex:1,padding:"10px 8px",borderRadius:10,border:"1.5px solid",borderColor:qcMode===m?"#1D9E75":"#e0e0e0",background:qcMode===m?"#E1F5EE":"#fff",color:qcMode===m?"#085041":"#666",fontSize:12,fontWeight:qcMode===m?600:400,cursor:"pointer" }}>
                      {lbl}
                    </button>
                  ))}
                </div>

                {/* ENTRY QC */}
                {qcMode==="entry"&&(
                  <div>
                    <div className="card">
                      <div className="card-title">🔍 Form QC Gate</div>
                      {/* Identitas mold */}
                      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
                        <div className="form-group">
                          <label className="form-label">Size mold *</label>
                          <input className="form-input" placeholder="cth: 205/65R15" value={qcForm.moldSize} onChange={e=>setQcForm(f=>({...f,moldSize:e.target.value}))} />
                        </div>
                        <div className="form-group">
                          <label className="form-label">Nomor seri / ID mold *</label>
                          <input className="form-input" placeholder="cth: M-20241001" value={qcForm.moldSerial} onChange={e=>setQcForm(f=>({...f,moldSerial:e.target.value}))} />
                        </div>
                        <div className="form-group">
                          <label className="form-label">Tanggal *</label>
                          <input type="date" className="form-input" value={qcForm.date} onChange={e=>setQcForm(f=>({...f,date:e.target.value}))} />
                        </div>
                        <div className="form-group">
                          <label className="form-label">Nama checker *</label>
                          <input className="form-input" placeholder="Nama checker QC" value={qcForm.checker} onChange={e=>setQcForm(f=>({...f,checker:e.target.value}))} />
                        </div>
                        <div className="form-group">
                          <label className="form-label">Jam mulai</label>
                          <input type="time" className="form-input" value={qcForm.jamMulai} onChange={e=>setQcForm(f=>({...f,jamMulai:e.target.value}))} />
                        </div>
                        <div className="form-group">
                          <label className="form-label">Jam selesai</label>
                          <input type="time" className="form-input" value={qcForm.jamSelesai} onChange={e=>setQcForm(f=>({...f,jamSelesai:e.target.value}))} />
                        </div>
                      </div>

                      {/* Kondisi cavity */}
                      <div className="form-group">
                        <label className="form-label">Kondisi cavity (deskripsi awal)</label>
                        <textarea className="form-input" rows={2} style={{ resize:"vertical" }} placeholder="Deskripsikan kondisi permukaan cavity mold..." value={qcForm.cavityCondition} onChange={e=>setQcForm(f=>({...f,cavityCondition:e.target.value}))} />
                      </div>

                      {/* Jenis cacat */}
                      <div className="form-group">
                        <label className="form-label">Jenis cacat yang ditemukan</label>
                        <div style={{ display:"flex",gap:8,flexWrap:"wrap" }}>
                          {["DMGM","Dirty Mold"].map(d=>{
                            const active=(qcForm.defects||[]).includes(d);
                            return(
                              <div key={d} onClick={()=>setQcForm(f=>({...f,defects:active?f.defects.filter(x=>x!==d):[...f.defects,d]}))}
                                style={{ padding:"8px 16px",borderRadius:8,border:`1.5px solid ${active?"#E24B4A":"#e0e0e0"}`,background:active?"#FCEBEB":"#f9f9f9",color:active?"#791F1F":"#666",fontSize:13,fontWeight:active?600:400,cursor:"pointer" }}>
                                {active?"✓ ":""}{d}
                              </div>
                            );
                          })}
                          <div onClick={()=>setQcForm(f=>({...f,defects:[]}))}
                            style={{ padding:"8px 16px",borderRadius:8,border:`1.5px solid ${qcForm.defects?.length===0?"#1D9E75":"#e0e0e0"}`,background:qcForm.defects?.length===0?"#E1F5EE":"#f9f9f9",color:qcForm.defects?.length===0?"#085041":"#666",fontSize:13,fontWeight:qcForm.defects?.length===0?600:400,cursor:"pointer" }}>
                            {qcForm.defects?.length===0?"✓ ":""}Tidak ada cacat
                          </div>
                        </div>
                      </div>

                      {/* Status hasil QC */}
                      <div className="form-group">
                        <label className="form-label">Status hasil QC *</label>
                        <div style={{ display:"flex",gap:8,flexWrap:"wrap" }}>
                          {[
                            ["ok","✅ OK — Langsung lolos","#1D9E75","#E1F5EE","#085041"],
                            ["minor","⚠️ Minor — Repair on-spot","#E8A020","#FEF3C7","#92400E"],
                            ["major","🚫 Major — Hold & area repair","#E24B4A","#FCEBEB","#791F1F"],
                          ].map(([val,lbl,border,bg,color])=>(
                            <div key={val} onClick={()=>setQcForm(f=>({...f,status:val}))}
                              style={{ flex:1,padding:"10px 8px",borderRadius:8,border:`1.5px solid ${qcForm.status===val?border:"#e0e0e0"}`,background:qcForm.status===val?bg:"#f9f9f9",color:qcForm.status===val?color:"#666",fontSize:12,fontWeight:qcForm.status===val?600:400,cursor:"pointer",textAlign:"center",minWidth:100 }}>
                              {lbl}
                            </div>
                          ))}
                        </div>
                      </div>

                      {/* Repair notes — muncul jika minor atau major */}
                      {(qcForm.status==="minor"||qcForm.status==="major")&&(
                        <div className="form-group">
                          <label className="form-label">{qcForm.status==="minor"?"Detail repair on-spot":"Keterangan hold / alasan major"}</label>
                          <textarea className="form-input" rows={3} style={{ resize:"vertical" }} placeholder={qcForm.status==="minor"?"Jelaskan repair yang dilakukan...":"Jelaskan kondisi cacat dan tindakan selanjutnya..."} value={qcForm.repairNotes} onChange={e=>setQcForm(f=>({...f,repairNotes:e.target.value}))} />
                        </div>
                      )}

                      <button className="btn-primary" onClick={submitQc}>Simpan Hasil QC</button>
                      <button className="btn-secondary" onClick={()=>setQcForm(emptyQcForm())}>Reset</button>
                    </div>
                  </div>
                )}

                {/* HISTORY QC */}
                {qcMode==="history"&&(
                  <div>
                    <div style={{ position:"relative",marginBottom:10 }}>
                      <span style={{ position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",color:"#999",fontSize:15 }}>🔍</span>
                      <input style={{ width:"100%",padding:"10px 12px 10px 36px",border:"1px solid #e0e0e0",borderRadius:8,fontSize:14,outline:"none",background:"#fff" }}
                        placeholder="Cari size mold atau nomor seri..."
                        value={qcSearch} onChange={e=>setQcSearch(e.target.value)}
                      />
                    </div>
                    {(()=>{
                      const list = qcRecords.filter(r=>!qcSearch||r.mold_size?.toLowerCase().includes(qcSearch.toLowerCase())||r.mold_serial?.toLowerCase().includes(qcSearch.toLowerCase()));
                      const statusCfg = {
                        ok:    { label:"✅ OK",    bg:"#E1F5EE", color:"#085041" },
                        minor: { label:"⚠️ Minor", bg:"#FEF3C7", color:"#92400E" },
                        major: { label:"🚫 Major", bg:"#FCEBEB", color:"#791F1F" },
                      };
                      return (
                        <>
                          <div style={{ fontSize:12,color:"#999",marginBottom:10 }}>{list.length} record QC</div>
                          {list.length===0&&<div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>Belum ada record QC.</div>}
                          {list.map(r=>{
                            const cfg = statusCfg[r.status]||statusCfg.ok;
                            return(
                              <div key={r.id} className="record-card">
                                <div className="record-card-header">
                                  <div>
                                    <div className="record-size">{r.mold_size}</div>
                                    <div className="record-type">SN: {r.mold_serial}</div>
                                  </div>
                                  <div style={{ textAlign:"right" }}>
                                    <div style={{ fontSize:11,color:"#999" }}>{r.date}</div>
                                    <div style={{ marginTop:4 }}>
                                      <span style={{ fontSize:11,padding:"2px 10px",borderRadius:20,fontWeight:600,background:cfg.bg,color:cfg.color }}>{cfg.label}</span>
                                    </div>
                                  </div>
                                </div>
                                <div style={{ fontSize:11,color:"#666",marginTop:4 }}>
                                  <div>Checker: <strong>{r.checker}</strong></div>
                                  {r.jam_mulai&&<div>Jam: {r.jam_mulai}–{r.jam_selesai}</div>}
                                  {r.defects?.length>0&&<div style={{ marginTop:4 }}>Cacat: {r.defects.map(d=><span key={d} style={{ background:"#FCEBEB",color:"#791F1F",fontSize:11,padding:"1px 8px",borderRadius:4,marginRight:4,fontWeight:500 }}>{d}</span>)}</div>}
                                  {r.cavity_condition&&<div style={{ marginTop:4,color:"#999" }}>Cavity: {r.cavity_condition}</div>}
                                  {r.repair_notes&&<div style={{ marginTop:4,color:"#999" }}>📝 {r.repair_notes}</div>}
                                </div>
                              </div>
                            );
                          })}
                        </>
                      );
                    })()}
                  </div>
                )}
              </div>
            )}

            {/* RAKIT MOLD - Under Development */}
            {page==="rakit"&&(
              <div style={{ display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minHeight:400,gap:16 }}>
                <div style={{ fontSize:56 }}>🔧</div>
                <div style={{ fontSize:18,fontWeight:700,color:"#111" }}>Rakit Mold</div>
                <div style={{ background:"#FEF3C7",color:"#92400E",padding:"8px 20px",borderRadius:20,fontSize:13,fontWeight:600 }}>🚧 Under Development</div>
                <div style={{ fontSize:13,color:"#999",textAlign:"center",maxWidth:280 }}>Fitur ini sedang dalam tahap pengembangan. Segera hadir!</div>
              </div>
            )}

            {/* NAIK MOLD */}
            {page==="naik"&&canNaik(role)&&(
              <div>
                {/* TAB NAVIGATION */}
                <div style={{ display:"flex",gap:8,marginBottom:16 }}>
                  {[["form","⬆️ Input"],["database","📋 Database"]].map(([id,lbl])=>(
                    <button key={id} onClick={()=>setNaikTab(id)}
                      style={{ flex:1,padding:"10px",borderRadius:8,border:"1.5px solid",cursor:"pointer",fontSize:13,fontWeight:600,
                        borderColor:naikTab===id?"#1D9E75":"#e0e0e0",
                        background:naikTab===id?"#1D9E75":"#fff",
                        color:naikTab===id?"#fff":"#666" }}>
                      {lbl}
                    </button>
                  ))}
                </div>

                {/* FORM TAB */}
                {naikTab==="form"&&(
                  <div>
                  <div className="card-title" style={{ textAlign:"center",fontSize:16 }}>⬆️ Naik Mold</div>
                  <div className="form-group">
                    <label className="form-label">Tanggal *</label>
                    <input type="date" className="form-input" value={naikForm.date} onChange={e=>setNaikForm(f=>({...f,date:e.target.value}))} />
                  </div>

                  {/* PLANT LINE MESIN */}
                  <div className="form-row">
                    <div className="form-group">
                      <label className="form-label">Plant *</label>
                      <select className="form-input" value={naikForm.plant} onChange={e=>setNaikForm(f=>({...f,plant:e.target.value,line:"",machine:""}))}>
                        <option value="">— Pilih —</option>
                        {PLANTS.map(p=><option key={p} value={p}>Plant {p}</option>)}
                      </select>
                    </div>
                    <div className="form-group">
                      <label className="form-label">Line *</label>
                      <select className="form-input" value={naikForm.line} onChange={e=>setNaikForm(f=>({...f,line:e.target.value,machine:""}))}>
                        <option value="">— Pilih —</option>
                        {LINES.map(l=><option key={l} value={l}>Line {l}</option>)}
                      </select>
                    </div>
                  </div>
                  <div className="form-group">
                    <label className="form-label">Mesin *</label>
                    <select className="form-input" value={naikForm.machine} onChange={e=>setNaikForm(f=>({...f,machine:e.target.value}))}>
                      <option value="">— Pilih —</option>
                      {(naikForm.plant==="D"?MACHINES_D:naikForm.plant==="K"?MACHINES_K:[]).map(m=><option key={m} value={m}>{m}</option>)}
                    </select>
                  </div>

                  {/* SIZE MOLD */}
                  <div className="form-row">
                    <div className="form-group">
                      <label className="form-label">Size Mold Naik *</label>
                      <input className="form-input" placeholder="cth: 205/65R15" value={naikForm.moldSizeNaik} onChange={e=>setNaikForm(f=>({...f,moldSizeNaik:e.target.value}))} />
                    </div>
                    <div className="form-group">
                      <label className="form-label">Size Mold Turun</label>
                      <input className="form-input" placeholder="cth: 195/65R15 (jika ada)" value={naikForm.moldSizeTurun} onChange={e=>setNaikForm(f=>({...f,moldSizeTurun:e.target.value}))} />
                    </div>
                  </div>

                  {/* NO MOLD */}
                  <div style={{ fontSize:11,color:"#999",fontWeight:600,marginBottom:6,marginTop:4 }}>NO MOLD</div>
                  <div className="form-row">
                    <div className="form-group">
                      <label className="form-label">No Mold L (Naik)</label>
                      <input className="form-input" placeholder="cth: ML-001" value={naikForm.moldNoL} onChange={e=>setNaikForm(f=>({...f,moldNoL:e.target.value}))} />
                    </div>
                    <div className="form-group">
                      <label className="form-label">No Mold R (Naik)</label>
                      <input className="form-input" placeholder="cth: MR-001" value={naikForm.moldNoR} onChange={e=>setNaikForm(f=>({...f,moldNoR:e.target.value}))} />
                    </div>
                  </div>
                  {naikForm.moldSizeTurun&&(
                    <div className="form-row">
                      <div className="form-group">
                        <label className="form-label">No Mold L (Turun)</label>
                        <input className="form-input" placeholder="cth: ML-001" value={naikForm.turunMoldNoL} onChange={e=>setNaikForm(f=>({...f,turunMoldNoL:e.target.value}))} />
                      </div>
                      <div className="form-group">
                        <label className="form-label">No Mold R (Turun)</label>
                        <input className="form-input" placeholder="cth: MR-001" value={naikForm.turunMoldNoR} onChange={e=>setNaikForm(f=>({...f,turunMoldNoR:e.target.value}))} />
                      </div>
                    </div>
                  )}

                  {/* NO CONTAINER */}
                  <div style={{ fontSize:11,color:"#999",fontWeight:600,marginBottom:6,marginTop:4 }}>NO CONTAINER</div>
                  <div className="form-row">
                    <div className="form-group">
                      <label className="form-label">No Container L (Naik)</label>
                      <input className="form-input" placeholder="cth: C-001L" value={naikForm.containerNoL} onChange={e=>setNaikForm(f=>({...f,containerNoL:e.target.value}))} />
                    </div>
                    <div className="form-group">
                      <label className="form-label">No Container R (Naik)</label>
                      <input className="form-input" placeholder="cth: C-001R" value={naikForm.containerNoR} onChange={e=>setNaikForm(f=>({...f,containerNoR:e.target.value}))} />
                    </div>
                  </div>
                  {naikForm.moldSizeTurun&&(
                    <div className="form-row">
                      <div className="form-group">
                        <label className="form-label">No Container L (Turun)</label>
                        <input className="form-input" placeholder="cth: C-001L" value={naikForm.turunContainerNoL} onChange={e=>setNaikForm(f=>({...f,turunContainerNoL:e.target.value}))} />
                      </div>
                      <div className="form-group">
                        <label className="form-label">No Container R (Turun)</label>
                        <input className="form-input" placeholder="cth: C-001R" value={naikForm.turunContainerNoR} onChange={e=>setNaikForm(f=>({...f,turunContainerNoR:e.target.value}))} />
                      </div>
                    </div>
                  )}

                  {/* JAM */}
                  <div className="form-group" style={{ marginTop:4 }}>
                    <label className="form-label">Jam pengerjaan</label>
                    <div style={{ display:"flex",gap:8,alignItems:"center" }}>
                      <input type="time" className="form-input" value={naikForm.jamMulai} onChange={e=>setNaikForm(f=>({...f,jamMulai:e.target.value}))} />
                      <span style={{ color:"#999" }}>–</span>
                      <input type="time" className="form-input" value={naikForm.jamSelesai} onChange={e=>setNaikForm(f=>({...f,jamSelesai:e.target.value}))} />
                    </div>
                  </div>

                  {/* PIC */}
                  <AutocompleteInput label="PIC / Operator *" value={naikForm.operator} onChange={v=>setNaikForm(f=>({...f,operator:v}))} suggestions={knownTechs} placeholder="Ketik nama PIC..."/>

                  {/* CATATAN */}
                  <div className="form-group">
                    <label className="form-label">Catatan</label>
                    <textarea className="form-input" rows={3} placeholder="Catatan tambahan..." value={naikForm.notes||""} onChange={e=>setNaikForm(f=>({...f,notes:e.target.value}))} style={{ resize:"vertical" }}/>
                  </div>

                  <button className="btn-primary" onClick={submitNaik}>⬆️ Simpan Naik Mold</button>
                  <button className="btn-secondary" style={{ marginTop:8 }} onClick={()=>setNaikForm(emptyNaikForm())}>Reset</button>
                  </div>
                )}

                {/* DATABASE TAB */}
                {naikTab==="database"&&(
                  <div>
                    <div style={{ marginBottom:10,fontSize:13,color:"#999" }}>{naikRecords.length} record ditemukan</div>
                    {naikRecords.length===0&&<div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>Belum ada record naik mold.</div>}
                    {naikRecords.map(r=>(
                      <div key={r.id} className="record-card">
                        <div className="record-card-header">
                          <div>
                            <div className="record-size">⬆️ {r.mold_size_naik}</div>
                            {r.mold_size_turun&&<div className="record-type">⬇️ Turun: {r.mold_size_turun}</div>}
                          </div>
                          <div style={{ fontSize:11,color:"#999",textAlign:"right" }}>
                            <div>{r.date}</div>
                            <div>{r.jam_mulai&&r.jam_selesai?`${r.jam_mulai}–${r.jam_selesai}`:""}</div>
                          </div>
                        </div>
                        {r.machine_code&&<div style={{ fontSize:12,fontWeight:600,color:"#1D9E75",marginBottom:6 }}><i className="ti ti-robot" style={{ fontSize:13,marginRight:4 }}></i>{r.machine_code}</div>}
                        {(r.container_no_l||r.container_no_r)&&(
                          <div style={{ fontSize:11,color:"#666",marginBottom:4,display:"flex",gap:12,flexWrap:"wrap" }}>
                            {r.container_no_l&&<span>Container L(↑): <strong>{r.container_no_l}</strong></span>}
                            {r.container_no_r&&<span>Container R(↑): <strong>{r.container_no_r}</strong></span>}
                            {r.turun_container_no_l&&<span>Container L(↓): <strong>{r.turun_container_no_l}</strong></span>}
                            {r.turun_container_no_r&&<span>Container R(↓): <strong>{r.turun_container_no_r}</strong></span>}
                          </div>
                        )}
                        {(r.mold_no_l||r.mold_no_r)&&(
                          <div style={{ fontSize:11,color:"#666",marginBottom:4,display:"flex",gap:12,flexWrap:"wrap" }}>
                            {r.mold_no_l&&<span>Mold L(↑): <strong>{r.mold_no_l}</strong></span>}
                            {r.mold_no_r&&<span>Mold R(↑): <strong>{r.mold_no_r}</strong></span>}
                            {r.turun_mold_no_l&&<span>Mold L(↓): <strong>{r.turun_mold_no_l}</strong></span>}
                            {r.turun_mold_no_r&&<span>Mold R(↓): <strong>{r.turun_mold_no_r}</strong></span>}
                          </div>
                        )}
                        {r.notes&&<div style={{ fontSize:11,color:"#666",marginBottom:4 }}>📝 {r.notes}</div>}
                        <div className="record-meta"><i className="ti ti-user" style={{ fontSize:12,marginRight:4 }}></i>{r.operator}</div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            {/* USER MANAGEMENT */}
            {page==="users"&&canManageUsers(role)&&(
              <UserManagementPage showToast={showToast} />
            )}

          </div>
        </div>

        {/* BOTTOM NAV mobile */}
        <nav className="bottom-nav">
          {navItems.map(([id,icon,lbl])=>(
            <button key={id} className={`bottom-nav-item${page===id?" active":""}`} onClick={()=>navTo(id)}>
              <i className={`ti ${icon}`}></i><span>{lbl}</span>
            </button>
          ))}
        </nav>
      </div>
    </>
  );
}
