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

const ROLES    = ["teknisi","persiapan","qcgate","rakit","naik","sh","analyst","adh","dh","admin"];

// Nomor WA per grup — ganti dengan nomor aktual
const WA_GROUPS = {
  "a": { persiapan:"6281234560001", qcgate:"6281234560002", rakit:"6281234560003", naik:"6281234560004", sh:"6281234560005" },
  "b": { persiapan:"6281234560011", qcgate:"6281234560012", rakit:"6281234560013", naik:"6281234560014", sh:"6281234560015" },
  "c": { persiapan:"6281234560021", qcgate:"6281234560022", rakit:"6281234560023", naik:"6281234560024", sh:"6281234560025" },
  "d": { persiapan:"6281234560031", qcgate:"6281234560032", rakit:"6281234560033", naik:"6281234560034", sh:"6281234560035" },
};

// Kirim notif WA — simpan ke antrian notif untuk ditampilkan ke user
const sendWA = (phone, msg) => {
  const url = `https://wa.me/${phone}?text=${encodeURIComponent(msg)}`;
  // Buka langsung — user harus allow popup, atau gunakan link manual
  const a = document.createElement("a");
  a.href = url; a.target = "_blank"; a.rel = "noopener noreferrer";
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
};

// Helper notifikasi per tahap
const notifPersiapan = (grup, moldSize, shift) => {
  const no = WA_GROUPS[grup]?.persiapan;
  if (!no) return;
  sendWA(no, "[MoldTrack] PERSIAPAN MOLD - Size: "+moldSize+" - Shift: "+shift+" - Mohon segera persiapan mold.");
};
const notifQCGate1 = (grup, moldSize) => {
  const no = WA_GROUPS[grup]?.qcgate;
  if (!no) return;
  sendWA(no, "[MoldTrack] QC GATE 1 - Size: "+moldSize+" - Mold selesai persiapan. Mohon cek visual.");
};
const notifRakit = (grup, moldSize) => {
  const no = WA_GROUPS[grup]?.rakit;
  if (!no) return;
  sendWA(no, "[MoldTrack] RAKIT MOLD - Size: "+moldSize+" - Lulus QC Gate 1. Mohon segera rakit mold.");
};
const notifQCGate2 = (grup, moldSize) => {
  const no = WA_GROUPS[grup]?.qcgate;
  if (!no) return;
  sendWA(no, "[MoldTrack] QC GATE 2 KALIBRASI - Size: "+moldSize+" - Mold sudah dirakit. Mohon kalibrasi (MOR/OOR/OF/OS).");
};
const notifNaik = (grup, moldSize) => {
  const no = WA_GROUPS[grup]?.naik;
  if (!no) return;
  sendWA(no, "[MoldTrack] NAIK MOLD - Size: "+moldSize+" - Lulus QC Gate 2. Siap dipasang ke mesin curing.");
};
const notifSH = (grup, moldSize, tahap, catatan) => {
  const no = WA_GROUPS[grup]?.sh;
  if (!no) return;
  sendWA(no, "[MoldTrack] INFO SECTION HEAD - Size: "+moldSize+" - Tahap: "+tahap+" - HOLD - "+catatan);
};

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
const canDatabase    = (r) => ["analyst","adh","dh","admin","teknisi","qcgate","naik","rakit","persiapan"].includes(r);
const canPersiapan   = (r) => ["persiapan","analyst","adh","dh","admin"].includes(r);
const canQCGate      = (r) => ["qcgate","analyst","adh","dh","admin"].includes(r);
const canRakit        = (r) => ["rakit","analyst","adh","dh","admin"].includes(r);
const canNaik        = (r) => ["naik","analyst","adh","dh","admin"].includes(r);
const canSH          = (r) => ["sh","analyst","adh","dh","admin"].includes(r);
  const canAnalyst     = (r) => ["analyst","admin","dh","adh"].includes(r);
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
  const [currentUser, setCurrentUser] = useState(() => {
    try { const s=localStorage.getItem("moldtrack_user"); return s?JSON.parse(s):null; } catch { return null; }
  });
  const [page, setPage] = useState(() => {
    try {
      const s = localStorage.getItem("moldtrack_user");
      if (!s) return "dashboard";
      const u = JSON.parse(s);
      if (u.role==="teknisi")   return "entry";
      if (u.role==="persiapan") return "persiapan";
      if (u.role==="qcgate")    return "qcgate";
      if (u.role==="naik")      return "naik";
      if (u.role==="rakit")     return "rakit";
      if (u.role==="sh")        return "shplan";
      return "dashboard";
    } catch { return "dashboard"; }
  });
  const [records, setRecords] = useState([]);
  const [loading, setLoading] = useState(false);
  const [form, setForm]       = useState(emptyForm());
  const [editId, setEditId]   = useState(null);
  const [detailId, setDetailId] = useState(null);
  const [search, setSearch]   = useState("");
  const [filterProblem, setFilterProblem] = useState("");
  const [filterTech, setFilterTech]       = useState("");
  const [toast, setToast]     = useState(null);
  const [filterMonth, setFilterMonth]     = useState(new Date().toISOString().slice(0,7));
  const [entryTab, setEntryTab]           = useState("form");
  const [dbTab, setDbTab]                 = useState("action");
  const [shTab, setShTab]                 = useState("form");
  const [shiftPlanRecords, setShiftPlanRecords] = useState([]);
  const emptyShiftPlan = () => ({ shift:"1", moldSizes:[""], date:new Date().toISOString().slice(0,10), catatan:"" });
  const [shiftPlan, setShiftPlan]         = useState(emptyShiftPlan());
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
  const emptyQcForm = () => ({ moldSize:"", moldSerial:"", cavityCondition:"", defects:[], status:"ok", repairNotes:"", checker:"",  date:new Date().toISOString().slice(0,10) });
  const [qcForm, setQcForm]         = useState(emptyQcForm());

  const emptyNaikForm = () => ({ moldSizeNaik:"", moldSizeTurun:"", plant:"", line:"", machine:"", press:[], makerContainerL:"", makerContainerR:"", typeContainerL:"", typeContainerR:"", containerNoL:"", containerNoR:"", moldNoL:"", moldNoR:"", turunMakerContainerL:"", turunMakerContainerR:"", turunTypeContainerL:"", turunTypeContainerR:"", turunContainerNoL:"", turunContainerNoR:"", turunMoldNoL:"", turunMoldNoR:"", operator:"", notes:"", date:new Date().toISOString().slice(0,10) });
  const [naikForm, setNaikForm]     = useState(emptyNaikForm());
  const [naikRecords, setNaikRecords] = useState([]);

  // rakit state
  const RAKIT_DATA = [{"tgl":"7-Oct","pic":"SIGIT - SATRIA","code":"AS138 - 03","size":"AS138","container":"","ket":"START","mc":"K08R","nilai_preload":"1..46","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-29","pic":"AGUS - UBAY","code":"XM380 - 07","size":"XM380","container":"","ket":"CLEANING","mc":"E07L","nilai_preload":"-","vmc":"-","shim_sr":"1.75","shim_pl":"-0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-29","pic":"FERDIKA - NAGITA","code":"XM380 - 03","size":"XM380","container":"","ket":"CLEANING","mc":"E07R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-05","pic":"PEBRI - UBAY","code":"BX731 - 16","size":"BX731","container":"","ket":"START","mc":"A02R","nilai_preload":"1.27","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.00 / BAWAH 1.90"},{"tgl":"2026-01-04","pic":"TOHANA - IDAD","code":"BX731 - 08","size":"BX731","container":"","ket":"START","mc":"A04L","nilai_preload":"-0.58","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 2.30 / BAWAH 3.00 + OOR"},{"tgl":"2026-01-02","pic":"TOHANA - IDAD","code":"BX776 - 05","size":"BX776","container":"","ket":"GACON","mc":"A07L","nilai_preload":"0.42","vmc":"-","shim_sr":"0.,70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.00 / BAWAH 1.80"},{"tgl":"2026-12-27","pic":"DWI - SATRIA","code":"BX775 - 01","size":"BX775","container":"","ket":"START","mc":"C10L","nilai_preload":"0.89","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-18","pic":"SYAIFUL - HARIS","code":"B466 - 01","size":"B466","container":"","ket":"START","mc":"A23L","nilai_preload":"-0.12","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OFSH","keterangan":""},{"tgl":"2026-12-18","pic":"HARIS - BADRU","code":"B466 - 02","size":"B466","container":"","ket":"START","mc":"A23R","nilai_preload":"-0.09","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-12-18","pic":"SIGIT - DIVA","code":"B739 - 02","size":"B739","container":"","ket":"START","mc":"A25L","nilai_preload":"-0.28","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OOR","keterangan":""},{"tgl":"2026-12-18","pic":"SIGIT - DIVA","code":"B739 - 01","size":"B739","container":"","ket":"START","mc":"A25R","nilai_preload":"0.50","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-12-18","pic":"IRFAN - AOP","code":"BX551 - 04","size":"BX551","container":"","ket":"CLEANING","mc":"B06L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"1.40","first_cure":"MOR + DMGM","keterangan":""},{"tgl":"2026-12-18","pic":"IRFAN - AOP","code":"BX551 - 24","size":"BX551","container":"","ket":"CLEANING","mc":"B06R","nilai_preload":"-","vmc":"-","shim_sr":"1.75","shim_pl":"2.10","first_cure":"OK","keterangan":""},{"tgl":"2026-12-18","pic":"HARIS - BADRU","code":"A293 - 01","size":"A293","container":"","ket":"START","mc":"B28L","nilai_preload":"-0.29","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK ( LL )","keterangan":""},{"tgl":"2026-12-18","pic":"AGUS - MUHAIMIN","code":"BX733 - 04","size":"BX733","container":"","ket":"CLEANING","mc":"C01L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"1.40","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-12-18","pic":"AGUS - MUHAIMIN","code":"BX733 - 06","size":"BX733","container":"","ket":"CLEANING","mc":"C01R","nilai_preload":"-","vmc":"-","shim_sr":"1.40","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-12-18","pic":"HARIS - BADRU","code":"XM379 - 01","size":"XM379","container":"","ket":"START","mc":"C09L","nilai_preload":"0.57","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-12-18","pic":"HARIS - BADRU","code":"XM384 - 03","size":"XM384","container":"","ket":"START","mc":"C09R","nilai_preload":"0.50","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-12-18","pic":"HARIS - BADRU","code":"B490 - 01","size":"B490","container":"","ket":"CLEANING","mc":"E01L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OFSH + DMGM","keterangan":""},{"tgl":"2026-01-09","pic":"DWI - SATRIA","code":"B256 - 05","size":"B256","container":"AZ STP - 18","ket":"CLEANING","mc":"A08R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"1.05","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.50 / BAWAH 2.10"},{"tgl":"2026-12-18","pic":"FERDIKA - NAGITA","code":"XM404 - 01","size":"XM404","container":"","ket":"START","mc":"E02L","nilai_preload":"-0.27","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-12-18","pic":"FERDIKA - NAGITA","code":"XM421 - 01","size":"XM421","container":"","ket":"START","mc":"E02R","nilai_preload":"-0.18","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-12-18","pic":"IRFAN - AOP","code":"BM397 - 01","size":"BM397","container":"","ket":"START","mc":"E04L","nilai_preload":"-0.28","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-01-09","pic":"SYAIFUL - HARIS","code":"BX776 - 02","size":"BX776","container":"SM STP - 144","ket":"CLEANING","mc":"A09L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"3.15","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.10 / BAWAH 1.90 + OS + OFSH"},{"tgl":"2026-12-18","pic":"FERDIKA - NAGITA","code":"B644 - 09","size":"B644","container":"","ket":"CLEANING","mc":"H15L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-12-18","pic":"FERDIKA - NAGITA","code":"B644 - 13","size":"B644","container":"","ket":"CLEANING","mc":"H15R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OFBR )","keterangan":""},{"tgl":"2026-01-23","pic":"PEBRI - DIKA","code":"BX731 - 10","size":"BX731","container":"HM STP - 561","ket":"START","mc":"A10R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR BAWAH ","keterangan":"BAWAH 2.30 + OS + DM"},{"tgl":"2026-12-18","pic":"DWI - SATRIA","code":"BX675 - 01","size":"BX675","container":"","ket":"START","mc":"J14L","nilai_preload":"-0.34","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-01-28","pic":"PEBRI - DIKA","code":"BX551 - 06","size":"BX551","container":"SM STP - 356","ket":"START","mc":"B05L","nilai_preload":"0.30","vmc":"1.05","shim_sr":"1.05","shim_pl":"0.35","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.05 / BAWAH 2.50"},{"tgl":"2026-12-27","pic":"SYAIFUL - HARIS","code":"BX551 - 13","size":"BX551","container":"","ket":"CLEANING","mc":"B08L","nilai_preload":"-","vmc":"-","shim_sr":"1.75","shim_pl":"1.75","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.00 / BAWAH 2.10"},{"tgl":"2026-12-27","pic":"SYAIFUL - HARIS","code":"BX551 - 26","size":"BX551","container":"","ket":"CLEANING","mc":"B08R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"1.75","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.10 / BAWAH 2.40"},{"tgl":"2026-12-18","pic":"IRFAN - AOP","code":"XS296 - 01","size":"XS296","container":"","ket":"START","mc":"K09L","nilai_preload":"-0.54","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OS","keterangan":""},{"tgl":"2026-01-20","pic":"PEBRI - DIKA","code":"XS236 - 01","size":"XS236","container":"SM STP - 580","ket":"START","mc":"B14L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -0.50 / BAWAH 1.00 / Spec MOR 1.20"},{"tgl":"2026-12-18","pic":"SIGIT - DIVA","code":"B799 - 02","size":"B799","container":"","ket":"CLEANING","mc":"L07L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-18","pic":"AGUS - AMIR","code":"AS122 - 02","size":"AS122","container":"","ket":"START","mc":"L26L","nilai_preload":"0.11","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OS + DM","keterangan":""},{"tgl":"2026-01-20","pic":"PEBRI - DIKA","code":"XS236 - 02","size":"XS236","container":"SM STP - 316","ket":"START","mc":"B14R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0 / BAWAH 1.00 / Spec MOR 1.20"},{"tgl":"2026-12-18","pic":"FERDIKA - NAGITA","code":"XS429 - 01","size":"XS429","container":"","ket":"START","mc":"M28L","nilai_preload":"0.47","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-17","pic":"AGUS - UBAY","code":"BX776 - 04","size":"BX776","container":"","ket":"START","mc":"A09R","nilai_preload":"-1.27","vmc":"-","shim_sr":"0.70","shim_pl":"1.75","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2026-12-17","pic":"TOHANA - IDAD","code":"BX551 - 16","size":"BX551","container":"","ket":"START","mc":"B03L","nilai_preload":"-0.30","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2026-12-17","pic":"TOHANA - IDAD","code":"BX51 - 09","size":"BX51","container":"","ket":"START","mc":"B03R","nilai_preload":"-0.55","vmc":"-","shim_sr":"0.70","shim_pl":"1.55","first_cure":"OK","keterangan":""},{"tgl":"2026-12-17","pic":"SIGIT - DIVA","code":"BM396 - 02","size":"BM396","container":"","ket":"START","mc":"D05L","nilai_preload":"-0.24","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-12-17","pic":"DWI - SATRIA","code":"BM396 - 01","size":"BM396","container":"","ket":"START","mc":"D05R","nilai_preload":"0.81","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-12-17","pic":"TOHANA - IDAD","code":"XM381 - 01","size":"XM381","container":"","ket":"START","mc":"D12L","nilai_preload":"-0.75","vmc":"0.65","shim_sr":"0.70","shim_pl":"1.40","first_cure":"OK","keterangan":""},{"tgl":"2026-12-17","pic":"TOHANA - IDAD","code":"XM381 - 01","size":"XM381","container":"","ket":"START","mc":"D12L","nilai_preload":"-0.75","vmc":"-","shim_sr":"0.70","shim_pl":"1.40","first_cure":"OK","keterangan":""},{"tgl":"2026-12-17","pic":"TOHANA - IDAD","code":"XM381 - 03","size":"XM381","container":"","ket":"START","mc":"D12R","nilai_preload":"0.32","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-12-17","pic":"TOHANA - IDAD","code":"XM381 - 03","size":"XM381","container":"","ket":"START","mc":"D12R","nilai_preload":"0.32","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-12-17","pic":"DWI - SATRIA","code":"B256 - 06","size":"B256","container":"","ket":"CLEANING","mc":"D15L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"2.10","first_cure":"OK","keterangan":""},{"tgl":"2026-12-17","pic":"SIGIT - DIVA","code":"B256 - 07","size":"B256","container":"","ket":"CLEANING","mc":"D15R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"2.10","first_cure":"OK ( OFBR + DMGM","keterangan":""},{"tgl":"2026-12-17","pic":"MUHAIMIN - DIAZ","code":"BX677 - 03","size":"BX677","container":"","ket":"CLEANING","mc":"F24L","nilai_preload":"-","vmc":"-","shim_sr":"1.40","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-22","pic":"AMIR - AGUS","code":"AS035 - 01","size":"AS035","container":"GT AZIII - 451","ket":"START","mc":"B24L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR BAWAH KELUAR","keterangan":"ATAS 0.50 / BAWAH 2.20 / Spec MOR 1.30"},{"tgl":"2026-12-17","pic":"AGUS - UBAY","code":"B476 - 01","size":"B476","container":"","ket":"START","mc":"G27L","nilai_preload":"0.49","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( LL )","keterangan":""},{"tgl":"2026-12-17","pic":"AGUS - UBAY","code":"B476 - 02","size":"B476","container":"","ket":"START","mc":"G27R","nilai_preload":"0.47","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( LL )","keterangan":""},{"tgl":"2026-01-22","pic":"AMIR - AGUS","code":"AS035 - 02","size":"AS035","container":"SH AZIII - 137","ket":"START","mc":"B24R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR BAWAH KELUAR","keterangan":"ATAS 1.1 / BAWAH 2.5 / Spec MOR 1.30"},{"tgl":"2026-01-13","pic":"IRFAN - AOP","code":"BX733 - 22","size":"BX733","container":"SM STP - 18","ket":"CLEANING","mc":"C03L","nilai_preload":"-","vmc":"-","shim_sr":"1.40","shim_pl":"0.35","first_cure":"MOR BAWAH KEDALAM","keterangan":"ATAS 1.70 / BAWAH 1.00"},{"tgl":"2026-12-17","pic":"DWI - SATRIA","code":"B484 - 01","size":"B484","container":"","ket":"START","mc":"L22L","nilai_preload":"-1.44","vmc":"-","shim_sr":"0.70","shim_pl":"1.75","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2026-12-17","pic":"FERDIKA - NAGITA","code":"B867 - 01","size":"B867","container":"","ket":"START","mc":"M21R","nilai_preload":"0.50","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-17","pic":"MUHAIMIN - DIAZ","code":"BX883 - 01","size":"BX883","container":"","ket":"START","mc":"M22L","nilai_preload":"-","vmc":"-","shim_sr":"-","shim_pl":"-","first_cure":"OFSH","keterangan":""},{"tgl":"2026-12-17","pic":"AGUS - UBAY","code":"AS093 - 01","size":"AS093","container":"","ket":"START","mc":"M26L","nilai_preload":"0.88","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK ( FM + LS )","keterangan":""},{"tgl":"2026-12-17","pic":"DWI - SATRIA","code":"XS225 - 01","size":"XS225","container":"","ket":"START","mc":"N10L","nilai_preload":"0.22","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OOR","keterangan":""},{"tgl":"2026-12-17","pic":"SIGIT - DIVA","code":"XS225 - 02","size":"XS225","container":"","ket":"START","mc":"N10R","nilai_preload":"0.15","vmc":"-","shim_sr":"1.05","shim_pl":"0.35","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-12-16","pic":"MUHAIMIN - DIAZ","code":"B356 - 06","size":"B356","container":"","ket":"CLEANING","mc":" A11L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-16","pic":"MUHAIMIN - DIAZ","code":"B356 - 16","size":"B356","container":"","ket":"CLEANING","mc":"A11R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OFSH + DMGM","keterangan":""},{"tgl":"2026-12-16","pic":"IRFAN - AOP","code":"A568 - 01","size":"A568","container":"","ket":"CLEANING","mc":"B21R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-12-16","pic":"SYAIFUL - IRFAN","code":"B644 - 03","size":"B644","container":"","ket":"CLEANING","mc":"G03L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-16","pic":"SYAIFUL - IRFAN","code":"B644 - 01","size":"B644","container":"","ket":"CLEANING","mc":"G03R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-12-16","pic":"AGUS - UBAY","code":"AS240 - 01","size":"AS240","container":"","ket":"CLEANING","mc":"G23L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-12-16","pic":"MUHAIMIN - DIAZ","code":"AA094 - 02","size":"AA094","container":"","ket":"CLEANING","mc":"K26L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-16","pic":"SYAIFUL - IRFAN","code":"AA094 - 01","size":"AA094","container":"","ket":"CLEANING","mc":"K29R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OFBR )","keterangan":""},{"tgl":"2026-12-16","pic":"SYAIFUL - IRFAN","code":"AA009 - 04","size":"AA009","container":"","ket":"START","mc":"M14L","nilai_preload":"0.57","vmc":"-","shim_sr":"-0","shim_pl":"0","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-12-16","pic":"SYAIFUL - IRFAN","code":"AA009 - 03","size":"AA009","container":"","ket":"START","mc":"M14R","nilai_preload":"0.81","vmc":"-","shim_sr":"0.35","shim_pl":"0","first_cure":"OS","keterangan":""},{"tgl":"2026-12-14","pic":"SATRIA - DIVA","code":"A538 - 01","size":"A538","container":"","ket":"START","mc":"A21L","nilai_preload":"0.75","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-14","pic":"SATRIA - DIVA","code":"A538 - 02","size":"A538","container":"","ket":"START","mc":"A21R","nilai_preload":"-0.31","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-12-14","pic":"TOHANA - IDAD","code":"AA090 - 01","size":"AA090","container":"","ket":"CLEANING","mc":"G22L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"-","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-12-14","pic":"IRFAN - AOP","code":"BM452 - 01","size":"BM452","container":"","ket":"START","mc":"H24L","nilai_preload":"0.10","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-12-14","pic":"IRFAN - AOP","code":"BM452 - 02","size":"BM452","container":"","ket":"START","mc":"H24R","nilai_preload":"0.49","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-12-14","pic":"AGUS - UBAY","code":"AA011 - 04","size":"AA011","container":"","ket":"START","mc":"K23L","nilai_preload":"-0.27","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2026-12-14","pic":"AGUS - UBAY","code":"AS243 - 01","size":"AS243","container":"","ket":"START","mc":"K23R","nilai_preload":"0.52","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OFSH + OS","keterangan":""},{"tgl":"2026-12-14","pic":"SYAIFUL - HARIS","code":"AS274 - 01","size":"AS274","container":"","ket":"START","mc":"K29L","nilai_preload":"0.33","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-12-14","pic":"AGUS - UBAY","code":"AS230 - 02","size":"AS230","container":"","ket":"CLEANING","mc":"L31L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2026-12-14","pic":"AGUS - UBAY","code":"AS230 - 01","size":"AS230","container":"","ket":"CLEANING","mc":"L31R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"OK ( DM + LSH )","keterangan":""},{"tgl":"2026-12-13","pic":"AGUS - UBAY","code":"B256 - 10","size":"B256","container":"","ket":"CLEANING","mc":"C13L","nilai_preload":"-","vmc":"-","shim_sr":"1.75","shim_pl":"1.40","first_cure":"OK","keterangan":""},{"tgl":"2026-12-13","pic":"AGUS - UBAY","code":"B256 - 12","size":"B256","container":"","ket":"CLEANING","mc":"C13R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-12-13","pic":"SATRIA - VICKY","code":"B468 - 01","size":"B468","container":"","ket":"START","mc":"H30L","nilai_preload":"1.29","vmc":"-","shim_sr":"1.40","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-13","pic":"SATRIA - VICKY","code":"B468 - 02","size":"B468","container":"","ket":"START","mc":"H30R","nilai_preload":"0.08","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OFSH","keterangan":""},{"tgl":"2026-12-13","pic":"IRFAN - AOP","code":"B354 - 04","size":"B354","container":"","ket":"START","mc":"K27L","nilai_preload":"0.53","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-12-13","pic":"SYAIFUL - HARIS","code":"B354 - 03","size":"B354","container":"","ket":"START","mc":"K27R","nilai_preload":"0.93","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-12-13","pic":"TOHANA - IDAD","code":"B437 - 01","size":"B437","container":"","ket":"START","mc":"L25R","nilai_preload":"0.18","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-12-12","pic":"MUHAIMIN - DIAZ","code":"BX677 - 02","size":"BX677","container":"","ket":"START","mc":"F21R","nilai_preload":"-0.21","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OFSH","keterangan":""},{"tgl":"2026-12-12","pic":"MUHAIMIN - DIAZ","code":"AS273 - 01","size":"AS273","container":"","ket":"START","mc":"K28L","nilai_preload":"0.54","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( FM )","keterangan":""},{"tgl":"2026-12-12","pic":"FERDIKA - NAGITA","code":"XS253 - 01","size":"XS253","container":"","ket":"START","mc":"L24L","nilai_preload":"0.51","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2026-12-12","pic":"DWI - SATRIA","code":"AS433 - 03","size":"AS433","container":"","ket":"CLEANING","mc":"L25L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0.70","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2026-12-12","pic":"DWI - SATRIA","code":"B056 - 02","size":"B056","container":"","ket":"START","mc":"L27L","nilai_preload":"0.16","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2026-12-12","pic":"SYAIFUL - VICKY","code":"B056 - 01","size":"B056","container":"","ket":"START","mc":"L27R","nilai_preload":"0.49","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2026-12-11","pic":"SYAIFUL - IRFAN","code":"BX677 - 12","size":"BX677","container":"","ket":"CLEANING","mc":"F23L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"1.40","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-12-11","pic":"SYAIFUL - IRFAN","code":"BX677 - 15","size":"BX677","container":"","ket":"CLEANING","mc":"F23R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"MOR BAWAH KELUAR","keterangan":""},{"tgl":"2026-12-11","pic":"AGUS - UBAY","code":"XS227 - 01","size":"XS227","container":"","ket":"START","mc":"F28L","nilai_preload":"-0.39","vmc":"-","shim_sr":"0","shim_pl":"1.05","first_cure":"OK ( LS )","keterangan":""},{"tgl":"2026-12-11","pic":"FERDIKA - NAGITA","code":"AS240 - 02","size":"AS240","container":"","ket":"START","mc":"G23R","nilai_preload":"0.25","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-12-11","pic":"SIGIT - DIVA","code":"B671 - 03","size":"B671","container":"","ket":"CLEANING","mc":"H27L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-11","pic":"SIGIT - DIVA","code":"BX671 - 01","size":"BX671","container":"","ket":"CLEANING","mc":"H27R","nilai_preload":"-","vmc":"-","shim_sr":"2.45","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-05","pic":"PEBRI - UBAY","code":"BX731 - 18","size":"BX731","container":"","ket":"START","mc":"A02L","nilai_preload":"0.67","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-13","pic":"SYAIFUL - HARIS","code":"BX733 - 01","size":"BX733","container":"SM STP - 210","ket":"CLEANING","mc":"C03R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0.70","first_cure":"MOR OUT SPEC","keterangan":"ATAS 0.70 / BAWAH 1.00 / SPEC MOR 1.70"},{"tgl":"2026-01-05","pic":"TOHANA - IDAD","code":"B644 - 09","size":"B644","container":"","ket":"START","mc":"G05L","nilai_preload":"-0.33","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-01-22","pic":"SYAIFUL - MUHAIMIN","code":"BX733 - 26","size":"BX733","container":"SM STP - 460","ket":"START","mc":"C05L","nilai_preload":"-","vmc":"","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.60 / BAWAH 2.0"},{"tgl":"2026-01-04","pic":"TOHANA - IDAD","code":"BX731 - 14","size":"BX731","container":"","ket":"START","mc":"A04R","nilai_preload":"0.06","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-01-04","pic":"DWI - SATRIA","code":"BX776 - 06","size":"BX776","container":"","ket":"CLEANING","mc":"A07R","nilai_preload":"-","vmc":"-","shim_sr":"0.35","shim_pl":"0","first_cure":"OK ( LS )","keterangan":""},{"tgl":"2026-01-04","pic":"IRFAN - AOP","code":"BX733 - 16","size":"BX733","container":"","ket":"CLEANING","mc":"C11L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK ( DM + DMGM )","keterangan":""},{"tgl":"2026-01-04","pic":"SYAIFUL - HARIS","code":"BX733 - 10","size":"BX733","container":"","ket":"CLEANING","mc":"C11R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK ( DM + DMGM )","keterangan":""},{"tgl":"2026-01-04","pic":"SIGIT - DIVA","code":"B356 - 12","size":"B356","container":"","ket":"START","mc":"D08L","nilai_preload":"-0.21","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-01-04","pic":"DWI - SATRIA","code":"B356 - 05","size":"B356","container":"","ket":"START","mc":"D08R","nilai_preload":"-0.76","vmc":"-","shim_sr":"0.70","shim_pl":"1.75","first_cure":"OFSH","keterangan":""},{"tgl":"2026-01-04","pic":"IRFAN - AOP","code":"XM394 - 03","size":"XM394","container":"","ket":"CLEANING","mc":"F09L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-04","pic":"SYAIFUL - HARIS","code":"XM394 - 06","size":"XM394","container":"","ket":"CLEANING","mc":"F09R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-01-04","pic":"AGUS - UBAY","code":"B471 - 02","size":"B471","container":"","ket":"START","mc":"G10L","nilai_preload":"0.44","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( IP )","keterangan":""},{"tgl":"2026-01-03","pic":"IRFAN - HARIS","code":"B356 - 19","size":"B356","container":"","ket":"CLEANING","mc":"A13L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-01-03","pic":"SYAIFUL - TEGAR","code":"B356 - 03","size":"B356","container":"","ket":"CLEANING","mc":"A13R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-03","pic":"SIGIT - DIVA","code":"B256 - 13","size":"B256","container":"","ket":"CLEANING","mc":"B13L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-01-03","pic":"SIGIT - DIVA","code":"B256 - 08","size":"B256","container":"","ket":"CLEANING","mc":"B13R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-03","pic":"NAGITA","code":"BX733 - 21","size":"BX733","container":"","ket":"START","mc":"C10R","nilai_preload":"-0.14","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-01-03","pic":"SYAIFUL - TEGAR","code":"XM399 - 01","size":"XM399","container":"","ket":"START","mc":"D09L","nilai_preload":"0.64","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-22","pic":"SYAIFUL - MUHAIMIN","code":"BX733 - 20","size":"BX733","container":"SM STP - 731","ket":"START","mc":"C05R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.80 / BAWAH 1.60"},{"tgl":"2026-01-20","pic":"PEBRI - DIKA","code":"XM377 - 02","size":"XM377","container":"SM STP - 368","ket":"START","mc":"C06L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -1.00 / BAWAH -0.30 / Spec MOR -0.60"},{"tgl":"2026-01-03","pic":"DWI - SATRIA","code":"AA061 - 04","size":"AA061","container":"","ket":"START","mc":"E12L","nilai_preload":"1.58","vmc":"-","shim_sr":"1.40","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-03","pic":"DWI - SATRIA","code":"AA061 - 01","size":"AA061","container":"","ket":"START","mc":"E12R","nilai_preload":"0.26","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-01-03","pic":"IRFAN - HARIS","code":"XM390 - 02","size":"XM390","container":"","ket":"CLEANING","mc":"F12L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-01-03","pic":"SIGIT - DIVA","code":"XM390 - 01","size":"XM390","container":"","ket":"CLEANING","mc":"F12R","nilai_preload":"--","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-01-20","pic":"PEBRI - DIKA","code":"XM377 - 04","size":"XM377","container":"AZ STP - 503","ket":"START","mc":"C06R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.80 / BAWAH -0.30 / Spec MOR -0.60"},{"tgl":"2026-01-20","pic":"MUHAIMIN - DIAZ","code":"XM423 - 02","size":"XM423","container":"SM STP - 799","ket":"START","mc":"C08L","nilai_preload":"-","vmc":"-","shim_sr":"?","shim_pl":"?","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 1.10 / 0.80 + DMGB"},{"tgl":"2026-01-03","pic":"DWI - SATRIA","code":"B471 - 01","size":"B471","container":"","ket":"START","mc":"G10R","nilai_preload":"0","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-01-08","pic":"FERDIKA - NAGITA","code":"AS270 - 01","size":"AS270","container":"SM STP - 13","ket":"START","mc":"C09R","nilai_preload":"-0.23","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 1.70 / BAWAH 0.30"},{"tgl":"2025-11-29","pic":"SYAIFUL - HARIS","code":"XS068 - 02","size":"XS068","container":"","ket":"START","mc":"G08R","nilai_preload":"0.44","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-29","pic":"AOP - IRFAN","code":"AS285 - 01","size":"AS285","container":"","ket":"START","mc":"K08L","nilai_preload":"0.53","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-29","pic":"FERDIKA - NAGITA","code":"BX735 - 02","size":"BX735","container":"","ket":"START","mc":"N06L","nilai_preload":"0.44","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-29","pic":"FERDIKA - NAGITA","code":"BX735 - 01","size":"BX735","container":"","ket":"START","mc":"N06R","nilai_preload":"0.62","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-28","pic":"SYAIFUL - HARIS","code":"B644 - 07","size":"B644","container":"","ket":"START","mc":"G06L","nilai_preload":"0.08","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2025-11-28","pic":"SYAIFUL - HARIS","code":"B644 - 08","size":"B644","container":"","ket":"START","mc":"G06R","nilai_preload":"-0.45","vmc":"-","shim_sr":"0","shim_pl":"1.05","first_cure":"OK ( DM + OFBR )","keterangan":""},{"tgl":"2025-11-28","pic":"MUHAIMIN - DIAZ","code":"XS068 - 01","size":"XS068","container":"","ket":"START","mc":"G08L","nilai_preload":"-0.21","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-27","pic":"SIGIT - DIVA","code":"XP173 - 01","size":"XP173","container":"","ket":"START","mc":"D08L","nilai_preload":"0.82","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-27","pic":"TOHANA - IDAD","code":"BX642 - 09","size":"BX642","container":"","ket":"START","mc":"D14R","nilai_preload":"-0.73","vmc":"0","shim_sr":"0","shim_pl":"1.40","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-27","pic":"FREDIKA -NAGITA","code":"XA030 - 01","size":"XA030","container":"","ket":"CLEANING","mc":"H07L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-27","pic":"SIGITI - DIVA","code":"XA030 - 02","size":"XA030","container":"","ket":"CLEANING","mc":"H07R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-27","pic":"TOHANA - IDAD","code":"BX254 - 05","size":"BX254","container":"","ket":"CLEANING","mc":"H14R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK ( OS + DMGM )","keterangan":""},{"tgl":"2025-11-26","pic":"AGUS - UBAY","code":"A291 - 03","size":"A291","container":"","ket":"START","mc":"G10L","nilai_preload":"0.60","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-26","pic":"AGUS - UBAY","code":"A291 - 02","size":"A291","container":"","ket":"START","mc":"G10R","nilai_preload":"-0.20","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK ( OS + OFSH )","keterangan":""},{"tgl":"2025-11-26","pic":"MUHAIMIN - DIAZ","code":"BX254 - 14","size":"BX254","container":"","ket":"CLEANING","mc":"H14L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"1.40","first_cure":"OK ( OS + OFSH )","keterangan":""},{"tgl":"2025-11-25","pic":"TOHANA - IDAD","code":"BX774 - 05","size":"BX774","container":"","ket":"START","mc":"D10L","nilai_preload":"-0.8","vmc":"0","shim_sr":"0","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-25","pic":"FERDIKA - NAGITA","code":"BX729 - 02","size":"BX729","container":"","ket":"START","mc":"F04R","nilai_preload":"0.52","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OOR","keterangan":""},{"tgl":"2025-11-25","pic":"SYAIFUL - TEGAR","code":"XM423 - 04","size":"XM423","container":"","ket":"CLEANING","mc":"F06L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK ( DMGB )","keterangan":""},{"tgl":"2025-11-25","pic":"AOP - IRFAN","code":"AS270 - 01","size":"AS270","container":"","ket":"CLEANING","mc":"F12R","nilai_preload":"-","vmc":"-","shim_sr":"1.75","shim_pl":"0","first_cure":"MOR ATAS MIRING","keterangan":""},{"tgl":"2025-11-25","pic":"TOHANA - IDAD","code":"AS264 - 02","size":"AS264","container":"","ket":"CLEANING","mc":"G12R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-25","pic":"MUHAIMIN - DIAZ","code":"B644 - 02","size":"B644","container":"","ket":"CLEANING","mc":"H03L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-25","pic":"FERDIKA - NAGITA","code":"B644 - 11","size":"B644","container":"","ket":"CLEANING","mc":"H03R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-24","pic":"SYAIFUL - HARIS","code":"B356 - 11","size":"B356","container":"","ket":"CLEANING","mc":"A12R","nilai_preload":"-","vmc":"-","shim_sr":"1.75","shim_pl":"2.8","first_cure":"OFSH + OFBR","keterangan":""},{"tgl":"2025-11-24","pic":"SIGIT - DIVA","code":"BX551 - 20","size":"BX551","container":"","ket":"CLEANING","mc":"B02L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"2.80","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-24","pic":"SIGIT - DIVA","code":"BX551 - 03","size":"BX551","container":"","ket":"CLEANING","mc":"B02R","nilai_preload":"-","vmc":"-","shim_sr":"1.4","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-24","pic":"SYAIFUL - HARIS","code":"BX551 - 18","size":"BX551","container":"","ket":"START","mc":"B14R","nilai_preload":"0.37","vmc":"0.25","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-24","pic":"AOP - IRFAN","code":"BX676 - 03","size":"BX676","container":"","ket":"START","mc":"D10R","nilai_preload":"-0.18","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-24","pic":"AOP - IRFAN","code":"XM423 - 02","size":"XM423","container":"","ket":"START","mc":"F06R","nilai_preload":"-0.19","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-24","pic":"SYAIFUL - HARIS","code":"XM376 - 02","size":"XM376","container":"","ket":"START","mc":"F10L","nilai_preload":"-0.60","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-24","pic":"AGUS - UBAY","code":"XM376 - 08","size":"XM376","container":"","ket":"START","mc":"F10R","nilai_preload":"1.2","vmc":"-","shim_sr":"1.4","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-24","pic":"DWI - SATRIA","code":"B644 - 10","size":"B644","container":"","ket":"CLEANING","mc":"G05L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-24","pic":"DWI - SATRIA","code":"B644 - 04","size":"B644","container":"","ket":"CLEANING","mc":"G05R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OFBR )","keterangan":""},{"tgl":"2025-11-23","pic":"TOHANA - IDAD","code":"XM376 - 04","size":"XM376","container":"","ket":"START","mc":"F01L","nilai_preload":"0.11","vmc":"0","shim_sr":"0","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-23","pic":"TOHANA - IDAD","code":"XM376 - 07","size":"XM376","container":"","ket":"START","mc":"F01R","nilai_preload":"1,08","vmc":"0","shim_sr":"1,05","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-23","pic":"SYAIFUL - HARIS","code":"B475 - 01","size":"B475","container":"","ket":"START","mc":"G11R","nilai_preload":"-0.16","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-22","pic":"SYAIFUL - HARIS","code":"B445 - 01","size":"B445","container":"","ket":"START","mc":"A223L","nilai_preload":"0.08","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-22","pic":"SYAIFUL - HARIS","code":"B445 - 04","size":"B445","container":"","ket":"START","mc":"A23R","nilai_preload":"0.91","vmc":"-","shim_sr":"0.35","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-22","pic":"FERDIKA - NAGITA","code":"BX551 - 19","size":"BX551","container":"","ket":"CLEANING","mc":"B07R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-22","pic":"FERDIKA - NAGITA","code":"BX551 - 19","size":"BX551","container":"","ket":"CLEANING","mc":"B07R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-22","pic":"AOP - IRFAN","code":"AS125 - 02","size":"AS125","container":"","ket":"START","mc":"B24L","nilai_preload":"0.41","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-22","pic":"AOP - IRFAN","code":"AS125 - 01","size":"AS125","container":"","ket":"START","mc":"B24R","nilai_preload":"0.58","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-22","pic":"FERDIKA - NAGITA","code":"XS203 - 02","size":"XS203","container":"","ket":"START","mc":"C09R","nilai_preload":"0.45","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-22","pic":"SYAIFUL - HARIS","code":"BX729 - 01","size":"BX729","container":"","ket":"START","mc":"F04L","nilai_preload":"-0.24","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-22","pic":"SYAIFUL - HARIS","code":"XM376 - 06","size":"XM376","container":"","ket":"START","mc":"F15L","nilai_preload":"-0.46","vmc":"-","shim_sr":"0","shim_pl":"1.05","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-22","pic":"AOP - IRFAN","code":"XM376 - 01","size":"XM376","container":"","ket":"START","mc":"F15R","nilai_preload":"-0.09","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-22","pic":"MUHAIMIN - DIAZ","code":"B869 - 01","size":"B869","container":"","ket":"START","mc":"K28L","nilai_preload":"0.17","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2025-11-22","pic":"FERDIKA - NAGITA","code":"B868 - 01","size":"B868","container":"","ket":"START","mc":"L27R","nilai_preload":"0.51","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-22","pic":"MUHAIMIN - DIAZ","code":"B541 - 01","size":"B541","container":"","ket":"START","mc":"L29R","nilai_preload":"0.57","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OS + OFSH )","keterangan":""},{"tgl":"2025-11-22","pic":"FERDIKA - NAGITA","code":"B800 - 02","size":"B800","container":"","ket":"START","mc":"L30R","nilai_preload":"0.61","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-22","pic":"MUHAIMIN - DIAZ","code":"B514 - 04","size":"B514","container":"","ket":"START","mc":"M02L","nilai_preload":"0.49","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-22","pic":"MUHAIMIN - DIAZ","code":"B514 - 02","size":"B514","container":"","ket":"START","mc":"M02R","nilai_preload":"0.46","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-22","pic":"DWI - SATRIA","code":"AS140 - 02","size":"AS140","container":"","ket":"START","mc":"M03L","nilai_preload":"1.47","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-22","pic":"DWI - SATRIA","code":"AS140 - 01","size":"AS140","container":"","ket":"START","mc":"M03R","nilai_preload":"0.85","vmc":"-","shim_sr":"0.35","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-22","pic":"SYAIFUL - HARIS","code":"B795 - 01","size":"B795","container":"","ket":"CLEANING","mc":"M21L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-22","pic":"MUHAIMIN - DIAZ","code":"B795 - 02","size":"B795","container":"","ket":"START","mc":"M21R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-22","pic":"AOP - IRFAN","code":"AS099 - 02","size":"AS099","container":"","ket":"START","mc":"N06L","nilai_preload":"0.44","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OS )","keterangan":""},{"tgl":"2025-11-22","pic":"DWI - SATRIA","code":"AS099 - 01","size":"AS099","container":"","ket":"START","mc":"N06R","nilai_preload":"0.12","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"DWI - SATIRA","code":"AA095 - 01","size":"AA095","container":"","ket":"START","mc":"A22L","nilai_preload":"0.24","vmc":"0.50","shim_sr":"0","shim_pl":"0.35","first_cure":"OK ( OFSH + FM )","keterangan":""},{"tgl":"2025-11-21","pic":"MUHAIMIN - DIAZ","code":"AA095 - 02","size":"AA095","container":"","ket":"START","mc":"A22R","nilai_preload":"-0.10","vmc":"0.50","shim_sr":"1.05","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"FERDIKA - NAGITA","code":"AS103 - 01","size":"AS103","container":"","ket":"START","mc":"A28L","nilai_preload":"0.49","vmc":"0.15","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-21","pic":"TOHANA - IDAD","code":"AS103 - 02","size":"AS103","container":"","ket":"START","mc":"A28R","nilai_preload":"0.58","vmc":"0.15","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DM + OS )","keterangan":""},{"tgl":"2025-11-21","pic":"PEBRI - DIKA","code":"BX551 - 07","size":"BX551","container":"","ket":"EX PROB","mc":"B07L","nilai_preload":"1.33","vmc":"0.8","shim_sr":"1.40","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-21","pic":"FERDIKA - NAGITA","code":"A323 - 01","size":"A323","container":"","ket":"START","mc":"B12L","nilai_preload":"0.63","vmc":"0.76","shim_sr":"1.4","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"SIGIT - DIVA","code":"B355 - 01","size":"B355","container":"","ket":"START","mc":"D13L","nilai_preload":"-0.72","vmc":"0","shim_sr":"0.70","shim_pl":"1.4","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"SIGIT - DIVA","code":"B355 - 01","size":"B355","container":"","ket":"START","mc":"D13L","nilai_preload":"-0.72","vmc":"0","shim_sr":"0.70","shim_pl":"1.4","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"PEBRI - DIKA","code":"XS073 - 03","size":"XS073","container":"","ket":"START","mc":"E10R","nilai_preload":"-0.27","vmc":"0","shim_sr":"0","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"TOHANA - IDAD","code":"AS159 - 01","size":"AS159","container":"","ket":"CLEANING","mc":"F26R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-21","pic":"AOP - HARIS","code":"BX512 - 02","size":"BX512","container":"","ket":"START","mc":"G01L","nilai_preload":"0.58","vmc":"0.60","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-21","pic":"AOP - HARIS","code":"B512 - 02","size":"B512","container":"","ket":"START","mc":"G01L","nilai_preload":"0.58","vmc":"0.60","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-21","pic":"AOP - HARIS","code":"BX512 - 05","size":"BX512","container":"","ket":"START","mc":"G01R","nilai_preload":"0.54","vmc":"0.60","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-21","pic":"AOP - HARIS","code":"B512 - 05","size":"B512","container":"","ket":"START","mc":"G01R","nilai_preload":"0.54","vmc":"0.60","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-21","pic":"MUHAIMIN - DIAZ","code":"A508 - 01","size":"A508","container":"","ket":"START","mc":"G15L","nilai_preload":"0.59","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"MUHAIMIN - DIAZ","code":"A508 - 01","size":"A508","container":"","ket":"START","mc":"G15L","nilai_preload":"0.59","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"MUHAIMIN - DIAZ","code":"A508 - 02","size":"A508","container":"","ket":"START","mc":"G15R","nilai_preload":"-0.09","vmc":"0","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"MUHAIMIN - DIAZ","code":"A508 - 02","size":"A508","container":"","ket":"START","mc":"G15R","nilai_preload":"-0.09","vmc":"0","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"SIGIT - DIVA","code":"A509 - 02","size":"A509","container":"","ket":"START","mc":"K22R","nilai_preload":"0.99","vmc":"0.27","shim_sr":"1.4","shim_pl":"0","first_cure":"OK ( OS + OF )","keterangan":""},{"tgl":"2025-11-21","pic":"AOP - HARIS","code":"AA074 - 01","size":"AA074","container":"","ket":"START","mc":"K25R","nilai_preload":"0.43","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-21","pic":"FERDIKA - NAGITA","code":"B541 - 02","size":"B541","container":"","ket":"START","mc":"K29L","nilai_preload":"0.59","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"DWI - SATRIA","code":"A883 - 02","size":"A883","container":"","ket":"START","mc":"K30L","nilai_preload":"-0.51","vmc":"0","shim_sr":"0","shim_pl":"1.05","first_cure":"OK ( OS + OFBR )","keterangan":""},{"tgl":"2025-11-21","pic":"AGUS - UBAY","code":"A883 - 04","size":"A883","container":"","ket":"START","mc":"K30R","nilai_preload":"0.47","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OS )","keterangan":""},{"tgl":"2025-11-21","pic":"DWI - SATRIA","code":"B700 - 02","size":"B700","container":"","ket":"START","mc":"L06R","nilai_preload":"0.49","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK ( MS + LL )","keterangan":""},{"tgl":"2025-11-21","pic":"AOP - HARIS","code":"XM366 - 01","size":"XM366","container":"","ket":"START","mc":"L08L","nilai_preload":"0.66","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2025-11-21","pic":"FERDIKA - NAGITA","code":"AA011 - 02","size":"AA011","container":"","ket":"START","mc":"L11L","nilai_preload":"0.48","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"FERDIKA - NAGITA","code":"AA011 - 04","size":"AA011","container":"","ket":"START","mc":"L11R","nilai_preload":"0.55","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-21","pic":"AGUS - UBAY","code":"AA081 - 01","size":"AA081","container":"","ket":"START","mc":"L12L","nilai_preload":"0.52","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"AGUS - UBAY","code":"AA081 - 03","size":"AA081","container":"","ket":"START","mc":"L12R","nilai_preload":"0.59","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"AGUS - UBAY","code":"BX526 - 01","size":"BX526","container":"","ket":"START","mc":"L21L","nilai_preload":"0.14","vmc":"0","shim_sr":"0.35","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-21","pic":"DWI - SATRIA","code":"BX536 - 02","size":"BX536","container":"","ket":"START","mc":"L25L","nilai_preload":"0.57","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OS )","keterangan":""},{"tgl":"2025-11-21","pic":"TOHANA - IDAD","code":"A509 - 01","size":"A509","container":"","ket":"START","mc":"L26L","nilai_preload":"0.54","vmc":"0.27","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR KELUAR","keterangan":""},{"tgl":"2025-11-21","pic":"SIGIT - DIVA","code":"B444 - 04","size":"B444","container":"","ket":"START","mc":"M05L","nilai_preload":"1.14","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK ( OS )","keterangan":""},{"tgl":"2025-11-21","pic":"SIGIT - DIVA","code":"B444 - 05","size":"B444","container":"","ket":"START","mc":"M05R","nilai_preload":"0.66","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"DWI - SATRIA","code":"AS088 - 01","size":"AS088","container":"","ket":"START","mc":"M07L","nilai_preload":"01.03","vmc":"0","shim_sr":"0.35","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"SIGIT - DIVA","code":"AS088 - 02","size":"AS088","container":"","ket":"START","mc":"M07R","nilai_preload":"0.82","vmc":"0","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"TOHANA - IDAD","code":"XS309 - 01","size":"XS309","container":"","ket":"START","mc":"M24R","nilai_preload":"0.81","vmc":"0.63","shim_sr":"1.4","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-21","pic":"TOHANA - IDAD","code":"B525 - 02","size":"B525","container":"","ket":"START","mc":"M27R","nilai_preload":"0.30","vmc":"0","shim_sr":"0","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2025-11-21","pic":"FERDIKA - NAGITA","code":"XS306 - 05","size":"XS306","container":"","ket":"START","mc":"N04L","nilai_preload":"0.49","vmc":"0.23","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-21","pic":"MUHAIMIN - DIAZ","code":"XS306 - 04","size":"XS306","container":"","ket":"START","mc":"N04R","nilai_preload":"0.42","vmc":"0.23","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-21","pic":"AGUS - UBAY","code":"B525 - 01","size":"B525","container":"","ket":"START","mc":"N26R","nilai_preload":"0.43","vmc":"0.80","shim_sr":"1.05","shim_pl":"0","first_cure":"OK ( OFSH⬆️⬇️ + OS )","keterangan":""},{"tgl":"2025-11-21","pic":"AOP - IRFAN","code":"B640 - 05","size":"B640","container":"","ket":"START","mc":"N27R","nilai_preload":"0.17","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-21","pic":"AGUS - UBAY","code":"A883 - 04","size":"A883","container":"","ket":"START","mc":"","nilai_preload":"0.47","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-21","pic":"MUHAIMIN - DIAZ","code":"XS306  - 04","size":"XS306","container":"","ket":"START","mc":"","nilai_preload":"0.42","vmc":"0.23","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-20","pic":"AGUS - UBAY","code":"B355 - 04","size":"B355","container":"","ket":"START","mc":"A02L","nilai_preload":"-0.78","vmc":"0.45","shim_sr":"1.05","shim_pl":"1.40","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"AGUS - UBAY","code":"B355 - 03","size":"B355","container":"","ket":"START","mc":"A02R","nilai_preload":"0.41","vmc":"0.45","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DMGM )","keterangan":""},{"tgl":"2025-11-20","pic":"SIGIT - DIVA","code":"B356 - 07","size":"B356","container":"","ket":"CLEANING","mc":"A07R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"SIGIT - DIVA","code":"B356 - 15","size":"B356","container":"","ket":"CLEANING","mc":"A07R","nilai_preload":"-","vmc":"-","shim_sr":"1.75","shim_pl":"1.40","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"MUHAIMIN - DIAZ","code":"BX551 - 29","size":"BX551","container":"","ket":"SPARE CM","mc":"B05L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"MUHAIMIN - DIAZ","code":"BX551 - 08","size":"BX551","container":"","ket":"SPARE CM","mc":"B05R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2025-11-20","pic":"FERDIKA - NAGITA","code":"BX733 - 25","size":"BX733","container":"","ket":"CLEANING","mc":"C08R","nilai_preload":"-","vmc":"-","shim_sr":"0.35","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"DWI - SATRIA","code":"XA061 - 01","size":"XA061","container":"","ket":"START","mc":"D06L","nilai_preload":"0.70","vmc":"0.50","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"DWI - SATRIA","code":"XA061 - 04","size":"XA061","container":"","ket":"START","mc":"D06R","nilai_preload":"-0.14","vmc":"0.50","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"TOHANA - IDAD","code":"BX677 - 10","size":"BX677","container":"","ket":"CLEANING","mc":"F25R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"SYAIFUL - DHONI","code":"AS048 - 01","size":"AS048","container":"","ket":"START","mc":"G09L","nilai_preload":"-0.21","vmc":"0","shim_sr":"0","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-20","pic":"TOHANA - IDAD","code":"AS048 - 02","size":"AS048","container":"","ket":"START","mc":"G09R","nilai_preload":"0.91","vmc":"0","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"TOHANA - IDAD","code":"AS048 - 02","size":"AS048","container":"","ket":"START","mc":"G09R","nilai_preload":"0.91","vmc":"0","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"SYAIFUL - DHONI","code":"XS029 - 05","size":"XS029","container":"","ket":"START","mc":"H10L","nilai_preload":"0.66","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK ( DMGB )","keterangan":""},{"tgl":"2025-11-20","pic":"SYAIFUL - DHONI","code":"XS029 - 05","size":"XS029","container":"","ket":"START","mc":"H10L","nilai_preload":"0.66","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK ( DMGM )","keterangan":""},{"tgl":"2025-11-20","pic":"SYAIFUL - DHONI","code":"XS029 - 03","size":"XS029","container":"","ket":"START","mc":"H10R","nilai_preload":"0.40","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"SYAIFUL - DHONI","code":"XS029 - 03","size":"XS029","container":"","ket":"START","mc":"H10R","nilai_preload":"0.40","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"TOHANA - IDAD","code":"AA081 - 04","size":"AA081","container":"","ket":"START","mc":"H31L","nilai_preload":"0.61","vmc":"0.50","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-20","pic":"SIGIT - DIVA","code":"AA081 - 02","size":"AA081","container":"","ket":"START","mc":"H31R","nilai_preload":"-0.63","vmc":"0.50","shim_sr":"1.05","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"SYAIFUL - DHONI","code":"A288 - 01","size":"A288","container":"","ket":"START","mc":"K07R","nilai_preload":"0.43","vmc":"0.25","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-20","pic":"AGUS - UBAY","code":"AA107 - 02","size":"AA107","container":"","ket":"START","mc":"K09R","nilai_preload":"0.7","vmc":"0","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-20","pic":"PEBRI - DIAZ","code":"AA095 - 09","size":"AA095","container":"","ket":"START","mc":"K22L","nilai_preload":"0.43","vmc":"0.25","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"PEBRI - DIAZ","code":"AA095 - 03","size":"AA095","container":"","ket":"START","mc":"K26R","nilai_preload":"0.19","vmc":"0.25","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK ( DM SW ⬇️ )","keterangan":""},{"tgl":"2025-11-20","pic":"SIGIT - DIVA","code":"A292 - 02","size":"A292","container":"","ket":"CLEANING","mc":"L02R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"SIGIT - DIVA","code":"B700 - 01","size":"B700","container":"","ket":"START","mc":"L07L","nilai_preload":"0.40","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"SYAIFUL - HARIS","code":"B255 - 02","size":"B255","container":"","ket":"START","mc":"L25L","nilai_preload":"0.37","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"AGUS - UBAY","code":"B255 - 01","size":"B255","container":"","ket":"START","mc":"L25R","nilai_preload":"0.97","vmc":"0","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-20","pic":"DWI - SATRIA","code":"XS179 - 01","size":"XS179","container":"","ket":"START","mc":"L29L","nilai_preload":"0.36","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DMGB )","keterangan":""},{"tgl":"2025-11-20","pic":"TOHANA - IDAD","code":"AX584 - 06","size":"AX584","container":"","ket":"START","mc":"M09L","nilai_preload":"0.55","vmc":"0.18","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-20","pic":"TOHANA - IDAD","code":"AX584 - 05","size":"AX584","container":"","ket":"START","mc":"M09R","nilai_preload":"-0.46","vmc":"0.18","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OK ( OFSH⬇️ )","keterangan":""},{"tgl":"2025-11-20","pic":"FERDIKA - NAGITA","code":"AS080 - 02","size":"AS080","container":"","ket":"START","mc":"M12R","nilai_preload":"0.51","vmc":"0.35","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-20","pic":"TOHANA - IDAD","code":"B515 - 04","size":"B515","container":"","ket":"START","mc":"M22R","nilai_preload":"0.91","vmc":"0.49","shim_sr":"1.4","shim_pl":"0","first_cure":"MOR ATAS MIRING","keterangan":""},{"tgl":"2025-11-20","pic":"TOHANA - IDAD","code":"B515 - 03","size":"B515","container":"","ket":"START","mc":"M25L","nilai_preload":"0.45","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"TOHANA - IDAD","code":"A294 - 01","size":"A294","container":"","ket":"START","mc":"N02L","nilai_preload":"0.45","vmc":"0.50","shim_sr":"1.05","shim_pl":"0","first_cure":"OK ( OS + OFSH + DM )","keterangan":""},{"tgl":"2025-11-20","pic":"DWI - SATRIA","code":"B762 - 01","size":"B762","container":"","ket":"START","mc":"N28L","nilai_preload":"0.36","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-20","pic":"FERDIKA - NAGITA","code":"BX733 - 26","size":"BX733","container":"","ket":"CLEANING","mc":"","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-19","pic":"FERDIKA - NAGITA","code":"B356 - 01","size":"B356","container":"","ket":"CLEANING","mc":"A15L","nilai_preload":"-","vmc":"-","shim_sr":"1.40","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2025-11-19","pic":"MUHAIMIN - DIAZ","code":"B356 - 08","size":"B356","container":"","ket":"CLEANING","mc":"A15R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OFSH + OS","keterangan":""},{"tgl":"2025-11-19","pic":"AOP - FAHMI","code":"XA058 - 02","size":"XA058","container":"","ket":"START","mc":"D08L","nilai_preload":"454","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-19","pic":"AOP - FAHMI","code":"XA058 - 01","size":"XA058","container":"","ket":"STARAT","mc":"D08R","nilai_preload":"0.41","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-19","pic":"TOHANA - IDAD","code":"B502 - 02","size":"B502","container":"","ket":"START","mc":"E02L","nilai_preload":"0.7","vmc":"0.65","shim_sr":"1.4","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-19","pic":"SYAIFUL - HARIS","code":"B502 - 01","size":"B502","container":"","ket":"START","mc":"E02R","nilai_preload":"-0.9","vmc":"-","shim_sr":"0","shim_pl":"1.40","first_cure":"OK","keterangan":""},{"tgl":"2025-11-19","pic":"FERDIKA - NAGITA","code":"A322 - 01","size":"A322","container":"","ket":"START","mc":"E04L","nilai_preload":"-0.19","vmc":"0","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-19","pic":"FERDIKA - NAGITA","code":"A322 - 02","size":"A322","container":"","ket":"START","mc":"E04R","nilai_preload":"0.05","vmc":"0","shim_sr":"0","shim_pl":"0.35","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-19","pic":"TOHANA - IDAD","code":"XA088 - 02","size":"XA088","container":"","ket":"START","mc":"E07L","nilai_preload":"0.49","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OS ( OK )","keterangan":""},{"tgl":"2025-11-19","pic":"TOHANA - IDAD","code":"XA088 - 01","size":"XA088","container":"","ket":"START","mc":"E07R","nilai_preload":"0.19","vmc":"0","shim_sr":"0","shim_pl":"0.35","first_cure":"MOR ATAS","keterangan":""},{"tgl":"2025-11-19","pic":"AOP - FAHMI","code":"XA062 - 02","size":"XA062","container":"","ket":"START","mc":"E09R","nilai_preload":"0.44","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-19","pic":"MUHAIMIN - DIAZ","code":"BX676 - 02","size":"BX676","container":"","ket":"START","mc":"F07L","nilai_preload":"0.40","vmc":"0.65","shim_sr":"1,05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-19","pic":"FERDIKA - NAGITA","code":"BX676 - 01","size":"BX676","container":"","ket":"START","mc":"F07R","nilai_preload":"-0.68","vmc":"0.65","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-19","pic":"AGUS - UBAY","code":"BX677 - 03","size":"BX677","container":"","ket":"CLEANING","mc":"F24L","nilai_preload":"-","vmc":"-","shim_sr":"1.40","shim_pl":"0.70","first_cure":"OK ( OFSH )","keterangan":""},{"tgl":"2025-11-19","pic":"AGUS - UBAY","code":"BX677 - 09","size":"BX677","container":"","ket":"CLEANING","mc":"F24R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-19","pic":"FERDIKA - NAGITA","code":"BX677 - 07","size":"BX677","container":"","ket":"CLEANING","mc":"F25L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OFSH )","keterangan":""},{"tgl":"2025-11-19","pic":"AGUS - UBAY","code":"AS057 - 02","size":"AS057","container":"","ket":"START","mc":"H09L","nilai_preload":"0.76","vmc":"0.50","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-19","pic":"MUHAIMIN - DIAZ","code":"AS057 - 01","size":"AS057","container":"","ket":"START","mc":"H09R","nilai_preload":"-0.19","vmc":"0.50","shim_sr":"1.05","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-19","pic":"SYAIFUL - HARIS","code":"B644 - 09","size":"B644","container":"","ket":"CLEANING","mc":"H15L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-19","pic":"SYAIFUL - HARIS","code":"B644 - 09","size":"B644","container":"","ket":"CLEANING","mc":"H15L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-19","pic":"SYAIFUL - HARIS","code":"B644 - 13","size":"B644","container":"","ket":"CLEANING","mc":"H15R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-19","pic":"SYAIFUL - HARIS","code":"B644 - 13","size":"B644","container":"","ket":"CLEANING","mc":"H15R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-19","pic":"PEBRI - DIKA","code":"B690 - 03","size":"B690","container":"","ket":"START","mc":"K01R","nilai_preload":"0.35","vmc":"0.30","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2025-11-19","pic":"AGUS - UBAY","code":"XM450 - 03","size":"XM450","container":"","ket":"START","mc":"K03L","nilai_preload":"-0.90","vmc":"0","shim_sr":"0","shim_pl":"1.05","first_cure":"OK ( OS + DM )","keterangan":""},{"tgl":"2025-11-19","pic":"PEBRI - DIKA","code":"B699 - 01","size":"B699","container":"","ket":"START","mc":"K03R","nilai_preload":"0.39","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-19","pic":"AGUS - UBAY","code":"A288 - 02","size":"A288","container":"","ket":"START","mc":"K07L","nilai_preload":"0.60","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-19","pic":"MUHAIMIN - DIAZ","code":"AA107 - 01","size":"AA107","container":"","ket":"START","mc":"M11L","nilai_preload":"0.40","vmc":"0.24","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-19","pic":"MUHAIMIN - DIAZ","code":"AA107 - 03","size":"AA107","container":"","ket":"START","mc":"M11R","nilai_preload":"0.42","vmc":"0.24","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-19","pic":"SYAIFUL - HARIS","code":"AS080 - 01","size":"AS080","container":"","ket":"START","mc":"M12L","nilai_preload":"0.63","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-19","pic":"TOHANA - IDAD","code":"A294 - 02","size":"A294","container":"","ket":"START","mc":"N02R","nilai_preload":"0.99","vmc":"0","shim_sr":"1.05","shim_pl":"0","first_cure":"OK ( OFSH⬇️ )","keterangan":""},{"tgl":"2025-11-19","pic":"TOHANA - IDAD","code":"A296 - 01","size":"A296","container":"","ket":"START","mc":"N14L","nilai_preload":"0.68","vmc":"0.25","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-18","pic":"FERDIKA - NAGITA","code":"BX551 - 23","size":"BX551","container":"","ket":"CLEANING","mc":"B09L","nilai_preload":"-","vmc":"-","shim_sr":"1.75","shim_pl":"1.40","first_cure":"OK","keterangan":""},{"tgl":"2025-11-18","pic":"MUHAIMIN - DIAZ","code":"BX551 - 20","size":"BX551","container":"","ket":"CLEANING","mc":"B09R","nilai_preload":"-","vmc":"-","shim_sr":"1.40","shim_pl":"0","first_cure":"OOR","keterangan":""},{"tgl":"2025-11-18","pic":"SYAIFUL - HARIS","code":"B491 - 02","size":"B491","container":"","ket":"START","mc":"C13R","nilai_preload":"-0.31","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK ( DMGM )","keterangan":""},{"tgl":"2025-11-18","pic":"MUHAIMIN - DIAZ","code":"BX774 - 03","size":"BX774","container":"","ket":"START","mc":"F03L","nilai_preload":"-0.30","vmc":"0.30","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-18","pic":"AGUS - UBAY","code":"AS264 - 01","size":"AS264","container":"","ket":"CLEANING","mc":"G12L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-18","pic":"AGUS - UBAY","code":"AS264 - 01","size":"AS264","container":"","ket":"CLEANING","mc":"G12L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-18","pic":"TOHANA - IDAD","code":"AS268 - 01","size":"AS268","container":"","ket":"CLEANING","mc":"G12R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2025-11-18","pic":"UHAIMIN - DIAZ","code":"B517 - 03","size":"B517","container":"","ket":"START","mc":"G24L","nilai_preload":"0.50","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-18","pic":"FERDIKA - NAGITA","code":"B517 - 02","size":"B517","container":"","ket":"START","mc":"G24R","nilai_preload":"0.57","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-18","pic":"PEBRI - RENDI","code":"XP175 - 01","size":"XP175","container":"","ket":"START","mc":"H1L","nilai_preload":"0.01","vmc":"0","shim_sr":"0.70","shim_pl":"0.70","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-18","pic":"TOHANA - IDAD","code":"A580 - 02","size":"A580","container":"","ket":"START","mc":"H29L","nilai_preload":"-0.23","vmc":"0","shim_sr":"0","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-18","pic":"TOHANA - IDAD","code":"A580 - 01","size":"A580","container":"","ket":"START","mc":"H29R","nilai_preload":"-0.20","vmc":"0","shim_sr":"0","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-18","pic":"SYAIFUL - HARIS","code":"XS438 - 01","size":"XS438","container":"","ket":"CLEANING","mc":"K25L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OSD )","keterangan":""},{"tgl":"2025-11-18","pic":"TOHANA - IDAD","code":"XM369 - 02","size":"XM369","container":"","ket":"START","mc":"K27L","nilai_preload":"0.57","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OS )","keterangan":""},{"tgl":"2025-11-18","pic":"TOHANA - IDAD","code":"XM369 - 01","size":"XM369","container":"","ket":"START","mc":"K27R","nilai_preload":"0.56","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-18","pic":"SYAIFUL - HARIS","code":"B436 - 02","size":"B436","container":"","ket":"START","mc":"K31R","nilai_preload":"0.45","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-18","pic":"FERDIKA - NAGITA","code":"B632 - 01","size":"B632","container":"","ket":"START","mc":"L02L","nilai_preload":"0.58","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-18","pic":"SYAIFUL - HARIS","code":"B436 - 02","size":"B436","container":"","ket":"START","mc":"L22R","nilai_preload":"0.45","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OS + OFSH )","keterangan":""},{"tgl":"2025-11-18","pic":"MUHAIMIN - DIAZ","code":"BM371 - 01","size":"BM371","container":"","ket":"START","mc":"M23R","nilai_preload":"0.51","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-18","pic":"AGUS - UBAY","code":"A296 - 02","size":"A296","container":"","ket":"START","mc":"N14R","nilai_preload":"0.60","vmc":"0.25","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-17","pic":"DWI - SATRIA","code":"AS127 - 02","size":"AS127","container":"","ket":"CLEANING","mc":"A21L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-17","pic":"SIGIT - DIVA","code":"AS127 - 01","size":"AS127","container":"","ket":"START","mc":"A21R","nilai_preload":"0.26","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-17","pic":"TOHANA - IDAD","code":"B739 - 03","size":"B739","container":"","ket":"CLEANING","mc":"A24L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK ( OFSH + DMGB )","keterangan":""},{"tgl":"2025-11-17","pic":"AGUS - UBAY","code":"B739 - 01","size":"B739","container":"","ket":"CLEANING","mc":"A24R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"1.40","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-17","pic":"SYAIFUL - HARIS","code":"B517 - 04","size":"B517","container":"","ket":"START","mc":"A25L","nilai_preload":"0.49","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-17","pic":"SYAIFUL - HARIS","code":"B517 - 01","size":"B517","container":"","ket":"START","mc":"A25R","nilai_preload":"0.60","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( DM SW⬆️ )","keterangan":""},{"tgl":"2025-11-17","pic":"TOHANA - IDAD","code":"BX551 - 02","size":"BX551","container":"","ket":"PERSIAPAN","mc":"B10R","nilai_preload":"0.65","vmc":"0","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-17","pic":"TOHANA - IDAD","code":"BX551 - 02","size":"BX551","container":"","ket":"START","mc":"B10R","nilai_preload":"0.65","vmc":"0","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-17","pic":"SYAIFUL - HARIS","code":"B356 - 04","size":"B356","container":"","ket":"CLEANING","mc":"D04L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0.35","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-17","pic":"AOP - IRFAN","code":"B356 - 17","size":"B356","container":"","ket":"CLEANING","mc":"D04R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2025-11-17","pic":"AGUS - UBAY","code":"B490 - 02","size":"B490","container":"","ket":"START","mc":"E01R","nilai_preload":"0.5","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-17","pic":"AGUS - UBAY","code":"B529 - 01","size":"B529","container":"","ket":"PERSIAPAN","mc":"E10L","nilai_preload":"-0.19","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-17","pic":"AGUS - UBAY","code":"B529 - 01","size":"B529","container":"","ket":"START","mc":"E10L","nilai_preload":"-0.19","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-17","pic":"TOHANA - IDAD","code":"BX774 - 06","size":"BX774","container":"","ket":"START","mc":"F03R","nilai_preload":"0.47","vmc":"0.30","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-17","pic":"DWI - SATRIA","code":"AS050 - 02","size":"AS050","container":"","ket":"START","mc":"G30L","nilai_preload":"0.80","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OOR  )","keterangan":""},{"tgl":"2025-11-17","pic":"DWI - SATRIA","code":"AS116 - 01","size":"AS116","container":"","ket":"START","mc":"H04L","nilai_preload":"0.44","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( DMGM + DM )","keterangan":""},{"tgl":"2025-11-17","pic":"DWI - SATRIA","code":"AS116 - 01","size":"AS116","container":"","ket":"START","mc":"H04L","nilai_preload":"0.44","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( DMGM + DM )","keterangan":""},{"tgl":"2025-11-17","pic":"SIGIT - DIVA","code":"AS116 - 02","size":"AS116","container":"","ket":"START","mc":"H04R","nilai_preload":"0.60","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( DMGM + DM )","keterangan":""},{"tgl":"2025-11-17","pic":"SIGIT - DIVA","code":"AS116 - 02","size":"AS116","container":"","ket":"START","mc":"H04R","nilai_preload":"0.60","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( DMGM + DM )","keterangan":""},{"tgl":"2025-11-17","pic":"DWI - SATRIA","code":"XS215 - 02","size":"XS215","container":"","ket":"START","mc":"H23L","nilai_preload":"0.47","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-17","pic":"SIGIT - DIVA","code":"XS215 - 01","size":"XS215","container":"","ket":"START","mc":"H23R","nilai_preload":"0.40","vmc":"--","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-17","pic":"AGUS - UBAY","code":"B463 - 01","size":"B463","container":"","ket":"START","mc":"L24R","nilai_preload":"0.60","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-17","pic":"SIGIT - DIVA","code":"B766 - 01","size":"B766","container":"","ket":"START","mc":"N21L","nilai_preload":"01.09","vmc":"0","shim_sr":"1.05","shim_pl":"0","first_cure":"OOR + OFSH⬆️","keterangan":""},{"tgl":"2025-11-17","pic":"TOHANA - IDAD","code":"B754 - 01","size":"B754","container":"","ket":"START","mc":"N21R","nilai_preload":"-0.09","vmc":"0.28","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-17","pic":"TOHANA - IDAD","code":"AA086 - 01","size":"AA086","container":"","ket":"START","mc":"N25L","nilai_preload":"0.45","vmc":"0.65","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-17","pic":"AOP - IRFAN","code":"B451 - 01","size":"B451","container":"","ket":"START","mc":"N27L","nilai_preload":"-0.26","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-16","pic":"DWI - SATRIA","code":"XS115 - 01","size":"XS115","container":"","ket":"START","mc":"D03L","nilai_preload":"-0.56","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-16","pic":"DWI - SATRIA","code":"XS115 - 02","size":"XS115","container":"","ket":"START","mc":"D03R","nilai_preload":"0.66","vmc":"-","shim_sr":"0.35","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-16","pic":"SIGIT - DIVA","code":"XS222 - 02","size":"XS222","container":"","ket":"CLEANING","mc":"D12L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"1.75","first_cure":"OK","keterangan":""},{"tgl":"2025-11-16","pic":"SIGIT - DIVA","code":"XS222 - 07","size":"XS222","container":"","ket":"CLEANING","mc":"D12R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-16","pic":"DWI - SATRIA","code":"XA019 - 01","size":"XA019","container":"","ket":"START","mc":"F09R","nilai_preload":"-0.14","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-15","pic":"AOP - IRFAN","code":"B356 - 06","size":"B356","container":"","ket":"CLEANING","mc":"A11L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2025-11-15","pic":"AOP - IRFAN","code":"B356 - 16","size":"B356","container":"","ket":"CLEANING","mc":"A11R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"2.10","first_cure":"OOR","keterangan":""},{"tgl":"2025-11-15","pic":"MUHAIMIN - DIAZ","code":"BX551 - 11","size":"BX551","container":"","ket":"CLEANING","mc":"B06L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-15","pic":"MUHAIMIN - DIAZ","code":"BX551 - 24","size":"BX551","container":"","ket":"CLEANING","mc":"B06R","nilai_preload":"-","vmc":"-","shim_sr":"1.75","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-15","pic":"SYAIFUL - HARIS","code":"B256 - 06","size":"B256","container":"","ket":"CLEANING","mc":"D15L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-15","pic":"SYAIFUL - HARIS","code":"B256 - 06","size":"B256","container":"","ket":"CLEANING","mc":"D15L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-15","pic":"SYAIFUL - HARIS","code":"B256 - 07","size":"B256","container":"","ket":"CLEANING","mc":"D15R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-15","pic":"DWI - SATRIA","code":"BX774 - 07","size":"BX774","container":"","ket":"START","mc":"F02L","nilai_preload":"0.35","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-15","pic":"SIGIT - DIVA","code":"BX774 - 04","size":"BX774","container":"","ket":"START","mc":"F02R","nilai_preload":"0.72","vmc":"0","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-15","pic":"AOP - IRFAN","code":"XA059 - 01","size":"XA059","container":"","ket":"START","mc":"F14L","nilai_preload":"-0.45","vmc":"0","shim_sr":"0","shim_pl":"1,05","first_cure":"OK","keterangan":""},{"tgl":"2025-11-15","pic":"AOP - IRFAN","code":"XA059 - 01","size":"XA059","container":"","ket":"START","mc":"F14L","nilai_preload":"-0.45","vmc":"0","shim_sr":"0","shim_pl":"1,05","first_cure":"OK","keterangan":""},{"tgl":"2025-11-15","pic":"AOP - IRFAN","code":"XA059 - 01","size":"XA059","container":"","ket":"START","mc":"F14L","nilai_preload":"-0.45","vmc":"0","shim_sr":"0","shim_pl":"1,05","first_cure":"OK","keterangan":""},{"tgl":"2025-11-15","pic":"AOP - IRFAN","code":"XA059 - 02","size":"XA059","container":"","ket":"START","mc":"F14R","nilai_preload":"-0.76","vmc":"0","shim_sr":"0","shim_pl":"1.4","first_cure":"OK","keterangan":""},{"tgl":"2025-11-15","pic":"DWI - SATRIA","code":"B644 - 03","size":"B644","container":"","ket":"START","mc":"G03L","nilai_preload":"0.37","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-15","pic":"DWI - SATRIA","code":"B644 - 01","size":"B644","container":"","ket":"START","mc":"G03R","nilai_preload":"0.11","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2025-11-15","pic":"SIGIT - DIVA","code":"BX675 - 06","size":"BX675","container":"","ket":"START","mc":"K08L","nilai_preload":"0.47","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-14","pic":"DWI - SATRIA","code":"BX256 - 02","size":"BX256","container":"","ket":"CLEANING","mc":"A04L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2025-11-14","pic":"SIGIT - DIVA","code":"BX256 - 04","size":"BX256","container":"","ket":"CLEANING","mc":"A04R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"OK ( DMGM + OFBR )","keterangan":""},{"tgl":"2025-11-14","pic":"FERDIKA - NAGITA","code":"XS222 - 05","size":"XS222","container":"","ket":"CLEANING","mc":"A06L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-14","pic":"FERDIKA - NAGITA","code":"XS222 - 08","size":"XS222","container":"","ket":"CLEANING","mc":"A06R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-14","pic":"DWI - SATRIA","code":"BX551 - 14","size":"BX551","container":"","ket":"CLEANING","mc":"B03L","nilai_preload":"-","vmc":"-","shim_sr":"1.4","shim_pl":"1.4","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-14","pic":"DWI - SATRIA","code":"BX551 - 32","size":"BX551","container":"","ket":"CLEANING","mc":"B03R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-14","pic":"SIGIT - DIVA","code":"B356 - 20","size":"B356","container":"","ket":"CLEANING","mc":"F05L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-14","pic":"SIGIT - DIVA","code":"B356 - 18","size":"B356","container":"","ket":"CLEANING","mc":"F05R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OFBR )","keterangan":""},{"tgl":"2025-11-14","pic":"DWI - SATRIA","code":"B644 - 03","size":"B644","container":"","ket":"START","mc":"G03L","nilai_preload":"0.37","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-14","pic":"DWI - SATRIA","code":"B644 - 01","size":"B644","container":"","ket":"START","mc":"G03R","nilai_preload":"0.11","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2025-11-13","pic":"SIGIT - SYAIFUL","code":"B501 - 01","size":"B501","container":"","ket":"START","mc":"F08L","nilai_preload":"0.48","vmc":"0","shim_sr":"0.35","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-13","pic":"AGUS - UBAY","code":"B501 - 02","size":"B501","container":"","ket":"START","mc":"F08R","nilai_preload":"0.66","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-13","pic":"TOHANA - IDAD","code":"BX675 - 05","size":"BX675","container":"","ket":"START","mc":"G07L","nilai_preload":"-0.58","vmc":"0","shim_sr":"0","shim_pl":"1.05","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-13","pic":"TOHANA - IDAD","code":"BX675 - 01","size":"BX675","container":"","ket":"START","mc":"G07R","nilai_preload":"-0.28","vmc":"0","shim_sr":"0","shim_pl":"0.70","first_cure":"OK ( OOR )","keterangan":""},{"tgl":"2025-11-13","pic":"DWI - SATRIA","code":"BX675 - 01","size":"BX675","container":"","ket":"START","mc":"G14L","nilai_preload":"0.34","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-13","pic":"FERDIKA - NAGITA","code":"BX675 - 04","size":"BX675","container":"","ket":"START","mc":"G14R","nilai_preload":"0.46","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-13","pic":"DWI - SATRIA","code":"BX254 - 01","size":"BX254","container":"","ket":"CLEANING","mc":"H13L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"2.10","first_cure":"OK","keterangan":""},{"tgl":"2025-11-13","pic":"NAGITA - FERDIKA","code":"XS222 - 01","size":"XS222","container":"","ket":"CLEANING","mc":"H13L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-13","pic":"DWI - SATRIA","code":"BX254 - 04","size":"BX254","container":"","ket":"CLEANING","mc":"H13R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"2.10","first_cure":"OK","keterangan":""},{"tgl":"2025-11-13","pic":"MUHAIMIN - DIAZ","code":"XS222 - 04","size":"XS222","container":"","ket":"CLEANING","mc":"H13R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK ( OFSH )","keterangan":""},{"tgl":"2025-11-12","pic":"NAGITA - DIAZ","code":"BX551 - 28","size":"BX551","container":"","ket":"CLEANING","mc":"B11L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-12","pic":"NAGITA - DIAZ","code":"BX551 - 22","size":"BX551","container":"","ket":"CLEANING","mc":"B11R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-11","pic":"AOP - IRFAN","code":"BX551 - 25","size":"BX551","container":"","ket":"CLEANING","mc":"B10L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0.70","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-11","pic":"DWI - SATRIA","code":"B256 - 10","size":"B256","container":"","ket":"CLEANING","mc":"D05L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK ( OS + OFSH )","keterangan":""},{"tgl":"2025-11-11","pic":"DWI - SATRIA","code":"B256 - 12","size":"B256","container":"","ket":"CLEANING","mc":"D05R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK ( OFSH )","keterangan":""},{"tgl":"2025-11-11","pic":"AOP - IRFAN","code":"XA019 - 02","size":"XA019","container":"","ket":"START","mc":"F09L","nilai_preload":"-0.64","vmc":"0.45","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OK ( OFSH )","keterangan":""},{"tgl":"2025-11-11","pic":"SYAIFUL - IRFAN","code":"B254 - 09","size":"B254","container":"","ket":"START","mc":"G02L","nilai_preload":"-0.54","vmc":"-","shim_sr":"0","shim_pl":"1.05","first_cure":"OK ( OFBR + DMGM )","keterangan":""},{"tgl":"2025-11-11","pic":"SYAIFUL - IRFAN","code":"B254 - 10","size":"B254","container":"","ket":"START","mc":"G02R","nilai_preload":"0.42","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OFSH )","keterangan":""},{"tgl":"2025-11-10","pic":"DWI - SATRIA","code":"BX733 - 12","size":"BX733","container":"","ket":"START","mc":"D07L","nilai_preload":"-0.65","vmc":"-","shim_sr":"0","shim_pl":"1.05","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2025-11-09","pic":"SIGIT - DIVA","code":"BX733 - 15","size":"BX733","container":"","ket":"START","mc":"D07R","nilai_preload":"-0.74","vmc":"-","shim_sr":"0","shim_pl":"1.4","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-09","pic":"DWI - SATRIA","code":"BX504 - 01","size":"BX504","container":"","ket":"START","mc":"E12R","nilai_preload":"0.21","vmc":"-","shim_sr":"0.35","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-09","pic":"MUHAIMIN - DIAZ","code":"XA031 - 01","size":"XA031","container":"","ket":"START","mc":"F11L","nilai_preload":"0.43","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OS + OFSH","keterangan":""},{"tgl":"2025-11-09","pic":"FERDIKA - NAGITA","code":"XA031 - 02","size":"XA031","container":"","ket":"START","mc":"F11R","nilai_preload":"0.54","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-11-09","pic":"DWI - SATRIA","code":"BX254 - 11","size":"BX254","container":"","ket":"START","mc":"H02L","nilai_preload":"-0.31","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-09","pic":"DWI - SATRIA","code":"BX254 - 04","size":"BX254","container":"","ket":"START","mc":"H02R","nilai_preload":"-0.66","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-08","pic":"AOP - IRFAN","code":"B256 - 11","size":"B256","container":"","ket":"CLEANING","mc":"F13L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( DMGM + DM )","keterangan":""},{"tgl":"2025-11-08","pic":"SYAIFUL - HARIS","code":"B256 - 14","size":"B256","container":"","ket":"CLEANING","mc":"F13R","nilai_preload":"-","vmc":"-","shim_sr":"2.1","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-08","pic":"FERDIKA - NAGITA","code":"BX254 - 08","size":"BX254","container":"","ket":"START","mc":"H12L","nilai_preload":"0.42","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OFSH )","keterangan":""},{"tgl":"2025-11-08","pic":"FERDIKA - NAGITA","code":"BX254 - 07","size":"BX254","container":"","ket":"START","mc":"H12R","nilai_preload":"-0.22","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-07","pic":"DWI - SATRIA","code":"BX551 - 30","size":"BX551","container":"","ket":"CLEANING","mc":"B04L","nilai_preload":"-","vmc":"-","shim_sr":"1.40","shim_pl":"2.45","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2025-11-07","pic":"DWI - SATRIA","code":"BX51 - 05","size":"BX51","container":"","ket":"CLEANING","mc":"B04R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-11-07","pic":"AGUS - UBAI","code":"BX642 - 07","size":"BX642","container":"","ket":"CLEANING","mc":"D11L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2025-11-07","pic":"FERDIKA - NAGITA","code":"XS222 - 03","size":"XS222","container":"","ket":"CLEANING","mc":"E13L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-07","pic":"FERDIKA - NAGITA","code":"XS222 - 06","size":"XS222","container":"","ket":"CLEANING","mc":"E13R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-06","pic":"TOHANA - IDAD","code":"BX551 - 35","size":"BX551","container":"","ket":"CLEANING","mc":"B01R","nilai_preload":"-","vmc":"-","shim_sr":"1.4","shim_pl":"1.05","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2025-11-05","pic":"MUHAIMIN - DIAZ","code":"B356 - 05","size":"B356","container":"","ket":"CLEANING","mc":"D02R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( OFSH )","keterangan":""},{"tgl":"2025-11-04","pic":"SIGIT - DIVA","code":"B644 - 05","size":"B644","container":"","ket":"START","mc":"G04L","nilai_preload":"0.11","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2025-11-04","pic":"SIGIT - DIVA","code":"B644 - 06","size":"B644","container":"","ket":"START","mc":"G04R","nilai_preload":"0.66","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-03","pic":"SIGIT - DIVA","code":"B356 - 13","size":"B356","container":"","ket":"CLEANING","mc":"A14R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"1.75","first_cure":"OK ( OFBR )","keterangan":""},{"tgl":"2025-11-03","pic":"AOP - IRFAN","code":"B490 - 01","size":"B490","container":"","ket":"START","mc":"E01L","nilai_preload":"-0.26","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK ( OFSH )","keterangan":""},{"tgl":"2025-11-03","pic":"DWI - SATRIA","code":"B644 - 14","size":"B644","container":"","ket":"START","mc":"G13L","nilai_preload":"0.68","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-11-02","pic":"AOP - IRFAN","code":"XS071 - 01","size":"XS071","container":"","ket":"START","mc":"C09L","nilai_preload":"0.29","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-10-31","pic":"RENDY - DIVA","code":"BX551 - 17","size":"BX551","container":"","ket":"CLEANING","mc":"B15L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2025-10-31","pic":"RENDY - DIVA","code":"BX551 - 26","size":"BX551","container":"","ket":"CLEANING","mc":"B15R","nilai_preload":"-","vmc":"-","shim_sr":"1.75","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2025-10-29","pic":"AOP - IRFAN","code":"BX733 - 16","size":"BX733","container":"","ket":"START","mc":"C11L","nilai_preload":"-0.24","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK ( DM + LL )","keterangan":""},{"tgl":"2025-10-29","pic":"AOP - IRFAN","code":"BX733 - 10","size":"BX733","container":"","ket":"START","mc":"C11R","nilai_preload":"-0.15","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK ( DMGM )","keterangan":""},{"tgl":"2025-10-29","pic":"AOP - IRFAN","code":"BX733 - 08","size":"BX733","container":"","ket":"START","mc":"C12L","nilai_preload":"-0.27","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2025-10-29","pic":"AOP - IRFAN","code":"BX733 - 13","size":"BX733","container":"","ket":"START","mc":"C12R","nilai_preload":"0.48","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2025-10-18","pic":"SYAIFUL - HARIS","code":"B491 - 01","size":"B491","container":"","ket":"START","mc":"C13L","nilai_preload":"-0.13","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2025-10-18","pic":"AOP - IRFAN","code":"B491 - 02","size":"B491","container":"","ket":"START","mc":"C13R","nilai_preload":"0.49","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK ( DMGM )","keterangan":""},{"tgl":"2025-09-17","pic":"DWI - SATRIA","code":"AS048 - 01","size":"AS048","container":"","ket":"START","mc":"G09L","nilai_preload":"0.56","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"","pic":"AGUS - UBAY","code":"BX776 - 05","size":"BX776","container":"","ket":"CLEANING","mc":"A09L","nilai_preload":"-","vmc":"-","shim_sr":"0.35","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"","pic":"DWI - SATRIA","code":"XS180 - 01","size":"XS180","container":"","ket":"CLEANING","mc":"A09R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"","pic":"FERDIKA - NAGITA","code":"BX551 - 17","size":"BX551","container":"","ket":"CLEANING","mc":"B08L","nilai_preload":"-","vmc":"-","shim_sr":"1.40","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"","pic":"FERDIKA - NAGITA","code":"BX551 - 38","size":"BX551","container":"","ket":"CLEANING","mc":"B08R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"","pic":"SIGIT - DIVA","code":"XM395 - 03","size":"XM395","container":"","ket":"START","mc":"D03L","nilai_preload":"0.36","vmc":"0.30","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR MIRING","keterangan":""},{"tgl":"","pic":"MUHAIMIN - DIAZ","code":"XM395 - 04","size":"XM395","container":"","ket":"START","mc":"D03R","nilai_preload":"0.45","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"","pic":"SIGIT - DIVA","code":"B392 - 01","size":"B392","container":"","ket":"CLEANING","mc":"D09L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"","pic":"SIGIT - DIVA","code":"B392 - 02","size":"B392","container":"","ket":"CLEANING","mc":"D09R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"","pic":"MUHAIMIN - DIAZ","code":"XM394 - 02","size":"XM394","container":"","ket":"START","mc":"F07L","nilai_preload":"0.44","vmc":"0.65","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"","pic":"FERDIKA - NAGITA","code":"XM394 - 05","size":"XM394","container":"","ket":"START","mc":"F07R","nilai_preload":"0.48","vmc":"0.65","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"","pic":"FERDIKA - NAGITA","code":"XM394 - 03","size":"XM394","container":"","ket":"START","mc":"F09L","nilai_preload":"0.51","vmc":"0.65","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"","pic":"FERDIKA - NAGITA","code":"XM394 - 06","size":"XM394","container":"","ket":"START","mc":"F09R","nilai_preload":"0.57","vmc":"0.65","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"","pic":"TOHANA - IDAD","code":"XM380 - 02","size":"XM380","container":"","ket":"START","mc":"F14L","nilai_preload":"-0.98","vmc":" 0","shim_sr":"0","shim_pl":"1.40","first_cure":"OK","keterangan":""},{"tgl":"","pic":"UBAI - AGUS","code":"XM380 - 10","size":"XM380","container":"","ket":"START","mc":"F14R","nilai_preload":"-0.40","vmc":"-","shim_sr":"0","shim_pl":"1.05","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"","pic":"SIGIT - DIVA","code":"B503 - 02","size":"B503","container":"","ket":"START","mc":"H04L","nilai_preload":"-0.10","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"","pic":"SIGIT - DIVA","code":"B503 - 01","size":"B503","container":"","ket":"START","mc":"H04R","nilai_preload":"-0.51","vmc":"-","shim_sr":"0","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"","pic":"FERDIKA - NAGITA","code":"XP173 - 01","size":"XP173","container":"","ket":"START","mc":"H08L","nilai_preload":"0.59","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"","pic":"AOP - IRFAN","code":"B628 - 01","size":"B628","container":"","ket":"START","mc":"L04R","nilai_preload":"0.11","vmc":"0.20","shim_sr":"0.35","shim_pl":"0.70","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"","pic":"SIGIT - DIVA","code":"AS260 - 201","size":"AS260","container":"","ket":"CLEANING","mc":"L07R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"","pic":"AOP - IRFAN","code":"B503 - 04","size":"B503","container":"","ket":"START","mc":"M04L","nilai_preload":"0.16","vmc":"0.50","shim_sr":"0","shim_pl":"1.05","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"","pic":"DWI - SATRIA","code":"B503 - 03","size":"B503","container":"","ket":"START","mc":"M04R","nilai_preload":"-0.15","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"","pic":"SYAIFUL - HARIS","code":"AS052 - 01","size":"AS052","container":"","ket":"START","mc":"M06R","nilai_preload":"0.10","vmc":"0.60","shim_sr":"1.05","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"","pic":"TOHANA - IDAD","code":"AX584 - 04","size":"AX584","container":"","ket":"START","mc":"M07L","nilai_preload":"0.57","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"","pic":"TOHANA - IDAD","code":"AX584 - 03","size":"AX584","container":"","ket":"START","mc":"M07R","nilai_preload":"0.53","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"","pic":"MUHAIMIN - DIAZ","code":"BX678 - 10","size":"BX678","container":"","ket":"START","mc":"N12L","nilai_preload":"-0.25","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"","pic":"MUHAIMIN - DIAZ","code":"BX678 - 07","size":"BX678","container":"","ket":"START","mc":"N12R","nilai_preload":"0.13","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"","pic":"TOHANA - IDAD","code":"B644 - 13","size":"B644","container":"","ket":"START","mc":"G05R","nilai_preload":"0.21","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-01-03","pic":"SYAIFUL - TEGAR","code":"BX777 - 01","size":"BX777","container":"","ket":"START","mc":"H15L","nilai_preload":"-0.50","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OK ( IP )","keterangan":""},{"tgl":"2026-01-28","pic":"TOHANA - IDAD","code":"B733 - 23","size":"B733","container":"SM STP - 688","ket":"START","mc":"C11L","nilai_preload":"-0.27","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.30 / BAWAH 1.40"},{"tgl":"2026-01-02","pic":"PEBRI - NAGITA","code":"B322 - 01","size":"B322","container":"","ket":"START","mc":"H11L","nilai_preload":"0.06","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OS + DM + DMGM","keterangan":""},{"tgl":"2026-01-02","pic":"PEBRI - NAGITA","code":"B322 - 04","size":"B322","container":"","ket":"START","mc":"H11R","nilai_preload":"0.12","vmc":"","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OS","keterangan":""},{"tgl":"2026-01-28","pic":"AGUS - UBAY","code":"B733 - 09","size":"B733","container":"GT STP - 286","ket":"START","mc":"C11R","nilai_preload":"-0.14","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.70","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.10 / BAWAH 1.60"},{"tgl":"2026-01-03","pic":"IRFAN - AOP","code":"B554 - 02","size":"B554","container":"","ket":"START","mc":"H05R","nilai_preload":"-0.22","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK ( CP )","keterangan":""},{"tgl":"2026-01-02","pic":"PEBRI - NAGITA","code":"A323 - 02","size":"A323","container":"","ket":"START","mc":"L05L","nilai_preload":"-0.52","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-01-05","pic":"TOHANA - IDAD","code":"B708 - 01","size":"B708","container":"","ket":"START","mc":"K08L","nilai_preload":"1.15","vmc":"-","shim_sr":"1.40","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-29","pic":"AGUS - UBAY","code":"B459 - 01","size":"B459","container":"","ket":"START","mc":"K09L","nilai_preload":"0.56","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-04","pic":"SIGIT - DIVA","code":"B601 - 03","size":"B601","container":"","ket":"CLEANING","mc":"K12L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-04","pic":"SIGIT - DIVA","code":"B601 - 04","size":"B601","container":"","ket":"CLEANING","mc":"K12R","nilai_preload":"-","vmc":"-","shim_sr":"2.10","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-29","pic":"HARIS - BADRU","code":"XS195 - 01","size":"XS195","container":"","ket":"START","mc":"L11R","nilai_preload":"0.52","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-07","pic":"AGUS - UBAY","code":"BX776 - 05","size":"BX776","container":"","ket":"EX REPAIR","mc":"C15L","nilai_preload":"0.42","vmc":"","shim_sr":"0.70","shim_pl":"0-","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.40 / BAWAH 1.80"},{"tgl":"2026-12-29","pic":"HARIS - BADRU","code":"B700 - 02","size":"B700","container":"","ket":"START","mc":"L07L","nilai_preload":"0.62","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-04","pic":"SYAIFUL - HARIS","code":"AS097 - 02","size":"AS097","container":"","ket":"START","mc":"M12L","nilai_preload":"0.56","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-04","pic":"TOHANA - IDAD","code":"AS097 - 01","size":"AS097","container":"","ket":"START","mc":"M12R","nilai_preload":"1.18","vmc":"-","shim_sr":"1.40","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-04","pic":"SIGIT - DIVA","code":"A291 - 02","size":"A291","container":"","ket":"START","mc":"N09L","nilai_preload":"0.28","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK ( DMGM )","keterangan":""},{"tgl":"2026-01-04","pic":"DWI - SATRIA","code":"A291 - 01","size":"A291","container":"","ket":"START","mc":"N09R","nilai_preload":"-1.06","vmc":"-","shim_sr":"0..70","shim_pl":"1.40","first_cure":"OK","keterangan":""},{"tgl":"2026-01-03","pic":"IRFAN - AOP ","code":"B444 - 03","size":"B444","container":"","ket":"START","mc":"N06L","nilai_preload":"0.62","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-04","pic":"SYAIFUL - HARIS","code":"B518 - 01","size":"B518","container":"","ket":"START","mc":"L21L","nilai_preload":"0.89","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-03","pic":"NAGITA","code":"XS226 - 01","size":"XS226","container":"","ket":"START","mc":"L25L","nilai_preload":"-0.55","vmc":"","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-01-04","pic":"IRFAN - AOP","code":"B518 - 03","size":"B518","container":"","ket":"START","mc":"L26L","nilai_preload":"0.18","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR MIRING","keterangan":""},{"tgl":"2026-01-02","pic":"MUHAIMIN - DIAZ","code":"AS136 - 04","size":"AS136","container":"","ket":"START","mc":"K27R","nilai_preload":"0.56","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( LS + FM + DM )","keterangan":""},{"tgl":"2026-12-18","pic":"DWI - SATRIA","code":"BX776 - 01","size":"BX776","container":"","ket":"PINDAHAN","mc":"C15R","nilai_preload":"-0.77","vmc":"-","shim_sr":"0.70","shim_pl":"1.40","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.90 / BAWAH 1.40 + OOR"},{"tgl":"2026-12-29","pic":"IRFAN - AOP","code":"AA095 - 06","size":"AA095","container":"","ket":"START","mc":"K24L","nilai_preload":"0.61","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-03","pic":"SIGIT - DIVA","code":"AA095 - 08","size":"AA095","container":"","ket":"START","mc":"K23R","nilai_preload":"0.91","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-29","pic":"AGUS - UBAY","code":"AS255 - 01","size":"AS255","container":"","ket":"CLEANING","mc":"G29L","nilai_preload":"-","vmc":"-","shim_sr":"2.45","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2026-12-29","pic":"AGUS - UBAY","code":"AS294 - 01","size":"AS294","container":"","ket":"CLEANING","mc":"G29L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-12-29","pic":"TOHANA - IDAD","code":"A883 - 02","size":"A883","container":"","ket":"START","mc":"G22L","nilai_preload":"-0.42","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OFSH","keterangan":""},{"tgl":"2026-12-29","pic":"TOHANA - IDAD","code":"A883 - 04","size":"A883","container":"","ket":"START","mc":"G22R","nilai_preload":"0.39","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-28","pic":"AOP - IRFAN","code":"XS178 - 02","size":"XS178","container":"","ket":"CLEANING","mc":"F30R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"1.75","first_cure":"OK","keterangan":""},{"tgl":"2026-12-29","pic":"MUHAIMIN - DIAZ","code":"A748 - 03","size":"A748","container":"","ket":"CLEANING","mc":"B27L","nilai_preload":"-","vmc":"-","shim_sr":"2.10","shim_pl":"1.40","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-12-29","pic":"FERDIKA - NAGITA","code":"A748 - 02","size":"A748","container":"","ket":"CLEANING","mc":"B27R","nilai_preload":"-","vmc":"-","shim_sr":"2.10","shim_pl":"0","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-01-05","pic":"TOHANA - IDAD","code":"AS238 - 01","size":"AS238","container":"","ket":"CLEANING","mc":"A29L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OS","keterangan":""},{"tgl":"2026-01-05","pic":"TOHANA - IDAD","code":"AS238 - 02","size":"AS238","container":"","ket":"CLEANING","mc":"A29R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-21","pic":"DWI - SATRIA","code":"AA096 - 01","size":"AA096","container":"","ket":"START","mc":"N26R","nilai_preload":"0.21","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-12-29","pic":"MUHAIMIN - DIAZ","code":"BX333 - 03","size":"BX333","container":"","ket":"START","mc":"H08L","nilai_preload":"0.42","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-12-29","pic":"MUHAIMIN - DIAZ","code":"BX333 - 04","size":"BX333","container":"","ket":"START","mc":"H08R","nilai_preload":"0.14","vmc":"--","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OFSH + DM","keterangan":""},{"tgl":"2026-12-29","pic":"MUHAIMIN - DIAZ","code":"A291 - 04","size":"A291","container":"","ket":"START","mc":"N01R","nilai_preload":"0.71","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OS + OFSH + OFBR","keterangan":""},{"tgl":"2026-12-28","pic":"TOHANA - IDAD","code":"BX638 - 05","size":"BX638","container":"","ket":"START","mc":"A27L ","nilai_preload":"0.16","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-12-29","pic":"IRFAN - AOP","code":"BX638 - 02","size":"BX638","container":"","ket":"EX REPAIR","mc":"A27R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-01-04","pic":"AGUS - UBAY","code":"AA081 - 03","size":"AA081","container":"","ket":"START","mc":"G26L ","nilai_preload":"1.65","vmc":"-","shim_sr":"1.75","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-04","pic":"TOHANA - IDAD","code":"AA081 - 01","size":"AA081","container":"","ket":"START","mc":"G26R","nilai_preload":"-0.41","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-01-09","pic":"SIGIT - DIVA","code":"BX776 - 04","size":"BX776","container":"SM STP - 75","ket":"START","mc":"C15R","nilai_preload":"-1.19","vmc":"-","shim_sr":"0.70","shim_pl":"1.75","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.50 / BAWAH 2.00"},{"tgl":"2026-01-03","pic":"MUHAIMIN - DIAZ","code":"BX328 - 08","size":"BX328","container":"","ket":"START","mc":"D05L","nilai_preload":"-0.46","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 3.00 / SPEC MOR 1.80 + OFSH"},{"tgl":"2026-01-03","pic":"MUHAIMIN - DIAZ","code":"BX328 - 09","size":"BX328","container":"","ket":"START","mc":"D05R","nilai_preload":"0.42","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 3.20 / SPEC MOR 1.80 + OFSH"},{"tgl":"2026-01-20","pic":"MUHAIMIN - DIAZ","code":"XM392 - 01","size":"XM392","container":"TYAN STP - 668","ket":"START","mc":"D06L","nilai_preload":"-","vmc":"-","shim_sr":"?","shim_pl":"?","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -0.90 / BAWAH 0"},{"tgl":"2026-01-07","pic":"MUHAIMIN - DIAZ","code":"BX643 - 06","size":"BX643","container":"","ket":"START","mc":"N08R","nilai_preload":"-","vmc":"-","shim_sr":"1.75","shim_pl":"0","first_cure":"OK ( DMGM + DM )","keterangan":""},{"tgl":"2026-01-10","pic":"SIGIT - DIVA","code":"BX328 - 06","size":"BX328","container":"SM STP - 498","ket":"CLEANING","mc":"D07L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.30 / BAWAH 1.40"},{"tgl":"2026-12-18","pic":"SIGIT - DIVA","code":"B490 - 02","size":"B490","container":"","ket":"CLEANING","mc":"E01R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.60 / BAWAH 1.40"},{"tgl":"2026-01-20","pic":"MUHAIMIN - DIAZ","code":"XM395 - 04","size":"XM395","container":"AZ STP - 340","ket":"START","mc":"E02R","nilai_preload":"-","vmc":"-","shim_sr":"?","shim_pl":"?","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.80 / BAWAH 1.09"},{"tgl":"2026-01-12","pic":"SYAIFUL - HARIS","code":"BX734 - 04","size":"BX734","container":"AZ STP - 92","ket":"START","mc":"E03R","nilai_preload":"1.13","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.40 / BAWAH 1.60"},{"tgl":"2026-12-18","pic":"HARIS - BADRU","code":"XM379 - 01","size":"XM379","container":"","ket":"START","mc":"M25L","nilai_preload":"0.57","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-08","pic":"SIGIT - DIVA","code":"A759 - 02","size":"A759","container":"","ket":"CLEANING","mc":"B30L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-08","pic":"DWI - SATRIA","code":"A759 - 01","size":"A759","container":"","ket":"CLEANING","mc":"B30R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-08","pic":"AGUS - UBAY","code":"BX677 - 01","size":"BX677","container":"","ket":"CLEANING","mc":"F28L","nilai_preload":"-0.23","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OFSH","keterangan":""},{"tgl":"2026-01-08","pic":"AGUS - UBAY","code":"BX677 - 06","size":"BX677","container":"","ket":"CLEANING","mc":"F28R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-01-06","pic":"MUHAIMIN - DIAZ","code":"AS264 - 01","size":"AS264","container":"","ket":"START","mc":"L26L","nilai_preload":"0.46","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-12-18","pic":"SYAIFUL - TEGAR","code":"BM397 - 02","size":"BM397","container":"","ket":"START","mc":"E04R","nilai_preload":"0.53","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -0.40 / BAWAH 2.10"},{"tgl":"2026-01-09","pic":"IRFAN - AOP","code":"B505 - 01","size":"B505","container":"AZ STP - 185","ket":"CLEANING","mc":"A06L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-01-09","pic":"SIGIT - DIVA","code":"B505 - 02","size":"B505","container":"SM STP - 459","ket":"START","mc":"A06R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"1.05","first_cure":"OK ( DM + MS )","keterangan":""},{"tgl":"2026-01-09","pic":"DWI - SATRIA","code":"B256 - 01","size":"B256","container":"AZ STP - 18","ket":"CLEANING","mc":"A08L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-01-12","pic":"DWI - SATRIA","code":"XM418 - 01","size":"XM418","container":"SM STP - 288","ket":"START","mc":"E11L","nilai_preload":"-0.71","vmc":"","shim_sr":"-0.70","shim_pl":"1,05","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -1.00 / BAWAH 0"},{"tgl":"2026-01-11","pic":"IDAD - DIVA","code":"BX776 - 02","size":"BX776","container":"SM STP - 444","ket":"START","mc":"A09L","nilai_preload":"-0.88","vmc":"-","shim_sr":"0.70","shim_pl":"1.40","first_cure":"OK","keterangan":""},{"tgl":"2026-01-09","pic":"SYAIFUL - HARIS","code":"XS180 - 01","size":"XS180","container":"HM AZIII/2 - 52","ket":"CLEANING","mc":"A09R","nilai_preload":"-","vmc":"-","shim_sr":"-","shim_pl":"-","first_cure":"MOR + OOR + DM","keterangan":""},{"tgl":"2026-01-11","pic":"AGUS - UBAY","code":"XS180 - 01","size":"XS180","container":"SM AZIII/2 - 589","ket":"GACON","mc":"A09R","nilai_preload":"0.55","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-30","pic":"PEBRI - DIKA","code":"B356 - 03","size":"B356","container":"SM STP - 693","ket":"START","mc":"E12R","nilai_preload":"0.50","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.00 / BAWAH 2.20"},{"tgl":"2026-01-11","pic":"SYAIFUL - HARIS","code":"B256 - 03","size":"B256","container":"SM STP - 338","ket":"CLEANING","mc":"C14L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"1.05","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-01-10","pic":"MUHAIMIN - DIAZ","code":"B256 - 09","size":"B256","container":"SM STP - 79","ket":"CLEANING","mc":"C14R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-13","pic":"PEBRI - DIKA","code":"BM397 - 01","size":"BM397","container":"SM STP - 431","ket":"START","mc":"F02L","nilai_preload":"0.24","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"MOR OUT SPEC","keterangan":"ATAS 0.60 / BAWAH -0.40"},{"tgl":"2026-01-09","pic":"DWI - SATRIA","code":"BX733 - 13","size":"BX733","container":"SM STP - 381","ket":"CLEANING","mc":"C07L","nilai_preload":"-0.02","vmc":"-","shim_sr":"0.35","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-01-09","pic":"SIGIT - DIVA","code":"BX733 - 24","size":"BX733","container":"AZ STP - 439","ket":"CLEANING","mc":"C07R","nilai_preload":"0","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-01-13","pic":"PEBRI - DIKA","code":"BM397 - 02","size":"BM397","container":"SM STP - 394","ket":"START","mc":"F02R","nilai_preload":"-0.19","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 1.00 / BAWAH -0.30"},{"tgl":"2026-01-03","pic":"DWI - SATRIA","code":"XM398 - -01","size":"","container":"","ket":"START","mc":"F11L","nilai_preload":"-0.67","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 0.20"},{"tgl":"2026-01-03","pic":"DWI - SATRIA","code":"XM398 - 02","size":"XM398","container":"","ket":"START","mc":"F11R","nilai_preload":"-1.00","vmc":"-","shim_sr":"0.70","shim_pl":"1.75","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 0.60"},{"tgl":"2026-01-10","pic":"DWI - SATRIA","code":"BX328 - 04","size":"BX328","container":"SM STP - 140","ket":"CLEANING","mc":"D07R","nilai_preload":"-","vmc":"--","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-13","pic":"SYAIFUL - HARIS","code":"XM391 - 02","size":"XM391","container":"SM STP - 380","ket":"START","mc":"D09R","nilai_preload":"-0.19","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"","keterangan":""},{"tgl":"2026-01-12","pic":"TOHANA - IDAD","code":"BX504 - 03","size":"BX504","container":"SM STP - 611","ket":"START","mc":"D12L","nilai_preload":"1.23","vmc":"-","shim_sr":"1.40","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-12","pic":"TOHANA - IDAD","code":"BX504 - 04","size":"BX504","container":"SM STP - 749","ket":"START","mc":"D12R","nilai_preload":"0.16","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-01-12","pic":"AGUS - UBAY","code":"BX475 - 01","size":"BX475","container":"AZ STP - 624","ket":"START","mc":"E13R","nilai_preload":"0.43","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DMGM )","keterangan":""},{"tgl":"2026-01-10","pic":"SIGIT - DIVA","code":"XM420 - 02","size":"XM420","container":"SM STP - 155","ket":"START","mc":"E12L","nilai_preload":"-0.84","vmc":"-","shim_sr":"0.70","shim_pl":"1.40","first_cure":"OK ( LL )","keterangan":""},{"tgl":"2026-01-10","pic":"DWI - SATRIA","code":"XM420 - 01","size":"XM420","container":"SM STP - 142","ket":"START","mc":"E12R","nilai_preload":"1.75","vmc":"-","shim_sr":"1.75","shim_pl":"0","first_cure":"OK ( LL )","keterangan":""},{"tgl":"2026-01-22","pic":"SYAIFUL - MUHAIMIN","code":"BX497 - 02","size":"BX497","container":"SM STP - 161","ket":"START","mc":"F21R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.2 / BAWH 1.50 + OF + DM"},{"tgl":"2026-01-12","pic":"SYAIFUL - HARIS","code":"BX734 - 05","size":"BX734","container":"SM STP - 374","ket":"START","mc":"E03L","nilai_preload":"0.46","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-17","pic":"FERDIKA - NAGITA","code":"BX677 - 09","size":"BX677","container":"","ket":"CLEANING","mc":"F24R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR BAWAH MIRING","keterangan":"MOR BAWAH 2.70 / Spec MOR 1.70"},{"tgl":"2026-12-29","pic":"IRFAN - HARIS","code":"BX726 - 01","size":"BX726","container":"","ket":"CLEANING","mc":"G08R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"1.05","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.50 / BAWAH 1.50"},{"tgl":"2026-01-10","pic":"SIGIT - DIVA","code":"BX499 - 01","size":"BX499","container":"SM STP - 116","ket":"CLEANING","mc":"G09R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.8 / BAWAH 2.00"},{"tgl":"2026-01-11","pic":"AOP - IRFAN","code":"XM395 - 02","size":"XM395","container":"SM STP - 86","ket":"START","mc":"F08L","nilai_preload":"-0.43","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-01-12","pic":"SIGIT - DIVA","code":"BX504 - 02","size":"BX504","container":"SM STP - 695","ket":"START","mc":"F11L","nilai_preload":"0.55","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-12","pic":"DWI - SATRIA","code":"BX504 - 01","size":"BX504","container":"GT STP - 272","ket":"START","mc":"F11R","nilai_preload":"-0.55","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-01-10","pic":"DWI - SATRIA","code":"BX499 - 02","size":"BX499","container":"GT STP","ket":"CLEANING","mc":"G09L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-01-22","pic":"PEBRI - DIKA","code":"BX553 - 06","size":"BX553","container":"AZ STP - 506","ket":"START","mc":"G09R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.50 / BAWAH 2.10 + OS"},{"tgl":"2026-01-19","pic":"AMIR - AGUS","code":"B845 - 02","size":"B845","container":"SM AZIII/2 - 790","ket":"START","mc":"G10R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.80 / BAWAH 2.00"},{"tgl":"2026-01-08","pic":"AGUS - UBAY","code":"BX675 - 04","size":"BX675","container":"GT AZIII - 136","ket":"START","mc":"G11L","nilai_preload":"0.56","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR OUT SPEC","keterangan":"ATAS -0.70 / BAWAH -0.40"},{"tgl":"2026-01-12","pic":"SYAIFUL - HARIS","code":"BX726 - 02","size":"BX726","container":"SM STP - 121","ket":"CLEANING","mc":"G08L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-12","pic":"SYAIFUL - HARIS","code":"BX726 - 01","size":"BX726","container":"GT STP - 320","ket":"CLEANING","mc":"G08R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-01-12","pic":"AOP - IRFAN","code":"BX254 - 09","size":"BX254","container":"GT STP - 371","ket":"CLEANING","mc":"G02L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"2.10","first_cure":"OFSH + DMGM","keterangan":""},{"tgl":"2026-01-12","pic":"AOP - IRFAN","code":"BX254 - 10","size":"BX254","container":"SM STP - 231","ket":"CLEANING","mc":"G02R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"1.05","first_cure":"OFSH + OS","keterangan":""},{"tgl":"2026-01-12","pic":"TOHANA - IDAD","code":"BX254 - 15","size":"BX254","container":"GT STP - 665","ket":"START","mc":"G01L","nilai_preload":"-0.19","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-01-12","pic":"AGUS - UBAY","code":"BX254 - 13","size":"BX254","container":"SM STP - 779","ket":"START","mc":"G01R","nilai_preload":"0.50","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OS","keterangan":""},{"tgl":"2026-01-08","pic":"DWI - SATRIA","code":"BX675 - 03","size":"BX675","container":"GT AZIII - 108","ket":"START","mc":"G11R","nilai_preload":"-0.06","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 1.10 / BAWAH -0.50"},{"tgl":"2026-01-30","pic":"DWI - SATRIA","code":"AS121 - 02","size":"AS121","container":"GT AZIV - 158","ket":"START","mc":"G24L","nilai_preload":"0.42","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR BAWAH KELUAR","keterangan":"ATAS 1.60 / BAWAH 2.90"},{"tgl":"2026-01-13","pic":"SYAIFUL - HARIS","code":"BX254 - 11","size":"BX254","container":"GT STP - 282","ket":"CLEANING","mc":"H02L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OS + OFSH + DMGM","keterangan":""},{"tgl":"2026-01-13","pic":"AOP - IRFAN","code":"BX254 - 04","size":"BX254","container":"SH STP - 69","ket":"CLEANING","mc":"H02R","nilai_preload":"-","vmc":"-","shim_sr":"0,.70","shim_pl":"0","first_cure":"OFSH + OS","keterangan":""},{"tgl":"2026-01-13","pic":"MUHAIMIN - DIAZ","code":"XS222 - 04","size":"XS222","container":"GT STP - 666","ket":"START","mc":"K03R","nilai_preload":"0.41","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OS + OFSH","keterangan":""},{"tgl":"2026-01-13","pic":"SYAIFUL - NAGITA","code":"XM380 - 10","size":"XM380","container":"SM STP - 149","ket":"START","mc":"K03L","nilai_preload":"0.55","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OFSH + DMGM","keterangan":""},{"tgl":"2026-12-18","pic":"DWI - SATRIA","code":"BX776 - 01","size":"BX776","container":"","ket":"PINDAHAN","mc":"C09L","nilai_preload":"-0.77","vmc":"-","shim_sr":"","shim_pl":"","first_cure":"OOR + DMGM + MOR","keterangan":""},{"tgl":"2026-01-08","pic":"FERDIKA - NAGITA","code":"AS270 - 01","size":"AS270","container":"SM STP - 13","ket":"START","mc":"C09R","nilai_preload":"-0.23","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"MOR + OS","keterangan":""},{"tgl":"2026-01-13","pic":"NAGITA - SYAIFUL","code":"B322 - 03","size":"B322","container":"HM AZIV - 663","ket":"START","mc":"H12L","nilai_preload":"0.48","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-13","pic":"NAGITA - SYAIFUL","code":"B322 - 02","size":"B322","container":"HM AZIV - 427","ket":"START","mc":"H12R","nilai_preload":"0.30","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-01-21","pic":"AOP - HARIS","code":"B495 - 02","size":"B495","container":"SM AZIII - 102","ket":"START","mc":"G30R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -0.50 / BAWAH 1.50 + OFSH"},{"tgl":"2026-12-17","pic":"FERDIKA - NAGITA","code":"XM369 - 01","size":"XM369","container":"","ket":"START","mc":"G31L","nilai_preload":"0.49","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS MIRING","keterangan":"MOR ATAS 1.53 / Spec MOR 0.50"},{"tgl":"2026-12-17","pic":"FERDIKA - NAGITA","code":"XM369 - 02","size":"XM369","container":"","ket":"START","mc":"G31R","nilai_preload":"0.54","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS MIRING","keterangan":"MOR ATAS 1.30 / Spec MOR  0.50"},{"tgl":"2026-01-11","pic":"SIGIT - DIVA","code":"BX675 - 01","size":"BX675","container":"SM AZIII - 577","ket":"START","mc":"H01L","nilai_preload":"0.35","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -0.80 / BAWAH -0.50"},{"tgl":"2026-01-06","pic":"AGUS - UBAY","code":"BX675 - 07","size":"BX675","container":"SM AZIII - 631","ket":"START","mc":"H01R","nilai_preload":"-0.73","vmc":"-","shim_sr":"0.70","shim_pl":"1.40","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -0.60 / BAWAH -0.50"},{"tgl":"2026-01-13","pic":"AGUS - UBAY","code":"BX254 - 07","size":"BX254","container":"SM STP - 676","ket":"CLEANING","mc":"H04L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"OS + OFSH","keterangan":""},{"tgl":"2026-01-13","pic":"AGUS - UBAY","code":"BX254 - 08","size":"BX254","container":"SM STP - 675","ket":"CLEANING","mc":"H04R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OS + OFSH","keterangan":""},{"tgl":"2026-01-09","pic":"AOP - IRFAN","code":"B644 - 02","size":"B644","container":"SM STP - 85","ket":"START","mc":"H03L","nilai_preload":"0.60","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( OFBR )","keterangan":""},{"tgl":"2026-01-03","pic":"IRFAN - AOP","code":"B554 - 01","size":"B554","container":"","ket":"START","mc":"H05L","nilai_preload":"0.61","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 3.20 / Max 2.10"},{"tgl":"2026-01-09","pic":"AOP - IRFAN","code":"BX254 - 03","size":"BX254","container":"SM STP - 334","ket":"START","mc":"J03L","nilai_preload":"-0.12","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK ( FM )","keterangan":""},{"tgl":"2026-01-08","pic":"MUHAIMIN - DIAZ","code":"BX254 - 02","size":"BX254","container":"SM STP - 295","ket":"START","mc":"L03R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK ( FM )","keterangan":""},{"tgl":"2026-01-07","pic":"MUHAIMIN - DIAZ","code":"XS222 - 07","size":"XS222","container":"SM STP - 629","ket":"START","mc":"J06L","nilai_preload":"0.15","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-01-10","pic":"SYAIFUL - HARIS","code":"B503 - 04","size":"B503","container":"SM STP - 198","ket":"CLEANING","mc":"H05L","nilai_preload":"-","vmc":"-","shim_sr":"1.-05","shim_pl":"1.40","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.00 / BAWAH 1.60"},{"tgl":"2026-01-08","pic":"SIGIT - DIAZ","code":"B491 - 02","size":"B491","container":"AZ STP - 512","ket":"START","mc":"J12L","nilai_preload":"1.44","vmc":"-","shim_sr":"1.40","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-08","pic":"FERDIKA - NAGITA","code":"B491 - 01","size":"B491","container":"SM STP - 60","ket":"START","mc":"J12R","nilai_preload":"0.44","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR + OFSH","keterangan":""},{"tgl":"2026-01-12","pic":"TOHANA - IDAD","code":"AS142 - 02","size":"AS142","container":"HM L46 - 62","ket":"START","mc":"K09L","nilai_preload":"0.59","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-01-13","pic":"AOP - IRFAN","code":"A580 - 02","size":"A580","container":"HM L46 - 77","ket":"START","mc":"K13L","nilai_preload":"0.55","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-12","pic":"PEBRI - DIKA","code":"B451 - 01","size":"B451","container":"AZ IV - 361","ket":"START","mc":"L11L","nilai_preload":"1.82","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-13","pic":"AGUS - UBAY","code":"A874 - 01","size":"A874","container":"SM L46 - 661","ket":"START","mc":"L10L","nilai_preload":"0.60","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-13","pic":"PEBRI - DIKA","code":"A874 - 03","size":"A874","container":"HM L46 - 108","ket":"START","mc":"L10R","nilai_preload":"0.11","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-09","pic":"MUHAIMIN - DIAZ","code":"A292 - 03","size":"A292","container":"GT L48 - 649","ket":"START","mc":"L08L","nilai_preload":"-0.18","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OS + DMGM + DM","keterangan":""},{"tgl":"2026-01-13","pic":"TOHANA - IDAD","code":"AS156 - 01","size":"AS156","container":"GT L48 - 339","ket":"START","mc":"L05R","nilai_preload":"0.50","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-01-09","pic":"FERDIKA - NAGITA","code":"A292 - 01","size":"A292","container":"SM L48 - 202","ket":"START","mc":"L01L","nilai_preload":"0.75","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-01-11","pic":"SIGIT - DIVA","code":"AA093 - 02","size":"AA093","container":"SM AZIV - 88","ket":"CLEANING","mc":"M03L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OS","keterangan":""},{"tgl":"2026-01-11","pic":"DWI - SATRIA","code":"AA093 - 03","size":"AA093","container":"SH AZIV - 40","ket":"CLEANING","mc":"M03R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-01-12","pic":"MUHAIMIN - DIAZ","code":"AS052 - 01","size":"AS052","container":"HM AZIII - 343","ket":"START","mc":"M11L","nilai_preload":"0.47","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-12","pic":"MUHAIMIN - DIAZ","code":"AS052 - 02","size":"AS052","container":"HM AZIII - 516","ket":"START","mc":"M11R","nilai_preload":"0.43","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-10","pic":"DWI - SATRIA","code":"B503 - 03","size":"B503","container":"SM STP - 331","ket":"START","mc":"H05R","nilai_preload":"-0.47","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.80 / BAWAH 1.60"},{"tgl":"2026-01-11","pic":"DWI - SATRIA","code":"B374 - 02","size":"B374","container":"GT AZIV - 184","ket":"CLEANING","mc":"N05R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"1.40","first_cure":"OK","keterangan":""},{"tgl":"2026-01-09","pic":"DWI - SATRIA","code":"BX726 - 04","size":"BX726","container":"GT STP - 260","ket":"CLEANING","mc":"H07R","nilai_preload":"-","vmc":"-","shim_sr":"1.40","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.80 / BAWAH 1.60"},{"tgl":"2026-01-10","pic":"AOP - IRFAN","code":"BX670 - 01","size":"BX670","container":"GT AZIV - 232","ket":"START","mc":"N03R","nilai_preload":"0.10","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-01-12","pic":"AOP - IRFAN","code":"B744 - 01","size":"B744","container":"AZ STP - 09","ket":"CLEANING","mc":"N02L","nilai_preload":"-","vmc":"-","shim_sr":"1.40","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-01-11","pic":"SYAIFUL - HARIS","code":"A291 - 04","size":"A291","container":"HM AZIII/2 - 63","ket":"START","mc":"N01R","nilai_preload":"0.46","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-09","pic":"FERDIKA - NAGITA","code":"BX728 - 02","size":"BX728","container":"AZ STP - 465","ket":"START","mc":"J10L","nilai_preload":"0.42","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR","keterangan":""},{"tgl":"2026-01-09","pic":"FERDIKA - NAGITA","code":"BX728 - 01","size":"BX728","container":"SM STP - 490","ket":"START","mc":"J10R","nilai_preload":"0.53","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR","keterangan":""},{"tgl":"2026-01-12","pic":"SIGIT - DIVA","code":"AS077 - 02","size":"AS077","container":"HM L46 - 542","ket":"START","mc":"J09L","nilai_preload":"0.44","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DM + MS + LSH )","keterangan":""},{"tgl":"2026-01-12","pic":"DWI - SATRIA","code":"AS077 - 04","size":"AS077","container":"SH L46 - 35","ket":"START","mc":"J09R","nilai_preload":"1.54","vmc":"-","shim_sr":"1.40","shim_pl":"0","first_cure":"OK ( DM + MS + LSH )","keterangan":""},{"tgl":"2026-01-11","pic":"AOP - IRFAN","code":"XS306 - 05","size":"XS306","container":"SM AZIII - 142","ket":"START","mc":"K14R","nilai_preload":"0.44","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR","keterangan":""},{"tgl":"2026-01-12","pic":"AOP - IRFAN","code":"XS306 - 01","size":"XS306","container":"SM AZIII - 122","ket":"GACON","mc":"K14L","nilai_preload":"-0.25","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-01-02","pic":"SIGIT - DIVA","code":"AA091 - 01","size":"AA091","container":"SM AZIV - 521","ket":"START","mc":"M04L","nilai_preload":"1.39","vmc":"-","shim_sr":"1.40","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-02","pic":"SIGIT - DIVA","code":"AA091 - 02","size":"AA091","container":"GT AZIV - 230","ket":"START","mc":"M04R","nilai_preload":"0.71","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-11","pic":"IDAD - DIVA","code":"BX639 - 05","size":"BX639","container":"GT AZIII/2 - 584","ket":"CLEANING","mc":"N04L","nilai_preload":"-","vmc":"-","shim_sr":"1.40","shim_pl":"0","first_cure":"OK ( FM )","keterangan":""},{"tgl":"2026-01-11","pic":"IDAD - DIVA","code":"BX639 - 02","size":"BX639","container":"AZIII/2 - 585","ket":"CLEANING","mc":"N04R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR + DMGM","keterangan":""},{"tgl":"2026-01-12","pic":"PEBRI - DIKA","code":"B450 - 04","size":"B450","container":"AZIV - 326","ket":"START","mc":"N06L","nilai_preload":"1.91","vmc":"-","shim_sr":"1.40","shim_pl":"0","first_cure":"OS + DM","keterangan":""},{"tgl":"2026-01-12","pic":"MUHAIMIN - DIAZ","code":"B450 - 02","size":"B450","container":"GT AZIV - 73","ket":"START","mc":"N06R","nilai_preload":"0.55","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR","keterangan":""},{"tgl":"2026-01-12","pic":"PEBRI - DIKA","code":"B362 - 02","size":"B362","container":"GT AZIII/2 - 607","ket":"START","mc":"K21R","nilai_preload":"0.56","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-12","pic":"PEBRI - DIKA","code":"B362 - 01","size":"B362","container":"SM AZIII/2 - 31","ket":"START","mc":"K24R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-12","pic":"AOP - IRFAN","code":"B378 - 01","size":"B378","container":"HM AZIII/2 - 52","ket":"START","mc":"K26L","nilai_preload":"0.56","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DMGM )","keterangan":""},{"tgl":"2026-01-13","pic":"MUHAIMIN - DIAZ","code":"B539 - 02","size":"B539","container":"GT L48 - 71","ket":"START","mc":"K27L","nilai_preload":"0.59","vmc":"","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-06","pic":"MUHAIMIN - DIAZ","code":"A488 - 03","size":"A488","container":"","ket":"START","mc":"H08L","nilai_preload":"0.50","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS BAWAH","keterangan":"MOR OUT SPEC / ATAS -1.10 / BAWAH -1.30"},{"tgl":"2026-01-06","pic":"FERDIKA - NAGITA","code":"A488 - 04","size":"A488","container":"","ket":"START","mc":"H08R","nilai_preload":"-0.16","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR ATAS BAWAH ","keterangan":"MOR OUT SPEC / ATAS -0.70 / BAWAH -0.60"},{"tgl":"2026-01-09","pic":"AOP - IRFAN","code":"BX740 - 01","size":"BX740","container":"AZ STP - 93","ket":"CLEANING","mc":"H10L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.50 / BAWAH 1.70"},{"tgl":"2026-01-09","pic":"SYAIFUL - HARIS","code":"BX740 - 08","size":"BX740","container":"SM STP - 212","ket":"CLEANING","mc":"H10R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.00 / BAWAH 1.50"},{"tgl":"2026-01-21","pic":"AOP - HARIS","code":"B697 - 01","size":"B697","container":"SM AZIV - 73","ket":"START","mc":"J08L","nilai_preload":"-","vmc":"-","shim_sr":"-","shim_pl":"-","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-01-03","pic":"SYAIFUL - TEGAR","code":"BX777 - 02","size":"BX777","container":"","ket":"START","mc":"H15R","nilai_preload":"-0.02","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.00 / BAWAH 1.60 / Spec MOR 1.70"},{"tgl":"2026-01-22","pic":"PEBRI - DIKA","code":"B612 - 02","size":"B612","container":"SM AZIII - 296","ket":"START","mc":"B28L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-20","pic":"PEBRI - DIKA","code":"BX677 - 17","size":"BX677","container":"AZ STP - 511","ket":"START","mc":"G27R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OOR + OFSH","keterangan":""},{"tgl":"2026-12-18","pic":"DWI - SATRIA","code":"BX776 - 01","size":"BX776","container":"","ket":"START","mc":"J06R","nilai_preload":"-0.77","vmc":"-","shim_sr":"0.70","shim_pl":"1.40","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.0 / BAWAH 1.70"},{"tgl":"2026-01-23","pic":"PEBRI - DIKA","code":"AS117 - 02","size":"AS117","container":"HM AZIII - 627","ket":"START","mc":"G29L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR + DM","keterangan":""},{"tgl":"2026-01-23","pic":"PEBRI - DIKA","code":"AS117 - 01","size":"AS117","container":"GT AZIII - 273","ket":"START","mc":"G29R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR + OS + DM","keterangan":""},{"tgl":"2026-01-22","pic":"SYAIFUL - MUHAIMIN","code":"B877 - 01","size":"B877","container":"HM V - 09","ket":"START","mc":"M26R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-18","pic":"AGUS - DIVA","code":"BX638 - 02","size":"BX638","container":"SM AZIII/2 - 330","ket":"START","mc":"A27R","nilai_preload":"-","vmc":"-","shim_sr":"?","shim_pl":"?","first_cure":"OK","keterangan":""},{"tgl":"2026-01-19","pic":"AMIR - AGUS","code":"AA096 - 03","size":"AA096","container":"GT L48 - 384","ket":"START","mc":"E25L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OS","keterangan":""},{"tgl":"2026-01-19","pic":"AMNIR - AGUS","code":"BX638 - 05","size":"BX638","container":"SM AZIII/2 - 266","ket":"START","mc":"A27L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-18","pic":"AGUS - DIVA","code":"XS425 - 01","size":"XS425","container":"GT L48 - 43","ket":"START","mc":"M25R","nilai_preload":"-","vmc":"-","shim_sr":"?","shim_pl":"?","first_cure":"OK","keterangan":""},{"tgl":"2026-01-18","pic":"AGUS - DIVA","code":"B477 - 01","size":"B477","container":"SM AZIII - 392","ket":"START","mc":"L22L","nilai_preload":"-","vmc":"-","shim_sr":"?","shim_pl":"?","first_cure":"OS + OFSH","keterangan":""},{"tgl":"2026-01-19","pic":"AMIR - AGUS","code":"AS050 - 02","size":"AS050","container":"SM AZIII - 322","ket":"START","mc":"L29L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-19","pic":"AMIR - AGUS","code":"AS050 - 01","size":"AS050","container":"SM AZIII","ket":"START","mc":"K29R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-21","pic":"AOP - HARIS","code":"AS049 - 02","size":"AS049","container":"SM AZIII - 102","ket":"START","mc":"K29L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-01-21","pic":"AOP - HARIS","code":"AS210 - 02","size":"AS210","container":"GT AZIV - 230","ket":"START","mc":"G26R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-07","pic":"MUHAIMIN - DIAZ","code":"XS222 - 02","size":"XS222","container":"SM STP - 651","ket":"START","mc":"J06R","nilai_preload":"0.45","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.50 / BAWAH 1.20"},{"tgl":"2026-01-22","pic":"SYAIFUL - MUHAIMIN","code":"BX497 - 01","size":"BX497","container":"GT STP - 142","ket":"START","mc":"F21L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-12-18","pic":"DWI - SATRIA","code":"BX675 - 05","size":"BX675","container":"","ket":"START","mc":"J14R","nilai_preload":"-0.41","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -1.0 / BAWAH 0.60"},{"tgl":"2026-01-22","pic":"PEBRI - DIKA","code":"AS057 - 02","size":"AS057","container":"AZ STP - 70","ket":"START","mc":"J14R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.60 / BAWAH 1.90"},{"tgl":"2026-01-18","pic":"AGUS - DIVA","code":"BX303 - 03","size":"BX303","container":"SM STP","ket":"START","mc":"N04R","nilai_preload":"-","vmc":"-","shim_sr":"?","shim_pl":"?","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-01-18","pic":"AGUS - DIVA","code":"BX303 - 04","size":"BX303","container":"SM STP - 616","ket":"START","mc":"N04L","nilai_preload":"-","vmc":"-","shim_sr":"?","shim_pl":"?","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-01-22","pic":"PEBRI","code":"BX731 - 06","size":"BX731","container":"SM STP - 253","ket":"START","mc":"A01R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-23","pic":"PEBRI - DIKA","code":"BX731 - 15","size":"BX731","container":"AZ STP - 78","ket":"START","mc":"A03R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( OFBR + DMGM )","keterangan":""},{"tgl":"2026-12-18","pic":"SIGIT - DIVA","code":"A322 - 02","size":"A322","container":"","ket":"CLEANING","mc":"K06L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.80 / BAWAH 1.30"},{"tgl":"2026-01-22","pic":"AMIR - AGUS","code":"B356 - 20","size":"B356","container":"GT STP - 60","ket":"START","mc":"A15R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OS","keterangan":""},{"tgl":"2026-01-22","pic":"AMIR - AGUS","code":"B356 - 12","size":"B356","container":"SM STP - 622","ket":"START","mc":"A15L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OS","keterangan":""},{"tgl":"2026-12-18","pic":"SYAIFUL - TEGAR","code":"A322 - 01","size":"A322","container":"","ket":"CLEANING","mc":"K06R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 1.70 / BAWAH 0.80"},{"tgl":"2026-12-18","pic":"IRFAN - AOP","code":"AS281 - 01","size":"AS281","container":"","ket":"START","mc":"K09R","nilai_preload":"1.07","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.00 / BAWAH 1.20"},{"tgl":"2026-01-07","pic":"TOHANA - IDAD","code":"B446 - 01","size":"B446","container":"","ket":"START","mc":"K13R","nilai_preload":"1.08","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.60 / BAWAH 1.30"},{"tgl":"2026-12-29","pic":"FERDIKA - NAGITA","code":"BX506 - 06","size":"BX506","container":"","ket":"START","mc":"K27L","nilai_preload":"-1.73","vmc":"-","shim_sr":"0.70","shim_pl":"2.10","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -0.59 / Spec MOR 1.07"},{"tgl":"2026-01-20","pic":"MUHAIMIN - DSIAZ","code":"XM412 - 02","size":"XM412","container":"SM STP - 24","ket":"START","mc":"E06R","nilai_preload":"-","vmc":"-","shim_sr":"?","shim_pl":"?","first_cure":"OK","keterangan":""},{"tgl":"2026-01-04","pic":"AGUS - UBAY","code":"B686 - 01","size":"B686","container":"","ket":"START","mc":"L04L","nilai_preload":"0.60","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.40 / BAWAH 1.80"},{"tgl":"2026-01-22","pic":"AMIR - AGUS","code":"B503 - 01","size":"B503","container":"SM STP - 431","ket":"START","mc":"F02R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-22","pic":"AMIR - AGUS","code":"B503 - 02","size":"B503","container":"AZ STP - 465","ket":"START","mc":"F02L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-12-29","pic":"HARIS - BADRU","code":"XS194 - 01","size":"XS194","container":"","ket":"START","mc":"L111L","nilai_preload":"0.53","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.10 / BAWAH 1.90"},{"tgl":"2026-01-19","pic":"AMIR - AGUS","code":"B845 - 01","size":"B845","container":"SM AZIII/2 - 33","ket":"START","mc":"G10L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OS + DMGM + DM","keterangan":""},{"tgl":"2026-01-23","pic":"FERDIKA - NAGITA 0T","code":"AS096 - 02","size":"AS096","container":"SM L48 - 125","ket":"START","mc":"L21L","nilai_preload":"0.59","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 4.66 / BAWAH 3.00"},{"tgl":"2026-01-22","pic":"PEBRI","code":"BX553 - 01","size":"BX553","container":"HM STP - 386","ket":"START","mc":"G09L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-06","pic":"MUHAIMIN - DIAZ","code":"AS264 - 02","size":"AS264","container":"","ket":"START","mc":"L26R","nilai_preload":"0.44","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -1.50"},{"tgl":"2026-01-18","pic":"AGUS - DIVA","code":"AX584 - 05","size":"AX584","container":"GT AZIV - 667","ket":"START","mc":"M13R","nilai_preload":"-","vmc":"-","shim_sr":"?","shim_pl":"?","first_cure":"OK","keterangan":""},{"tgl":"2026-01-28","pic":"TOHANA - IDAD","code":"B511 - 06","size":"B511","container":"HM AZIII - 404","ket":"START","mc":"G25L","nilai_preload":"0.23","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-01-26","pic":"SYAIFUL - HARIS","code":"B651 - 02","size":"B651","container":"GT AZIII - 83","ket":"START","mc":"L28R","nilai_preload":"0.06","vmc":"","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR MIRING","keterangan":"ATAS 1.1 / BAWAH 2.50"},{"tgl":"2026-12-18","pic":"SYAIFUL - TEGAR","code":"XS220 - 02","size":"XS220","container":"","ket":"START","mc":"M07L","nilai_preload":"0.11","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0 / BAWAH 1.30"},{"tgl":"2026-01-21","pic":"AOP - HARIS","code":"AA107 - 01","size":"AA107","container":"GT L48 - 26","ket":"START","mc":"M09L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OFSH","keterangan":"OFSH SIDERING BAWAH 2.00"},{"tgl":"2026-01-28","pic":"TOHANA - IDAD","code":"B356 - 14","size":"B356","container":"SM STP - 246","ket":"START","mc":"E14R","nilai_preload":"-1.20","vmc":"0.70","shim_sr":"0.70","shim_pl":"1.75","first_cure":"OK ( DMGM )","keterangan":""},{"tgl":"2026-01-28","pic":"TOHANA - IDAD","code":"BX775 - 01","size":"BX775","container":"SM STP - 682","ket":"START","mc":"J13R","nilai_preload":"1.10","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-28","pic":"SIGIT - DIVA","code":"A294 - 01","size":"A294","container":"GT L48 - 234","ket":"START","mc":"K03L","nilai_preload":"0.57","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DM + LL )","keterangan":""},{"tgl":"2026-01-28","pic":"DWI - SATRIA","code":"AS064 - 02","size":"AS064","container":"SM STP - 334","ket":"START","mc":"H07R","nilai_preload":"0.17","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-01-28","pic":"DWI - SATRIA","code":"BX679 - 10","size":"BX679","container":"HM AZIII - 155","ket":"START","mc":"M15L","nilai_preload":"0.49","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-29","pic":"AOP - IRFAN","code":"XM450 - 03","size":"XM450","container":"HM AZIV - 120","ket":"START","mc":"E24R","nilai_preload":"0.61","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2026-01-29","pic":"SYAIFUL - HARIS","code":"BX671 - 01","size":"BX671","container":"SM AZIV - 60","ket":"START","mc":"M23L","nilai_preload":"0.09","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK ( FM )","keterangan":""},{"tgl":"2026-01-29","pic":"SIGIT - DIVA","code":"XM379 - 01","size":"XM379","container":"SM AZIII - 322","ket":"START","mc":"F09L","nilai_preload":"0.56","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-29","pic":"DWI - SATRIA","code":"XM450 - 02","size":"XM450","container":"HM AZIV - 138","ket":"START","mc":"E24L","nilai_preload":"0.16","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-01-30","pic":"AOP - IRFAN","code":"AS121 - 01","size":"AS121","container":"HM AZIV - 196","ket":"START","mc":"G24L","nilai_preload":"0.57","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-30","pic":"AOP - IRFAN","code":"AS087 - 02","size":"AS087","container":"HM AZIV - 200","ket":"START","mc":"A28R","nilai_preload":"0.55","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-28","pic":"AGUS - UBAY","code":"B375 - 02","size":"B375","container":"SM AZIV - 375","ket":"START","mc":"M22R","nilai_preload":"0.48","vmc":"0","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.80 / BAWAH 2.00"},{"tgl":"2026-01-30","pic":"PEBRI - DIKA","code":"B356 - 09","size":"B356","container":"SM STP - 502","ket":"START","mc":"E12L","nilai_preload":"0.10","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-01-30","pic":"PEBRI - DIKA","code":"B362 - 01","size":"B362","container":"GT AZIII/2 - 69","ket":"START","mc":"K24L","nilai_preload":"0.37","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-30","pic":"PEBRI - DIKA","code":"B603 - 01","size":"B603","container":"SM AZIII - 362","ket":"START","mc":"M10L","nilai_preload":"0.07","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK ( IP )","keterangan":""},{"tgl":"2026-01-30","pic":"PEBRI - DIKA","code":"B551 - 35","size":"B551","container":"SH STP - 699","ket":"START","mc":"B13R","nilai_preload":"0.27","vmc":"1.05","shim_sr":"1.05","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-01-30","pic":"PEBRI - DIKA","code":"B551 - 37","size":"B551","container":"SM STP - 169","ket":"START","mc":"B13L","nilai_preload":"1.92","vmc":"1.05","shim_sr":"1.40","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-30","pic":"SIGIT - DIVA","code":"B602 - 01","size":"B602","container":"GT AZIII - 336","ket":"START","mc":"M06L","nilai_preload":"0.72","vmc":"1.05","shim_sr":"1.05","shim_pl":"0","first_cure":"OS","keterangan":""},{"tgl":"2026-01-30","pic":"SIGIT - DIVA","code":"XA066 - 04","size":"XA066","container":"SM TP - 687","ket":"START","mc":"J11L","nilai_preload":"0.27","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-01-05","pic":"IRFAN - AOP","code":"B875 - 01","size":"B875","container":"","ket":"START","mc":"M23R","nilai_preload":"0.13","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 1.70 / BAWAH 0.90"},{"tgl":"2026-01-30","pic":"DWI - SATRIA","code":"XM421 - 01","size":"XM421","container":"SM STP - 122","ket":"START","mc":"C10R","nilai_preload":"1.31","vmc":"1.05","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-30","pic":"RENDY - DIVA","code":"XS309 - 01","size":"XS309","container":"HM L46 - 108","ket":"START","mc":"M24R","nilai_preload":"0.24","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 1.60 / BAWAH 0.50"},{"tgl":"2026-01-24","pic":"AOP - AGUS","code":"B643 - 09","size":"B643","container":"GT AZIII/2 - 42","ket":"CLEANING","mc":"N13R","nilai_preload":"-","vmc":"-","shim_sr":"2.45","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-01-24","pic":"SYAIFUL - HARIS","code":"B643 - 03","size":"B643","container":"AZIII/2 - 24","ket":"CLEANING","mc":"N13L","nilai_preload":"-","vmc":"-","shim_sr":"1.75","shim_pl":"0","first_cure":"MOR + OF + OS","keterangan":""},{"tgl":"2026-01-10","pic":"AOP - IRFAN","code":"BX670 - 02","size":"BX670","container":"GT AZIV - 134","ket":"START","mc":"N03L","nilai_preload":"-0.49","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 3.00 / BAWAH 1.80"},{"tgl":"2026-01-28","pic":"MUHAIMIN - DIAZ","code":"XS226 - 02","size":"XS226","container":"SM STP - 310","ket":"CLEANING","mc":"L25R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( FM )","keterangan":""},{"tgl":"2026-01-07","pic":"MUHAIMIN - DIAZ","code":"BX643 - 02","size":"BX643","container":"","ket":"CLEANING","mc":"N08L","nilai_preload":"-","vmc":"-","shim_sr":"2.10","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.90 / BAWAH 1.90"},{"tgl":"2026-01-27","pic":"AGUS - UBAY","code":"B629 - 01","size":"B629","container":"SM L48 - 96","ket":"CLEANING","mc":"L30R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-28","pic":"FERDIKA - NAGITA","code":"AA116 - 01","size":"AA116","container":"SM L48 - 192","ket":"START","mc":"L29R","nilai_preload":"1.11","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-01-25","pic":"TOHANA - IDAD","code":"BX731 - 13","size":"BX731","container":"AZ STP - 416","ket":"START","mc":"B15L","nilai_preload":"-0.90","vmc":"-","shim_sr":"0.70","shim_pl":"1.40","first_cure":"OK","keterangan":""},{"tgl":"2026-01-26","pic":"SYAIFUL - HARIS","code":"BX731 - 07","size":"BX731","container":"SM STP - 307","ket":"START","mc":"B15R","nilai_preload":"-0.19","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-01-12","pic":"AGUS - UBAY","code":"AA093 - 05","size":"AA093","container":"GT AZIV - 532","ket":"CLEANING","mc":"N14L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.00 / BAWAH 3.00"},{"tgl":"2026-01-26","pic":"TOHANA - IDAD","code":"BX733 - 06","size":"BX733","container":"SM STP - 244","ket":"START","mc":"B01L","nilai_preload":"-0.29","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK ( LL )","keterangan":""},{"tgl":"2026-01-26","pic":"AGUS - UBAY","code":"BX733 - 04","size":"BX733","container":"SM STP - 824","ket":"START","mc":"B01R","nilai_preload":"-0.16","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"AGUS - UBAY","code":"B355 - 01","size":"B355","container":"SM STP - 704","ket":"Spare UC","mc":"B09L","nilai_preload":"-0.20","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"FERDIKA - NAGITA","code":"XM381 - 02","size":"XM381","container":"HM STP - 707","ket":"START","mc":"C04R","nilai_preload":"0.50","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"FERDIKA - NAGITA","code":"XM402 - 06","size":"XM402","container":"AZ STP - 506","ket":"START","mc":"D07L","nilai_preload":"-1.13","vmc":"0.70","shim_sr":"0.70","shim_pl":"1.75","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"TOHANA - IDAD","code":"B356 - 04","size":"B356","container":"AZ STP - 464","ket":"START","mc":"D12L","nilai_preload":"1.25","vmc":"1.05","shim_sr":"1.40","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"TOHANA - IDAD","code":"B356 - 17","size":"B356","container":"AZ STP - 80","ket":"START","mc":"D12R","nilai_preload":"0.96","vmc":"1.05","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"AOP - IRFAN","code":"B256 - 02","size":"B256","container":"SM STP - 82","ket":"START","mc":"E01L","nilai_preload":"-0.08","vmc":"1.05","shim_sr":"1.05","shim_pl":"0.70","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.00 / BAWAH 2.60"},{"tgl":"2026-02-02","pic":"AOP - IRFAN","code":"B256 - 04","size":"B256","container":"SM STP - 253","ket":"START","mc":"E01R","nilai_preload":"0.19","vmc":"1.05","shim_sr":"1.05","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"SYAIFUL - HARIS","code":"XM400 - 02","size":"XM400","container":"SM STP - 316","ket":"START","mc":"F02L","nilai_preload":"0.52","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"SYAIFUL - HARIS","code":"XM400 - -03","size":"","container":"SM STP - 354","ket":"START","mc":"F02R","nilai_preload":"-0.64","vmc":"0.70","shim_sr":"0.70","shim_pl":"1.05","first_cure":"MOR MIRING","keterangan":"ATAS 0.80 / BAWAH 1.20"},{"tgl":"2026-02-02","pic":"AGUS - UBAY","code":"B644 -- 14","size":"","container":"SM STP - 364","ket":"CLEANING","mc":"G13L","nilai_preload":"-","vmc":"-","shim_sr":"-","shim_pl":"-","first_cure":"OK ( FM )","keterangan":""},{"tgl":"2026-02-02","pic":"AGUS - UBAY","code":"B644 - 12","size":"B644","container":"SM STP - 378","ket":"CLEANING","mc":"G13R","nilai_preload":"-","vmc":"-","shim_sr":"-","shim_pl":"-","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"FERDIKA - NAGITA","code":"B690 - 01","size":"B690","container":"HM AZIII - 630","ket":"CLEANING","mc":"J05L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"-","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"MUHAIMIN - DIAZ","code":"B690 - 04","size":"B690","container":"HM AZIII - 173","ket":"CLEANING","mc":"J05R","nilai_preload":"-","vmc":"-","shim_sr":"1.40","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"SYAIFUL - IRFAN","code":"BX740 - 05","size":"BX740","container":"SM STP - 130","ket":"CLEANING","mc":"J04L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OS","keterangan":""},{"tgl":"2026-02-02","pic":"SYAIFUL - IRFAN","code":"BX740 - 07","size":"BX740","container":"SM STP - 387","ket":"CLEANING","mc":"J04R","nilai_preload":"-","vmc":"-","shim_sr":"1.75","shim_pl":"1.75","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"MUHAIMIN - DIAZ","code":"B429 - 01","size":"B429","container":"SM AZIV - 59","ket":"START","mc":"J14R","nilai_preload":"-0.15","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"MUHAIMIN - DIAZ","code":"B430 - 03","size":"B430","container":"SM L48 - 365","ket":"START","mc":"K14L","nilai_preload":"0.68","vmc":"0","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"AGUS - UBAY","code":"B430 - 01","size":"B430","container":"GT L48 - 201","ket":"START","mc":"K14R","nilai_preload":"0.52","vmc":"0","shim_sr":"0.70","shim_pl":"0","first_cure":"OS","keterangan":""},{"tgl":"2026-02-02","pic":"FERDIKA - NAGITA","code":"B853 - 01","size":"B853","container":"SM STP - 651","ket":"START","mc":"K09L","nilai_preload":"-0.19","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 1.70 / BAWAH 0.60"},{"tgl":"2026-02-02","pic":"AOP - IRFAN","code":"AS100 - 02","size":"AS100","container":"GT AZIV - 133","ket":"START","mc":"M02L","nilai_preload":"0.45","vmc":"0.35","shim_sr":"0.35","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"AOP - IRFAN","code":"AS100 - 01","size":"AS100","container":"SM AZIV - 61","ket":"START","mc":"M02R","nilai_preload":"0.58","vmc":"0.35","shim_sr":"0.35","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 3.60 / BAWAH 2.80"},{"tgl":"2026-02-02","pic":"SYAIFUL - HARIS","code":"XM410 - 01","size":"XM410","container":"SM STP - 240","ket":"START","mc":"F08R","nilai_preload":"0.65","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -0.90 / BAWAH -0.48"},{"tgl":"2026-02-02","pic":"FERDIKA - NAGITA","code":"B705 - 01","size":"B705","container":"HM L46 - 469","ket":"START","mc":"L21L","nilai_preload":"0.60","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.50"},{"tgl":"2026-02-02","pic":"PEBRI - DIKA","code":"AS278 - 01","size":"AS278","container":"GT AZIII - 191","ket":"START","mc":"K26R","nilai_preload":"0.43","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.20 / BAWAH 1.10"},{"tgl":"2026-02-02","pic":"SYAIFUL - IRFAN","code":"A291 - 01","size":"A291","container":"HM AZIII - 49","ket":"CLEANING","mc":"N09L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"1.40","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"AGUS - UBAY","code":"AS077 - 02","size":"AS077","container":"SH L46 - 65","ket":"START","mc":"M12R","nilai_preload":"0.56","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"TOHANA - IDAD","code":"AS077 - 03","size":"AS077","container":"L46 - 425","ket":"START","mc":"M12L","nilai_preload":"0.56","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-02-02","pic":"SYAIFUL - HARIS","code":"B603 - 02","size":"B603","container":"SH AZIII - 56","ket":"START","mc":"M10R","nilai_preload":"0.24","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"SYAIFUL - HARIS","code":"B700 - 01","size":"B700","container":"HM L46 - 51","ket":"CLEANING","mc":"L06L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"MUHAIMIN - DIAZ","code":"BX827 - 01","size":"BX827","container":"SM STP - 615","ket":"START","mc":"J10R","nilai_preload":"-0.64","vmc":"0","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-02-02","pic":"MUHAIMIN - DIAZ","code":"BX816 - 01","size":"BX816","container":"SM STP - 614","ket":"START","mc":"K05L","nilai_preload":"0.46","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-02-02","pic":"SYAIFUL - HARIS","code":"BX821 - 01","size":"BX821","container":"HM STP - 701","ket":"START","mc":"M27L","nilai_preload":"0.76","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.70 / BAWAH 1.90"},{"tgl":"2026-02-02","pic":"AGUS - UBAY","code":"XS402 - 01","size":"XS402","container":"GT AZIII - 573","ket":"START","mc":"M25L","nilai_preload":"0.58","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":""},{"tgl":"2026-02-02","pic":"TOHANA - IDAD","code":"A488 - -03","size":"","container":"SM STP - 184","ket":"CLEANING","mc":"H08L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -0.50 / BAWAH -0.80"},{"tgl":"2026-02-02","pic":"TOHANA - IDAD","code":"A488 - 04","size":"A488","container":"SM STP 471","ket":"CLEANING","mc":"H08R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -0.20 / BAWAH -0.70"},{"tgl":"2026-02-02","pic":"PEBRI - DIKA","code":"A291 - 02","size":"A291","container":"AZ III/2 - 216","ket":"CLEANING","mc":"N09R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"PEBRI - DWI","code":"AA007 - 03","size":"AA007","container":"SM AZIV - 66","ket":"START","mc":"A21L","nilai_preload":"0.31","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OFSH","keterangan":""},{"tgl":"2026-02-03","pic":"SIGIT - DIVA","code":"AA007 - 02","size":"AA007","container":"SM AZIV - 328","ket":"START","mc":"A21R","nilai_preload":"1.18","vmc":"1.05","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"PEBRI - DWI","code":"B747 - 02","size":"B747","container":"SM AZIII - 25","ket":"START","mc":"G22R","nilai_preload":"1.98","vmc":"1.05","shim_sr":"1.40","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.00 / BAWAH 2.00"},{"tgl":"2026-02-03","pic":"DWI","code":"B747 - 01","size":"B747","container":"AZIII/2 - 13","ket":"START","mc":"G22L","nilai_preload":"-0.22","vmc":"1.05","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OFSH","keterangan":""},{"tgl":"2026-01-03","pic":"PEBRI - DWI ","code":"BX678 - 09","size":"BX678","container":"HM AZIII - 173","ket":"CLEANING","mc":"N15L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"PEBRI - DWI","code":"AS238 - 02","size":"AS238","container":"GT AZIV - 244","ket":"CLEANING","mc":"A29R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"SIGIT - DIVA","code":"BX638 - 02","size":"BX638","container":"HM AZIII - 414","ket":"START","mc":"B24L","nilai_preload":"1.20","vmc":"1.05","shim_sr":"1.40","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -0.90 / BAWAH 1.60"},{"tgl":"2026-02-03","pic":"SIGIT - DIVA","code":"BX638 - 05","size":"BX638","container":"HM AZIII - 415","ket":"START","mc":"B24R","nilai_preload":"0.21","vmc":"1.05","shim_sr":"1.05","shim_pl":"0.35","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.90 / BAWAH 1.30"},{"tgl":"2026-02-03","pic":"SIGIT - DIVA","code":"BX678 - 01","size":"BX678","container":"GT AZIII - 533","ket":"CLEANING","mc":"N15R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.00 / BAWAH 1.60"},{"tgl":"2026-02-03","pic":"SIGIT - DIVA","code":"AS238 - 01","size":"AS238","container":"HM AZIV - 709","ket":"CLEANING","mc":"A29L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"FERDIKA - NAGITA","code":"B858 - 01","size":"B858","container":"GT AZIII - 188","ket":"START","mc":"N26L","nilai_preload":"0.60","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"FERDIKA - NAGITA","code":"XA024 - 02","size":"XA024","container":"SM STP - 749","ket":"START","mc":"K10L","nilai_preload":"0.23","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 1.80 / BAWAH 1.10"},{"tgl":"2026-02-03","pic":"FERDIKA - NAGITA","code":"XM404 - 01","size":"XM404","container":"SM STP - 611","ket":"START","mc":"M28L","nilai_preload":"0.45","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-02-03","pic":"FERIDKA - NAGITA","code":"B643 - 05","size":"B643","container":"HM AZIII - 86","ket":"SPARE CM","mc":"N08L","nilai_preload":"0.45","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"FERDIKA - NAGITA","code":"B675 - 06","size":"B675","container":"SM AZIII - 07","ket":"CLEANING","mc":"G07L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"MUHAIMIN - DIAZ","code":"B501 - 02","size":"B501","container":"SM STP - 711","ket":"START","mc":"B25R","nilai_preload":"-0.45","vmc":"0.70","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OFSH","keterangan":""},{"tgl":"2026-02-03","pic":"MUHAIMIN - DIAZ","code":"B501 - 01","size":"B501","container":"GT STP - 179","ket":"START","mc":"B25L","nilai_preload":"-0.62","vmc":"0.70","shim_sr":"0.70","shim_pl":"1.05","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.40 / BAWAH 2.20"},{"tgl":"2026-02-03","pic":"MUHAIMIN - DIAZ","code":"B804 - 01","size":"B804","container":"HM L46 - 14","ket":"START","mc":"K22L","nilai_preload":"0.21","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0 / BAWAH 2.00"},{"tgl":"2026-02-03","pic":"MUHAIMIN - DIAZ","code":"B468 - 02","size":"B468","container":"SM STP - 417","ket":"START","mc":"L31L","nilai_preload":"0.13","vmc":"0","shim_sr":"0","shim_pl":"0.35","first_cure":"MOR ATAS MIRING","keterangan":"ATAS 0.80 / BAWAH 1.60"},{"tgl":"2026-02-03","pic":"MUHAIMIN - DIAZ","code":"AB001 - 01","size":"AB001","container":"HM L46 - 94","ket":"START","mc":"M26L","nilai_preload":"0.59","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-02-03","pic":"AOP - HARIS","code":"A227 - 01","size":"A227","container":"SM AZIII - 70","ket":"START","mc":"G24L","nilai_preload":"0.09","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.40 / BAWAH 1.60"},{"tgl":"2026-02-03","pic":"AOP - HARIS","code":"A227 - 02","size":"A227","container":"HM AZIII - 172","ket":"START","mc":"G24R","nilai_preload":"0.27","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.70 / BAWAH 1.40"},{"tgl":"2026-02-03","pic":"AOP - HARIS","code":"B462 - 02","size":"B462","container":"HM AZIV - 76","ket":"START","mc":"L03L","nilai_preload":"0.24","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"AOP - HARIS","code":"B468 - 01","size":"B468","container":"SM STP - 137","ket":"START","mc":"L31R","nilai_preload":"0.56","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"AOP - HARIS","code":"AA007 - 04","size":"AA007","container":"GT AZIV - 160","ket":"START","mc":"L30R","nilai_preload":"0.08","vmc":"0","shim_sr":"0","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"AGUS - UBAY","code":"BX551 - 25","size":"BX551","container":"SM STP - 713","ket":"START","mc":"B04R","nilai_preload":"-0.49","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"AGUS - UBAY","code":"BX551 - 16","size":"BX551","container":"AZ STP - 712","ket":"START","mc":"B04L","nilai_preload":"0.40","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.70 / BAWAH 1.00"},{"tgl":"2026-02-03","pic":"AGUS - UBAY","code":"XM397 - 01","size":"XM397","container":"GT STP - 508","ket":"START","mc":"C15R","nilai_preload":"0.49","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"AGUS - UBAY","code":"XM397 - 02","size":"XM397","container":"SM STP - 76","ket":"START","mc":"C15L","nilai_preload":"-1.18","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"AGUS - UBAY","code":"B297 - 02","size":"B297","container":"SM STP - 380","ket":"START","mc":"J13L","nilai_preload":"0.74","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"TOHANA IDAD","code":"BX823 - 02","size":"BX823","container":"SM STP -149","ket":"START","mc":"K01L","nilai_preload":"0.48","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"TOHANA IDAD","code":"BX504 - 04","size":"BX504","container":"SM STP - 200","ket":"START","mc":"H09L","nilai_preload":"0.76","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"TOHANA IDAD","code":"BX504 - 03","size":"BX504","container":"SM STP - 710","ket":"START","mc":"H09R","nilai_preload":"1.14","vmc":"-","shim_sr":"1.4-0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-03","pic":"TOHANA IDAD","code":"B462 - 03","size":"B462","container":"SM AZIV - 88","ket":"START","mc":"L12R","nilai_preload":"0.15","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OS","keterangan":""},{"tgl":"2026-02-03","pic":"TOHANA IDAD","code":"B742 - 01","size":"B742","container":"SM STP - 394","ket":"START","mc":"M26R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"-","first_cure":"OK ( IP )","keterangan":""},{"tgl":"2026-02-04","pic":"PEBRI - DIKA","code":"BM405 - 02","size":"BM405","container":"AZ STP - 382","ket":"START","mc":"E13R","nilai_preload":"2.11","vmc":"0.70","shim_sr":"1.40","shim_pl":"0","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-02-04","pic":"PEBRI - DIKA","code":"BM384 - 07","size":"BM384","container":"SM AZIII - 605","ket":"START","mc":"D10R","nilai_preload":"0.45","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( LL )","keterangan":""},{"tgl":"2026-02-04","pic":"PEBRI - DIKA","code":"BM380 - 01","size":"BM380","container":"SM STP - 75","ket":"START","mc":"F05L","nilai_preload":"0.35","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK ( DMGB )","keterangan":""},{"tgl":"2026-02-04","pic":"SIGIT - DIVA","code":"AS094 - 01","size":"AS094","container":"SM L48 - 345","ket":"START","mc":"E25L","nilai_preload":"0.44","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-04","pic":"SIGIT - DIVA","code":"AS094 - 02","size":"AS094","container":"GT L48 - 140","ket":"START","mc":"E25R","nilai_preload":"0.82","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK ( MS )","keterangan":""},{"tgl":"2026-02-04","pic":"SIGIT - DIVA","code":"XM405 - 01","size":"XM405","container":"SM STP - 190","ket":"TART","mc":"F06L","nilai_preload":"1.10","vmc":"0.70","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-04","pic":"SIGIT - DIVA","code":"BX642 - 08","size":"BX642","container":"SM STP - 659","ket":"START","mc":"M03L","nilai_preload":"-0.22","vmc":"0.70-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-02-04","pic":"SIGIT - DIVA","code":"BX776 - 04","size":"BX776","container":"SM STP - 686","ket":"START","mc":"N11L","nilai_preload":"0.06","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-02-04","pic":"DIW - SATRIA","code":"XM380 - 11","size":"XM380","container":"SM STP - 191","ket":"START","mc":"F05R","nilai_preload":"0.31","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-04","pic":"DIW - SATRIA","code":"XM384 - 02","size":"XM384","container":"SM AZIII - 34","ket":"START","mc":"D09L","nilai_preload":"1.32","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-04","pic":"DIW - SATRIA","code":"XM384 - 03","size":"XM384","container":"SM AZIII - 81","ket":"START","mc":"D09R","nilai_preload":"-0.26","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-02-04","pic":"DIW - SATRIA","code":"XM391 - 02","size":"XM391","container":"AZ STP - 64","ket":"START","mc":"D03L","nilai_preload":"-0.06","vmc":"0.70","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK ( DMGM )","keterangan":""},{"tgl":"2026-02-04","pic":"DIW - SATRIA","code":"B819 - 01","size":"B819","container":"SM STP - 390","ket":"START","mc":"M28L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.00 / BAWAH 1.0"},{"tgl":"2026-02-04","pic":"FERDIKA - NAGITA","code":"XM396 - 01","size":"XM396","container":"SM STP - 611","ket":"START","mc":"C09R","nilai_preload":"0.47","vmc":"0","shim_sr":"0.70","shim_pl":"0","first_cure":"OOR","keterangan":""},{"tgl":"2026-02-04","pic":"FERDIKA - NAGITA","code":"AS239 - 01","size":"AS239","container":"GT AZIV - 519","ket":"START","mc":"G21R","nilai_preload":"0.54","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR","keterangan":""},{"tgl":"2026-02-04","pic":"FERDIKA - NAGITA","code":"BX817 - 01","size":"BX817","container":"AZ STP - 91","ket":"CLEANING","mc":"K03L","nilai_preload":"-","vmc":"--","shim_sr":"0","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 1.40 / BAWAH 2.70"},{"tgl":"2026-02-04","pic":"FERDIKA - NAGITA","code":"XM417 - 01","size":"XM417","container":"SM STP - 466","ket":"CLEANING","mc":"F13L","nilai_preload":"-","vmc":"","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-02-04","pic":"FERDIKA - NAGITA","code":"XM395 - 01","size":"XM395","container":"GT STP - 581","ket":"CLEANING","mc":"F13R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-04","pic":"MUHAIMIN - DIAZ","code":"XM384 - 06","size":"XM384","container":"AZIII - 527","ket":"START","mc":"D10L","nilai_preload":"0.50","vmc":"0.70","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-04","pic":"MUHAIMIN - DIAZ","code":"AS295 - 01","size":"AS295","container":"GT AZIV - 103","ket":"START","mc":"J07R","nilai_preload":"0.22","vmc":"-","shim_sr":"0","shim_pl":"0.35","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-02-04","pic":"MUHAIMIN - DIAZ","code":"AA074 - 01","size":"AA074","container":"GT AZIV - 73","ket":"CLEANING","mc":"L27L","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-04","pic":"MUHAIMIN - DIAZ","code":"AA099 - 02","size":"AA099","container":"HM V - 347","ket":"CLEANING","mc":"L02L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"K","keterangan":""},{"tgl":"2026-02-04","pic":"MUHAIMIN - DIAZ","code":"AA099 - 03","size":"AA099","container":"HM V - 242","ket":" CLEANING","mc":"L02R","nilai_preload":"-","vmc":"","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-04","pic":"SYAIFUL - IRFAN","code":"BM406 - 01","size":"BM406","container":"AZ STP - 15","ket":"START","mc":"E15L","nilai_preload":"1.89","vmc":"0.70","shim_sr":"1.40","shim_pl":"0","first_cure":"OOR","keterangan":""},{"tgl":"2026-02-04","pic":"SYAIFUL - IRFAN","code":"AS299 - 02","size":"AS299","container":"GT AZIV - 158","ket":"START","mc":"G21L","nilai_preload":"0.10","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"MOR","keterangan":""},{"tgl":"2026-02-04","pic":"SYAIFUL - IRFAN","code":"A013- 01","size":"A013","container":"HM AZIV - 535","ket":"START","mc":"G29L","nilai_preload":"0.28","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK ( LL )","keterangan":""},{"tgl":"2026-02-04","pic":"SYAIFUL - IRFAN","code":"A013 - 02","size":"A013","container":"SH AZIV - 252","ket":"START","mc":"G29R","nilai_preload":"0.21","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK ( LL )","keterangan":""},{"tgl":"2026-02-04","pic":"SYAIFUL - IRFAN","code":"AA074 - 02","size":"AA074","container":"SM AZIV - 33","ket":"CLEANING","mc":"L27R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-04","pic":"AGUS - UBAY","code":"AS264 - 01","size":"AS264","container":"GT STP - 647","ket":"CLEANING ","mc":"L26L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-02-04","pic":"AGUS - UBAY","code":"XM406 - 02","size":"XM406","container":"SM STP - 245","ket":"START","mc":"E15R","nilai_preload":"-0.57","vmc":"-0.70","shim_sr":"0.70","shim_pl":"1.05","first_cure":"MOR ATAS MIRING","keterangan":"ATAS 0.70"},{"tgl":"2026-02-04","pic":"AGUS - UBAY","code":"BX677 - 05","size":"BX677","container":"GT STP - 666","ket":"START","mc":"F22L","nilai_preload":"0.52","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.90 / SPEC MOR 1.70"},{"tgl":"2026-02-04","pic":"AGUS - UBAY","code":"B297 - 01","size":"B297","container":"SM STP - 118","ket":"START","mc":"J07L","nilai_preload":"0.57","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-04","pic":"TOHANA - IDAD","code":"BX776 - 05","size":"BX776","container":"AZ STP - 694","ket":"START","mc":"B11L","nilai_preload":"-0.95","vmc":"-","shim_sr":"0.70","shim_pl":"1.40","first_cure":"OK","keterangan":""},{"tgl":"2026-02-04","pic":"TOHANA - IDAD","code":"B475 - 01","size":"B475","container":"AZ STP - 624","ket":"CLEANING","mc":"K11L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OS","keterangan":""},{"tgl":"2026-02-04","pic":"TOHANA - IDAD","code":"B462 - 03","size":"B462","container":"SM AZIV - 88","ket":"CLEANING","mc":"L12R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OS","keterangan":""},{"tgl":"2026-02-04","pic":"TOHANA - IDAD","code":"AS264 - 02","size":"AS264","container":"GT STP - 262","ket":"CLEANING","mc":"L26R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-05","pic":"SIGIT - DIVA","code":"B644 - 13","size":"B644","container":"GT STP - 624","ket":"CLEANING","mc":"G05R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OFSH","keterangan":""},{"tgl":"2026-02-05","pic":"DWI - SATRIA","code":"B644 - 09","size":"B644","container":"SM STP - 355","ket":"CLEANING","mc":"G05L","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-02-05","pic":"SIGIT - DIVA","code":"XS077 - 01","size":"XS077","container":"SM AZIII - 137","ket":"START","mc":"L24L","nilai_preload":"0.33","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK ( LSH )","keterangan":""},{"tgl":"2026-02-05","pic":"SIGIT - DIVA","code":"BM459 - 04","size":"BM459","container":"HM L46 - 44","ket":"START","mc":"K28R","nilai_preload":"0.51","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( FM )","keterangan":""},{"tgl":"2026-02-05","pic":"SIGIT - DIVA","code":"BX741 - 01","size":"BX741","container":"SM STP - 362","ket":"START","mc":"N04L","nilai_preload":"-0.45","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-02-05","pic":"SIGIT - DIVA","code":"BX741 - 02","size":"BX741","container":"GT STP - 157","ket":"START","mc":"N04R","nilai_preload":"-0.23","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-02-05","pic":"DWI - SATRIA","code":"BM455 - 01","size":"BM455","container":"HM L46 - 69","ket":"START","mc":"K29L","nilai_preload":"0.46","vmc":"-","shim_sr":"0.70","shim_pl":"0.35","first_cure":"OK","keterangan":""},{"tgl":"2026-02-05","pic":"DWI - SATRIA","code":"BM455 - 02","size":"BM455","container":"HM L46 - 87","ket":"START","mc":"K29R","nilai_preload":"0.69","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS -0.80 / BAWAH 1.00"},{"tgl":"2026-02-05","pic":"DWI - SATRIA","code":"BX642 - 01","size":"BX642","container":"SM STP - 238","ket":"START","mc":"M03R","nilai_preload":"-0.42","vmc":"-","shim_sr":"0.70","shim_pl":"1.05","first_cure":"OK","keterangan":""},{"tgl":"2026-02-05","pic":"FERDIKA - NAGITA","code":"AB003 - 01","size":"AB003","container":"GT AZIV - 255","ket":"START","mc":"M23R","nilai_preload":"0.50","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.60 / BAWAH 1.60"},{"tgl":"2026-02-05","pic":"FERDIKA - NAGITA","code":"B861 - 01","size":"B861","container":"GT AZIII - 189","ket":"START","mc":"M26R","nilai_preload":"0.46","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-05","pic":"FERDIKA - NAGITA","code":"B797 - 01","size":"B797","container":"HM V - 571","ket":"START","mc":"M21L","nilai_preload":"0.65","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-05","pic":"FERDIKA - NAGITA","code":"AS251 - 01","size":"AS251","container":"GT AZIV - 377","ket":"CLEANING","mc":"L23R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.40 / BAWAH 1.30"},{"tgl":"2026-02-05","pic":"FERDIKA - NAGITA","code":"AS294 - 02","size":"AS294","container":"GT AZIV - 532","ket":"CLEANING","mc":"L28R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KEDALAM","keterangan":"ATAS 0.70 / BAWAH 1.50"},{"tgl":"2026-02-05","pic":"MUHAIMIN - DIAZ","code":"B643 - 01","size":"B643","container":"SM AZIII/2 - 97","ket":"START","mc":"N08R","nilai_preload":"0.49","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-02-05","pic":"MUHAIMIN - DIAZ","code":"B748 - 01","size":"B748","container":"GT AZIII - 187","ket":"START","mc":"N22L","nilai_preload":"0.40","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":""},{"tgl":"2026-02-05","pic":"MUHAIMIN - DIAZ","code":"BM460 - 02","size":"BM460","container":"HM L46 - 109","ket":"START","mc":"G25R","nilai_preload":"0.56","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-02-05","pic":"MUHAIMIN - DIAZ","code":"BM460 - 04","size":"BM460","container":"HM L46 - 250","ket":"START","mc":"G25L","nilai_preload":"0.49","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( DMGM )","keterangan":""},{"tgl":"2026-02-05","pic":"AGUS - UBAY","code":"BM372 - 01","size":"BM372","container":"SM STP - 393","ket":"START","mc":"M24R","nilai_preload":"-","vmc":"-","shim_sr":"0","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-05","pic":"AGUS - UBAY","code":"AB002 - 01","size":"AB002","container":"HM L46 - 542","ket":"START","mc":"M26L","nilai_preload":"0.53","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-05","pic":"AGUS - UBAY","code":"B643 - 01","size":"B643","container":"GT AZIII/2 - 68","ket":"START","mc":"N08R","nilai_preload":"0.48","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-02-05","pic":"AGUS - UBAY","code":"BX823 - 01","size":"BX823","container":"SM STP - 525","ket":"START","mc":"K05R","nilai_preload":"-0.22","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OOR","keterangan":""},{"tgl":"2026-02-05","pic":"AGUS - UBAY","code":"AS252 - 01","size":"AS252","container":"GT STP - 371","ket":"START","mc":"C09R","nilai_preload":"-0.20","vmc":"-","shim_sr":"0.70","shim_pl":"0.70","first_cure":"OK","keterangan":""},{"tgl":"2026-02-05","pic":"FERDIKA - IRFAN","code":"B796 - 02","size":"B796","container":"AZ L46 - 263","ket":"START","mc":"IIIR","nilai_preload":"0.43","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OFSH","keterangan":""},{"tgl":"2026-02-05","pic":"FERDIKA - IRFAN","code":"BX679 - 15","size":"BX679","container":"HM AZIII - 609","ket":"START","mc":"M06R","nilai_preload":"0.46","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 3.00 / BAWAH 1.80"},{"tgl":"2026-02-05","pic":"FERDIKA - IRFAN","code":"AS159 - 01","size":"AS159","container":"SM L48 - 677","ket":"START","mc":"K24L","nilai_preload":"0.57","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK ( FM + LL )","keterangan":""},{"tgl":"2026-02-05","pic":"FERDIKA - IRFAN","code":"B256 - 01","size":"B256","container":"AZ STP - 656","ket":"CLEANING","mc":"A08R","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-02-05","pic":"FERDIKA - IRFAN","code":"AS434 - 01","size":"AS434","container":"HM AZIV - 323","ket":"CLEANING","mc":"L30L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-05","pic":"SYAIFUL - HARIS","code":"BM459 - 01","size":"BM459","container":"AZ L46 - 23","ket":"START","mc":"K27L","nilai_preload":"0.52","vmc":"0","shim_sr":"0","shim_pl":"0","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-02-05","pic":"SYAIFUL - HARIS","code":"B679 - 16","size":"B679","container":"SM AZIII - 24","ket":"START","mc":"M06L","nilai_preload":"-0.29","vmc":"0.25","shim_sr":"0.70","shim_pl":"0.70","first_cure":"MOR ATAS KELUAR","keterangan":"ATAS 2.50"},{"tgl":"2026-02-05","pic":"SYAIFUL - HARIS","code":"B675 - 05","size":"B675","container":"HM AZIII - 343","ket":"START","mc":"I13R","nilai_preload":"0.54","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""},{"tgl":"2026-02-05","pic":"SYAIFUL - HARIS","code":"B256 - 05","size":"B256","container":"SM STP - 816","ket":"CLEANING","mc":"A08L","nilai_preload":"-","vmc":"-","shim_sr":"1.05","shim_pl":"2.45","first_cure":"OK ( DM )","keterangan":""},{"tgl":"2026-02-05","pic":"SYAIFUL - HARIS","code":"AS137 - 01","size":"AS137","container":"SM AZIV - 327","ket":"CLEANING","mc":"K22R","nilai_preload":"-","vmc":"-","shim_sr":"0.70","shim_pl":"0","first_cure":"OK","keterangan":""}];

const RAKIT_PARTS = ["Cavity Atas","Cavity Bawah","Bead Ring Atas","Bead Ring Bawah","Container L","Container R","Spacer Ring","Segmen"];
  const emptyRakitForm = () => ({ moldSize:"", moldSerial:"", plant:"", line:"", machine:"", partsChecked:[], kondisiCavity:"", kondisiContainer:"", hasilRakit:"ok", catatan:"", operator:"",  date:new Date().toISOString().slice(0,10) });
  const [rakitTab, setRakitTab]       = useState("search");
  const [rakitSearch, setRakitSearch] = useState("");
  const [firstCureData, setFirstCureData]   = useState([]);
  const [firstCureLoaded, setFirstCureLoaded] = useState(false);
  const [uploadFile, setUploadFile]         = useState(null);
  const [uploadPreview, setUploadPreview]   = useState([]);
  const [uploadLoading, setUploadLoading]   = useState(false);
  const [uploadMsg, setUploadMsg]           = useState("");
  const [rakitForm, setRakitForm]     = useState(emptyRakitForm());
  const [rakitRecords, setRakitRecords] = useState([]);

  const role = currentUser?.role || "";
  const showToast = (msg, type="success") => { setToast({msg,type}); setTimeout(()=>setToast(null),2800); };

  // session restore sudah dilakukan via lazy useState init

  const handleLogin = (user) => {
    localStorage.setItem("moldtrack_user", JSON.stringify(user));
    let startPage = "dashboard";
    if (user.role === "teknisi")        startPage = "entry";
    else if (user.role === "persiapan") startPage = "persiapan";
    else if (user.role === "qcgate")    startPage = "qcgate";
    else if (user.role === "naik")      startPage = "naik";
    else if (user.role === "rakit")     startPage = "rakit";
    else if (user.role === "sh")         startPage = "shplan";
    setPage(startPage);
    setCurrentUser(user);
  };

  const handleLogout = () => {
    setCurrentUser(null);
    localStorage.removeItem("moldtrack_user");
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

  const loadShiftPlanRecords = useCallback(async () => {
    const { data } = await supabase.from("shift_plan_records").select("*").order("created_at",{ascending:false});
    setShiftPlanRecords(data||[]);
  },[]);

  const loadFirstCureData = useCallback(async () => {
    const { data } = await supabase.from("first_cure_data").select("*").order("uploaded_at", { ascending: false });
    if (data) { setFirstCureData(data); setFirstCureLoaded(true); }
  }, []);

  const uploadFirstCureData = async (rows) => {
    setUploadLoading(true);
    setUploadMsg("Menghapus data lama...");
    await supabase.from("first_cure_data").delete().neq("id", "00000000-0000-0000-0000-000000000000");
    setUploadMsg("Menyimpan " + rows.length + " record...");
    const chunkSize = 100;
    for (let i = 0; i < rows.length; i += chunkSize) {
      const chunk = rows.slice(i, i + chunkSize).map(r => ({...r, uploaded_by: currentUser?.id}));
      await supabase.from("first_cure_data").insert(chunk);
      setUploadMsg("Menyimpan... " + Math.min(i + chunkSize, rows.length) + "/" + rows.length);
    }
    setUploadLoading(false);
    setUploadMsg("✅ Berhasil upload " + rows.length + " record!");
    setUploadFile(null);
    setUploadPreview([]);
    loadFirstCureData();
  };

  const loadRakitRecords = useCallback(async () => {
    const { data } = await supabase.from("rakit_mold_records").select("*").order("created_at",{ascending:false});
    setRakitRecords(data||[]);
  },[]);

  useEffect(()=>{ if(currentUser){ loadRecords(); loadPrepRecords(); loadQcRecords(); loadNaikRecords(); loadRakitRecords(); loadShiftPlanRecords(); loadFirstCureData(); } },[loadRecords, loadPrepRecords, loadQcRecords, loadNaikRecords, loadRakitRecords, loadShiftPlanRecords, loadFirstCureData, currentUser]);

  useEffect(()=>{
    if (!currentUser) return;
    const ch1 = supabase.channel("rr").on("postgres_changes",{event:"*",schema:"public",table:"repair_records"},()=>loadRecords()).subscribe();
    const ch2 = supabase.channel("pr").on("postgres_changes",{event:"*",schema:"public",table:"preparation_records"},()=>loadPrepRecords()).subscribe();
    const ch3 = supabase.channel("qr").on("postgres_changes",{event:"*",schema:"public",table:"qc_records"},()=>loadQcRecords()).subscribe();
    const ch4 = supabase.channel("nr").on("postgres_changes",{event:"*",schema:"public",table:"naik_mold_records"},()=>loadNaikRecords()).subscribe();
    const ch5 = supabase.channel("rkr").on("postgres_changes",{event:"*",schema:"public",table:"rakit_mold_records"},()=>loadRakitRecords()).subscribe();
    const ch6 = supabase.channel("spr").on("postgres_changes",{event:"*",schema:"public",table:"shift_plan_records"},()=>loadShiftPlanRecords()).subscribe();
    return ()=>{ supabase.removeChannel(ch1); supabase.removeChannel(ch2); supabase.removeChannel(ch3); supabase.removeChannel(ch4); supabase.removeChannel(ch5); supabase.removeChannel(ch6); };
  },[loadRecords, loadPrepRecords, loadQcRecords, loadNaikRecords, loadRakitRecords, currentUser]);

  const submitNaik = async () => {
    if (!naikForm.moldSizeNaik.trim()||!naikForm.plant||!naikForm.line||!naikForm.machine||!naikForm.operator.trim()) {
      showToast("Lengkapi: size mold naik, mesin, dan operator.","error"); return;
    }
    const payload = {
      mold_size_naik:        naikForm.moldSizeNaik.trim().toUpperCase(),
      mold_size_turun:       naikForm.moldSizeTurun.trim().toUpperCase()||null,
      press:                 naikForm.press||[],
      plant:                 naikForm.plant,
      line:                  naikForm.line,
      machine:               naikForm.machine,
      machine_code:          `${naikForm.plant}-${naikForm.line}${naikForm.machine}`,
      maker_container_l:     naikForm.makerContainerL||null,
      maker_container_r:     naikForm.makerContainerR||null,
      type_container_l:      naikForm.typeContainerL||null,
      type_container_r:      naikForm.typeContainerR||null,
      container_no_l:        naikForm.containerNoL.trim().toUpperCase()||null,
      container_no_r:        naikForm.containerNoR.trim().toUpperCase()||null,
      mold_no_l:             naikForm.moldNoL.trim().toUpperCase()||null,
      mold_no_r:             naikForm.moldNoR.trim().toUpperCase()||null,
      turun_maker_container_l: naikForm.turunMakerContainerL||null,
      turun_maker_container_r: naikForm.turunMakerContainerR||null,
      turun_type_container_l:  naikForm.turunTypeContainerL||null,
      turun_type_container_r:  naikForm.turunTypeContainerR||null,
      turun_container_no_l:  naikForm.turunContainerNoL.trim().toUpperCase()||null,
      turun_container_no_r:  naikForm.turunContainerNoR.trim().toUpperCase()||null,
      turun_mold_no_l:       naikForm.turunMoldNoL.trim().toUpperCase()||null,
      turun_mold_no_r:       naikForm.turunMoldNoR.trim().toUpperCase()||null,
      operator:              naikForm.operator.trim(),
      date:                  naikForm.date,
      notes:                 naikForm.notes||null,
      created_by:            currentUser?.id,
      grup:                  getGrup(currentUser?.username),
    };
    const { error } = await supabase.from("naik_mold_records").insert(payload);
    if (error) { showToast("Gagal simpan: "+error.message,"error"); return; }
    const grup = getGrup(currentUser?.username);
    notifSH(grup, naikForm.moldSizeNaik, "Naik Mold", `Mold ${naikForm.moldSizeNaik} berhasil dipasang ke mesin ${naikForm.plant}-${naikForm.line}${naikForm.machine}`);
    showToast("Naik mold berhasil! Notif WA ke Section Head dikirim ✅");
    setNaikForm(emptyNaikForm());
    setNaikTab("database");
  };

  const submitShiftPlan = async () => {
    const validSizes = shiftPlan.moldSizes.map(s=>s.trim().toUpperCase()).filter(Boolean);
    if (!validSizes.length) { showToast("Masukkan minimal 1 size mold.","error"); return; }
    const grup = getGrup(currentUser?.username);
    const payload = validSizes.map(size => ({
      shift:       shiftPlan.shift,
      mold_size:   size,
      date:        shiftPlan.date,
      catatan:     shiftPlan.catatan||null,
      status:      "pending",
      flow:        { persiapan:null, qc1:null, rakit:null, qc2:null, naik:null },
      created_by:  currentUser?.id,
      grup,
    }));
    const { error } = await supabase.from("shift_plan_records").insert(payload);
    if (error) { showToast("Gagal: "+error.message,"error"); return; }
    // Notif WA ke tim Persiapan
    validSizes.forEach(size => notifPersiapan(grup, size, `Shift ${shiftPlan.shift}`));
    showToast(`${validSizes.length} size mold dijadwalkan! Notif WA dikirim ke tim Persiapan ✅`);
    setShiftPlan(emptyShiftPlan());
    setShTab("tracker");
  };

  const submitRakit = async () => {
    if (!rakitForm.moldSize.trim()||!rakitForm.plant||!rakitForm.line||!rakitForm.machine||!rakitForm.operator.trim()) {
      showToast("Lengkapi: size mold, mesin, dan operator.","error"); return;
    }
    const payload = {
      mold_size:        rakitForm.moldSize.trim().toUpperCase(),
      mold_serial:      rakitForm.moldSerial.trim().toUpperCase()||null,
      plant:            rakitForm.plant,
      line:             rakitForm.line,
      machine:          rakitForm.machine,
      machine_code:     `${rakitForm.plant}-${rakitForm.line}${rakitForm.machine}`,
      parts_checked:    rakitForm.partsChecked,
      kondisi_cavity:   rakitForm.kondisiCavity.trim()||null,
      kondisi_container:rakitForm.kondisiContainer.trim()||null,
      hasil_rakit:      rakitForm.hasilRakit,
      catatan:          rakitForm.catatan.trim()||null,
      operator:         rakitForm.operator.trim(),
      date:             rakitForm.date,
      created_by:       currentUser?.id,
      grup:             getGrup(currentUser?.username),
    };
    const { error } = await supabase.from("rakit_mold_records").insert(payload);
    if (error) { showToast("Gagal simpan: "+error.message,"error"); return; }
    const grup = getGrup(currentUser?.username);
    if (rakitForm.hasilRakit === "ok") {
      notifQCGate2(grup, rakitForm.moldSize);
      showToast("Rakit selesai! Notif WA ke QC Gate 2 (kalibrasi) dikirim ✅");
    } else if (rakitForm.hasilRakit === "ditahan") {
      notifSH(grup, rakitForm.moldSize, "Rakit Mold", rakitForm.catatan||"Branding berbeda / ada masalah");
      showToast("Rakit HOLD — Notif WA ke Section Head dikirim ⚠️");
    } else {
      showToast("Rakit dicatat — Perlu review lebih lanjut ✅");
    }
    setRakitForm(emptyRakitForm());
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
      date:form.date,
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
    setForm(emptyForm());setEditId(null);setEntryTab("database");
  };

  const startEdit = (rec) => {
    const pressVal = Array.isArray(rec.press) ? rec.press : (rec.press ? [rec.press] : []);
    setForm({moldSize:rec.mold_size,makerContainerL:rec.maker_container_l||"",makerContainerR:rec.maker_container_r||"",typeContainerL:rec.type_container_l||"",typeContainerR:rec.type_container_r||"",containerNoL:rec.container_no_l||"",containerNoR:rec.container_no_r||"",moldNoL:rec.mold_no_l||"",moldNoR:rec.mold_no_r||"",date:rec.date,technician:rec.technician,plant:rec.plant||"",line:rec.line||"",machine:rec.machine||"",press:pressVal,problems:rec.problems||[],problemDetails:rec.problem_details||{},notes:rec.notes||""});
    setEditId(rec.id);setPage("entry");setEntryTab("form");
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

  // build nav items based on role
  const navItems = [
    ...(canDashboard(role)   ? [["dashboard","ti-layout-dashboard","Dashboard"]] : []),
    ...(canSH(role)          ? [["shplan","ti-clipboard-list","Shift Plan"]] : []),
    ...(canEntry(role)       ? [["entry","ti-plus","Action Problem"]] : []),
    ...(canDatabase(role)    ? [["database","ti-database","Database"]] : []),
    ...(canPersiapan(role)   ? [["persiapan","ti-package","Persiapan"]] : []),
    ...(canQCGate(role)      ? [["qcgate","ti-clipboard-check","QC Gate"]] : []),
    ...(canRakit(role)        ? [["rakit","ti-tools","Rakit Mold"]] : []),
    ...(canNaik(role)        ? [["naik","ti-arrow-up","Naik Mold"]] : []),
    ...(canManageUsers(role) ? [["users","ti-users","Users"]] : []),
  ];
  // Reset tab ke "form" saat user navigasi ke halaman baru
  const navTo = (p) => {
    setPage(p);
    if(p==="shplan") setShTab("form");
    if(p==="entry"){ setForm(emptyForm()); setEditId(null); setEntryTab("form"); }
    if(p==="qcgate") setQcMode("entry");
    if(p==="rakit")  setRakitTab("search");
    if(p==="naik")   setNaikTab("form");
  };

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
                  {page==="shplan"&&"Shift Plan"}
                  {page==="entry"&&canEntry(role)&&(editId?"Edit Record":"Action Problem")}
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
              <button className="btn-sm" onClick={()=>{loadRecords();loadPrepRecords();loadQcRecords();loadNaikRecords();loadRakitRecords();}} style={{ fontSize:11 }}>↻ Refresh</button>
              <button className="btn-sm" onClick={handleLogout} style={{ color:"#E24B4A",borderColor:"#E24B4A",fontSize:11 }}>Keluar</button>
            </div>
          </div>

          <div className="content">

            {/* SHIFT PLAN — Section Head */}
            {page==="shplan"&&canSH(role)&&(
              <div>
              <div>
                {/* TAB */}
                <div style={{ display:"flex",gap:8,marginBottom:16 }}>
                  {[["form","📋 Input Shift"],["tracker","📊 Tracker"]].map(([id,lbl])=>(
                    <button key={id} onClick={()=>setShTab(id)}
                      style={{ flex:1,padding:"10px",borderRadius:8,border:"1.5px solid",cursor:"pointer",fontSize:13,fontWeight:600,
                        borderColor:shTab===id?"#1D9E75":"#e0e0e0",
                        background:shTab===id?"#1D9E75":"#fff",
                        color:shTab===id?"#fff":"#666" }}>
                      {lbl}
                    </button>
                  ))}
                </div>

                {/* FORM INPUT SHIFT */}
                {shTab==="form"&&(
                  <div>
                    <div className="card-title" style={{ textAlign:"center",fontSize:16,marginBottom:16 }}>📋 Input Shift Plan</div>

                    <div className="form-row">
                      <div className="form-group">
                        <label className="form-label">Tanggal *</label>
                        <input type="date" className="form-input" value={shiftPlan.date} onChange={e=>setShiftPlan(f=>({...f,date:e.target.value}))} />
                      </div>
                      <div className="form-group">
                        <label className="form-label">Shift *</label>
                        <select className="form-input" value={shiftPlan.shift} onChange={e=>setShiftPlan(f=>({...f,shift:e.target.value}))}>
                          <option value="1">Shift 1</option>
                          <option value="2">Shift 2</option>
                          <option value="3">Shift 3</option>
                        </select>
                      </div>
                    </div>

                    {/* Size mold list — bisa tambah multiple */}
                    <div className="form-group">
                      <label className="form-label">Size Mold yang akan Naik *</label>
                      {shiftPlan.moldSizes.map((size,idx)=>(
                        <div key={idx} style={{ display:"flex",gap:8,marginBottom:8 }}>
                          <input className="form-input" style={{ flex:1,margin:0 }}
                            placeholder={`cth: BX551`}
                            value={size}
                            onChange={e=>{
                              const arr=[...shiftPlan.moldSizes];
                              arr[idx]=e.target.value;
                              setShiftPlan(f=>({...f,moldSizes:arr}));
                            }}
                          />
                          {shiftPlan.moldSizes.length>1&&(
                            <button onClick={()=>setShiftPlan(f=>({...f,moldSizes:f.moldSizes.filter((_,i)=>i!==idx)}))}
                              style={{ padding:"0 12px",borderRadius:8,border:"1.5px solid #E24B4A",background:"#FCEBEB",color:"#E24B4A",cursor:"pointer",fontSize:16,flexShrink:0 }}>✕</button>
                          )}
                        </div>
                      ))}
                      <button onClick={()=>setShiftPlan(f=>({...f,moldSizes:[...f.moldSizes,""]}))}
                        style={{ width:"100%",padding:"8px",borderRadius:8,border:"1.5px dashed #1D9E75",background:"#E1F5EE",color:"#1D9E75",cursor:"pointer",fontSize:12,fontWeight:600,marginTop:4 }}>
                        + Tambah Size
                      </button>
                    </div>

                    <div className="form-group">
                      <label className="form-label">Catatan</label>
                      <textarea className="form-input" rows={2} placeholder="Catatan shift..." value={shiftPlan.catatan} onChange={e=>setShiftPlan(f=>({...f,catatan:e.target.value}))} style={{ resize:"vertical" }}/>
                    </div>

                    {/* Preview */}
                    {shiftPlan.moldSizes.filter(Boolean).length>0&&(
                      <div style={{ background:"#E1F5EE",borderRadius:8,padding:"12px 14px",marginBottom:14 }}>
                        <div style={{ fontSize:11,color:"#085041",fontWeight:600,marginBottom:8 }}>📋 Preview — {shiftPlan.moldSizes.filter(Boolean).length} size akan dijadwalkan:</div>
                        {shiftPlan.moldSizes.filter(Boolean).map((s,i)=>(
                          <div key={i} style={{ display:"flex",alignItems:"center",gap:8,marginBottom:4 }}>
                            <span style={{ width:6,height:6,borderRadius:"50%",background:"#1D9E75",flexShrink:0,display:"inline-block" }}></span>
                            <span style={{ fontSize:13,fontWeight:600,color:"#085041" }}>{s.toUpperCase()}</span>
                          </div>
                        ))}
                        <div style={{ marginTop:8,fontSize:11,color:"#666" }}>📱 Notif WA otomatis → Tim Persiapan Grup {getGrup(currentUser?.username)?.toUpperCase()}</div>
                      </div>
                    )}

                    <button className="btn-primary" onClick={submitShiftPlan}>
                      📋 Submit Shift Plan + Kirim Notif WA
                    </button>
                    <div style={{ marginTop:10,padding:"10px 14px",background:"#FEF3C7",borderRadius:8,fontSize:11,color:"#92400E" }}>
                      ⚠️ <strong>Catatan:</strong> Jika notif WA tidak terbuka otomatis, pastikan browser mengizinkan popup dari situs ini, atau klik tombol di tab Tracker untuk kirim manual.
                    </div>
                  </div>
                )}

                {/* TRACKER */}
                {shTab==="tracker"&&(()=>{
                  const myGrup = getGrup(currentUser?.username);
                  const myPlans = shiftPlanRecords.filter(p=> !myGrup || p.grup===myGrup);
                  return(
                  <div>
                    <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12 }}>
                      <div style={{ fontSize:13,fontWeight:600,color:"#111" }}>
                        📊 Progress Tracker — Grup {myGrup?.toUpperCase()||"Semua"}
                      </div>
                      <div style={{ fontSize:11,color:"#999" }}>{myPlans.length} size terjadwal</div>
                    </div>
                    {myPlans.length===0&&(
                      <div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>
                        Belum ada shift plan. Silakan input di tab Input Shift.
                      </div>
                    )}
                    {myPlans.map(plan=>{
                      const STEPS = [
                        { id:"persiapan", label:"Persiapan",  icon:"📦" },
                        { id:"qc1",       label:"QC Gate 1",  icon:"🔍" },
                        { id:"rakit",     label:"Rakit Mold", icon:"🔨" },
                        { id:"qc2",       label:"QC Gate 2",  icon:"📐" },
                        { id:"naik",      label:"Naik Mold",  icon:"⬆️" },
                      ];
                      // Hitung status tiap tahap: "done" | "active" | "hold" | "pending"
                      const hasPersiapan  = prepRecords.some(r=>r.mold_size===plan.mold_size&&r.date===plan.date);
                      const qc1Rec        = qcRecords.find(r=>r.mold_size===plan.mold_size&&r.date===plan.date);
                      const hasQC1        = !!qc1Rec;
                      const qc1OK         = qc1Rec?.status==="ok";
                      const qc1Hold       = qc1Rec?.status==="major";
                      const rakitRec      = rakitRecords.find(r=>r.mold_size===plan.mold_size&&r.date===plan.date);
                      const hasRakit      = !!rakitRec;
                      const rakitOK       = rakitRec?.hasil_rakit==="ok";
                      const rakitHold     = rakitRec?.hasil_rakit==="ditahan";
                      const hasNaik       = naikRecords.some(r=>r.mold_size_naik===plan.mold_size&&r.date===plan.date);

                      // Status per step: "done"=hijau, "active"=kuning, "hold"=merah, "pending"=abu
                      const getStatus = (stepId) => {
                        if (stepId==="persiapan") {
                          if (hasPersiapan) return "done";
                          if (!hasPersiapan) return "active"; // selalu aktif pertama
                          return "pending";
                        }
                        if (stepId==="qc1") {
                          if (qc1Hold) return "hold";
                          if (qc1OK)   return "done";
                          if (hasQC1)  return "active"; // minor — sedang diproses
                          if (hasPersiapan) return "active";
                          return "pending";
                        }
                        if (stepId==="rakit") {
                          if (rakitHold) return "hold";
                          if (rakitOK)   return "done";
                          if (hasRakit)  return "active";
                          if (qc1OK)     return "active";
                          return "pending";
                        }
                        if (stepId==="qc2") {
                          if (hasNaik)   return "done";
                          if (rakitOK)   return "active";
                          return "pending";
                        }
                        if (stepId==="naik") {
                          if (hasNaik)   return "done";
                          if (rakitOK)   return "active";
                          return "pending";
                        }
                        return "pending";
                      };
                      const statusColor = { done:"#1D9E75", active:"#E8A020", hold:"#E24B4A", pending:"#e0e0e0" };
                      const statusBorder= { done:"#1D9E75", active:"#E8A020", hold:"#E24B4A", pending:"#ccc" };
                      const statusText  = { done:"DONE", active:"NOW", hold:"HOLD", pending:"" };
                      const statusGlow  = { done:"none", active:"0 0 0 3px #FEF3C7", hold:"0 0 0 3px #FCEBEB", pending:"none" };
                      const doneCount   = STEPS.filter(s=>getStatus(s.id)==="done").length;

                      return(
                        <div key={plan.id} className="card" style={{ marginBottom:12 }}>
                          {/* Header */}
                          <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12 }}>
                            <div>
                              <div style={{ fontSize:18,fontWeight:700,color:"#111" }}>{plan.mold_size}</div>
                              <div style={{ fontSize:11,color:"#999" }}>Shift {plan.shift} · {plan.date} · Grup {plan.grup?.toUpperCase()}</div>
                            </div>
                            <div style={{ textAlign:"right" }}>
                              <div style={{ fontSize:11,fontWeight:600,
                                color:STEPS.some(s=>getStatus(s.id)==="hold")?"#E24B4A":doneCount===5?"#1D9E75":"#E8A020" }}>
                                {STEPS.some(s=>getStatus(s.id)==="hold")?"🚫 Ada HOLD":doneCount===5?"✅ Selesai":`⏳ ${doneCount}/5 tahap`}
                              </div>
                            </div>
                          </div>

                          {/* Progress bar */}
                          <div style={{ position:"relative",marginBottom:16 }}>
                            {/* Garis */}
                            <div style={{ position:"absolute",top:14,left:"10%",right:"10%",height:3,background:"#e0e0e0",borderRadius:2,zIndex:0 }}></div>
                            <div style={{ position:"absolute",top:14,left:"10%",height:3,background:STEPS.some(s=>getStatus(s.id)==="hold")?"#E24B4A":"#1D9E75",borderRadius:2,zIndex:1,
                              width:`${Math.max(0,(doneCount/5)*80)}%`,transition:"width 0.5s" }}></div>
                            {/* Dots */}
                            <div style={{ display:"flex",justifyContent:"space-between",position:"relative",zIndex:2 }}>
                              {STEPS.map((step)=>{
                                const st = getStatus(step.id);
                                const c  = statusColor[st];
                                return(
                                  <div key={step.id} style={{ display:"flex",flexDirection:"column",alignItems:"center",gap:4,flex:1 }}>
                                    <div style={{ width:28,height:28,borderRadius:"50%",
                                      background:c,
                                      display:"flex",alignItems:"center",justifyContent:"center",
                                      fontSize:st==="done"?14:13,
                                      border:`2px solid ${statusBorder[st]}`,
                                      boxShadow:statusGlow[st] }}>
                                      {st==="done"?"✓":st==="hold"?"!":step.icon}
                                    </div>
                                    <div style={{ fontSize:9,color:c,fontWeight:st!=="pending"?600:400,textAlign:"center",lineHeight:1.2 }}>
                                      {step.label}
                                    </div>
                                    <div style={{ fontSize:9,color:c,fontWeight:700 }}>
                                      {statusText[st]}
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          </div>

                          {plan.catatan&&<div style={{ fontSize:11,color:"#666",marginTop:4 }}>📝 {plan.catatan}</div>}
                                                    {/* Tombol WA manual per tahap */}
                          {(()=>{
                            const waUrl = (phone, msg) => "https://wa.me/"+(phone||"")+"?text="+encodeURIComponent(msg);
                            const btnStyle = (bg,clr,bdr) => ({ fontSize:11,padding:"4px 10px",borderRadius:6,background:bg,color:clr,border:"1px solid "+bdr,textDecoration:"none",fontWeight:600 });
                            return(
                              <div style={{ marginTop:8,display:"flex",gap:6,flexWrap:"wrap" }}>
                                {!hasPersiapan&&<a href={waUrl(WA_GROUPS[plan.grup]?.persiapan,"[MoldTrack] PERSIAPAN MOLD - Size: "+plan.mold_size+" - Shift: "+plan.shift+" - Mohon segera persiapan mold.")} target="_blank" rel="noopener noreferrer" style={btnStyle("#E1F5EE","#085041","#1D9E75")}>📱 WA Persiapan</a>}
                                {hasPersiapan&&!hasQC1&&<a href={waUrl(WA_GROUPS[plan.grup]?.qcgate,"[MoldTrack] QC GATE 1 - Size: "+plan.mold_size+" - Mold selesai persiapan. Mohon cek visual.")} target="_blank" rel="noopener noreferrer" style={btnStyle("#FEF3C7","#92400E","#E8A020")}>📱 WA QC Gate 1</a>}
                                {hasQC1&&!hasRakit&&<a href={waUrl(WA_GROUPS[plan.grup]?.rakit,"[MoldTrack] RAKIT MOLD - Size: "+plan.mold_size+" - Lulus QC Gate 1. Mohon segera rakit mold.")} target="_blank" rel="noopener noreferrer" style={btnStyle("#EDE9FE","#5B21B6","#8B5CF6")}>📱 WA Rakit</a>}
                                {hasRakit&&!hasNaik&&<a href={waUrl(WA_GROUPS[plan.grup]?.naik,"[MoldTrack] NAIK MOLD - Size: "+plan.mold_size+" - Lulus QC Gate 2. Siap naik ke mesin.")} target="_blank" rel="noopener noreferrer" style={btnStyle("#FCEBEB","#791F1F","#E24B4A")}>📱 WA Naik</a>}
                              </div>
                            );
                          })()}
                        </div>
                      );
                    })}
                  </div>
                  );
                })()}
              </div>
              </div>
            )}

            {/* DASHBOARD */}
            {page==="dashboard"&&canDashboard(role)&&(
              <div>
                {/* FILTER PERIODE */}
                {(()=>{
                  const now = new Date();
                  const thisMonth = now.toISOString().slice(0,7);
                  // Filter semua tabel berdasarkan filterMonth
                  const applyFilter = (data) => filterMonth==="all" ? data : data.filter(r=>r.date?.startsWith(filterMonth));
                  const filtered     = applyFilter(records);
                  const filteredPrep = applyFilter(prepRecords);
                  const filteredQc   = applyFilter(qcRecords);
                  const filteredNaik = applyFilter(naikRecords);
                  const filteredRakit = applyFilter(rakitRecords);
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

                  // Tren 7 hari terakhir — gabungkan semua proses
                  const last7 = Array.from({length:7},(_,i)=>{
                    const d = new Date(now); d.setDate(d.getDate()-i);
                    const ds = d.toISOString().slice(0,10);
                    return {
                      date: ds.slice(5),
                      action: records.filter(r=>r.date===ds).length,
                      prep:   prepRecords.filter(r=>r.date===ds).length,
                      qc:     qcRecords.filter(r=>r.date===ds).length,
                      naik:   naikRecords.filter(r=>r.date===ds).length,
                      rakit:  rakitRecords.filter(r=>r.date===ds).length,
                      total:  records.filter(r=>r.date===ds).length + prepRecords.filter(r=>r.date===ds).length + qcRecords.filter(r=>r.date===ds).length + naikRecords.filter(r=>r.date===ds).length + rakitRecords.filter(r=>r.date===ds).length,
                    };
                  }).reverse();
                  const maxTren = Math.max(...last7.map(d=>d.total), 1);



                  // QC status summary
                  const qcOk    = filteredQc.filter(r=>r.status==="ok").length;
                  const qcMinor = filteredQc.filter(r=>r.status==="minor").length;
                  const qcMajor = filteredQc.filter(r=>r.status==="major").length;

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
                      <div className="stat-grid" style={{ gridTemplateColumns:"repeat(2,1fr)",marginBottom:16,gap:8 }}>
                        {[
                          {label:"Action Problem",val:filtered.length,sub:filterMonth==="all"?"semua waktu":filterMonth,icon:"ti-tools",color:"#1D9E75"},
                          {label:"Persiapan Mold",val:filteredPrep.length,sub:"keluar gudang",icon:"ti-package",color:"#3B82F6"},
                          {label:"QC Gate",val:filteredQc.length,sub:`OK:${qcOk} Minor:${qcMinor} Major:${qcMajor}`,icon:"ti-clipboard-check",color:"#F59E0B"},
                          {label:"Naik Mold",val:filteredNaik.length,sub:filterMonth==="all"?"semua waktu":filterMonth,icon:"ti-arrow-up",color:"#EF4444"},
                          {label:"Rakit Mold",val:filteredRakit.length,sub:filterMonth==="all"?"semua waktu":filterMonth,icon:"ti-tools",color:"#8B5CF6"},
                        ].map((c,i)=>(
                          <div key={i} className="stat-card" style={{ textAlign:"center" }}>
                            <i className={`ti ${c.icon}`} style={{ fontSize:22,color:c.color,marginBottom:6,display:"block" }}></i>
                            <div className="stat-label">{c.label}</div>
                            <div className="stat-val" style={{ color:c.color }}>{c.val}</div>
                            <div className="stat-sub">{c.sub}</div>
                          </div>
                        ))}
                      </div>
                      {/* STAT ROW 2 */}
                      <div className="stat-grid" style={{ gridTemplateColumns:"repeat(3,1fr)",marginBottom:16 }}>
                        {[
                          {label:"Hari ini (Action)",val:todayFiltered.length,sub:todayStr,icon:"ti-calendar-today",color:"#1D9E75"},

                          {label:"Teknisi aktif",val:techCounts.length,sub:"orang",icon:"ti-users",color:"#6B7280"},
                        ].map((c,i)=>(
                          <div key={i} className="stat-card" style={{ textAlign:"center" }}>
                            <i className={`ti ${c.icon}`} style={{ fontSize:20,color:c.color,marginBottom:4,display:"block" }}></i>
                            <div className="stat-label">{c.label}</div>
                            <div className="stat-val" style={{ fontSize:18 }}>{c.val}</div>
                            <div className="stat-sub">{c.sub}</div>
                          </div>
                        ))}
                      </div>

                      {/* TRACKER HARI INI */}
                      {(()=>{
                        const todayPlans = shiftPlanRecords.filter(p=>p.date===todayStr);
                        if (!todayPlans.length) return null;
                        return(
                          <div className="card" style={{ marginBottom:16 }}>
                            <div className="card-title">🚦 Progress Mold Hari Ini ({todayStr})</div>
                            {todayPlans.map(plan=>{
                              const hasP = prepRecords.some(r=>r.mold_size===plan.mold_size&&r.date===plan.date);
                              const hasQ = qcRecords.some(r=>r.mold_size===plan.mold_size&&r.date===plan.date);
                              const hasR = rakitRecords.some(r=>r.mold_size===plan.mold_size&&r.date===plan.date);
                              const hasN = naikRecords.some(r=>r.mold_size_naik===plan.mold_size&&r.date===plan.date);
                              const steps = [{l:"Persiapan",d:hasP},{l:"QC 1",d:hasQ},{l:"Rakit",d:hasR},{l:"QC 2",d:hasR},{l:"Naik",d:hasN}];
                              const done = steps.filter(s=>s.d).length;
                              return(
                                <div key={plan.id} style={{ marginBottom:12,paddingBottom:12,borderBottom:"1px solid #f0f0f0" }}>
                                  <div style={{ display:"flex",justifyContent:"space-between",marginBottom:8 }}>
                                    <div style={{ fontWeight:700,fontSize:14,color:"#111" }}>{plan.mold_size}</div>
                                    <div style={{ fontSize:11,color:"#999" }}>Shift {plan.shift} · Grup {plan.grup?.toUpperCase()} · {done}/5</div>
                                  </div>
                                  <div style={{ display:"flex",gap:4 }}>
                                    {steps.map((s,i)=>(
                                      <div key={i} style={{ flex:1,textAlign:"center" }}>
                                        <div style={{ height:6,borderRadius:3,background:s.d?"#1D9E75":"#e0e0e0",marginBottom:3 }}></div>
                                        <div style={{ fontSize:9,color:s.d?"#1D9E75":"#999" }}>{s.l}</div>
                                      </div>
                                    ))}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        );
                      })()}

                      {/* TREN 7 HARI */}
                      <div className="card" style={{ marginBottom:16 }}>
                        <div className="card-title">📈 Aktivitas 7 Hari Terakhir (Semua Proses)</div>
                        {/* Legend */}
                        <div style={{ display:"flex",gap:12,marginBottom:10,flexWrap:"wrap" }}>
                          {[["#1D9E75","Action Problem"],["#3B82F6","Persiapan"],["#F59E0B","QC Gate"],["#EF4444","Naik Mold"],["#8B5CF6","Rakit"]].map(([c,l])=>(
                            <div key={l} style={{ display:"flex",alignItems:"center",gap:4,fontSize:10,color:"#666" }}>
                              <div style={{ width:8,height:8,borderRadius:2,background:c }}></div>{l}
                            </div>
                          ))}
                        </div>
                        <div style={{ display:"flex",alignItems:"flex-end",gap:8,height:90,marginTop:8 }}>
                          {last7.map((d,i)=>(
                            <div key={i} style={{ flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:4 }}>
                              <div style={{ fontSize:11,fontWeight:600,color:"#555" }}>{d.total||""}</div>
                              <div style={{ width:"100%",display:"flex",flexDirection:"column",gap:1 }}>
                                {[["#8B5CF6",d.rakit],["#EF4444",d.naik],["#F59E0B",d.qc],["#3B82F6",d.prep],["#1D9E75",d.action]].map(([c,v],j)=>
                                  v>0?<div key={j} style={{ width:"100%",background:c,height:`${Math.round((v/maxTren)*60)}px`,minHeight:2 }}></div>:null
                                )}
                                {d.total===0&&<div style={{ width:"100%",background:"#e0e0e0",height:4,borderRadius:"2px 2px 0 0" }}></div>}
                              </div>
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
                          {label:"Persiapan Mold", data:filteredPrep},
                          {label:"QC Gate",        data:filteredQc},
                          {label:"Naik Mold",      data:filteredNaik},
                          {label:"Rakit Mold",     data:filteredRakit},
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
                        {canEntry(role)&&<button className="btn-primary" style={{ width:"auto",padding:"8px 16px",margin:0 }} onClick={()=>navTo("entry")}>
                          <i className="ti ti-plus" style={{ marginRight:4 }}></i>Entry baru
                        </button>}
                      </div>
                      {filtered.length===0&&<div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>Belum ada record.</div>}
                      {filtered.slice(0,5).map(r=>(
                        <div key={r.id} className="record-card" onClick={()=>{setDetailId(r.id);setPage("detail");}}>
                          <div className="record-card-header">
                            <div><div className="record-size">{r.mold_size}</div><div className="record-type">{r.maker_container_l||r.maker_container_r||"-"}</div></div>
                            <div style={{ fontSize:11,color:"#999",textAlign:"right" }}><div>{r.date}</div></div>
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
            {/* SHIFT PLAN — Section Head */}
            {page==="shplan"&&canSH(role)&&(
              <div>
                {/* TAB */}
                <div style={{ display:"flex",gap:8,marginBottom:16 }}>
                  {[["form","📋 Input Shift"],["tracker","📊 Tracker"]].map(([id,lbl])=>(
                    <button key={id} onClick={()=>setShTab(id)}
                      style={{ flex:1,padding:"10px",borderRadius:8,border:"1.5px solid",cursor:"pointer",fontSize:13,fontWeight:600,
                        borderColor:shTab===id?"#1D9E75":"#e0e0e0",
                        background:shTab===id?"#1D9E75":"#fff",
                        color:shTab===id?"#fff":"#666" }}>
                      {lbl}
                    </button>
                  ))}
                </div>

                {/* FORM INPUT SHIFT */}
                {shTab==="form"&&(
                  <div>
                    <div className="card-title" style={{ textAlign:"center",fontSize:16,marginBottom:16 }}>📋 Input Shift Plan</div>

                    <div className="form-row">
                      <div className="form-group">
                        <label className="form-label">Tanggal *</label>
                        <input type="date" className="form-input" value={shiftPlan.date} onChange={e=>setShiftPlan(f=>({...f,date:e.target.value}))} />
                      </div>
                      <div className="form-group">
                        <label className="form-label">Shift *</label>
                        <select className="form-input" value={shiftPlan.shift} onChange={e=>setShiftPlan(f=>({...f,shift:e.target.value}))}>
                          <option value="1">Shift 1</option>
                          <option value="2">Shift 2</option>
                          <option value="3">Shift 3</option>
                        </select>
                      </div>
                    </div>

                    {/* Size mold list — bisa tambah multiple */}
                    <div className="form-group">
                      <label className="form-label">Size Mold yang akan Naik *</label>
                      {shiftPlan.moldSizes.map((size,idx)=>(
                        <div key={idx} style={{ display:"flex",gap:8,marginBottom:8 }}>
                          <input className="form-input" style={{ flex:1,margin:0 }}
                            placeholder={`cth: BX551`}
                            value={size}
                            onChange={e=>{
                              const arr=[...shiftPlan.moldSizes];
                              arr[idx]=e.target.value;
                              setShiftPlan(f=>({...f,moldSizes:arr}));
                            }}
                          />
                          {shiftPlan.moldSizes.length>1&&(
                            <button onClick={()=>setShiftPlan(f=>({...f,moldSizes:f.moldSizes.filter((_,i)=>i!==idx)}))}
                              style={{ padding:"0 12px",borderRadius:8,border:"1.5px solid #E24B4A",background:"#FCEBEB",color:"#E24B4A",cursor:"pointer",fontSize:16,flexShrink:0 }}>✕</button>
                          )}
                        </div>
                      ))}
                      <button onClick={()=>setShiftPlan(f=>({...f,moldSizes:[...f.moldSizes,""]}))}
                        style={{ width:"100%",padding:"8px",borderRadius:8,border:"1.5px dashed #1D9E75",background:"#E1F5EE",color:"#1D9E75",cursor:"pointer",fontSize:12,fontWeight:600,marginTop:4 }}>
                        + Tambah Size
                      </button>
                    </div>

                    <div className="form-group">
                      <label className="form-label">Catatan</label>
                      <textarea className="form-input" rows={2} placeholder="Catatan shift..." value={shiftPlan.catatan} onChange={e=>setShiftPlan(f=>({...f,catatan:e.target.value}))} style={{ resize:"vertical" }}/>
                    </div>

                    {/* Preview */}
                    {shiftPlan.moldSizes.filter(Boolean).length>0&&(
                      <div style={{ background:"#E1F5EE",borderRadius:8,padding:"12px 14px",marginBottom:14 }}>
                        <div style={{ fontSize:11,color:"#085041",fontWeight:600,marginBottom:8 }}>📋 Preview — {shiftPlan.moldSizes.filter(Boolean).length} size akan dijadwalkan:</div>
                        {shiftPlan.moldSizes.filter(Boolean).map((s,i)=>(
                          <div key={i} style={{ display:"flex",alignItems:"center",gap:8,marginBottom:4 }}>
                            <span style={{ width:6,height:6,borderRadius:"50%",background:"#1D9E75",flexShrink:0,display:"inline-block" }}></span>
                            <span style={{ fontSize:13,fontWeight:600,color:"#085041" }}>{s.toUpperCase()}</span>
                          </div>
                        ))}
                        <div style={{ marginTop:8,fontSize:11,color:"#666" }}>📱 Notif WA otomatis → Tim Persiapan Grup {getGrup(currentUser?.username)?.toUpperCase()}</div>
                      </div>
                    )}

                    <button className="btn-primary" onClick={submitShiftPlan}>
                      📋 Submit Shift Plan + Kirim Notif WA
                    </button>
                    <div style={{ marginTop:10,padding:"10px 14px",background:"#FEF3C7",borderRadius:8,fontSize:11,color:"#92400E" }}>
                      ⚠️ <strong>Catatan:</strong> Jika notif WA tidak terbuka otomatis, pastikan browser mengizinkan popup dari situs ini, atau klik tombol di tab Tracker untuk kirim manual.
                    </div>
                  </div>
                )}

                {/* TRACKER */}
                {shTab==="tracker"&&(()=>{
                  const myGrup = getGrup(currentUser?.username);
                  const myPlans = shiftPlanRecords.filter(p=> !myGrup || p.grup===myGrup);
                  return(
                  <div>
                    <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12 }}>
                      <div style={{ fontSize:13,fontWeight:600,color:"#111" }}>
                        📊 Progress Tracker — Grup {myGrup?.toUpperCase()||"Semua"}
                      </div>
                      <div style={{ fontSize:11,color:"#999" }}>{myPlans.length} size terjadwal</div>
                    </div>
                    {myPlans.length===0&&(
                      <div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>
                        Belum ada shift plan. Silakan input di tab Input Shift.
                      </div>
                    )}
                    {shiftPlanRecords.map(plan=>{
                      const STEPS = [
                        { id:"persiapan", label:"Persiapan",  icon:"📦" },
                        { id:"qc1",       label:"QC Gate 1",  icon:"🔍" },
                        { id:"rakit",     label:"Rakit Mold", icon:"🔨" },
                        { id:"qc2",       label:"QC Gate 2",  icon:"📐" },
                        { id:"naik",      label:"Naik Mold",  icon:"⬆️" },
                      ];
                      // Hitung status tiap tahap: "done" | "active" | "hold" | "pending"
                      const hasPersiapan  = prepRecords.some(r=>r.mold_size===plan.mold_size&&r.date===plan.date);
                      const qc1Rec        = qcRecords.find(r=>r.mold_size===plan.mold_size&&r.date===plan.date);
                      const hasQC1        = !!qc1Rec;
                      const qc1OK         = qc1Rec?.status==="ok";
                      const qc1Hold       = qc1Rec?.status==="major";
                      const rakitRec      = rakitRecords.find(r=>r.mold_size===plan.mold_size&&r.date===plan.date);
                      const hasRakit      = !!rakitRec;
                      const rakitOK       = rakitRec?.hasil_rakit==="ok";
                      const rakitHold     = rakitRec?.hasil_rakit==="ditahan";
                      const hasNaik       = naikRecords.some(r=>r.mold_size_naik===plan.mold_size&&r.date===plan.date);

                      // Status per step: "done"=hijau, "active"=kuning, "hold"=merah, "pending"=abu
                      const getStatus = (stepId) => {
                        if (stepId==="persiapan") {
                          if (hasPersiapan) return "done";
                          if (!hasPersiapan) return "active"; // selalu aktif pertama
                          return "pending";
                        }
                        if (stepId==="qc1") {
                          if (qc1Hold) return "hold";
                          if (qc1OK)   return "done";
                          if (hasQC1)  return "active"; // minor — sedang diproses
                          if (hasPersiapan) return "active";
                          return "pending";
                        }
                        if (stepId==="rakit") {
                          if (rakitHold) return "hold";
                          if (rakitOK)   return "done";
                          if (hasRakit)  return "active";
                          if (qc1OK)     return "active";
                          return "pending";
                        }
                        if (stepId==="qc2") {
                          if (hasNaik)   return "done";
                          if (rakitOK)   return "active";
                          return "pending";
                        }
                        if (stepId==="naik") {
                          if (hasNaik)   return "done";
                          if (rakitOK)   return "active";
                          return "pending";
                        }
                        return "pending";
                      };
                      const statusColor = { done:"#1D9E75", active:"#E8A020", hold:"#E24B4A", pending:"#e0e0e0" };
                      const statusBorder= { done:"#1D9E75", active:"#E8A020", hold:"#E24B4A", pending:"#ccc" };
                      const statusText  = { done:"DONE", active:"NOW", hold:"HOLD", pending:"" };
                      const statusGlow  = { done:"none", active:"0 0 0 3px #FEF3C7", hold:"0 0 0 3px #FCEBEB", pending:"none" };
                      const doneCount   = STEPS.filter(s=>getStatus(s.id)==="done").length;

                      return(
                        <div key={plan.id} className="card" style={{ marginBottom:12 }}>
                          {/* Header */}
                          <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12 }}>
                            <div>
                              <div style={{ fontSize:18,fontWeight:700,color:"#111" }}>{plan.mold_size}</div>
                              <div style={{ fontSize:11,color:"#999" }}>Shift {plan.shift} · {plan.date} · Grup {plan.grup?.toUpperCase()}</div>
                            </div>
                            <div style={{ textAlign:"right" }}>
                              <div style={{ fontSize:11,fontWeight:600,
                                color:STEPS.some(s=>getStatus(s.id)==="hold")?"#E24B4A":doneCount===5?"#1D9E75":"#E8A020" }}>
                                {STEPS.some(s=>getStatus(s.id)==="hold")?"🚫 Ada HOLD":doneCount===5?"✅ Selesai":`⏳ ${doneCount}/5 tahap`}
                              </div>
                            </div>
                          </div>

                          {/* Progress bar */}
                          <div style={{ position:"relative",marginBottom:16 }}>
                            {/* Garis */}
                            <div style={{ position:"absolute",top:14,left:"10%",right:"10%",height:3,background:"#e0e0e0",borderRadius:2,zIndex:0 }}></div>
                            <div style={{ position:"absolute",top:14,left:"10%",height:3,background:STEPS.some(s=>getStatus(s.id)==="hold")?"#E24B4A":"#1D9E75",borderRadius:2,zIndex:1,
                              width:`${Math.max(0,(doneCount/5)*80)}%`,transition:"width 0.5s" }}></div>
                            {/* Dots */}
                            <div style={{ display:"flex",justifyContent:"space-between",position:"relative",zIndex:2 }}>
                              {STEPS.map((step)=>{
                                const st = getStatus(step.id);
                                const c  = statusColor[st];
                                return(
                                  <div key={step.id} style={{ display:"flex",flexDirection:"column",alignItems:"center",gap:4,flex:1 }}>
                                    <div style={{ width:28,height:28,borderRadius:"50%",
                                      background:c,
                                      display:"flex",alignItems:"center",justifyContent:"center",
                                      fontSize:st==="done"?14:13,
                                      border:`2px solid ${statusBorder[st]}`,
                                      boxShadow:statusGlow[st] }}>
                                      {st==="done"?"✓":st==="hold"?"!":step.icon}
                                    </div>
                                    <div style={{ fontSize:9,color:c,fontWeight:st!=="pending"?600:400,textAlign:"center",lineHeight:1.2 }}>
                                      {step.label}
                                    </div>
                                    <div style={{ fontSize:9,color:c,fontWeight:700 }}>
                                      {statusText[st]}
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          </div>

                          {plan.catatan&&<div style={{ fontSize:11,color:"#666",marginTop:4 }}>📝 {plan.catatan}</div>}
                        </div>
                      );
                    })}
                  </div>
                  );
                })()}
              </div>
            )}

            {page==="entry"&&canEntry(role)&&(
              <div>
                {/* TAB NAVIGATION */}
                <div style={{ display:"flex",gap:8,marginBottom:16 }}>
                  {[["form","✏️ Input"],["database","📋 Database"]].map(([id,lbl])=>(
                    <button key={id} onClick={()=>setEntryTab(id)}
                      style={{ flex:1,padding:"10px",borderRadius:8,border:"1.5px solid",cursor:"pointer",fontSize:13,fontWeight:600,
                        borderColor:entryTab===id?"#1D9E75":"#e0e0e0",
                        background:entryTab===id?"#1D9E75":"#fff",
                        color:entryTab===id?"#fff":"#666" }}>
                      {lbl}
                    </button>
                  ))}
                </div>
                {entryTab==="form"&&<div>
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
                </div>}
                {entryTab==="database"&&                <div>
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
                      <div style={{ fontSize:11,color:"#999",textAlign:"right" }}><div>{r.date}</div></div>
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
                </div>}
              </div>
            )}

            {/* DATABASE — tab per proses untuk admin/adh/dh/analyst */}
            {page==="database"&&canDatabase(role)&&(
              <div>
                {/* TAB PROSES */}
                <div style={{ display:"flex",gap:6,marginBottom:14,overflowX:"auto",paddingBottom:2 }}>
                  {[
                    ["action","🔧 Action Problem", records.length],
                    ["persiapan","📦 Persiapan", prepRecords.length],
                    ["qc","🔍 QC Gate", qcRecords.length],
                    ["rakit","🔨 Rakit Mold", rakitRecords.length],
                    ["naik","⬆️ Naik Mold", naikRecords.length],
                    ["tracker","🚦 Shift Tracker", shiftPlanRecords.length],
                  ].map(([id,lbl,cnt])=>(
                    <button key={id} onClick={()=>setDbTab(id)}
                      style={{ flexShrink:0,padding:"8px 14px",borderRadius:8,border:"1.5px solid",cursor:"pointer",fontSize:12,fontWeight:600,
                        borderColor:dbTab===id?"#1D9E75":"#e0e0e0",
                        background:dbTab===id?"#1D9E75":"#fff",
                        color:dbTab===id?"#fff":"#666" }}>
                      {lbl} <span style={{ fontSize:10,opacity:0.8,marginLeft:2 }}>({cnt})</span>
                    </button>
                  ))}
                </div>

                {/* TAB: ACTION PROBLEM */}
                {dbTab==="action"&&(
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
                      <button style={{ padding:"8px 14px",borderRadius:8,border:"1.5px solid #1D9E75",background:"#E1F5EE",color:"#085041",fontSize:12,fontWeight:600,cursor:"pointer",flexShrink:0 }} onClick={()=>exportToExcel(records)}>⬇ Excel</button>
                    </div>
                    <div style={{ fontSize:12,color:"#999",marginBottom:10 }}>{filtered.length} record</div>
                    {filtered.length===0&&<div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>Belum ada data.</div>}
                    {filtered.map(r=>(
                      <div key={r.id} className="record-card">
                        <div className="record-card-header" onClick={()=>{setDetailId(r.id);setPage("detail");}}>
                          <div><div className="record-size">{r.mold_size}</div><div className="record-type">{r.mold_type==="segmented"?"Segmented":"Two Piece"}</div></div>
                          <div style={{ fontSize:11,color:"#999",textAlign:"right" }}><div>{r.date}</div></div>
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

                {/* TAB: PERSIAPAN */}
                {dbTab==="persiapan"&&(
                  <div>
                    <div style={{ fontSize:12,color:"#999",marginBottom:10 }}>{prepRecords.length} record</div>
                    {prepRecords.length===0&&<div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>Belum ada data.</div>}
                    {prepRecords.map(r=>(
                      <div key={r.id} className="record-card">
                        <div className="record-card-header">
                          <div>
                            <div className="record-size">{r.mold_size}</div>
                            <div className="record-type">{r.mold_serial}</div>
                          </div>
                          <div style={{ textAlign:"right" }}>
                            <div style={{ fontSize:11,color:"#999" }}>{r.date}</div>
                            <div style={{ marginTop:4 }}>
                              <span style={{ fontSize:11,padding:"2px 10px",borderRadius:20,fontWeight:600,
                                background:r.status==="out"?"#FEF3C7":"#E1F5EE",
                                color:r.status==="out"?"#92400E":"#085041" }}>
                                {r.status==="out"?"📤 Keluar":"📥 Kembali"}
                              </span>
                            </div>
                          </div>
                        </div>
                        {r.slot_location&&<div style={{ fontSize:11,color:"#666",marginBottom:4 }}>📍 {r.slot_location}</div>}
                        <div className="record-meta"><i className="ti ti-user" style={{ fontSize:12,marginRight:4 }}></i>{r.operator} {r.grup?`· Grup ${r.grup}`:""}</div>
                      </div>
                    ))}
                  </div>
                )}

                {/* TAB: QC GATE */}
                {dbTab==="qc"&&(()=>{
                  const statusCfg = {
                    ok:    { label:"✅ OK",    bg:"#E1F5EE", color:"#085041" },
                    minor: { label:"⚠️ Minor", bg:"#FEF3C7", color:"#92400E" },
                    major: { label:"🚫 Major", bg:"#FCEBEB", color:"#791F1F" },
                  };
                  return(
                    <div>
                      <div style={{ fontSize:12,color:"#999",marginBottom:10 }}>{qcRecords.length} record</div>
                      {qcRecords.length===0&&<div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>Belum ada data.</div>}
                      {qcRecords.map(r=>{
                        const cfg=statusCfg[r.status]||statusCfg.ok;
                        return(
                          <div key={r.id} className="record-card">
                            <div className="record-card-header">
                              <div>
                                <div className="record-size">{r.mold_size}</div>
                                {r.mold_serial&&<div className="record-type">SN: {r.mold_serial}</div>}
                              </div>
                              <div style={{ textAlign:"right" }}>
                                <div style={{ fontSize:11,color:"#999" }}>{r.date}</div>
                                <div style={{ marginTop:4 }}><span style={{ fontSize:11,padding:"2px 10px",borderRadius:20,fontWeight:600,background:cfg.bg,color:cfg.color }}>{cfg.label}</span></div>
                              </div>
                            </div>
                            {r.machine_code&&<div style={{ fontSize:12,fontWeight:600,color:"#1D9E75",marginBottom:4 }}><i className="ti ti-robot" style={{ fontSize:13,marginRight:4 }}></i>{r.machine_code}</div>}
                            {r.defects?.length>0&&<div style={{ fontSize:11,color:"#666",marginBottom:4 }}>Cacat: {r.defects.map(d=><span key={d} style={{ background:"#FCEBEB",color:"#791F1F",fontSize:11,padding:"1px 8px",borderRadius:4,marginRight:4 }}>{d}</span>)}</div>}
                            {r.repair_notes&&<div style={{ fontSize:11,color:"#666",marginBottom:4 }}>📝 {r.repair_notes}</div>}
                            <div className="record-meta"><i className="ti ti-user" style={{ fontSize:12,marginRight:4 }}></i>{r.checker} {r.grup?`· Grup ${r.grup}`:""}</div>
                          </div>
                        );
                      })}
                    </div>
                  );
                })()}

                {/* TAB: RAKIT MOLD */}
                {dbTab==="rakit"&&(()=>{
                  const hasilCfg = {
                    ok:           { label:"✅ OK",           bg:"#E1F5EE", color:"#085041" },
                    perlu_review: { label:"⚠️ Perlu Review", bg:"#FEF3C7", color:"#92400E" },
                    ditahan:      { label:"🚫 Ditahan",       bg:"#FCEBEB", color:"#791F1F" },
                  };
                  return(
                    <div>
                      <div style={{ fontSize:12,color:"#999",marginBottom:10 }}>{rakitRecords.length} record</div>
                      {rakitRecords.length===0&&<div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>Belum ada data.</div>}
                      {rakitRecords.map(r=>{
                        const cfg=hasilCfg[r.hasil_rakit]||hasilCfg.ok;
                        return(
                          <div key={r.id} className="record-card">
                            <div className="record-card-header">
                              <div>
                                <div className="record-size">🔨 {r.mold_size}</div>
                                {r.mold_serial&&<div className="record-type">SN: {r.mold_serial}</div>}
                              </div>
                              <div style={{ textAlign:"right" }}>
                                <div style={{ fontSize:11,color:"#999" }}>{r.date}</div>
                                <div style={{ marginTop:4 }}><span style={{ fontSize:11,padding:"2px 10px",borderRadius:20,fontWeight:600,background:cfg.bg,color:cfg.color }}>{cfg.label}</span></div>
                              </div>
                            </div>
                            {r.machine_code&&<div style={{ fontSize:12,fontWeight:600,color:"#1D9E75",marginBottom:4 }}><i className="ti ti-robot" style={{ fontSize:13,marginRight:4 }}></i>{r.machine_code}</div>}
                            {r.parts_checked?.length>0&&<div style={{ fontSize:11,color:"#666",marginBottom:4 }}>✅ Komponen: {r.parts_checked.length}/8 dicek</div>}
                            {r.catatan&&<div style={{ fontSize:11,color:"#666",marginBottom:4 }}>📝 {r.catatan}</div>}
                            <div className="record-meta"><i className="ti ti-user" style={{ fontSize:12,marginRight:4 }}></i>{r.operator} {r.grup?`· Grup ${r.grup}`:""}</div>
                          </div>
                        );
                      })}
                    </div>
                  );
                })()}

                {/* TAB: SHIFT TRACKER */}
                {dbTab==="tracker"&&(
                  <div>
                    <div style={{ fontSize:12,color:"#999",marginBottom:12 }}>{shiftPlanRecords.length} shift plan tercatat</div>
                    {shiftPlanRecords.length===0&&<div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>Belum ada data shift plan.</div>}
                    {shiftPlanRecords.map(plan=>{
                      const hasP = prepRecords.some(r=>r.mold_size===plan.mold_size&&r.date===plan.date);
                      const hasQ = qcRecords.some(r=>r.mold_size===plan.mold_size&&r.date===plan.date);
                      const hasR = rakitRecords.some(r=>r.mold_size===plan.mold_size&&r.date===plan.date);
                      const hasN = naikRecords.some(r=>r.mold_size_naik===plan.mold_size&&r.date===plan.date);
                      const steps = [
                        {l:"Persiapan",d:hasP,icon:"📦"},
                        {l:"QC Gate 1",d:hasQ,icon:"🔍"},
                        {l:"Rakit",d:hasR,icon:"🔨"},
                        {l:"QC Gate 2",d:hasR,icon:"📐"},
                        {l:"Naik Mold",d:hasN,icon:"⬆️"},
                      ];
                      const done = steps.filter(s=>s.d).length;
                      return(
                        <div key={plan.id} className="card" style={{ marginBottom:12 }}>
                          <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10 }}>
                            <div>
                              <div style={{ fontSize:16,fontWeight:700,color:"#111" }}>{plan.mold_size}</div>
                              <div style={{ fontSize:11,color:"#999" }}>Shift {plan.shift} · {plan.date} · Grup {plan.grup?.toUpperCase()}</div>
                            </div>
                            <span style={{ fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600,
                              background:done===5?"#E1F5EE":done>0?"#FEF3C7":"#f0f0f0",
                              color:done===5?"#085041":done>0?"#92400E":"#999" }}>
                              {done===5?"✅ Selesai":done>0?`⏳ ${done}/5`:"🔴 Belum mulai"}
                            </span>
                          </div>
                          <div style={{ display:"flex",gap:4,marginBottom:8 }}>
                            {steps.map((s,i)=>(
                              <div key={i} style={{ flex:1,textAlign:"center" }}>
                                <div style={{ height:6,borderRadius:3,background:s.d?"#1D9E75":"#e0e0e0",marginBottom:3,transition:"background 0.3s" }}></div>
                                <div style={{ fontSize:9,color:s.d?"#1D9E75":"#999",fontWeight:s.d?600:400 }}>{s.l}</div>
                              </div>
                            ))}
                          </div>
                          {plan.catatan&&<div style={{ fontSize:11,color:"#666" }}>📝 {plan.catatan}</div>}
                        </div>
                      );
                    })}
                  </div>
                )}

                {/* TAB: NAIK MOLD */}
                {dbTab==="naik"&&(
                  <div>
                    <div style={{ fontSize:12,color:"#999",marginBottom:10 }}>{naikRecords.length} record</div>
                    {naikRecords.length===0&&<div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>Belum ada data.</div>}
                    {naikRecords.map(r=>(
                      <div key={r.id} className="record-card">
                        <div className="record-card-header">
                          <div>
                            <div className="record-size">⬆️ {r.mold_size_naik}</div>
                            {r.mold_size_turun&&<div className="record-type">⬇️ Turun: {r.mold_size_turun}</div>}
                          </div>
                          <div style={{ fontSize:11,color:"#999",textAlign:"right" }}><div>{r.date}</div></div>
                        </div>
                        {r.machine_code&&<div style={{ fontSize:12,fontWeight:600,color:"#1D9E75",marginBottom:4 }}><i className="ti ti-robot" style={{ fontSize:13,marginRight:4 }}></i>{r.machine_code}</div>}
                        {(r.press||[]).length>0&&<div style={{ fontSize:11,color:"#666",marginBottom:2 }}>Press: <strong>{r.press.join(" & ")}</strong></div>}
                        {(r.maker_container_l||r.maker_container_r)&&<div style={{ fontSize:11,color:"#666",marginBottom:2 }}>Maker: {[r.maker_container_l,r.maker_container_r].filter(Boolean).join(" / ")}</div>}
                        {(r.container_no_l||r.container_no_r)&&<div style={{ fontSize:11,color:"#666",marginBottom:2 }}>Container: {[r.container_no_l,r.container_no_r].filter(Boolean).join(" / ")}</div>}
                        {r.notes&&<div style={{ fontSize:11,color:"#666",marginBottom:4 }}>📝 {r.notes}</div>}
                        <div className="record-meta"><i className="ti ti-user" style={{ fontSize:12,marginRight:4 }}></i>{r.operator} {r.grup?`· Grup ${r.grup}`:""}</div>
                      </div>
                    ))}
                  </div>
                )}
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
                        <div style={{ fontSize:11,color:"#999",marginBottom:6 }}>{r.date}{r.machine_code?` · ${r.machine_code}`:""}</div>
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
                {/* ANTRIAN SHIFT PLAN */}
                {(()=>{
                  const myGrup = getGrup(currentUser?.username);
                  const today = new Date().toISOString().slice(0,10);
                  const antrian = shiftPlanRecords.filter(p=>p.grup===myGrup&&p.date===today&&(!prepRecords.some(r=>r.mold_size===p.mold_size&&r.date===today)));
                  if (!antrian.length) return null;
                  return(
                    <div style={{ background:"#E1F5EE",borderRadius:10,padding:"10px 14px",marginBottom:14,border:"1.5px solid #1D9E75" }}>
                      <div style={{ fontSize:12,fontWeight:700,color:"#085041",marginBottom:8 }}>
                        📦 Antrian Shift Plan hari ini ({antrian.length} size)
                      </div>
                      <div style={{ display:"flex",flexWrap:"wrap",gap:6 }}>
                        {antrian.map(p=>(
                          <span key={p.id}
                            style={{ background:"#fff",border:"1.5px solid #1D9E75",borderRadius:20,padding:"4px 12px",fontSize:12,fontWeight:700,color:"#085041",cursor:"pointer" }}
                            onClick={()=>{ setPrepForm(f=>({...f,moldSize:p.mold_size})) }}>
                            {p.mold_size}
                          </span>
                        ))}
                      </div>
                      <div style={{ fontSize:10,color:"#666",marginTop:6 }}>💡 Klik size untuk auto-isi form</div>
                    </div>
                  );
                })()}
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
                {/* ANTRIAN SHIFT PLAN */}
                {(()=>{
                  const myGrup = getGrup(currentUser?.username);
                  const today = new Date().toISOString().slice(0,10);
                  const antrian = shiftPlanRecords.filter(p=>p.grup===myGrup&&p.date===today&&(prepRecords.some(r=>r.mold_size===p.mold_size&&r.date===today)&&!qcRecords.some(r=>r.mold_size===p.mold_size&&r.date===today)));
                  if (!antrian.length) return null;
                  return(
                    <div style={{ background:"#FEF9E7",borderRadius:10,padding:"10px 14px",marginBottom:14,border:"1.5px solid #1D9E75" }}>
                      <div style={{ fontSize:12,fontWeight:700,color:"#085041",marginBottom:8 }}>
                        🔍 Antrian Shift Plan hari ini ({antrian.length} size)
                      </div>
                      <div style={{ display:"flex",flexWrap:"wrap",gap:6 }}>
                        {antrian.map(p=>(
                          <span key={p.id}
                            style={{ background:"#fff",border:"1.5px solid #1D9E75",borderRadius:20,padding:"4px 12px",fontSize:12,fontWeight:700,color:"#085041",cursor:"pointer" }}
                            onClick={()=>{ setQcForm(f=>({...f,moldSize:p.mold_size}));setQcMode('entry') }}>
                            {p.mold_size}
                          </span>
                        ))}
                      </div>
                      <div style={{ fontSize:10,color:"#666",marginTop:6 }}>💡 Klik size untuk auto-isi form</div>
                    </div>
                  );
                })()}

                {/* TABS */}
                <div style={{ display:"flex",gap:8,marginBottom:16 }}>
                  {[["entry","✏️ Input QC"],["history","📋 Database"]].map(([m,lbl])=>(
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

            {/* RAKIT MOLD */}
            {page==="rakit"&&canRakit(role)&&(
              <div>
                {/* ANTRIAN SHIFT PLAN */}
                {(()=>{
                  const myGrup = getGrup(currentUser?.username);
                  const today = new Date().toISOString().slice(0,10);
                  const antrian = shiftPlanRecords.filter(p=>p.grup===myGrup&&p.date===today&&(qcRecords.some(r=>r.mold_size===p.mold_size&&r.date===today&&r.status==="ok")&&!rakitRecords.some(r=>r.mold_size===p.mold_size&&r.date===today)));
                  if (!antrian.length) return null;
                  return(
                    <div style={{ background:"#EDE9FE",borderRadius:10,padding:"10px 14px",marginBottom:14,border:"1.5px solid #1D9E75" }}>
                      <div style={{ fontSize:12,fontWeight:700,color:"#085041",marginBottom:8 }}>
                        🔨 Antrian Shift Plan hari ini ({antrian.length} size)
                      </div>
                      <div style={{ display:"flex",flexWrap:"wrap",gap:6 }}>
                        {antrian.map(p=>(
                          <span key={p.id}
                            style={{ background:"#fff",border:"1.5px solid #1D9E75",borderRadius:20,padding:"4px 12px",fontSize:12,fontWeight:700,color:"#085041",cursor:"pointer" }}
                            onClick={()=>{ setRakitForm(f=>({...f,moldSize:p.mold_size}));setRakitTab('form') }}>
                            {p.mold_size}
                          </span>
                        ))}
                      </div>
                      <div style={{ fontSize:10,color:"#666",marginTop:6 }}>💡 Klik size untuk auto-isi form</div>
                    </div>
                  );
                })()}

                {/* TAB NAVIGATION */}
                <div style={{ display:"flex",gap:8,marginBottom:16 }}>
                  {[["search","🔍 Cari Data"], ...(canAnalyst(role)?[["upload","📤 Upload Data"]]:[])].map(([id,lbl])=>(
                    <button key={id} onClick={()=>setRakitTab(id)}
                      style={{ flex:1,padding:"10px",borderRadius:8,border:"1.5px solid",cursor:"pointer",fontSize:13,fontWeight:600,
                        borderColor:rakitTab===id?"#1D9E75":"#e0e0e0",
                        background:rakitTab===id?"#1D9E75":"#fff",
                        color:rakitTab===id?"#fff":"#666" }}>
                      {lbl}
                    </button>
                  ))}
                </div>

                {/* SEARCH TAB */}
                {rakitTab==="search"&&(()=>{
                  const q = rakitSearch.trim().toUpperCase();
                  const sourceData = firstCureLoaded && firstCureData.length > 0 ? firstCureData : RAKIT_DATA;
                  const results = q.length>=2
                    ? sourceData.filter(r=>r.size?.toUpperCase().includes(q)||r.code?.toUpperCase().includes(q)||r.mc?.toUpperCase().includes(q))
                    : [];
                  const firstCureColor = (fc) => {
                    if (!fc) return { bg:"#f0f0f0", color:"#999" };
                    const f = fc.toUpperCase();
                    if (f==="OK"||f==="OK ( LL )") return { bg:"#E1F5EE", color:"#085041" };
                    if (f.includes("MOR")) return { bg:"#FEF3C7", color:"#92400E" };
                    if (f.includes("OOR")) return { bg:"#FEF3C7", color:"#92400E" };
                    if (f.includes("OFSH")) return { bg:"#FCEBEB", color:"#791F1F" };
                    if (f.includes("DMG")) return { bg:"#FCEBEB", color:"#791F1F" };
                    return { bg:"#F3F4F6", color:"#555" };
                  };
                  return(
                    <div>
                      {/* Search box */}
                      <div style={{ position:"relative",marginBottom:12 }}>
                        <i className="ti ti-search" style={{ position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",color:"#999",fontSize:16 }}></i>
                        <input
                          className="search-input"
                          style={{ paddingLeft:36 }}
                          placeholder="Ketik size mold... (min 2 karakter, cth: BX551)"
                          value={rakitSearch}
                          onChange={e=>setRakitSearch(e.target.value)}
                          autoFocus
                        />
                        {rakitSearch&&<button onClick={()=>setRakitSearch("")}
                          style={{ position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:"#999",fontSize:16 }}>✕</button>}
                      </div>

                      {/* Info */}
                      {q.length<2&&(
                        <div style={{ textAlign:"center",padding:"32px 16px",color:"#999" }}>
                          <div style={{ fontSize:32,marginBottom:8 }}>🔍</div>
                          <div style={{ fontSize:13,fontWeight:600,marginBottom:4 }}>Cari Data First Cure</div>
                          <div style={{ fontSize:11 }}>Ketik size mold untuk melihat history data rakit & kalibrasi</div>
                          <div style={{ marginTop:12,fontSize:11,color:"#bbb" }}>Total data: {sourceData.length} record · {[...new Set(sourceData.map(r=>r.size).filter(Boolean))].length} size {firstCureLoaded&&firstCureData.length>0?"(dari Supabase)":"(dari cache)"}</div>
                        </div>
                      )}

                      {/* Hasil */}
                      {q.length>=2&&results.length===0&&(
                        <div style={{ textAlign:"center",padding:"24px",color:"#999",fontSize:13 }}>
                          Tidak ada data untuk "<strong>{q}</strong>"
                        </div>
                      )}

                      {results.length>0&&(
                        <div>
                          <div style={{ fontSize:12,color:"#999",marginBottom:10 }}>
                            {results.length} record ditemukan untuk "<strong>{q}</strong>"
                          </div>

                          {/* Summary stat */}
                          {(()=>{
                            const ok    = results.filter(r=>r.first_cure?.toUpperCase()==="OK"||r.first_cure?.toUpperCase()==="OK ( LL )").length;
                            const mor   = results.filter(r=>r.first_cure?.toUpperCase().includes("MOR")).length;
                            const oor   = results.filter(r=>r.first_cure?.toUpperCase().includes("OOR")).length;
                            const ofsh  = results.filter(r=>r.first_cure?.toUpperCase().includes("OFSH")).length;
                            const dmgm  = results.filter(r=>r.first_cure?.toUpperCase().includes("DMG")).length;
                            return(
                              <div style={{ display:"flex",gap:6,flexWrap:"wrap",marginBottom:12 }}>
                                {[
                                  ["OK",ok,"#E1F5EE","#085041"],
                                  ["MOR",mor,"#FEF3C7","#92400E"],
                                  ["OOR",oor,"#FEF3C7","#92400E"],
                                  ["OFSH",ofsh,"#FCEBEB","#791F1F"],
                                  ["DMGM",dmgm,"#FCEBEB","#791F1F"],
                                ].filter(([,n])=>n>0).map(([l,n,bg,c])=>(
                                  <div key={l} style={{ padding:"4px 12px",borderRadius:20,background:bg,color:c,fontSize:11,fontWeight:700 }}>
                                    {l}: {n} <span style={{ fontSize:10,opacity:0.75 }}>({results.length>0?Math.round(n/results.length*100):0}%)</span>
                                  </div>
                                ))}
                              </div>
                            );
                          })()}

                          {/* Record cards */}
                          {results.map((r,i)=>{
                            const fc = firstCureColor(r.first_cure);
                            return(
                              <div key={i} className="record-card" style={{ marginBottom:10 }}>
                                <div className="record-card-header">
                                  <div>
                                    <div className="record-size">{r.code}</div>
                                    <div className="record-type">{r.ket} · {r.mc}</div>
                                  </div>
                                  <div style={{ textAlign:"right" }}>
                                    <div style={{ fontSize:11,color:"#999",marginBottom:4 }}>{r.tgl}</div>
                                    <span style={{ fontSize:11,padding:"2px 10px",borderRadius:20,fontWeight:700,background:fc.bg,color:fc.color }}>
                                      {r.first_cure||"-"}
                                    </span>
                                  </div>
                                </div>

                                {/* Detail grid */}
                                <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginTop:8,marginBottom:6 }}>
                                  {[
                                    ["Nilai Preload", r.nilai_preload||"-"],
                                    ["VMC",           r.vmc||"-"],
                                    ["Shim SR",       r.shim_sr||"-"],
                                    ["Shim PL",       r.shim_pl||"-"],
                                    ["Container",     r.container||"-"],
                                    ["PIC",           r.pic||"-"],
                                  ].map(([label,val])=>(
                                    <div key={label} style={{ background:"#f9f9f9",borderRadius:6,padding:"6px 8px" }}>
                                      <div style={{ fontSize:9,color:"#999",marginBottom:2 }}>{label}</div>
                                      <div style={{ fontSize:12,fontWeight:600,color:"#111" }}>{val}</div>
                                    </div>
                                  ))}
                                </div>

                                {r.keterangan&&(
                                  <div style={{ fontSize:11,color:"#666",background:"#FEF3C7",padding:"4px 8px",borderRadius:6 }}>
                                    📝 {r.keterangan}
                                  </div>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  );
                })()}

                {/* UPLOAD TAB */}
                {rakitTab==="upload"&&canAnalyst(role)&&(
                  <div>
                    {/* Info */}
                    <div style={{ background:"#EDE9FE",borderRadius:10,padding:"12px 16px",marginBottom:16,fontSize:13 }}>
                      <div style={{ fontWeight:700,color:"#5B21B6",marginBottom:4 }}>📤 Upload Data First Cure</div>
                      <div style={{ color:"#6B7280",fontSize:12 }}>Pilih file Excel (sheet "History Data"). Data lama akan diganti seluruhnya dengan data baru.</div>
                      {firstCureLoaded&&<div style={{ marginTop:6,fontSize:11,color:"#8B5CF6" }}>Data aktif di Supabase: <strong>{firstCureData.length} record</strong></div>}
                    </div>

                    {/* File picker */}
                    <div style={{ border:"2px dashed #D1D5DB",borderRadius:10,padding:"24px",textAlign:"center",marginBottom:16,background:"#FAFAFA" }}>
                      <div style={{ fontSize:28,marginBottom:8 }}>📂</div>
                      <div style={{ fontSize:13,color:"#666",marginBottom:12 }}>Pilih file Excel dari komputer Anda</div>
                      <input type="file" accept=".xlsx,.xls" style={{ display:"none" }} id="fc-upload-input"
                        onChange={async (e)=>{
                          const file = e.target.files[0];
                          if (!file) return;
                          setUploadFile(file);
                          setUploadMsg("Membaca file...");
                          try {
                            // Load SheetJS dari CDN jika belum ada
                          if (!window.XLSX) {
                            await new Promise((resolve, reject) => {
                              const s = document.createElement("script");
                              s.src = "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";
                              s.onload = resolve; s.onerror = reject;
                              document.head.appendChild(s);
                            });
                          }
                          const XLSX = window.XLSX;
                          const buf = await file.arrayBuffer();
                          const wb = XLSX.read(buf, {type:"array",cellDates:true});
                            // Cari sheet yang relevan
                            const sheetName = wb.SheetNames.find(s=>s.toLowerCase().includes("first cure")||s.toLowerCase().includes("history")) || wb.SheetNames[0];
                            const ws = wb.Sheets[sheetName];
                            const rows = XLSX.utils.sheet_to_json(ws, {defval:""});
                            // Map kolom
                            const mapped = rows.map(r=>{
                              const findVal = (keys) => {
                                for (const k of keys) {
                                  const found = Object.keys(r).find(rk => rk.toLowerCase().replace(/[\s_]/g,"").includes(k));
                                  if (found && r[found]!==undefined) return String(r[found]).trim();
                                }
                                return "";
                              };
                              const code = findVal(["code"]);
                              let size = "";
                              const m = code.match(/^([A-Za-z0-9]+)\s*-\s*\d+/);
                              if (m) size = m[1].toUpperCase();
                              return {
                                tgl: findVal(["tglrakit","tanggal","date","tgl"]),
                                pic: findVal(["pic"]),
                                code,
                                size,
                                container: findVal(["container"]),
                                ket: findVal(["ket","keterangan2","status"]),
                                mc: findVal(["mc"]),
                                nilai_preload: findVal(["nilaipreload","preload"]),
                                vmc: findVal(["vmc"]),
                                shim_sr: findVal(["shimsidering","shimsr","sidering"]),
                                shim_pl: findVal(["shimpreload","shimpl"]),
                                first_cure: findVal(["hasilfirstcure","firstcure","hasil"]),
                                keterangan: findVal(["ket.firstcure","keteranganfirst","ketfirst","keterangan"]),
                              };
                            }).filter(r=>r.code);
                            setUploadPreview(mapped);
                            setUploadMsg("Preview: " + mapped.length + " record siap diupload.");
                          } catch(err) {
                            setUploadMsg("❌ Error membaca file: " + err.message);
                          }
                        }}
                      />
                      <label htmlFor="fc-upload-input" style={{ cursor:"pointer",background:"#1D9E75",color:"#fff",padding:"8px 20px",borderRadius:8,fontWeight:600,fontSize:13 }}>
                        Pilih File Excel
                      </label>
                      {uploadFile&&<div style={{ marginTop:8,fontSize:12,color:"#666" }}>📄 {uploadFile.name}</div>}
                    </div>

                    {/* Status message */}
                    {uploadMsg&&(
                      <div style={{ padding:"10px 14px",borderRadius:8,background: uploadMsg.startsWith("✅")?"#E1F5EE":uploadMsg.startsWith("❌")?"#FCEBEB":"#FEF3C7",
                        color: uploadMsg.startsWith("✅")?"#085041":uploadMsg.startsWith("❌")?"#791F1F":"#92400E",
                        fontSize:13,marginBottom:12,fontWeight:600 }}>
                        {uploadMsg}
                      </div>
                    )}

                    {/* Preview */}
                    {uploadPreview.length>0&&(
                      <div>
                        <div style={{ fontWeight:700,fontSize:13,marginBottom:8,color:"#333" }}>
                          Preview {Math.min(5,uploadPreview.length)} dari {uploadPreview.length} record:
                        </div>
                        <div style={{ overflow:"auto",borderRadius:8,border:"1px solid #e0e0e0",marginBottom:16 }}>
                          <table style={{ width:"100%",borderCollapse:"collapse",fontSize:11 }}>
                            <thead>
                              <tr style={{ background:"#f5f5f5" }}>
                                {["Tgl","PIC","Code","Size","MC","Nilai PL","VMC","Shim SR","Shim PL","First Cure"].map(h=>(
                                  <th key={h} style={{ padding:"6px 8px",textAlign:"left",borderBottom:"1px solid #e0e0e0",whiteSpace:"nowrap" }}>{h}</th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              {uploadPreview.slice(0,5).map((r,i)=>(
                                <tr key={i} style={{ borderBottom:"1px solid #f0f0f0" }}>
                                  {[r.tgl,r.pic,r.code,r.size,r.mc,r.nilai_preload,r.vmc,r.shim_sr,r.shim_pl,r.first_cure].map((v,j)=>(
                                    <td key={j} style={{ padding:"6px 8px",whiteSpace:"nowrap",maxWidth:100,overflow:"hidden",textOverflow:"ellipsis" }}>{v||"-"}</td>
                                  ))}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>

                        {/* Tombol simpan */}
                        <button
                          onClick={()=>{ if(window.confirm("Hapus semua data lama dan upload "+uploadPreview.length+" record baru?")) uploadFirstCureData(uploadPreview); }}
                          disabled={uploadLoading}
                          style={{ width:"100%",padding:"12px",borderRadius:8,background:uploadLoading?"#ccc":"#1D9E75",color:"#fff",border:"none",fontWeight:700,fontSize:14,cursor:uploadLoading?"not-allowed":"pointer" }}>
                          {uploadLoading?"⏳ Sedang upload...":"💾 Simpan ke Database ("+uploadPreview.length+" record)"}
                        </button>
                        <div style={{ fontSize:11,color:"#999",textAlign:"center",marginTop:6 }}>
                          ⚠️ Data lama akan dihapus dan diganti dengan data baru
                        </div>
                      </div>
                    )}
                  </div>
                )}

                {/* FORM TAB */}
                
              </div>
            )}

            {/* NAIK MOLD */}
            {page==="naik"&&canNaik(role)&&(
              <div>
                {/* ANTRIAN SHIFT PLAN */}
                {(()=>{
                  const myGrup = getGrup(currentUser?.username);
                  const today = new Date().toISOString().slice(0,10);
                  const antrian = shiftPlanRecords.filter(p=>p.grup===myGrup&&p.date===today&&(rakitRecords.some(r=>r.mold_size===p.mold_size&&r.date===today)&&!naikRecords.some(r=>r.mold_size_naik===p.mold_size&&r.date===today)));
                  if (!antrian.length) return null;
                  return(
                    <div style={{ background:"#FCEBEB",borderRadius:10,padding:"10px 14px",marginBottom:14,border:"1.5px solid #1D9E75" }}>
                      <div style={{ fontSize:12,fontWeight:700,color:"#085041",marginBottom:8 }}>
                        ⬆️ Antrian Shift Plan hari ini ({antrian.length} size)
                      </div>
                      <div style={{ display:"flex",flexWrap:"wrap",gap:6 }}>
                        {antrian.map(p=>(
                          <span key={p.id}
                            style={{ background:"#fff",border:"1.5px solid #1D9E75",borderRadius:20,padding:"4px 12px",fontSize:12,fontWeight:700,color:"#085041",cursor:"pointer" }}
                            onClick={()=>{ setNaikForm(f=>({...f,moldSizeNaik:p.mold_size}));setNaikTab('form') }}>
                            {p.mold_size}
                          </span>
                        ))}
                      </div>
                      <div style={{ fontSize:10,color:"#666",marginTop:6 }}>💡 Klik size untuk auto-isi form</div>
                    </div>
                  );
                })()}

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
                      <input className="form-input" placeholder="cth: 195/65R15 (opsional)" value={naikForm.moldSizeTurun} onChange={e=>setNaikForm(f=>({...f,moldSizeTurun:e.target.value}))} />
                    </div>
                  </div>

                  {/* PRESS */}
                  <div className="form-group">
                    <label className="form-label">Press <span style={{ fontSize:10,color:"#999",fontWeight:400 }}>(bisa pilih keduanya)</span></label>
                    <div style={{ display:"flex",gap:8 }}>
                      {["L","R"].map(p=>{
                        const selected=(naikForm.press||[]).includes(p);
                        return(
                          <button key={p} className={`type-btn${selected?" active":""}`} style={{ flex:1 }}
                            onClick={()=>{ const cur=naikForm.press||[]; setNaikForm(f=>({...f,press:cur.includes(p)?cur.filter(x=>x!==p):[...cur,p]})); }}>
                            {p==="L"?"◀ Left (L)":"Right (R) ▶"}
                          </button>
                        );
                      })}
                    </div>
                    {(naikForm.press||[]).length>0&&<div style={{ marginTop:6,fontSize:11,color:"#1D9E75",fontWeight:500 }}>Dipilih: {naikForm.press.join(" & ")}</div>}
                  </div>

                  {/* MOLD NAIK — Maker, Type, No Container, No Mold */}
                  <div style={{ fontSize:11,color:"#1D9E75",fontWeight:700,marginBottom:6,marginTop:8,background:"#E1F5EE",padding:"6px 10px",borderRadius:6 }}>⬆️ MOLD NAIK</div>
                  <div className="form-row">
                    <div className="form-group">
                      <label className="form-label">Maker Container L</label>
                      <select className="form-input" value={naikForm.makerContainerL} onChange={e=>setNaikForm(f=>({...f,makerContainerL:e.target.value}))}>
                        <option value="">— Pilih —</option>
                        {MAKER_CONTAINER.map(m=><option key={m} value={m}>{m}</option>)}
                      </select>
                    </div>
                    <div className="form-group">
                      <label className="form-label">Maker Container R</label>
                      <select className="form-input" value={naikForm.makerContainerR} onChange={e=>setNaikForm(f=>({...f,makerContainerR:e.target.value}))}>
                        <option value="">— Pilih —</option>
                        {MAKER_CONTAINER.map(m=><option key={m} value={m}>{m}</option>)}
                      </select>
                    </div>
                  </div>
                  <div className="form-row">
                    <div className="form-group">
                      <label className="form-label">Type Container L</label>
                      <select className="form-input" value={naikForm.typeContainerL} onChange={e=>setNaikForm(f=>({...f,typeContainerL:e.target.value}))}>
                        <option value="">— Pilih —</option>
                        {TYPE_CONTAINER.map(t=><option key={t} value={t}>{t}</option>)}
                      </select>
                    </div>
                    <div className="form-group">
                      <label className="form-label">Type Container R</label>
                      <select className="form-input" value={naikForm.typeContainerR} onChange={e=>setNaikForm(f=>({...f,typeContainerR:e.target.value}))}>
                        <option value="">— Pilih —</option>
                        {TYPE_CONTAINER.map(t=><option key={t} value={t}>{t}</option>)}
                      </select>
                    </div>
                  </div>
                  <div className="form-row">
                    <div className="form-group">
                      <label className="form-label">No Container L</label>
                      <input className="form-input" placeholder="cth: C-001L" value={naikForm.containerNoL} onChange={e=>setNaikForm(f=>({...f,containerNoL:e.target.value}))} />
                    </div>
                    <div className="form-group">
                      <label className="form-label">No Container R</label>
                      <input className="form-input" placeholder="cth: C-001R" value={naikForm.containerNoR} onChange={e=>setNaikForm(f=>({...f,containerNoR:e.target.value}))} />
                    </div>
                  </div>
                  <div className="form-row">
                    <div className="form-group">
                      <label className="form-label">No Mold L</label>
                      <input className="form-input" placeholder="cth: ML-001" value={naikForm.moldNoL} onChange={e=>setNaikForm(f=>({...f,moldNoL:e.target.value}))} />
                    </div>
                    <div className="form-group">
                      <label className="form-label">No Mold R</label>
                      <input className="form-input" placeholder="cth: MR-001" value={naikForm.moldNoR} onChange={e=>setNaikForm(f=>({...f,moldNoR:e.target.value}))} />
                    </div>
                  </div>

                  {/* MOLD TURUN — hanya tampil jika ada size turun */}
                  {naikForm.moldSizeTurun&&(
                    <>
                      <div style={{ fontSize:11,color:"#E24B4A",fontWeight:700,marginBottom:6,marginTop:8,background:"#FCEBEB",padding:"6px 10px",borderRadius:6 }}>⬇️ MOLD TURUN</div>
                      <div className="form-row">
                        <div className="form-group">
                          <label className="form-label">Maker Container L</label>
                          <select className="form-input" value={naikForm.turunMakerContainerL} onChange={e=>setNaikForm(f=>({...f,turunMakerContainerL:e.target.value}))}>
                            <option value="">— Pilih —</option>
                            {MAKER_CONTAINER.map(m=><option key={m} value={m}>{m}</option>)}
                          </select>
                        </div>
                        <div className="form-group">
                          <label className="form-label">Maker Container R</label>
                          <select className="form-input" value={naikForm.turunMakerContainerR} onChange={e=>setNaikForm(f=>({...f,turunMakerContainerR:e.target.value}))}>
                            <option value="">— Pilih —</option>
                            {MAKER_CONTAINER.map(m=><option key={m} value={m}>{m}</option>)}
                          </select>
                        </div>
                      </div>
                      <div className="form-row">
                        <div className="form-group">
                          <label className="form-label">Type Container L</label>
                          <select className="form-input" value={naikForm.turunTypeContainerL} onChange={e=>setNaikForm(f=>({...f,turunTypeContainerL:e.target.value}))}>
                            <option value="">— Pilih —</option>
                            {TYPE_CONTAINER.map(t=><option key={t} value={t}>{t}</option>)}
                          </select>
                        </div>
                        <div className="form-group">
                          <label className="form-label">Type Container R</label>
                          <select className="form-input" value={naikForm.turunTypeContainerR} onChange={e=>setNaikForm(f=>({...f,turunTypeContainerR:e.target.value}))}>
                            <option value="">— Pilih —</option>
                            {TYPE_CONTAINER.map(t=><option key={t} value={t}>{t}</option>)}
                          </select>
                        </div>
                      </div>
                      <div className="form-row">
                        <div className="form-group">
                          <label className="form-label">No Container L</label>
                          <input className="form-input" placeholder="cth: C-001L" value={naikForm.turunContainerNoL} onChange={e=>setNaikForm(f=>({...f,turunContainerNoL:e.target.value}))} />
                        </div>
                        <div className="form-group">
                          <label className="form-label">No Container R</label>
                          <input className="form-input" placeholder="cth: C-001R" value={naikForm.turunContainerNoR} onChange={e=>setNaikForm(f=>({...f,turunContainerNoR:e.target.value}))} />
                        </div>
                      </div>
                      <div className="form-row">
                        <div className="form-group">
                          <label className="form-label">No Mold L</label>
                          <input className="form-input" placeholder="cth: ML-001" value={naikForm.turunMoldNoL} onChange={e=>setNaikForm(f=>({...f,turunMoldNoL:e.target.value}))} />
                        </div>
                        <div className="form-group">
                          <label className="form-label">No Mold R</label>
                          <input className="form-input" placeholder="cth: MR-001" value={naikForm.turunMoldNoR} onChange={e=>setNaikForm(f=>({...f,turunMoldNoR:e.target.value}))} />
                        </div>
                      </div>
                    </>
                  )}

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
                            
                          </div>
                        </div>
                        {r.machine_code&&<div style={{ fontSize:12,fontWeight:600,color:"#1D9E75",marginBottom:6 }}><i className="ti ti-robot" style={{ fontSize:13,marginRight:4 }}></i>{r.machine_code}</div>}
                        {(r.press||[]).length>0&&<div style={{ fontSize:11,color:"#666",marginBottom:4 }}>Press: <strong>{r.press.join(" & ")}</strong></div>}
                        {/* MOLD NAIK info */}
                        {(r.maker_container_l||r.maker_container_r||r.type_container_l||r.type_container_r)&&(
                          <div style={{ fontSize:11,color:"#1D9E75",fontWeight:600,marginBottom:2 }}>⬆️ Naik</div>
                        )}
                        {(r.maker_container_l||r.maker_container_r)&&(
                          <div style={{ fontSize:11,color:"#666",marginBottom:2,display:"flex",gap:12,flexWrap:"wrap" }}>
                            {r.maker_container_l&&<span>Maker L: <strong>{r.maker_container_l}</strong></span>}
                            {r.maker_container_r&&<span>Maker R: <strong>{r.maker_container_r}</strong></span>}
                          </div>
                        )}
                        {(r.type_container_l||r.type_container_r)&&(
                          <div style={{ fontSize:11,color:"#666",marginBottom:2,display:"flex",gap:12,flexWrap:"wrap" }}>
                            {r.type_container_l&&<span>Type L: <strong>{r.type_container_l}</strong></span>}
                            {r.type_container_r&&<span>Type R: <strong>{r.type_container_r}</strong></span>}
                          </div>
                        )}
                        {(r.container_no_l||r.container_no_r)&&(
                          <div style={{ fontSize:11,color:"#666",marginBottom:2,display:"flex",gap:12,flexWrap:"wrap" }}>
                            {r.container_no_l&&<span>No Container L: <strong>{r.container_no_l}</strong></span>}
                            {r.container_no_r&&<span>No Container R: <strong>{r.container_no_r}</strong></span>}
                          </div>
                        )}
                        {(r.mold_no_l||r.mold_no_r)&&(
                          <div style={{ fontSize:11,color:"#666",marginBottom:4,display:"flex",gap:12,flexWrap:"wrap" }}>
                            {r.mold_no_l&&<span>No Mold L: <strong>{r.mold_no_l}</strong></span>}
                            {r.mold_no_r&&<span>No Mold R: <strong>{r.mold_no_r}</strong></span>}
                          </div>
                        )}
                        {/* MOLD TURUN info */}
                        {r.mold_size_turun&&(
                          <>
                            <div style={{ fontSize:11,color:"#E24B4A",fontWeight:600,marginBottom:2,marginTop:4 }}>⬇️ Turun: {r.mold_size_turun}</div>
                            {(r.turun_maker_container_l||r.turun_maker_container_r)&&(
                              <div style={{ fontSize:11,color:"#666",marginBottom:2,display:"flex",gap:12,flexWrap:"wrap" }}>
                                {r.turun_maker_container_l&&<span>Maker L: <strong>{r.turun_maker_container_l}</strong></span>}
                                {r.turun_maker_container_r&&<span>Maker R: <strong>{r.turun_maker_container_r}</strong></span>}
                              </div>
                            )}
                            {(r.turun_type_container_l||r.turun_type_container_r)&&(
                              <div style={{ fontSize:11,color:"#666",marginBottom:2,display:"flex",gap:12,flexWrap:"wrap" }}>
                                {r.turun_type_container_l&&<span>Type L: <strong>{r.turun_type_container_l}</strong></span>}
                                {r.turun_type_container_r&&<span>Type R: <strong>{r.turun_type_container_r}</strong></span>}
                              </div>
                            )}
                            {(r.turun_container_no_l||r.turun_container_no_r)&&(
                              <div style={{ fontSize:11,color:"#666",marginBottom:2,display:"flex",gap:12,flexWrap:"wrap" }}>
                                {r.turun_container_no_l&&<span>No Container L: <strong>{r.turun_container_no_l}</strong></span>}
                                {r.turun_container_no_r&&<span>No Container R: <strong>{r.turun_container_no_r}</strong></span>}
                              </div>
                            )}
                            {(r.turun_mold_no_l||r.turun_mold_no_r)&&(
                              <div style={{ fontSize:11,color:"#666",marginBottom:4,display:"flex",gap:12,flexWrap:"wrap" }}>
                                {r.turun_mold_no_l&&<span>No Mold L: <strong>{r.turun_mold_no_l}</strong></span>}
                                {r.turun_mold_no_r&&<span>No Mold R: <strong>{r.turun_mold_no_r}</strong></span>}
                              </div>
                            )}
                          </>
                        )}
                        {r.notes&&<div style={{ fontSize:11,color:"#666",marginBottom:4 }}>📝 {r.notes}</div>}
                        <div className="record-meta"><i className="ti ti-user" style={{ fontSize:12,marginRight:4 }}></i>{r.operator} {r.grup?`· Grup ${r.grup}`:""}</div>
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
