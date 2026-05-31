import { useState, useMemo, useRef, useEffect, useCallback } from "react";
import { createClient } from "@supabase/supabase-js";

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

const SECTORS = ["A","B","C","D","E","F","G","H"];

const emptyProblemDetail = (pid) => {
  if (pid.startsWith("MOR"))      return { shimAction:"tambah", shimJumlah:"", shimSectors:[] };
  if (pid.startsWith("OVERFLOW")) return { shimAction:"tambah", shimJumlah:"", shimLokasi:"" };
  if (pid === "OS")               return { shimAction:"tambah", shimJumlah:"", shimLokasi:"" };
  if (pid === "OOR")              return { oorType:"segmented", oorSectors:[] };
  return {};
};

const emptyForm = () => ({
  moldSize:"", moldType:"segmented",
  date: new Date().toISOString().slice(0,10),
  jamMulai:"", jamSelesai:"",
  technician:"", problems:[], problemDetails:{}, notes:"",
});

const CSS = `
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:system-ui,-apple-system,sans-serif;background:#f5f5f5;}
  .app{display:flex;flex-direction:column;min-height:100vh;}
  
  /* BOTTOM NAV - mobile */
  .bottom-nav{position:fixed;bottom:0;left:0;right:0;background:#fff;border-top:1px solid #e5e5e5;display:flex;z-index:50;padding-bottom:env(safe-area-inset-bottom);}
  .bottom-nav-item{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:8px 4px;cursor:pointer;font-size:10px;color:#999;gap:3px;border:none;background:none;}
  .bottom-nav-item.active{color:#1D9E75;}
  .bottom-nav-item i{font-size:20px;}
  
  /* TOPBAR */
  .topbar{background:#fff;border-bottom:1px solid #e5e5e5;padding:12px 16px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:40;}
  .topbar-logo{display:flex;align-items:center;gap:8px;}
  .logo-dot{width:28px;height:28px;background:#1D9E75;border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0;}
  .topbar-title{font-size:14px;font-weight:600;color:#111;}
  .topbar-sub{font-size:10px;color:#999;}
  
  /* CONTENT */
  .content{flex:1;overflow-y:auto;padding:16px;padding-bottom:80px;}
  
  /* CARDS */
  .card{background:#fff;border-radius:12px;padding:16px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.06);}
  .card-title{font-size:13px;font-weight:600;color:#111;margin-bottom:12px;}
  
  /* STAT GRID */
  .stat-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px;}
  .stat-card{background:#f8f8f8;border-radius:10px;padding:12px;}
  .stat-label{font-size:10px;color:#999;margin-bottom:4px;}
  .stat-val{font-size:22px;font-weight:700;color:#111;}
  .stat-sub{font-size:10px;color:#999;margin-top:2px;}
  
  /* PROBLEM BADGE */
  .pbadge{font-size:11px;padding:2px 8px;border-radius:4px;display:inline-block;margin-right:3px;margin-bottom:3px;font-weight:500;}
  
  /* FORM */
  .form-label{font-size:11px;color:#666;font-weight:500;margin-bottom:5px;display:block;}
  .form-input{width:100%;padding:10px 12px;border:1px solid #e0e0e0;border-radius:8px;font-size:14px;color:#111;background:#fff;outline:none;}
  .form-input:focus{border-color:#1D9E75;}
  .form-group{margin-bottom:14px;}
  .form-row{display:grid;grid-template-columns:1fr 1fr;gap:10px;}
  
  /* TYPE BUTTONS */
  .type-btn-group{display:grid;grid-template-columns:1fr 1fr;gap:8px;}
  .type-btn{padding:10px 8px;border-radius:8px;border:1.5px solid #e0e0e0;background:#fff;font-size:12px;color:#666;cursor:pointer;text-align:center;font-weight:400;}
  .type-btn.active{border-color:#1D9E75;background:#E1F5EE;color:#085041;font-weight:500;}
  
  /* PROBLEM CARDS */
  .prob-card{border:1.5px solid #e0e0e0;border-radius:10px;margin-bottom:8px;overflow:hidden;background:#f9f9f9;}
  .prob-card.active{border-color:#1D9E75;background:#f0faf5;}
  .prob-card-header{display:flex;align-items:center;gap:10px;padding:12px;cursor:pointer;}
  .prob-check{width:18px;height:18px;border-radius:5px;border:1.5px solid #ccc;background:#fff;display:flex;align-items:center;justify-content:center;flex-shrink:0;}
  .prob-check.active{background:#1D9E75;border-color:#1D9E75;}
  .prob-card-detail{padding:0 12px 14px;border-top:1px solid #c8eedd;}
  
  /* SHIM TOGGLE */
  .shim-toggle-group{display:flex;gap:8px;margin-bottom:10px;}
  .shim-toggle{flex:1;padding:9px;border-radius:8px;border:1.5px solid #e0e0e0;background:#f9f9f9;font-size:12px;color:#666;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:6px;}
  .shim-toggle.active{border-color:#1D9E75;background:#E1F5EE;color:#085041;font-weight:500;}
  .radio-dot{width:12px;height:12px;border-radius:50%;background:#ccc;flex-shrink:0;}
  .radio-dot.active{background:#1D9E75;}
  
  /* SECTOR BUTTONS */
  .sector-grid{display:flex;flex-wrap:wrap;gap:8px;}
  .sector-btn{width:40px;height:40px;border-radius:8px;border:1.5px solid #e0e0e0;background:#f9f9f9;font-size:14px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;color:#444;}
  .sector-btn.active{background:#1D9E75;border-color:#1D9E75;color:#fff;}
  
  /* BUTTONS */
  .btn-primary{background:#1D9E75;color:#fff;border:none;padding:12px 20px;border-radius:10px;font-size:14px;font-weight:600;cursor:pointer;width:100%;margin-bottom:8px;}
  .btn-secondary{background:#fff;color:#444;border:1.5px solid #e0e0e0;padding:11px 20px;border-radius:10px;font-size:14px;cursor:pointer;width:100%;margin-bottom:8px;}
  .btn-danger{background:#fff;color:#E24B4A;border:1.5px solid #E24B4A;padding:11px 20px;border-radius:10px;font-size:14px;cursor:pointer;width:100%;}
  .btn-sm{padding:6px 12px;border-radius:6px;border:1px solid #e0e0e0;background:#fff;font-size:12px;cursor:pointer;color:#444;}
  .btn-sm-danger{padding:6px 12px;border-radius:6px;border:1px solid #E24B4A;background:#fff;font-size:12px;cursor:pointer;color:#E24B4A;}
  
  /* TABLE - mobile: card-based */
  .record-card{background:#fff;border-radius:10px;padding:14px;margin-bottom:10px;box-shadow:0 1px 3px rgba(0,0,0,0.06);}
  .record-card-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px;}
  .record-size{font-size:16px;font-weight:700;color:#111;}
  .record-type{font-size:10px;color:#999;margin-top:1px;}
  .record-meta{font-size:11px;color:#999;margin-top:6px;}
  .record-actions{display:flex;gap:6px;margin-top:10px;padding-top:10px;border-top:1px solid #f0f0f0;}
  
  /* DETAIL */
  .detail-section{background:#f8f8f8;border-radius:8px;padding:12px;margin-bottom:10px;}
  .detail-label{font-size:10px;color:#999;margin-bottom:4px;}
  .detail-val{font-size:13px;font-weight:500;color:#111;}
  
  /* TOAST */
  .toast{position:fixed;top:16px;left:50%;transform:translateX(-50%);z-index:999;padding:10px 20px;border-radius:10px;font-size:13px;font-weight:500;white-space:nowrap;box-shadow:0 4px 12px rgba(0,0,0,0.15);}
  .toast.success{background:#E1F5EE;color:#085041;border:1px solid #9FE1CB;}
  .toast.error{background:#FCEBEB;color:#791F1F;border:1px solid #F09595;}
  
  /* BAR CHART */
  .bar-row{margin-bottom:10px;}
  .bar-row-header{display:flex;justify-content:space-between;font-size:11px;color:#666;margin-bottom:4px;}
  .bar-track{height:6px;background:#f0f0f0;border-radius:3px;}
  .bar-fill{height:6px;background:#1D9E75;border-radius:3px;}
  
  /* SEARCH */
  .search-wrap{position:relative;margin-bottom:10px;}
  .search-icon{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:#999;font-size:16px;}
  .search-input{width:100%;padding:10px 12px 10px 36px;border:1px solid #e0e0e0;border-radius:8px;font-size:14px;background:#fff;outline:none;}
  .filter-row{display:flex;gap:8px;margin-bottom:12px;overflow-x:auto;padding-bottom:2px;}
  .filter-select{padding:8px 10px;border:1px solid #e0e0e0;border-radius:8px;font-size:12px;background:#fff;white-space:nowrap;flex-shrink:0;}
  
  /* AUTOCOMPLETE */
  .autocomplete-wrap{position:relative;}
  .autocomplete-dropdown{position:absolute;top:100%;left:0;right:0;background:#fff;border:1px solid #e0e0e0;border-radius:8px;z-index:100;max-height:160px;overflow-y:auto;margin-top:2px;box-shadow:0 4px 12px rgba(0,0,0,0.1);}
  .autocomplete-item{padding:10px 12px;font-size:13px;cursor:pointer;color:#111;}
  .autocomplete-item:hover{background:#f5f5f5;}
  
  /* STATUS indicator */
  .status-dot{width:8px;height:8px;border-radius:50%;background:#1D9E75;display:inline-block;margin-right:5px;}
  
  /* GROUP LABEL */
  .group-label{font-size:10px;color:#999;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;margin-bottom:6px;margin-top:4px;}

  /* DESKTOP OVERRIDE */
  @media (min-width:768px) {
    .app{flex-direction:row;}
    .bottom-nav{display:none;}
    .sidebar{width:200px;height:100vh;position:sticky;top:0;background:#fff;border-right:1px solid #e5e5e5;display:flex;flex-direction:column;flex-shrink:0;}
    .sidebar-logo{padding:18px 16px;border-bottom:1px solid #e5e5e5;}
    .sidebar-nav{padding:10px 0;flex:1;}
    .sidebar-nav-item{display:flex;align-items:center;gap:10px;padding:10px 16px;cursor:pointer;font-size:13px;color:#666;border-left:2px solid transparent;}
    .sidebar-nav-item.active{color:#1D9E75;background:#f0faf5;border-left-color:#1D9E75;}
    .sidebar-nav-item i{font-size:17px;flex-shrink:0;}
    .sidebar-footer{padding:12px 16px;border-top:1px solid #e5e5e5;font-size:10px;color:#999;}
    .main{flex:1;min-width:0;display:flex;flex-direction:column;}
    .content{padding:20px;padding-bottom:20px;}
    .stat-grid{grid-template-columns:repeat(4,1fr);}
    .form-row-desktop{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
    .dashboard-charts{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:12px;}
    .topbar{padding:12px 20px;}
  }
`;

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
          {filtered.map(s=>(
            <div key={s} className="autocomplete-item" onMouseDown={()=>{onChange(s);setOpen(false);}}>{s}</div>
          ))}
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

function ProblemDetailForm({ pid, detail, onChange }) {
  const isMOR=pid.startsWith("MOR"), isOverflow=pid.startsWith("OVERFLOW"), isOS=pid==="OS", isOOR=pid==="OOR";
  const toggleArr = (arr, key, s) => {
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
        <input type="number" min="1" className="form-input" placeholder="cth: 2" value={detail.shimJumlah||""} onChange={e=>onChange({...detail,shimJumlah:e.target.value})} style={{ width:140 }} />
      </div>
      <div className="form-group">
        <label className="form-label">Pilih sektor</label>
        <div className="sector-grid">
          {SECTORS.map(s=><div key={s} className={`sector-btn${(detail.shimSectors||[]).includes(s)?" active":""}`} onClick={()=>toggleArr(detail,"shimSectors",s)}>{s}</div>)}
        </div>
        {(detail.shimSectors||[]).length>0 && <div style={{ marginTop:8, fontSize:11, color:"#666" }}>Dipilih: <strong>{detail.shimSectors.join(", ")}</strong></div>}
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
            {SECTORS.map(s=><div key={s} className={`sector-btn${(detail.oorSectors||[]).includes(s)?" active":""}`} onClick={()=>toggleArr(detail,"oorSectors",s)}>{s}</div>)}
          </div>
          {(detail.oorSectors||[]).length>0&&<div style={{ marginTop:8, fontSize:11, color:"#666" }}>Dipilih: <strong>{detail.oorSectors.join(", ")}</strong></div>}
        </div>
      )}
      {detail.oorType==="two_piece"&&(
        <div className="form-group">
          <label className="form-label">Pilih bagian yang diperbaiki</label>
          <div className="shim-toggle-group">
            {["Mold Atas","Mold Bawah"].map(s=>(
              <div key={s} className={`shim-toggle${(detail.oorSectors||[]).includes(s)?" active":""}`} onClick={()=>toggleArr(detail,"oorSectors",s)}>
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
              <span style={{ fontSize:11, color:"#666", display:"block", marginTop:2, paddingLeft:4 }}>
                Shim {det.shimAction} {det.shimJumlah} lbr
                {isMOR&&det.shimSectors?.length>0?` — Sektor: ${det.shimSectors.join(", ")}`:det.shimLokasi?` (${det.shimLokasi})`:""}
              </span>
            )}
            {isOOR&&det.oorSectors?.length>0&&(
              <span style={{ fontSize:11, color:"#666", display:"block", marginTop:2, paddingLeft:4 }}>Sektor: {det.oorSectors.join(", ")}</span>
            )}
          </div>
        );
      })}
    </div>
  );
};

export default function App() {
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
  const isMobile = typeof window !== "undefined" && window.innerWidth < 768;

  const showToast = (msg, type="success") => { setToast({msg,type}); setTimeout(()=>setToast(null),2800); };

  const loadRecords = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase.from("repair_records").select("*").order("created_at",{ascending:false});
    if (error) showToast("Gagal memuat data: "+error.message,"error");
    else setRecords(data||[]);
    setLoading(false);
  },[]);

  useEffect(()=>{loadRecords();},[loadRecords]);

  useEffect(()=>{
    const ch = supabase.channel("rr").on("postgres_changes",{event:"*",schema:"public",table:"repair_records"},()=>loadRecords()).subscribe();
    return ()=>supabase.removeChannel(ch);
  },[loadRecords]);

  const knownSizes = useMemo(()=>[...new Set(records.map(r=>r.mold_size))].sort(),[records]);
  const knownTechs = useMemo(()=>[...new Set(records.map(r=>r.technician))].sort(),[records]);

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
    const payload={mold_size:form.moldSize.trim().toUpperCase(),mold_type:form.moldType,date:form.date,jam_mulai:form.jamMulai,jam_selesai:form.jamSelesai,technician:form.technician.trim(),problems:form.problems,problem_details:form.problemDetails,notes:form.notes};
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
    setForm({moldSize:rec.mold_size,moldType:rec.mold_type,date:rec.date,jamMulai:rec.jam_mulai||"",jamSelesai:rec.jam_selesai||"",technician:rec.technician,problems:rec.problems||[],problemDetails:rec.problem_details||{},notes:rec.notes||""});
    setEditId(rec.id);setPage("entry");
  };

  const deleteRecord = async (id) => {
    if(!window.confirm("Hapus record ini?")) return;
    const{error}=await supabase.from("repair_records").delete().eq("id",id);
    if(error){showToast("Gagal hapus: "+error.message,"error");return;}
    showToast("Record dihapus.");
    if(detailId===id){setDetailId(null);setPage("database");}
  };

  const filtered = useMemo(()=>records.filter(r=>{
    if(search&&!r.mold_size?.toLowerCase().includes(search.toLowerCase())&&!r.technician?.toLowerCase().includes(search.toLowerCase())) return false;
    if(filterProblem&&!(r.problems||[]).includes(filterProblem)) return false;
    if(filterTech&&r.technician!==filterTech) return false;
    return true;
  }),[records,search,filterProblem,filterTech]);

  const today=new Date().toISOString().slice(0,10);
  const todayRecs=records.filter(r=>r.date===today);
  const problemCounts=PROBLEMS.map(p=>({...p,count:records.filter(r=>(r.problems||[]).includes(p.id)).length})).sort((a,b)=>b.count-a.count);
  const detailRec=records.find(r=>r.id===detailId);
  const sameSize=detailRec?records.filter(r=>r.mold_size===detailRec.mold_size&&r.id!==detailRec.id).slice(0,5):[];

  const navTo = (p) => { setPage(p); if(p==="entry"){setForm(emptyForm());setEditId(null);} };

  if(loading) return (
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
        {/* SIDEBAR desktop only */}
        <div className="sidebar" style={{ display:"none" }}>
          <div className="sidebar-logo">
            <div style={{ display:"flex",alignItems:"center",gap:8 }}>
              <div className="logo-dot"><svg width="14" height="14" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="6" stroke="white" strokeWidth="1.5"/><circle cx="8" cy="8" r="2" fill="white"/></svg></div>
              <div><div style={{ fontSize:13,fontWeight:600,color:"#111" }}>MoldTrack</div><div style={{ fontSize:10,color:"#999" }}>Mold Tire Action</div></div>
            </div>
          </div>
          <div className="sidebar-nav">
            {[["dashboard","ti-layout-dashboard","Dashboard"],["entry","ti-plus","Entry Perbaikan"],["database","ti-database","Database Record"]].map(([id,icon,lbl])=>(
              <div key={id} className={`sidebar-nav-item${page===id?" active":""}`} onClick={()=>navTo(id)}>
                <i className={`ti ${icon}`}></i><span>{lbl}</span>
              </div>
            ))}
          </div>
          <div className="sidebar-footer"><span className="status-dot"></span>Terhubung ke Supabase</div>
        </div>

        {/* MAIN */}
        <div className="main">
          {/* TOPBAR */}
          <div className="topbar">
            <div className="topbar-logo">
              <div className="logo-dot"><svg width="14" height="14" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="6" stroke="white" strokeWidth="1.5"/><circle cx="8" cy="8" r="2" fill="white"/></svg></div>
              <div style={{ marginLeft:8 }}>
                <div className="topbar-title">
                  {page==="dashboard"&&"Dashboard"}
                  {page==="entry"&&(editId?"Edit Record":"Entry Perbaikan")}
                  {page==="database"&&"Database Record"}
                  {page==="detail"&&"Detail Record"}
                </div>
                <div className="topbar-sub">{today}</div>
              </div>
            </div>
            <div style={{ display:"flex",alignItems:"center",gap:8 }}>
              <span className="status-dot"></span>
              <button className="btn-sm" onClick={loadRecords}><i className="ti ti-refresh"></i></button>
            </div>
          </div>

          {/* CONTENT */}
          <div className="content">

            {/* ── DASHBOARD ── */}
            {page==="dashboard"&&(
              <div>
                <div className="stat-grid">
                  {[
                    {label:"Total record",val:records.length,sub:`${todayRecs.length} hari ini`},
                    {label:"Teknisi aktif",val:knownTechs.length,sub:"total teknisi"},
                    {label:"Size tercatat",val:knownSizes.length,sub:"jenis mold"},
                    {label:"Bulan ini",val:records.filter(r=>r.date?.startsWith(new Date().toISOString().slice(0,7))).length,sub:new Date().toLocaleString("id-ID",{month:"short",year:"numeric"})},
                  ].map((c,i)=>(
                    <div key={i} className="stat-card">
                      <div className="stat-label">{c.label}</div>
                      <div className="stat-val">{c.val}</div>
                      <div className="stat-sub">{c.sub}</div>
                    </div>
                  ))}
                </div>

                <div className="dashboard-charts">
                  <div className="card">
                    <div className="card-title">Distribusi problem</div>
                    {problemCounts.filter(p=>p.count>0).length===0&&<div style={{ fontSize:12,color:"#999" }}>Belum ada data.</div>}
                    {problemCounts.filter(p=>p.count>0).map(p=>(
                      <div key={p.id} className="bar-row">
                        <div className="bar-row-header"><span>{p.label}</span><span style={{ fontWeight:600,color:"#111" }}>{p.count}</span></div>
                        <div className="bar-track"><div className="bar-fill" style={{ width:`${Math.round((p.count/records.length)*100)}%` }}></div></div>
                      </div>
                    ))}
                  </div>
                  <div className="card">
                    <div className="card-title">Aktivitas teknisi</div>
                    {knownTechs.length===0&&<div style={{ fontSize:12,color:"#999" }}>Belum ada data.</div>}
                    {knownTechs.map(t=>{
                      const cnt=records.filter(r=>r.technician===t).length;
                      return(
                        <div key={t} className="bar-row">
                          <div className="bar-row-header"><span>{t}</span><span style={{ fontWeight:600,color:"#111" }}>{cnt}</span></div>
                          <div className="bar-track"><div className="bar-fill" style={{ width:records.length?`${Math.round((cnt/records.length)*100)}%`:"0%" }}></div></div>
                        </div>
                      );
                    })}
                  </div>
                </div>

                <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10 }}>
                  <div style={{ fontSize:13,fontWeight:600,color:"#111" }}>Record terbaru</div>
                  <button className="btn-primary" style={{ width:"auto",padding:"8px 16px",margin:0 }} onClick={()=>navTo("entry")}>
                    <i className="ti ti-plus" style={{ marginRight:4 }}></i>Entry baru
                  </button>
                </div>
                {records.length===0&&<div className="card" style={{ textAlign:"center",color:"#999",fontSize:13,padding:24 }}>Belum ada record. Mulai entry perbaikan pertama!</div>}
                {records.slice(0,5).map(r=>(
                  <div key={r.id} className="record-card" onClick={()=>{setDetailId(r.id);setPage("detail");}}>
                    <div className="record-card-header">
                      <div><div className="record-size">{r.mold_size}</div><div className="record-type">{r.mold_type==="segmented"?"Segmented":"Two Piece"}</div></div>
                      <div style={{ fontSize:11,color:"#999",textAlign:"right" }}><div>{r.date}</div><div>{r.jam_mulai&&r.jam_selesai?`${r.jam_mulai}–${r.jam_selesai}`:""}</div></div>
                    </div>
                    <DetailSummary rec={{...r,problemDetails:r.problem_details,problems:r.problems||[]}}/>
                    <div className="record-meta"><i className="ti ti-user" style={{ fontSize:12,marginRight:4 }}></i>{r.technician}</div>
                  </div>
                ))}
              </div>
            )}

            {/* ── ENTRY FORM ── */}
            {page==="entry"&&(
              <div>
                <div className="card">
                  <div className="card-title">Identitas perbaikan</div>
                  <div className="form-row-desktop" style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
                    <AutocompleteInput label="Size mold *" value={form.moldSize} onChange={v=>setForm(f=>({...f,moldSize:v}))} suggestions={knownSizes} placeholder="cth: 205/65R15"/>
                    <div className="form-group">
                      <label className="form-label">Tanggal *</label>
                      <input type="date" className="form-input" value={form.date} onChange={e=>setForm(f=>({...f,date:e.target.value}))}/>
                    </div>
                  </div>
                  <div className="form-group">
                    <label className="form-label">Tipe mold *</label>
                    <div className="type-btn-group">
                      {[["segmented","Segmented (A–H)"],["two_piece","Two Piece"]].map(([val,lbl])=>(
                        <button key={val} className={`type-btn${form.moldType===val?" active":""}`} onClick={()=>setForm(f=>({...f,moldType:val}))}>
                          {lbl}
                        </button>
                      ))}
                    </div>
                  </div>
                  <AutocompleteInput label="Nama teknisi *" value={form.technician} onChange={v=>setForm(f=>({...f,technician:v}))} suggestions={knownTechs} placeholder="Ketik nama kamu..."/>
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
                <button className="btn-secondary" onClick={()=>setPage("database")}>Batal</button>
              </div>
            )}

            {/* ── DATABASE ── */}
            {page==="database"&&(
              <div>
                <div className="search-wrap">
                  <i className="ti ti-search search-icon"></i>
                  <input className="search-input" placeholder="Cari size mold atau teknisi..." value={search} onChange={e=>setSearch(e.target.value)}/>
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
                    <i className="ti ti-plus"></i>
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
                    <div style={{ marginBottom:6 }}><DetailSummary rec={{...r,problemDetails:r.problem_details,problems:r.problems||[]}}/></div>
                    <div className="record-meta" style={{ marginBottom:8 }}><i className="ti ti-user" style={{ fontSize:12,marginRight:4 }}></i>{r.technician}</div>
                    <div className="record-actions">
                      <button className="btn-sm" onClick={()=>{setDetailId(r.id);setPage("detail");}}>Detail</button>
                      <button className="btn-sm" onClick={()=>startEdit(r)}>Edit</button>
                      <button className="btn-sm-danger" onClick={()=>deleteRecord(r.id)}>Hapus</button>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {/* ── DETAIL ── */}
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
                  <div style={{ marginTop:16,display:"flex",gap:8 }}>
                    <button className="btn-primary" style={{ margin:0 }} onClick={()=>startEdit(detailRec)}>Edit</button>
                    <button className="btn-danger" onClick={()=>deleteRecord(detailRec.id)}>Hapus</button>
                  </div>
                </div>

                {sameSize.length>0&&(
                  <div className="card">
                    <div className="card-title"><i className="ti ti-history" style={{ marginRight:6,color:"#999" }}></i>Riwayat size {detailRec.mold_size}</div>
                    {sameSize.map(r=>(
                      <div key={r.id} style={{ paddingBottom:10,marginBottom:10,borderBottom:"1px solid #f0f0f0",cursor:"pointer" }} onClick={()=>setDetailId(r.id)}>
                        <div style={{ fontSize:11,color:"#999",marginBottom:6 }}>{r.date}{r.jam_mulai?` · ${r.jam_mulai}–${r.jam_selesai}`:""}</div>
                        <DetailSummary rec={{...r,problemDetails:r.problem_details,problems:r.problems||[]}}/>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}
          </div>
        </div>

        {/* BOTTOM NAV mobile */}
        <nav className="bottom-nav">
          {[["dashboard","ti-layout-dashboard","Dashboard"],["entry","ti-plus","Entry"],["database","ti-database","Database"]].map(([id,icon,lbl])=>(
            <button key={id} className={`bottom-nav-item${page===id?" active":""}`} onClick={()=>navTo(id)}>
              <i className={`ti ${icon}`}></i><span>{lbl}</span>
            </button>
          ))}
        </nav>
      </div>
    </>
  );
}
