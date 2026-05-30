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
  technician:"",
  problems:[],
  problemDetails:{},
  notes:"",
});

// ── helpers ──────────────────────────────────────────────────────────────────
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
    <div ref={ref} style={{ position:"relative" }}>
      {label && <label style={{ fontSize:11, color:"var(--color-text-secondary)", marginBottom:6, display:"block", fontWeight:500 }}>{label}</label>}
      <input value={value} onChange={e=>{onChange(e.target.value);setOpen(true);}} onFocus={()=>setOpen(true)}
        placeholder={placeholder} style={{ width:"100%", marginBottom:0 }} autoComplete="off" />
      {open && filtered.length > 0 && (
        <div style={{ position:"absolute", top:"100%", left:0, right:0, background:"var(--color-background-primary)", border:"0.5px solid var(--color-border-secondary)", borderRadius:"var(--border-radius-md)", zIndex:100, maxHeight:160, overflowY:"auto", marginTop:2 }}>
          {filtered.map(s => (
            <div key={s} onMouseDown={()=>{onChange(s);setOpen(false);}}
              style={{ padding:"8px 12px", fontSize:13, cursor:"pointer", color:"var(--color-text-primary)" }}
              onMouseEnter={e=>e.currentTarget.style.background="var(--color-background-secondary)"}
              onMouseLeave={e=>e.currentTarget.style.background="transparent"}>{s}</div>
          ))}
        </div>
      )}
    </div>
  );
}

const ProblemBadge = ({ pid }) => {
  const p = PROBLEMS.find(x=>x.id===pid);
  const colors = { MOR:["#EAF3DE","#27500A"], OOR:["#E6F1FB","#0C447C"], Overflow:["#FAEEDA","#633806"], OS:["#FAECE7","#712B13"] };
  const [bg,fg] = colors[p?.group] || ["#F1EFE8","#444441"];
  return <span style={{ background:bg, color:fg, fontSize:11, padding:"2px 8px", borderRadius:4, marginRight:4, display:"inline-block", marginBottom:2 }}>{p?.label}</span>;
};

function ProblemDetailForm({ pid, detail, onChange }) {
  const isMOR      = pid.startsWith("MOR");
  const isOverflow = pid.startsWith("OVERFLOW");
  const isOS       = pid === "OS";
  const isOOR      = pid === "OOR";

  const S = {
    label:  { fontSize:11, color:"var(--color-text-secondary)", marginBottom:4, display:"block", fontWeight:500 },
    inp:    { width:"100%", marginBottom:0 },
    toggle: (a) => ({ display:"flex", alignItems:"center", gap:8, padding:"7px 14px", borderRadius:"var(--border-radius-md)", border:a?"1.5px solid #1D9E75":"0.5px solid var(--color-border-tertiary)", background:a?"#E1F5EE":"var(--color-background-secondary)", color:a?"#085041":"var(--color-text-secondary)", fontSize:12, cursor:"pointer", userSelect:"none" }),
    sector: (a) => ({ width:36, height:36, borderRadius:"var(--border-radius-md)", border:a?"1.5px solid #1D9E75":"0.5px solid var(--color-border-tertiary)", background:a?"#1D9E75":"var(--color-background-secondary)", color:a?"#fff":"var(--color-text-primary)", fontSize:13, fontWeight:500, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center" }),
  };

  const toggleArr = (arr, key, s) => {
    const cur = detail[key] || [];
    onChange({ ...detail, [key]: cur.includes(s) ? cur.filter(x=>x!==s) : [...cur, s] });
  };

  if (isMOR) return (
    <div style={{ marginTop:10 }}>
      <div style={{ display:"flex", gap:8, marginBottom:12 }}>
        {["tambah","kurangi"].map(a => (
          <div key={a} style={S.toggle(detail.shimAction===a)} onClick={()=>onChange({...detail,shimAction:a})}>
            <div style={{ width:12, height:12, borderRadius:"50%", background:detail.shimAction===a?"#1D9E75":"var(--color-border-secondary)", flexShrink:0 }}></div>
            {a==="tambah"?"Tambah shim":"Kurangi shim"}
          </div>
        ))}
      </div>
      <div style={{ marginBottom:10 }}>
        <label style={S.label}>Jumlah shim (lembar)</label>
        <input type="number" min="1" style={{ ...S.inp, width:160 }} placeholder="cth: 2"
          value={detail.shimJumlah||""} onChange={e=>onChange({...detail,shimJumlah:e.target.value})} />
      </div>
      <div>
        <label style={S.label}>Pilih sektor</label>
        <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
          {SECTORS.map(s=><div key={s} style={S.sector((detail.shimSectors||[]).includes(s))} onClick={()=>toggleArr(detail,"shimSectors",s)}>{s}</div>)}
        </div>
        {(detail.shimSectors||[]).length>0 && <div style={{ marginTop:8, fontSize:11, color:"var(--color-text-secondary)" }}>Dipilih: <span style={{ fontWeight:500, color:"var(--color-text-primary)" }}>{detail.shimSectors.join(", ")}</span></div>}
      </div>
    </div>
  );

  if (isOverflow || isOS) {
    const opts = isOverflow ? ["Shoulder Atas","Shoulder Bawah","Atas & Bawah"] : ["Atas","Bawah","Atas & Bawah"];
    return (
      <div style={{ marginTop:10 }}>
        <div style={{ display:"flex", gap:8, marginBottom:10 }}>
          {["tambah","kurangi"].map(a=>(
            <div key={a} style={S.toggle(detail.shimAction===a)} onClick={()=>onChange({...detail,shimAction:a})}>
              <div style={{ width:12, height:12, borderRadius:"50%", background:detail.shimAction===a?"#1D9E75":"var(--color-border-secondary)", flexShrink:0 }}></div>
              {a==="tambah"?"Tambah shim":"Kurangi shim"}
            </div>
          ))}
        </div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
          <div>
            <label style={S.label}>Jumlah shim (lembar)</label>
            <input type="number" min="1" style={S.inp} placeholder="cth: 2"
              value={detail.shimJumlah||""} onChange={e=>onChange({...detail,shimJumlah:e.target.value})} />
          </div>
          <div>
            <label style={S.label}>Lokasi shim</label>
            <select style={S.inp} value={detail.shimLokasi||""} onChange={e=>onChange({...detail,shimLokasi:e.target.value})}>
              <option value="">— Pilih lokasi —</option>
              {opts.map(o=><option key={o}>{o}</option>)}
            </select>
          </div>
        </div>
      </div>
    );
  }

  if (isOOR) return (
    <div style={{ marginTop:10 }}>
      <div style={{ display:"flex", gap:8, marginBottom:10 }}>
        {[["segmented","ti-grid-dots","Segmented (A–H)"],["two_piece","ti-layout-rows","Two Piece"]].map(([val,icon,lbl])=>(
          <div key={val} style={S.toggle(detail.oorType===val)} onClick={()=>onChange({...detail,oorType:val,oorSectors:[]})}>
            <i className={`ti ${icon}`} style={{ fontSize:14 }}></i>{lbl}
          </div>
        ))}
      </div>
      {detail.oorType==="segmented" && (
        <>
          <div style={{ fontSize:11, color:"var(--color-text-secondary)", marginBottom:8 }}>Pilih sektor yang diperbaiki:</div>
          <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
            {SECTORS.map(s=><div key={s} style={S.sector((detail.oorSectors||[]).includes(s))} onClick={()=>toggleArr(detail,"oorSectors",s)}>{s}</div>)}
          </div>
          {(detail.oorSectors||[]).length>0 && <div style={{ marginTop:8, fontSize:11, color:"var(--color-text-secondary)" }}>Dipilih: <span style={{ fontWeight:500, color:"var(--color-text-primary)" }}>{detail.oorSectors.join(", ")}</span></div>}
        </>
      )}
      {detail.oorType==="two_piece" && (
        <>
          <div style={{ fontSize:11, color:"var(--color-text-secondary)", marginBottom:8 }}>Pilih bagian yang diperbaiki:</div>
          <div style={{ display:"flex", gap:8 }}>
            {["Mold Atas","Mold Bawah"].map(s=>(
              <div key={s} style={S.toggle((detail.oorSectors||[]).includes(s))} onClick={()=>toggleArr(detail,"oorSectors",s)}>
                <div style={{ width:14, height:14, borderRadius:4, border:(detail.oorSectors||[]).includes(s)?"none":"1px solid var(--color-border-secondary)", background:(detail.oorSectors||[]).includes(s)?"#1D9E75":"transparent", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
                  {(detail.oorSectors||[]).includes(s)&&<i className="ti ti-check" style={{ fontSize:10, color:"#fff" }}></i>}
                </div>
                {s}
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  );

  return null;
}

const DetailSummary = ({ rec }) => {
  const d = rec.problemDetails || {};
  return (
    <>
      {rec.problems.map(pid=>{
        const det = d[pid];
        if (!det) return null;
        const isMOR=pid.startsWith("MOR"), isOverflow=pid.startsWith("OVERFLOW"), isOS=pid==="OS", isOOR=pid==="OOR";
        return (
          <div key={pid} style={{ marginBottom:4 }}>
            <ProblemBadge pid={pid} />
            {(isMOR||isOverflow||isOS)&&det.shimJumlah&&(
              <span style={{ fontSize:11, color:"var(--color-text-secondary)", marginLeft:2 }}>
                Shim {det.shimAction} {det.shimJumlah} lbr
                {isMOR&&det.shimSectors?.length>0?` — Sektor: ${det.shimSectors.join(", ")}`:det.shimLokasi?` (${det.shimLokasi})`:""}
              </span>
            )}
            {isOOR&&det.oorSectors?.length>0&&(
              <span style={{ fontSize:11, color:"var(--color-text-secondary)", marginLeft:2 }}>Sektor: {det.oorSectors.join(", ")}</span>
            )}
          </div>
        );
      })}
    </>
  );
};

// ── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [page, setPage]         = useState("dashboard");
  const [records, setRecords]   = useState([]);
  const [loading, setLoading]   = useState(true);
  const [form, setForm]         = useState(emptyForm());
  const [editId, setEditId]     = useState(null);
  const [detailId, setDetailId] = useState(null);
  const [search, setSearch]     = useState("");
  const [filterProblem, setFilterProblem] = useState("");
  const [filterTech, setFilterTech]       = useState("");
  const [sidebarOpen, setSidebarOpen]     = useState(true);
  const [toast, setToast]       = useState(null);

  const showToast = (msg, type="success") => { setToast({msg,type}); setTimeout(()=>setToast(null),2800); };

  // ── Supabase: load records ─────────────────────────────────────────────────
  const loadRecords = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase
      .from("repair_records")
      .select("*")
      .order("created_at", { ascending: false });
    if (error) { showToast("Gagal memuat data: " + error.message, "error"); }
    else { setRecords(data || []); }
    setLoading(false);
  }, []);

  useEffect(() => { loadRecords(); }, [loadRecords]);

  // ── Supabase: realtime subscription ───────────────────────────────────────
  useEffect(() => {
    const channel = supabase
      .channel("repair_records_changes")
      .on("postgres_changes", { event:"*", schema:"public", table:"repair_records" }, () => loadRecords())
      .subscribe();
    return () => supabase.removeChannel(channel);
  }, [loadRecords]);

  const knownSizes = useMemo(() => [...new Set(records.map(r=>r.mold_size))].sort(), [records]);
  const knownTechs = useMemo(() => [...new Set(records.map(r=>r.technician))].sort(), [records]);

  const toggleProblem = (pid) => {
    setForm(f => {
      const active = f.problems.includes(pid);
      const problems = active ? f.problems.filter(x=>x!==pid) : [...f.problems, pid];
      const problemDetails = { ...f.problemDetails };
      if (active) delete problemDetails[pid];
      else problemDetails[pid] = emptyProblemDetail(pid);
      return { ...f, problems, problemDetails };
    });
  };

  const updateDetail = (pid, detail) => setForm(f=>({...f, problemDetails:{...f.problemDetails,[pid]:detail}}));

  // ── Supabase: save record ─────────────────────────────────────────────────
  const submitForm = async () => {
    if (!form.moldSize.trim()||!form.technician.trim()||form.problems.length===0) {
      showToast("Lengkapi: size mold, nama teknisi, dan minimal 1 problem.","error"); return;
    }
    const payload = {
      mold_size:       form.moldSize.trim().toUpperCase(),
      mold_type:       form.moldType,
      date:            form.date,
      jam_mulai:       form.jamMulai,
      jam_selesai:     form.jamSelesai,
      technician:      form.technician.trim(),
      problems:        form.problems,
      problem_details: form.problemDetails,
      notes:           form.notes,
    };
    if (editId) {
      const { error } = await supabase.from("repair_records").update(payload).eq("id", editId);
      if (error) { showToast("Gagal update: "+error.message,"error"); return; }
      showToast("Record berhasil diupdate.");
    } else {
      const { error } = await supabase.from("repair_records").insert(payload);
      if (error) { showToast("Gagal simpan: "+error.message,"error"); return; }
      showToast("Record berhasil disimpan.");
    }
    setForm(emptyForm()); setEditId(null); setPage("database");
  };

  // ── Supabase: delete record ───────────────────────────────────────────────
  const deleteRecord = async (id) => {
    const { error } = await supabase.from("repair_records").delete().eq("id", id);
    if (error) { showToast("Gagal hapus: "+error.message,"error"); return; }
    showToast("Record dihapus.");
    if (detailId===id) { setDetailId(null); setPage("database"); }
  };

  const startEdit = (rec) => {
    setForm({
      moldSize: rec.mold_size, moldType: rec.mold_type,
      date: rec.date, jamMulai: rec.jam_mulai||"", jamSelesai: rec.jam_selesai||"",
      technician: rec.technician,
      problems: rec.problems||[],
      problemDetails: rec.problem_details||{},
      notes: rec.notes||"",
    });
    setEditId(rec.id); setPage("entry");
  };

  const filtered = useMemo(() => records.filter(r=>{
    if (search && !r.mold_size?.toLowerCase().includes(search.toLowerCase()) && !r.technician?.toLowerCase().includes(search.toLowerCase())) return false;
    if (filterProblem && !(r.problems||[]).includes(filterProblem)) return false;
    if (filterTech && r.technician!==filterTech) return false;
    return true;
  }), [records, search, filterProblem, filterTech]);

  const today = new Date().toISOString().slice(0,10);
  const todayRecs = records.filter(r=>r.date===today);
  const problemCounts = PROBLEMS.map(p=>({...p, count:records.filter(r=>(r.problems||[]).includes(p.id)).length})).sort((a,b)=>b.count-a.count);

  const detailRec = records.find(r=>r.id===detailId);
  const sameSize  = detailRec ? records.filter(r=>r.mold_size===detailRec.mold_size && r.id!==detailRec.id).slice(0,5) : [];

  const nav = [
    { id:"dashboard", icon:"ti-layout-dashboard", label:"Dashboard" },
    { id:"entry",     icon:"ti-plus",             label:"Entry Perbaikan" },
    { id:"database",  icon:"ti-database",          label:"Database Record" },
  ];

  const S = {
    app:    { display:"flex", minHeight:"100vh", background:"var(--color-background-tertiary)", fontFamily:"var(--font-sans)" },
    sb:     { width:sidebarOpen?210:52, background:"var(--color-background-primary)", borderRight:"0.5px solid var(--color-border-tertiary)", display:"flex", flexDirection:"column", transition:"width 0.2s", flexShrink:0 },
    sTop:   { padding:sidebarOpen?"18px 16px 12px":"18px 10px 12px", borderBottom:"0.5px solid var(--color-border-tertiary)", display:"flex", alignItems:"center", gap:10 },
    logo:   { width:28, height:28, background:"#1D9E75", borderRadius:6, display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 },
    nav:    (a)=>({ display:"flex", alignItems:"center", gap:10, padding:sidebarOpen?"9px 16px":"9px 14px", cursor:"pointer", fontSize:13, color:a?"var(--color-text-success)":"var(--color-text-secondary)", background:a?"var(--color-background-success)":"transparent", borderLeft:a?"2px solid var(--color-border-success)":"2px solid transparent", whiteSpace:"nowrap", overflow:"hidden" }),
    main:   { flex:1, display:"flex", flexDirection:"column", overflow:"hidden", minWidth:0 },
    topbar: { padding:"12px 20px", borderBottom:"0.5px solid var(--color-border-tertiary)", background:"var(--color-background-primary)", display:"flex", alignItems:"center", justifyContent:"space-between" },
    ct:     { flex:1, overflow:"auto", padding:20 },
    card:   { background:"var(--color-background-primary)", border:"0.5px solid var(--color-border-tertiary)", borderRadius:"var(--border-radius-lg)", padding:"16px 20px", marginBottom:16 },
    stat:   { background:"var(--color-background-secondary)", borderRadius:"var(--border-radius-md)", padding:"12px 16px" },
    lbl:    { fontSize:11, color:"var(--color-text-secondary)", marginBottom:6, display:"block", fontWeight:500 },
    inp:    { width:"100%", marginBottom:0 },
    btnP:   { padding:"8px 18px", borderRadius:"var(--border-radius-md)", fontSize:13, cursor:"pointer", fontWeight:500, border:"none", background:"#1D9E75", color:"#fff" },
    btn:    { padding:"8px 18px", borderRadius:"var(--border-radius-md)", fontSize:13, cursor:"pointer", fontWeight:500, border:"0.5px solid var(--color-border-secondary)", background:"transparent", color:"var(--color-text-primary)" },
    btnSm:  { padding:"4px 10px", borderRadius:"var(--border-radius-md)", fontSize:11, cursor:"pointer", border:"0.5px solid var(--color-border-secondary)", background:"transparent", color:"var(--color-text-primary)" },
    th:     { textAlign:"left", padding:"8px 12px", fontSize:11, color:"var(--color-text-secondary)", fontWeight:500, borderBottom:"0.5px solid var(--color-border-tertiary)", whiteSpace:"nowrap" },
    td:     { padding:"10px 12px", fontSize:12, color:"var(--color-text-primary)", borderBottom:"0.5px solid var(--color-border-tertiary)", verticalAlign:"top" },
    pCard:  (a)=>({ border:a?"1.5px solid #1D9E75":"0.5px solid var(--color-border-tertiary)", borderRadius:"var(--border-radius-md)", cursor:"pointer", background:a?"#E1F5EE":"var(--color-background-secondary)", overflow:"hidden" }),
    sec:    { background:"var(--color-background-secondary)", borderRadius:"var(--border-radius-md)", padding:"14px 16px", marginBottom:14 },
    typeB:  (a)=>({ flex:1, padding:"8px 12px", borderRadius:"var(--border-radius-md)", border:a?"1.5px solid #1D9E75":"0.5px solid var(--color-border-tertiary)", background:a?"#E1F5EE":"transparent", color:a?"#085041":"var(--color-text-secondary)", fontSize:12, fontWeight:a?500:400, cursor:"pointer", textAlign:"center" }),
  };

  if (loading) return (
    <div style={{ display:"flex", alignItems:"center", justifyContent:"center", minHeight:"100vh", flexDirection:"column", gap:12, background:"var(--color-background-tertiary)" }}>
      <div style={{ width:32, height:32, border:"3px solid #E1F5EE", borderTop:"3px solid #1D9E75", borderRadius:"50%", animation:"spin 0.8s linear infinite" }}></div>
      <div style={{ fontSize:13, color:"var(--color-text-secondary)" }}>Memuat data...</div>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );

  return (
    <div style={S.app}>
      {toast && (
        <div style={{ position:"fixed", top:16, right:20, zIndex:999, background:toast.type==="error"?"var(--color-background-danger)":"var(--color-background-success)", color:toast.type==="error"?"var(--color-text-danger)":"var(--color-text-success)", padding:"10px 18px", borderRadius:"var(--border-radius-md)", fontSize:13, border:`0.5px solid ${toast.type==="error"?"var(--color-border-danger)":"var(--color-border-success)"}`, boxShadow:"0 2px 8px rgba(0,0,0,0.08)" }}>
          {toast.msg}
        </div>
      )}

      {/* SIDEBAR */}
      <div style={S.sb}>
        <div style={S.sTop}>
          <div style={S.logo}><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="6" stroke="white" strokeWidth="1.5"/><circle cx="8" cy="8" r="2" fill="white"/></svg></div>
          {sidebarOpen&&<div><div style={{ fontSize:13, fontWeight:500, color:"var(--color-text-primary)" }}>MoldTrack</div><div style={{ fontSize:10, color:"var(--color-text-secondary)" }}>Mold Tire Action</div></div>}
          <button onClick={()=>setSidebarOpen(o=>!o)} style={{ marginLeft:"auto", background:"none", border:"none", cursor:"pointer", color:"var(--color-text-secondary)", padding:2 }}>
            <i className={`ti ${sidebarOpen?"ti-chevron-left":"ti-chevron-right"}`} style={{ fontSize:15 }}></i>
          </button>
        </div>
        <div style={{ padding:"10px 0", flex:1 }}>
          {nav.map(n=>(
            <div key={n.id} style={S.nav(page===n.id)} onClick={()=>{setPage(n.id);if(n.id==="entry"){setForm(emptyForm());setEditId(null);}}}>
              <i className={`ti ${n.icon}`} style={{ fontSize:17, flexShrink:0 }}></i>
              {sidebarOpen&&<span>{n.label}</span>}
            </div>
          ))}
        </div>
        {sidebarOpen&&(
          <div style={{ padding:"12px 16px", borderTop:"0.5px solid var(--color-border-tertiary)" }}>
            <div style={{ fontSize:10, color:"var(--color-text-secondary)", display:"flex", alignItems:"center", gap:6 }}>
              <div style={{ width:6, height:6, borderRadius:"50%", background:"#1D9E75" }}></div>
              Terhubung ke Supabase
            </div>
          </div>
        )}
      </div>

      {/* MAIN */}
      <div style={S.main}>
        <div style={S.topbar}>
          <div style={{ fontSize:14, fontWeight:500, color:"var(--color-text-primary)" }}>
            {page==="dashboard"&&"Dashboard"}
            {page==="entry"&&(editId?"Edit record perbaikan":"Entry perbaikan baru")}
            {page==="database"&&"Database record perbaikan"}
            {page==="detail"&&"Detail record"}
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:12 }}>
            <span style={{ fontSize:12, color:"var(--color-text-secondary)" }}>{today}</span>
            <button style={{ ...S.btnSm, fontSize:12 }} onClick={loadRecords} title="Refresh data">
              <i className="ti ti-refresh" style={{ fontSize:13 }}></i>
            </button>
          </div>
        </div>

        <div style={S.ct}>

          {/* ── DASHBOARD ── */}
          {page==="dashboard"&&(
            <div>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(140px, 1fr))", gap:12, marginBottom:20 }}>
                {[
                  { label:"Total record",    val:records.length,   sub:`${todayRecs.length} hari ini` },
                  { label:"Teknisi aktif",   val:knownTechs.length, sub:"total teknisi" },
                  { label:"Size tercatat",   val:knownSizes.length, sub:"jenis mold" },
                  { label:"Record bulan ini",val:records.filter(r=>r.date?.startsWith(new Date().toISOString().slice(0,7))).length, sub:new Date().toLocaleString("id-ID",{month:"long",year:"numeric"}) },
                ].map((c,i)=>(
                  <div key={i} style={S.stat}>
                    <div style={{ fontSize:11, color:"var(--color-text-secondary)", marginBottom:6 }}>{c.label}</div>
                    <div style={{ fontSize:24, fontWeight:500, color:"var(--color-text-primary)" }}>{c.val}</div>
                    <div style={{ fontSize:11, color:"var(--color-text-secondary)" }}>{c.sub}</div>
                  </div>
                ))}
              </div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:16 }}>
                <div style={S.card}>
                  <div style={{ fontSize:12, fontWeight:500, color:"var(--color-text-primary)", marginBottom:14 }}>Distribusi jenis problem</div>
                  {problemCounts.filter(p=>p.count>0).map(p=>(
                    <div key={p.id} style={{ marginBottom:10 }}>
                      <div style={{ display:"flex", justifyContent:"space-between", fontSize:11, color:"var(--color-text-secondary)", marginBottom:4 }}>
                        <span>{p.label}</span><span style={{ fontWeight:500, color:"var(--color-text-primary)" }}>{p.count}</span>
                      </div>
                      <div style={{ height:5, background:"var(--color-background-secondary)", borderRadius:3 }}>
                        <div style={{ height:5, background:"#1D9E75", borderRadius:3, width:`${Math.round((p.count/records.length)*100)}%` }}></div>
                      </div>
                    </div>
                  ))}
                  {problemCounts.every(p=>p.count===0)&&<div style={{ fontSize:12, color:"var(--color-text-secondary)" }}>Belum ada data.</div>}
                </div>
                <div style={S.card}>
                  <div style={{ fontSize:12, fontWeight:500, color:"var(--color-text-primary)", marginBottom:14 }}>Aktivitas per teknisi</div>
                  {knownTechs.length===0&&<div style={{ fontSize:12, color:"var(--color-text-secondary)" }}>Belum ada data.</div>}
                  {knownTechs.map(t=>{
                    const cnt=records.filter(r=>r.technician===t).length;
                    return (
                      <div key={t} style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10 }}>
                        <div style={{ width:28, height:28, borderRadius:"50%", background:"var(--color-background-info)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:10, fontWeight:500, color:"var(--color-text-info)", flexShrink:0 }}>
                          {t.split(" ").map(w=>w[0]).join("").slice(0,2).toUpperCase()}
                        </div>
                        <div style={{ flex:1 }}>
                          <div style={{ fontSize:12, color:"var(--color-text-primary)", marginBottom:3 }}>{t}</div>
                          <div style={{ height:4, background:"var(--color-background-secondary)", borderRadius:2 }}>
                            <div style={{ height:4, background:"#1D9E75", borderRadius:2, width:records.length?`${Math.round((cnt/records.length)*100)}%`:"0%" }}></div>
                          </div>
                        </div>
                        <span style={{ fontSize:11, color:"var(--color-text-secondary)", minWidth:20, textAlign:"right" }}>{cnt}</span>
                      </div>
                    );
                  })}
                </div>
              </div>
              <div style={S.card}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
                  <div style={{ fontSize:12, fontWeight:500, color:"var(--color-text-primary)" }}>Record terbaru</div>
                  <button style={S.btnP} onClick={()=>{setPage("entry");setForm(emptyForm());setEditId(null);}}>
                    <i className="ti ti-plus" style={{ fontSize:13, marginRight:4 }}></i>Entry baru
                  </button>
                </div>
                {records.length===0&&<div style={{ fontSize:13, color:"var(--color-text-secondary)", textAlign:"center", padding:24 }}>Belum ada record. Mulai entry perbaikan pertama!</div>}
                {records.length>0&&(
                  <div style={{ overflowX:"auto" }}>
                    <table style={{ width:"100%", borderCollapse:"collapse" }}>
                      <thead><tr>{["Size mold","Problem","Teknisi","Tanggal","Jam",""].map((h,i)=><th key={i} style={S.th}>{h}</th>)}</tr></thead>
                      <tbody>
                        {records.slice(0,8).map(r=>(
                          <tr key={r.id} style={{ cursor:"pointer" }} onClick={()=>{setDetailId(r.id);setPage("detail");}}>
                            <td style={S.td}><span style={{ fontWeight:500 }}>{r.mold_size}</span><br/><span style={{ fontSize:10, color:"var(--color-text-secondary)" }}>{r.mold_type==="segmented"?"Segmented":"Two Piece"}</span></td>
                            <td style={S.td}>{(r.problems||[]).slice(0,2).map(p=><ProblemBadge key={p} pid={p}/>)}{(r.problems||[]).length>2&&<span style={{ fontSize:10, color:"var(--color-text-secondary)" }}>+{r.problems.length-2}</span>}</td>
                            <td style={S.td}>{r.technician}</td>
                            <td style={S.td}>{r.date}</td>
                            <td style={S.td}>{r.jam_mulai&&r.jam_selesai?`${r.jam_mulai}–${r.jam_selesai}`:"-"}</td>
                            <td style={S.td}><button style={S.btnSm} onClick={e=>{e.stopPropagation();setDetailId(r.id);setPage("detail");}}>Detail</button></td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* ── ENTRY FORM ── */}
          {page==="entry"&&(
            <div style={{ maxWidth:680 }}>
              <div style={S.card}>
                <div style={{ fontSize:13, fontWeight:500, marginBottom:14, color:"var(--color-text-primary)" }}>Identitas perbaikan</div>
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:12 }}>
                  <AutocompleteInput label="Size mold *" value={form.moldSize} onChange={v=>setForm(f=>({...f,moldSize:v}))} suggestions={knownSizes} placeholder="cth: 205/65R15" />
                  <div>
                    <label style={S.lbl}>Tanggal perbaikan *</label>
                    <input type="date" style={S.inp} value={form.date} onChange={e=>setForm(f=>({...f,date:e.target.value}))} />
                  </div>
                </div>
                <div style={{ marginBottom:12 }}>
                  <label style={S.lbl}>Tipe mold *</label>
                  <div style={{ display:"flex", gap:8 }}>
                    {[["segmented","ti-grid-dots","Segmented (Sektor A–H)"],["two_piece","ti-layout-rows","Two Piece (Atas/Bawah)"]].map(([val,icon,lbl])=>(
                      <button key={val} onClick={()=>setForm(f=>({...f,moldType:val}))} style={S.typeB(form.moldType===val)}>
                        <i className={`ti ${icon}`} style={{ fontSize:14, marginRight:6 }}></i>{lbl}
                      </button>
                    ))}
                  </div>
                </div>
                <div style={{ marginBottom:12 }}>
                  <AutocompleteInput label="Nama teknisi *" value={form.technician} onChange={v=>setForm(f=>({...f,technician:v}))} suggestions={knownTechs} placeholder="Ketik nama kamu..." />
                </div>
                <div>
                  <label style={S.lbl}>Jam pengerjaan</label>
                  <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                    <input type="time" style={{ ...S.inp, flex:1 }} value={form.jamMulai} onChange={e=>setForm(f=>({...f,jamMulai:e.target.value}))} />
                    <span style={{ fontSize:12, color:"var(--color-text-secondary)", flexShrink:0 }}>sampai</span>
                    <input type="time" style={{ ...S.inp, flex:1 }} value={form.jamSelesai} onChange={e=>setForm(f=>({...f,jamSelesai:e.target.value}))} />
                    {form.jamMulai&&form.jamSelesai&&(()=>{
                      const [h1,m1]=form.jamMulai.split(":").map(Number);
                      const [h2,m2]=form.jamSelesai.split(":").map(Number);
                      const diff=(h2*60+m2)-(h1*60+m1);
                      if(diff>0) return <span style={{ fontSize:11, color:"var(--color-text-secondary)", flexShrink:0, whiteSpace:"nowrap" }}>{Math.floor(diff/60)>0?`${Math.floor(diff/60)} jam `:""}{diff%60>0?`${diff%60} mnt`:""}</span>;
                      return null;
                    })()}
                  </div>
                </div>
              </div>

              <div style={S.card}>
                <div style={{ fontSize:13, fontWeight:500, marginBottom:14, color:"var(--color-text-primary)" }}>Jenis problem <span style={{ fontSize:11, color:"var(--color-text-secondary)", fontWeight:400 }}>(bisa lebih dari 1)</span></div>
                {["MOR","OOR","Overflow","OS"].map(grp=>(
                  <div key={grp} style={{ marginBottom:16 }}>
                    <div style={{ fontSize:11, color:"var(--color-text-secondary)", marginBottom:8, fontWeight:500, letterSpacing:"0.5px" }}>{grp}</div>
                    <div style={{ display:"flex", flexDirection:"column", gap:6 }}>
                      {PROBLEMS.filter(p=>p.group===grp).map(p=>{
                        const active=form.problems.includes(p.id);
                        const det=form.problemDetails[p.id];
                        return (
                          <div key={p.id} style={S.pCard(active)}>
                            <div style={{ display:"flex", alignItems:"center", gap:8, padding:"9px 12px" }} onClick={()=>toggleProblem(p.id)}>
                              <div style={{ width:16, height:16, borderRadius:4, border:active?"none":"1px solid var(--color-border-secondary)", background:active?"#1D9E75":"transparent", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
                                {active&&<i className="ti ti-check" style={{ fontSize:11, color:"#fff" }}></i>}
                              </div>
                              <span style={{ fontSize:12, color:active?"#085041":"var(--color-text-primary)", fontWeight:active?500:400 }}>{p.label}</span>
                            </div>
                            {active&&det&&(
                              <div style={{ padding:"0 12px 12px", borderTop:"0.5px solid #9FE1CB" }}>
                                <ProblemDetailForm pid={p.id} detail={det} onChange={d=>updateDetail(p.id,d)} />
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                ))}
              </div>

              <div style={S.card}>
                <div style={{ fontSize:13, fontWeight:500, marginBottom:14, color:"var(--color-text-primary)" }}>Catatan perbaikan</div>
                <textarea rows={4} style={{ ...S.inp, resize:"vertical" }} placeholder="Deskripsikan langkah perbaikan yang dilakukan..." value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} />
              </div>

              <div style={{ display:"flex", gap:10 }}>
                <button style={S.btnP} onClick={submitForm}>{editId?"Update record":"Simpan record"}</button>
                <button style={S.btn} onClick={()=>{setForm(emptyForm());setEditId(null);}}>Reset</button>
                <button style={S.btn} onClick={()=>setPage("database")}>Batal</button>
              </div>
            </div>
          )}

          {/* ── DATABASE ── */}
          {page==="database"&&(
            <div>
              <div style={{ display:"flex", gap:10, marginBottom:16, flexWrap:"wrap", alignItems:"center" }}>
                <div style={{ position:"relative", flex:"1 1 200px" }}>
                  <i className="ti ti-search" style={{ position:"absolute", left:10, top:"50%", transform:"translateY(-50%)", fontSize:15, color:"var(--color-text-secondary)" }}></i>
                  <input style={{ paddingLeft:34, width:"100%" }} placeholder="Cari size mold atau teknisi..." value={search} onChange={e=>setSearch(e.target.value)} />
                </div>
                <select style={{ minWidth:150 }} value={filterProblem} onChange={e=>setFilterProblem(e.target.value)}>
                  <option value="">Semua problem</option>
                  {PROBLEMS.map(p=><option key={p.id} value={p.id}>{p.label}</option>)}
                </select>
                <select style={{ minWidth:150 }} value={filterTech} onChange={e=>setFilterTech(e.target.value)}>
                  <option value="">Semua teknisi</option>
                  {knownTechs.map(t=><option key={t}>{t}</option>)}
                </select>
                <button style={S.btnP} onClick={()=>{setPage("entry");setForm(emptyForm());setEditId(null);}}>
                  <i className="ti ti-plus" style={{ fontSize:13, marginRight:4 }}></i>Entry baru
                </button>
              </div>
              <div style={{ fontSize:12, color:"var(--color-text-secondary)", marginBottom:10 }}>{filtered.length} record ditemukan</div>
              <div style={{ ...S.card, padding:0, overflow:"hidden" }}>
                <div style={{ overflowX:"auto" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse" }}>
                    <thead>
                      <tr style={{ background:"var(--color-background-secondary)" }}>
                        {["#","Size mold","Problem & tindakan","Teknisi","Tanggal","Jam","Aksi"].map((h,i)=><th key={i} style={S.th}>{h}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {filtered.length===0&&<tr><td colSpan={7} style={{ ...S.td, textAlign:"center", padding:32, color:"var(--color-text-secondary)" }}>Tidak ada record yang sesuai filter.</td></tr>}
                      {filtered.map((r,i)=>(
                        <tr key={r.id} style={{ cursor:"pointer" }} onClick={()=>{setDetailId(r.id);setPage("detail");}}>
                          <td style={{ ...S.td, color:"var(--color-text-secondary)", fontSize:11 }}>{i+1}</td>
                          <td style={S.td}><span style={{ fontWeight:500, fontSize:13 }}>{r.mold_size}</span><br/><span style={{ fontSize:10, color:"var(--color-text-secondary)" }}>{r.mold_type==="segmented"?"Segmented":"Two Piece"}</span></td>
                          <td style={{ ...S.td, maxWidth:220 }}><DetailSummary rec={{...r, problemDetails:r.problem_details, problems:r.problems||[]}} /></td>
                          <td style={S.td}>{r.technician}</td>
                          <td style={{ ...S.td, whiteSpace:"nowrap" }}>{r.date}</td>
                          <td style={{ ...S.td, whiteSpace:"nowrap" }}>{r.jam_mulai&&r.jam_selesai?`${r.jam_mulai}–${r.jam_selesai}`:"-"}</td>
                          <td style={S.td}>
                            <div style={{ display:"flex", flexDirection:"column", gap:4 }} onClick={e=>e.stopPropagation()}>
                              <button style={S.btnSm} onClick={()=>{setDetailId(r.id);setPage("detail");}}>Detail</button>
                              <button style={S.btnSm} onClick={()=>startEdit(r)}>Edit</button>
                              <button style={{ ...S.btnSm, color:"var(--color-text-danger)", borderColor:"var(--color-border-danger)" }} onClick={()=>deleteRecord(r.id)}>Hapus</button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* ── DETAIL ── */}
          {page==="detail"&&detailRec&&(
            <div style={{ maxWidth:680 }}>
              <button style={{ ...S.btn, marginBottom:16, display:"flex", alignItems:"center", gap:6 }} onClick={()=>setPage("database")}>
                <i className="ti ti-arrow-left" style={{ fontSize:14 }}></i> Kembali
              </button>
              <div style={S.card}>
                <div style={{ marginBottom:16 }}>
                  <div style={{ fontSize:20, fontWeight:500, color:"var(--color-text-primary)" }}>{detailRec.mold_size}</div>
                  <div style={{ fontSize:12, color:"var(--color-text-secondary)", marginTop:2 }}>{detailRec.mold_type==="segmented"?"Segmented (A–H)":"Two Piece"} · {detailRec.date}</div>
                </div>
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:12, marginBottom:16 }}>
                  <div style={S.stat}><div style={{ fontSize:11, color:"var(--color-text-secondary)", marginBottom:4 }}>Teknisi</div><div style={{ fontSize:13, fontWeight:500 }}>{detailRec.technician}</div></div>
                  <div style={S.stat}><div style={{ fontSize:11, color:"var(--color-text-secondary)", marginBottom:4 }}>Tanggal</div><div style={{ fontSize:13, fontWeight:500 }}>{detailRec.date}</div></div>
                  <div style={S.stat}><div style={{ fontSize:11, color:"var(--color-text-secondary)", marginBottom:4 }}>Jam pengerjaan</div><div style={{ fontSize:13, fontWeight:500 }}>{detailRec.jam_mulai&&detailRec.jam_selesai?`${detailRec.jam_mulai} – ${detailRec.jam_selesai}`:"-"}</div></div>
                </div>
                <div style={{ fontSize:12, color:"var(--color-text-secondary)", marginBottom:10 }}>Detail tindakan per problem</div>
                {(detailRec.problems||[]).map(pid=>{
                  const det=(detailRec.problem_details||{})[pid];
                  const isMOR=pid.startsWith("MOR"), isOverflow=pid.startsWith("OVERFLOW"), isOS=pid==="OS", isOOR=pid==="OOR";
                  return (
                    <div key={pid} style={{ ...S.sec, marginBottom:10 }}>
                      <div style={{ marginBottom:8 }}><ProblemBadge pid={pid}/></div>
                      {det&&(isMOR||isOverflow||isOS)&&(
                        <div style={{ fontSize:13, display:"flex", flexDirection:"column", gap:6 }}>
                          <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                            <i className="ti ti-layers-subtract" style={{ fontSize:15, color:"var(--color-text-secondary)" }}></i>
                            Shim <span style={{ fontWeight:500 }}>{det.shimAction}</span> {det.shimJumlah||"—"} lembar
                            {!isMOR&&det.shimLokasi&&<span> — lokasi <span style={{ fontWeight:500 }}>{det.shimLokasi}</span></span>}
                          </div>
                          {isMOR&&det.shimSectors?.length>0&&(
                            <div style={{ paddingLeft:22 }}>
                              Sektor: {det.shimSectors.map(s=>(
                                <span key={s} style={{ display:"inline-block", background:"#EAF3DE", color:"#27500A", fontSize:11, padding:"1px 8px", borderRadius:4, marginRight:4, fontWeight:500 }}>{s}</span>
                              ))}
                            </div>
                          )}
                        </div>
                      )}
                      {det&&isOOR&&(
                        <div style={{ fontSize:13 }}>
                          <div style={{ display:"flex", alignItems:"center", gap:6, marginBottom:4 }}>
                            <i className="ti ti-grid-dots" style={{ fontSize:15, color:"var(--color-text-secondary)" }}></i>
                            Tipe: <span style={{ fontWeight:500 }}>{det.oorType==="segmented"?"Segmented":"Two Piece"}</span>
                          </div>
                          {det.oorSectors?.length>0&&(
                            <div style={{ paddingLeft:22 }}>
                              Sektor: {det.oorSectors.map(s=>(
                                <span key={s} style={{ display:"inline-block", background:"#E1F5EE", color:"#085041", fontSize:11, padding:"1px 8px", borderRadius:4, marginRight:4, fontWeight:500 }}>{s}</span>
                              ))}
                            </div>
                          )}
                        </div>
                      )}
                      {!det&&<div style={{ fontSize:12, color:"var(--color-text-secondary)" }}>Tidak ada detail tindakan.</div>}
                    </div>
                  );
                })}
                {detailRec.notes&&(
                  <div style={{ marginTop:12 }}>
                    <div style={{ fontSize:11, color:"var(--color-text-secondary)", marginBottom:6 }}>Catatan perbaikan</div>
                    <div style={{ fontSize:13, lineHeight:1.7 }}>{detailRec.notes}</div>
                  </div>
                )}
                <div style={{ display:"flex", gap:10, paddingTop:16, borderTop:"0.5px solid var(--color-border-tertiary)", marginTop:16 }}>
                  <button style={S.btnP} onClick={()=>startEdit(detailRec)}><i className="ti ti-edit" style={{ fontSize:13, marginRight:4 }}></i>Edit</button>
                  <button style={{ ...S.btn, color:"var(--color-text-danger)", borderColor:"var(--color-border-danger)" }} onClick={()=>deleteRecord(detailRec.id)}>Hapus</button>
                </div>
              </div>
              {sameSize.length>0&&(
                <div style={S.card}>
                  <div style={{ fontSize:13, fontWeight:500, marginBottom:14 }}>
                    <i className="ti ti-history" style={{ fontSize:15, marginRight:6, color:"var(--color-text-secondary)" }}></i>
                    Riwayat size {detailRec.mold_size} sebelumnya
                  </div>
                  {sameSize.map(r=>(
                    <div key={r.id} style={{ padding:"10px 0", borderBottom:"0.5px solid var(--color-border-tertiary)", cursor:"pointer" }} onClick={()=>setDetailId(r.id)}>
                      <div style={{ fontSize:12, color:"var(--color-text-secondary)", marginBottom:6 }}>{r.date}{r.jam_mulai?` · ${r.jam_mulai}–${r.jam_selesai}`:""}</div>
                      <DetailSummary rec={{...r, problemDetails:r.problem_details, problems:r.problems||[]}} />
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

        </div>
      </div>
    </div>
  );
}
