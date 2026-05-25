import { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Utilities ────────────────────────────────────────────────────────────────
const uid = () => Math.random().toString(36).slice(2,8) + Date.now().toString(36);
const fmtDate = (d) => !d ? "—" : new Date(d + "T00:00:00").toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"});
const daysLeft = (d) => !d ? null : Math.ceil((new Date(d+"T00:00:00") - Date.now()) / 86400000);
const todayStr = () => new Date().toISOString().slice(0,10);
const clamp = (v,min,max) => Math.max(min,Math.min(max,v));

// ─── API ──────────────────────────────────────────────────────────────────────
export const API = {
  get: async (table) => { try { const r = await fetch(`http://localhost:3000/api/${table}`); return await r.json(); } catch(e) { console.error(e); return []; } },
  create: async (table, data) => { await fetch(`http://localhost:3000/api/${table}`, { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(data) }); },
  update: async (table, id, data) => { await fetch(`http://localhost:3000/api/${table}/${id}`, { method: 'PUT', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(data) }); },
  delete: async (table, id) => { await fetch(`http://localhost:3000/api/${table}/${id}`, { method: 'DELETE' }); }
};

// ─── Constants ────────────────────────────────────────────────────────────────
const SCH_STATUSES = ["open","closed","applied","awarded","rejected"];
const APP_STATUSES = ["researching","in-progress","submitted","interview","accepted","waitlisted","rejected"];
const VERDICTS = ["pending","interested","positive","negative","no-response","declined"];
const DEFAULT_DOCS = [
  "Statement of Purpose","CV / Resume","Official Transcripts",
  "Recommendation Letter 1","Recommendation Letter 2","Recommendation Letter 3",
  "English Test (IELTS/TOEFL)","Research Proposal","Writing Sample",
  "Passport Copy","Financial Proof","Application Fee Receipt"
];

// ─── Theme ────────────────────────────────────────────────────────────────────
const T = {
  bg0:"var(--bg0)", bg1:"var(--bg1)", bg2:"var(--bg2)", bg3:"var(--bg3)",
  bg4:"var(--bg4)", border:"var(--border)", borderEm:"var(--borderEm)",
  tp:"var(--tp)", ts:"var(--ts)", tm:"var(--tm)",
  accent:"#f59e0b", accentDim:"rgba(245,158,11,0.15)", accentHover:"#d97706",
  teal:"#14b8a6", tealDim:"rgba(20,184,166,0.15)",
  red:"#ef4444", redDim:"rgba(239,68,68,0.15)",
  green:"#22c55e", greenDim:"rgba(34,197,94,0.15)",
  blue:"#60a5fa", blueDim:"rgba(96,165,250,0.15)",
  purple:"#a78bfa", purpleDim:"rgba(167,139,250,0.15)",
  orange:"#fb923c", orangeDim:"rgba(251,146,60,0.15)",
};

const SC = {
  open:{bg:T.tealDim,tx:T.teal}, closed:{bg:"rgba(148,163,184,.12)",tx:"#94a3b8"},
  applied:{bg:T.blueDim,tx:T.blue}, awarded:{bg:T.greenDim,tx:T.green},
  rejected:{bg:T.redDim,tx:T.red}, researching:{bg:T.purpleDim,tx:T.purple},
  "in-progress":{bg:T.blueDim,tx:T.blue}, submitted:{bg:T.tealDim,tx:T.teal},
  interview:{bg:T.orangeDim,tx:T.orange}, accepted:{bg:T.greenDim,tx:T.green},
  waitlisted:{bg:T.accentDim,tx:T.accent}, pending:{bg:"rgba(148,163,184,.12)",tx:"#94a3b8"},
  interested:{bg:T.purpleDim,tx:T.purple}, positive:{bg:T.greenDim,tx:T.green},
  negative:{bg:T.redDim,tx:T.red}, "no-response":{bg:T.accentDim,tx:T.accent},
  declined:{bg:T.redDim,tx:T.red},
};

// ─── Shared UI Components ─────────────────────────────────────────────────────

const Badge = ({v}) => {
  const c = SC[v]||{bg:"rgba(148,163,184,.12)",tx:"#94a3b8"};
  return <span style={{background:c.bg,color:c.tx,fontSize:11,fontWeight:600,padding:"2px 10px",borderRadius:20,whiteSpace:"nowrap",letterSpacing:.3}}>{(v||"—").replace(/-/g," ")}</span>;
};

const DaysChip = ({date}) => {
  const d = daysLeft(date);
  if(d===null) return <span style={{color:T.tm}}>—</span>;
  if(d<0) return <span style={{color:T.red,fontSize:11,fontWeight:600}}>{Math.abs(d)}d overdue</span>;
  if(d===0) return <span style={{color:T.accent,fontSize:11,fontWeight:700}}>TODAY!</span>;
  const color = d<=7?T.red:d<=14?T.orange:d<=30?T.accent:T.ts;
  return <span style={{color,fontSize:11,fontWeight:d<=30?600:400}}>{d}d left</span>;
};

const Btn = ({children,onClick,variant="default",size="md",style:sx={},...rest}) => {
  const base = {cursor:"pointer",border:"none",borderRadius:8,fontWeight:600,fontFamily:"inherit",transition:"all .15s",display:"inline-flex",alignItems:"center",gap:6,...sx};
  const sizes = {sm:{padding:"5px 12px",fontSize:12}, md:{padding:"8px 16px",fontSize:13}, lg:{padding:"10px 20px",fontSize:14}};
  const variants = {
    default:{background:T.bg3,color:T.tp,border:`1px solid ${T.border}`},
    primary:{background:T.accent,color:"#000"},
    danger:{background:T.redDim,color:T.red,border:`1px solid rgba(239,68,68,.3)`},
    ghost:{background:"transparent",color:T.ts,border:`1px solid ${T.border}`},
    teal:{background:T.tealDim,color:T.teal,border:`1px solid rgba(20,184,166,.3)`},
  };
  return <button {...rest} onClick={onClick} style={{...base,...sizes[size],...variants[variant],...sx}}>{children}</button>;
};

const Input = ({style:sx={},...props}) => (
  <input {...props} style={{background:T.bg3,border:`1px solid ${T.border}`,borderRadius:8,color:T.tp,padding:"8px 12px",fontSize:13,fontFamily:"inherit",width:"100%",outline:"none",boxSizing:"border-box",...sx}} />
);

const Select = ({children,style:sx={},...props}) => (
  <select {...props} style={{background:T.bg3,border:`1px solid ${T.border}`,borderRadius:8,color:T.tp,padding:"8px 12px",fontSize:13,fontFamily:"inherit",width:"100%",outline:"none",boxSizing:"border-box",...sx}}>
    {children}
  </select>
);

const Textarea = ({style:sx={},...props}) => (
  <textarea {...props} style={{background:T.bg3,border:`1px solid ${T.border}`,borderRadius:8,color:T.tp,padding:"8px 12px",fontSize:13,fontFamily:"inherit",width:"100%",outline:"none",resize:"vertical",minHeight:80,boxSizing:"border-box",...sx}} />
);

const Label = ({children}) => <div style={{fontSize:11,fontWeight:600,color:T.ts,textTransform:"uppercase",letterSpacing:.8,marginBottom:5}}>{children}</div>;

const FormGroup = ({label,children}) => <div style={{marginBottom:14}}><Label>{label}</Label>{children}</div>;

// ─── Modal ────────────────────────────────────────────────────────────────────
const Modal = ({title,onClose,children,width=600}) => (
  <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.7)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,backdropFilter:"blur(4px)",padding:16}} onClick={e=>{if(e.target===e.currentTarget)onClose()}}>
    <div style={{background:T.bg2,border:`1px solid ${T.borderEm}`,borderRadius:14,width:"100%",maxWidth:width,maxHeight:"90vh",overflow:"hidden",display:"flex",flexDirection:"column",boxShadow:"0 24px 80px rgba(0,0,0,.5)"}}>
      <div style={{padding:"18px 22px",borderBottom:`1px solid ${T.border}`,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <span style={{fontWeight:700,fontSize:16,color:T.tp}}>{title}</span>
        <button onClick={onClose} style={{background:"none",border:"none",color:T.ts,cursor:"pointer",fontSize:20,lineHeight:1,padding:4}}>✕</button>
      </div>
      <div style={{overflow:"auto",flex:1,padding:"20px 22px"}}>{children}</div>
    </div>
  </div>
);

// ─── Confirm Dialog ───────────────────────────────────────────────────────────
const Confirm = ({message,onConfirm,onCancel}) => (
  <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.75)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1100}}>
    <div style={{background:T.bg2,border:`1px solid ${T.borderEm}`,borderRadius:12,padding:28,maxWidth:380,width:"90%",textAlign:"center"}}>
      <div style={{fontSize:28,marginBottom:12}}>🗑️</div>
      <p style={{color:T.tp,marginBottom:20,fontSize:14}}>{message}</p>
      <div style={{display:"flex",gap:10,justifyContent:"center"}}>
        <Btn onClick={onCancel} variant="ghost">Cancel</Btn>
        <Btn onClick={onConfirm} variant="danger">Delete</Btn>
      </div>
    </div>
  </div>
);

// ─── Table wrapper ────────────────────────────────────────────────────────────
const TH = ({children,onClick,sorted,style:sx={}}) => (
  <th onClick={onClick} style={{padding:"10px 12px",fontSize:11,fontWeight:600,color:T.ts,textTransform:"uppercase",letterSpacing:.7,textAlign:"left",whiteSpace:"nowrap",cursor:onClick?"pointer":"default",userSelect:"none",background:T.bg1,borderBottom:`1px solid ${T.border}`,...sx}}>
    {children}{sorted==1?" ↑":sorted==-1?" ↓":onClick?" ↕":""}
  </th>
);
const TD = ({children,style:sx={}}) => <td style={{padding:"10px 12px",fontSize:13,color:T.tp,borderBottom:`1px solid ${T.border}`,verticalAlign:"middle",...sx}}>{children}</td>;

// ─── Dashboard ────────────────────────────────────────────────────────────────
function Dashboard({scholarships,applications,emails,setTab}) {
  const now = Date.now();
  const upcomingSch = [...scholarships].filter(s=>s.deadline&&daysLeft(s.deadline)>=0&&daysLeft(s.deadline)<=60&&s.status==="open").sort((a,b)=>new Date(a.deadline)-new Date(b.deadline)).slice(0,5);
  const runningApps = applications.filter(a=>["in-progress","submitted","interview"].includes(a.status)).sort((a,b)=>new Date(a.deadline)-new Date(b.deadline)).slice(0,5);
  const pendingEmails = emails.filter(e=>e.verdict==="pending"||!e.verdict);
  const responsesReceived = emails.filter(e=>e.responseDate);
  const positiveEmails = emails.filter(e=>["positive","interested"].includes(e.verdict));
  const followUpsNeeded = emails.filter(e=>e.mailedDate && !e.responseDate && daysLeft(e.mailedDate)!==null && daysLeft(e.mailedDate)<-14);

  const StatCard = ({label,value,sub,color,onClick}) => (
    <div onClick={onClick} style={{background:T.bg2,border:`1px solid ${T.border}`,borderRadius:12,padding:"18px 20px",cursor:onClick?"pointer":"default",transition:"border-color .2s"}}
      onMouseEnter={e=>onClick&&(e.currentTarget.style.borderColor=T.borderEm)}
      onMouseLeave={e=>e.currentTarget.style.borderColor=T.border}>
      <div style={{fontSize:28,fontWeight:800,color:color||T.accent,marginBottom:4,fontVariantNumeric:"tabular-nums"}}>{value}</div>
      <div style={{fontSize:13,fontWeight:600,color:T.tp,marginBottom:2}}>{label}</div>
      {sub&&<div style={{fontSize:11,color:T.ts}}>{sub}</div>}
    </div>
  );

  const SectionHead = ({title,tab}) => (
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,marginTop:24}}>
      <span style={{fontWeight:700,fontSize:15,color:T.tp}}>{title}</span>
      <Btn onClick={()=>setTab(tab)} size="sm" variant="ghost">View all →</Btn>
    </div>
  );

  return (
    <div>
      <div style={{marginBottom:24}}>
        <h2 style={{margin:0,fontSize:22,fontWeight:800,color:T.tp,marginBottom:6}}>Academic Hub</h2>
        <p style={{margin:0,color:T.ts,fontSize:13}}>Track your scholarships, applications, and professor outreach.</p>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:12,marginBottom:8}}>
        <StatCard label="Scholarships" value={scholarships.length} sub={`${scholarships.filter(s=>s.status==="open").length} open`} color={T.teal} onClick={()=>setTab("scholarships")} />
        <StatCard label="Applications" value={applications.length} sub={`${applications.filter(a=>["accepted","awarded"].includes(a.status)).length} accepted`} color={T.blue} onClick={()=>setTab("applications")} />
        <StatCard label="Emails Sent" value={emails.length} sub={`${responsesReceived.length} responses received`} color={T.purple} onClick={()=>setTab("emails")} />
        <StatCard label="Positive Responses" value={positiveEmails.length} sub={`${followUpsNeeded.length} need follow-up`} color={T.green} onClick={()=>setTab("emails")} />
      </div>

      {followUpsNeeded.length > 0 && (
        <div style={{background:"rgba(239,68,68,.08)",border:`1px solid rgba(239,68,68,.25)`,borderRadius:10,padding:"12px 16px",marginTop:16,display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:18}}>⚠️</span>
          <span style={{color:T.red,fontSize:13,fontWeight:600}}>{followUpsNeeded.length} professor email{followUpsNeeded.length>1?"s":""} haven't responded in 14+ days — consider sending a follow-up.</span>
          <Btn onClick={()=>setTab("emails")} size="sm" variant="danger" style={{marginLeft:"auto"}}>View</Btn>
        </div>
      )}

      <SectionHead title="🎓 Upcoming Scholarship Deadlines" tab="scholarships" />
      {upcomingSch.length===0 ? <div style={{color:T.tm,fontSize:13,padding:"16px 0"}}>No open scholarships with upcoming deadlines. Add some! →</div> :
        <div style={{background:T.bg2,border:`1px solid ${T.border}`,borderRadius:12,overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>
              <TH>Scholarship</TH><TH>Country</TH><TH>Amount</TH><TH>Deadline</TH><TH>Time Left</TH>
            </tr></thead>
            <tbody>{upcomingSch.map(s=>(
              <tr key={s.id}>
                <TD><span style={{fontWeight:600,color:T.tp}}>{s.name}</span></TD>
                <TD><span style={{color:T.ts}}>{s.country||"—"}</span></TD>
                <TD><span style={{color:T.accent}}>{s.amount||"—"}</span></TD>
                <TD style={{color:T.ts,fontSize:12}}>{fmtDate(s.deadline)}</TD>
                <TD><DaysChip date={s.deadline}/></TD>
              </tr>
            ))}</tbody>
          </table>
        </div>
      }

      <SectionHead title="📋 Active Applications" tab="applications" />
      {runningApps.length===0 ? <div style={{color:T.tm,fontSize:13,padding:"16px 0"}}>No active applications. Start tracking! →</div> :
        <div style={{background:T.bg2,border:`1px solid ${T.border}`,borderRadius:12,overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>
              <TH>University</TH><TH>Program</TH><TH>Deadline</TH><TH>Status</TH><TH>Docs</TH>
            </tr></thead>
            <tbody>{runningApps.map(a=>{
              const done = (a.docs||[]).filter(d=>d.done).length;
              const total = (a.docs||[]).length;
              return <tr key={a.id}>
                <TD><span style={{fontWeight:600}}>{a.university}</span></TD>
                <TD><span style={{color:T.ts,fontSize:12}}>{a.program||"—"}</span></TD>
                <TD><DaysChip date={a.deadline}/></TD>
                <TD><Badge v={a.status}/></TD>
                <TD><span style={{fontSize:12,color:done===total&&total>0?T.green:T.ts}}>{done}/{total}</span></TD>
              </tr>;
            })}</tbody>
          </table>
        </div>
      }

      <SectionHead title="📧 Recent Professor Outreach" tab="emails" />
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:10}}>
        {VERDICTS.map(v=>{
          const cnt = emails.filter(e=>(e.verdict||"pending")===v).length;
          return <div key={v} style={{background:T.bg2,border:`1px solid ${T.border}`,borderRadius:10,padding:"12px 14px"}}>
            <div style={{fontSize:20,fontWeight:800,color:(SC[v]||{tx:T.ts}).tx,marginBottom:4}}>{cnt}</div>
            <div style={{fontSize:11,color:T.ts,textTransform:"capitalize"}}>{v.replace(/-/g," ")}</div>
          </div>;
        })}
      </div>
    </div>
  );
}

// ─── Scholarships ─────────────────────────────────────────────────────────────
const emptySch = () => ({id:uid(),name:"",country:"",provider:"",amount:"",deadline:"",status:"open",requirements:"",requiredDocuments:"",link:"",notes:""});

function ScholarshipsView({data,setData}) {
  const [modal, setModal] = useState(null); // null | {mode:"add"|"edit", item}
  const [confirm, setConfirm] = useState(null);
  const [search, setSearch] = useState("");
  const [filterStatus, setFilterStatus] = useState("all");
  const [form, setForm] = useState({});

  const filtered = useMemo(()=>{
    return data.filter(s=>{
      const q = search.toLowerCase();
      const matchQ = !q || s.name.toLowerCase().includes(q) || (s.country||"").toLowerCase().includes(q) || (s.provider||"").toLowerCase().includes(q);
      const matchS = filterStatus==="all" || s.status===filterStatus;
      return matchQ && matchS;
    }).sort((a,b)=>new Date(a.deadline||0)-new Date(b.deadline||0));
  },[data,search,filterStatus]);

  const openAdd = () => { const item=emptySch(); setForm(item); setModal({mode:"add",item}); };
  const openEdit = (item) => { setForm({...item}); setModal({mode:"edit",item}); };
  const save = async () => {
    if(!form.name.trim()){alert("Name is required");return;}
    if(modal.mode==="add"){
      await API.create('scholarships', form);
      setData([...data,{...form}]);
    } else {
      await API.update('scholarships', form.id, form);
      setData(data.map(s=>s.id===form.id?{...form}:s));
    }
    setModal(null);
  };
  const del = async (id) => { await API.delete('scholarships', id); setData(data.filter(s=>s.id!==id)); setConfirm(null); };
  const F = (k) => ({value:form[k]||"", onChange:e=>setForm(f=>({...f,[k]:e.target.value}))});

  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:20,flexWrap:"wrap"}}>
        <div style={{flex:1,minWidth:200}}>
          <Input placeholder="Search by name, country, provider…" value={search} onChange={e=>setSearch(e.target.value)} />
        </div>
        <Select style={{width:140}} value={filterStatus} onChange={e=>setFilterStatus(e.target.value)}>
          <option value="all">All statuses</option>
          {SCH_STATUSES.map(s=><option key={s} value={s}>{s}</option>)}
        </Select>
        <Btn onClick={openAdd} variant="primary">+ Add Scholarship</Btn>
      </div>
      <div style={{fontSize:12,color:T.ts,marginBottom:12}}>{filtered.length} scholarship{filtered.length!==1?"s":""}</div>
      <div style={{background:T.bg2,border:`1px solid ${T.border}`,borderRadius:12,overflow:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",minWidth:800}}>
          <thead><tr>
            <TH>Name</TH><TH>Country</TH><TH>Provider</TH><TH>Amount</TH><TH>Deadline</TH><TH>Time Left</TH><TH>Status</TH><TH style={{width:80}}>Actions</TH>
          </tr></thead>
          <tbody>
            {filtered.length===0 && <tr><TD colSpan={8}><span style={{color:T.tm}}>No scholarships found. Click "+ Add Scholarship" to get started.</span></TD></tr>}
            {filtered.map(s=>(
              <tr key={s.id} style={{cursor:"pointer"}} onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.03)"} onMouseLeave={e=>e.currentTarget.style.background=""}>
                <TD><span style={{fontWeight:600,color:T.tp}}>{s.name}</span>{s.link&&<a href={s.link} target="_blank" rel="noreferrer" style={{color:T.blue,fontSize:11,marginLeft:6}}>↗</a>}</TD>
                <TD>{s.country||<span style={{color:T.tm}}>—</span>}</TD>
                <TD><span style={{color:T.ts}}>{s.provider||"—"}</span></TD>
                <TD><span style={{color:T.accent,fontWeight:600}}>{s.amount||"—"}</span></TD>
                <TD style={{fontSize:12,color:T.ts}}>{fmtDate(s.deadline)}</TD>
                <TD><DaysChip date={s.deadline}/></TD>
                <TD><Badge v={s.status}/></TD>
                <TD>
                  <div style={{display:"flex",gap:6}}>
                    <Btn size="sm" variant="ghost" onClick={()=>openEdit(s)}>✏️</Btn>
                    <Btn size="sm" variant="danger" onClick={()=>setConfirm(s.id)}>🗑</Btn>
                  </div>
                </TD>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {modal && (
        <Modal title={modal.mode==="add"?"Add Scholarship":"Edit Scholarship"} onClose={()=>setModal(null)} width={640}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            <FormGroup label="Scholarship Name *"><Input {...F("name")} placeholder="e.g. DAAD Fellowship" /></FormGroup>
            <FormGroup label="Country"><Input {...F("country")} placeholder="e.g. Germany" /></FormGroup>
            <FormGroup label="Provider / Organization"><Input {...F("provider")} placeholder="e.g. German Academic Exchange Service" /></FormGroup>
            <FormGroup label="Amount / Funding"><Input {...F("amount")} placeholder="e.g. €1,000/month" /></FormGroup>
            <FormGroup label="Deadline"><Input type="date" {...F("deadline")} /></FormGroup>
            <FormGroup label="Status">
              <Select {...F("status")}>{SCH_STATUSES.map(s=><option key={s} value={s}>{s}</option>)}</Select>
            </FormGroup>
            <FormGroup label="Application Link"><Input {...F("link")} placeholder="https://…" /></FormGroup>
          </div>
          <FormGroup label="Eligibility Requirements"><Textarea {...F("requirements")} placeholder="GPA, nationality, field of study, etc." /></FormGroup>
          <FormGroup label="Required Documents"><Textarea {...F("requiredDocuments")} placeholder="List documents needed for this scholarship…" /></FormGroup>
          <FormGroup label="Notes"><Textarea {...F("notes")} style={{minHeight:60}} placeholder="Personal notes, tips, contacts…" /></FormGroup>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end",paddingTop:8}}>
            <Btn onClick={()=>setModal(null)} variant="ghost">Cancel</Btn>
            <Btn onClick={save} variant="primary">💾 Save Scholarship</Btn>
          </div>
        </Modal>
      )}
      {confirm && <Confirm message="Delete this scholarship? This cannot be undone." onConfirm={()=>del(confirm)} onCancel={()=>setConfirm(null)} />}
    </div>
  );
}

// ─── Applications ─────────────────────────────────────────────────────────────
const emptyApp = () => ({
  id:uid(), university:"", country:"", program:"", degree:"", deadline:"",
  status:"researching", notes:"", link:"",
  docs: DEFAULT_DOCS.map((name,i)=>({id:i,name,done:false}))
});

function ApplicationsView({data,setData}) {
  const [modal, setModal] = useState(null);
  const [confirm, setConfirm] = useState(null);
  const [docsModal, setDocsModal] = useState(null);
  const [search, setSearch] = useState("");
  const [filterStatus, setFilterStatus] = useState("all");
  const [form, setForm] = useState({});
  const fileRef = useRef();

  const filtered = useMemo(()=>
    data.filter(a=>{
      const q=search.toLowerCase();
      const mQ=!q||a.university.toLowerCase().includes(q)||(a.program||"").toLowerCase().includes(q)||(a.country||"").toLowerCase().includes(q);
      const mS=filterStatus==="all"||a.status===filterStatus;
      return mQ&&mS;
    }).sort((a,b)=>new Date(a.deadline||0)-new Date(b.deadline||0))
  ,[data,search,filterStatus]);

  const openAdd = () => { setForm(emptyApp()); setModal({mode:"add"}); };
  const openEdit = (a) => { setForm({...a,docs:[...(a.docs||[]).map(d=>({...d}))]}); setModal({mode:"edit"}); };
  const save = async () => {
    if(!form.university.trim()){alert("University name required");return;}
    if(modal.mode==="add"){
      await API.create('applications', form);
      setData([...data,{...form}]);
    } else {
      await API.update('applications', form.id, form);
      setData(data.map(a=>a.id===form.id?{...form}:a));
    }
    setModal(null);
  };
  const del = async (id) => { await API.delete('applications', id); setData(data.filter(a=>a.id!==id)); setConfirm(null); };
  const toggleDoc = async (appId, docId) => {
    const app = data.find(a=>a.id===appId);
    if(!app) return;
    const newDocs = app.docs.map(d=>d.id!==docId?d:{...d,done:!d.done});
    const newApp = {...app, docs: newDocs};
    await API.update('applications', appId, newApp);
    setData(data.map(a=>a.id!==appId?a:newApp));
    setDocsModal(m=>m?{...m,item:newApp}:null);
  };
  const F = (k) => ({value:form[k]||"", onChange:e=>setForm(f=>({...f,[k]:e.target.value}))});

  // Export
  const exportXLSX = () => {
    const rows = data.map(a=>({
      University:a.university, Country:a.country||"", Program:a.program||"",
      Degree:a.degree||"", Deadline:a.deadline||"", Status:a.status,
      Notes:a.notes||"", Link:a.link||"",
      ...Object.fromEntries((a.docs||[]).map(d=>[d.name,d.done?"✓":""]))
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Applications");
    XLSX.writeFile(wb,"applications.xlsx");
  };

  // Import
  const importXLSX = (e) => {
    const file = e.target.files[0]; if(!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result,{type:"binary"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws);
        const imported = rows.map(r=>({
          id:uid(),
          university:r.University||r.university||"Unknown",
          country:r.Country||r.country||"",
          program:r.Program||r.program||"",
          degree:r.Degree||r.degree||"",
          deadline:r.Deadline||r.deadline||"",
          status:APP_STATUSES.includes(r.Status||r.status)?r.Status||r.status:"researching",
          notes:r.Notes||r.notes||"",
          link:r.Link||r.link||"",
          docs:DEFAULT_DOCS.map((name,i)=>({id:i,name,done:r[name]==="✓"||r[name]===true}))
        }));
        if(window.confirm(`Import ${imported.length} applications? Existing data will be kept.`)){
          setData(prev=>[...prev,...imported]);
        }
      } catch(err){ alert("Error reading XLSX: "+err.message); }
    };
    reader.readAsBinaryString(file);
    e.target.value="";
  };

  return (
    <div>
      <div style={{display:"flex",gap:12,marginBottom:20,flexWrap:"wrap",alignItems:"center"}}>
        <div style={{flex:1,minWidth:200}}><Input placeholder="Search university, program, country…" value={search} onChange={e=>setSearch(e.target.value)} /></div>
        <Select style={{width:140}} value={filterStatus} onChange={e=>setFilterStatus(e.target.value)}>
          <option value="all">All statuses</option>
          {APP_STATUSES.map(s=><option key={s} value={s}>{s}</option>)}
        </Select>
        <Btn onClick={openAdd} variant="primary">+ Add Application</Btn>
        <Btn onClick={exportXLSX} variant="teal">⬇ Export XLSX</Btn>
        <Btn onClick={()=>fileRef.current.click()} variant="ghost">⬆ Import XLSX</Btn>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={importXLSX} />
      </div>
      <div style={{fontSize:12,color:T.ts,marginBottom:12}}>{filtered.length} application{filtered.length!==1?"s":""}</div>

      <div style={{background:T.bg2,border:`1px solid ${T.border}`,borderRadius:12,overflow:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",minWidth:900}}>
          <thead><tr>
            <TH>University</TH><TH>Program</TH><TH>Country</TH><TH>Degree</TH><TH>Deadline</TH><TH>Status</TH><TH>Docs</TH><TH style={{width:110}}>Actions</TH>
          </tr></thead>
          <tbody>
            {filtered.length===0 && <tr><TD colSpan={8}><span style={{color:T.tm}}>No applications found. Add your first one!</span></TD></tr>}
            {filtered.map(a=>{
              const done = (a.docs||[]).filter(d=>d.done).length;
              const total = (a.docs||[]).length;
              const pct = total>0 ? Math.round(done/total*100) : 0;
              return (
                <tr key={a.id} onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.03)"} onMouseLeave={e=>e.currentTarget.style.background=""}>
                  <TD><span style={{fontWeight:600}}>{a.university}</span>{a.link&&<a href={a.link} target="_blank" rel="noreferrer" style={{color:T.blue,fontSize:11,marginLeft:6}}>↗</a>}</TD>
                  <TD><span style={{color:T.ts,fontSize:12}}>{a.program||"—"}</span></TD>
                  <TD><span style={{color:T.ts,fontSize:12}}>{a.country||"—"}</span></TD>
                  <TD><span style={{color:T.ts,fontSize:12}}>{a.degree||"—"}</span></TD>
                  <TD><DaysChip date={a.deadline}/><div style={{fontSize:11,color:T.tm,marginTop:2}}>{fmtDate(a.deadline)}</div></TD>
                  <TD><Badge v={a.status}/></TD>
                  <TD>
                    <button onClick={()=>setDocsModal({appId:a.id})} style={{background:"none",border:"none",cursor:"pointer",textAlign:"left",padding:0}}>
                      <div style={{display:"flex",alignItems:"center",gap:6}}>
                        <div style={{height:6,width:60,background:T.bg3,borderRadius:3,overflow:"hidden"}}>
                          <div style={{height:"100%",width:`${pct}%`,background:pct===100?T.green:pct>50?T.teal:T.accent,borderRadius:3,transition:"width .3s"}} />
                        </div>
                        <span style={{fontSize:11,color:done===total&&total>0?T.green:T.ts}}>{done}/{total}</span>
                      </div>
                    </button>
                  </TD>
                  <TD>
                    <div style={{display:"flex",gap:6}}>
                      <Btn size="sm" variant="ghost" onClick={()=>openEdit(a)}>✏️</Btn>
                      <Btn size="sm" variant="ghost" onClick={()=>setDocsModal({appId:a.id})}>📋</Btn>
                      <Btn size="sm" variant="danger" onClick={()=>setConfirm(a.id)}>🗑</Btn>
                    </div>
                  </TD>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Application form modal */}
      {modal && (
        <Modal title={modal.mode==="add"?"Add Application":"Edit Application"} onClose={()=>setModal(null)} width={660}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            <FormGroup label="University *"><Input {...F("university")} placeholder="e.g. TU Munich" /></FormGroup>
            <FormGroup label="Country"><Input {...F("country")} placeholder="e.g. Germany" /></FormGroup>
            <FormGroup label="Program / Major"><Input {...F("program")} placeholder="e.g. Computer Science" /></FormGroup>
            <FormGroup label="Degree Level">
              <Select {...F("degree")}>
                <option value="">Select…</option>
                {["Bachelor","Master","PhD","Postdoc","Other"].map(d=><option key={d} value={d}>{d}</option>)}
              </Select>
            </FormGroup>
            <FormGroup label="Deadline"><Input type="date" {...F("deadline")} /></FormGroup>
            <FormGroup label="Status">
              <Select {...F("status")}>{APP_STATUSES.map(s=><option key={s} value={s}>{s}</option>)}</Select>
            </FormGroup>
            <FormGroup label="Application Portal Link"><Input {...F("link")} placeholder="https://…" /></FormGroup>
          </div>
          <FormGroup label="Notes"><Textarea {...F("notes")} style={{minHeight:70}} placeholder="Interview dates, contacts, requirements…" /></FormGroup>

          <div style={{marginTop:16}}>
            <Label>Document Checklist</Label>
            <div style={{background:T.bg1,borderRadius:10,padding:14,display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              {(form.docs||[]).map(doc=>(
                <label key={doc.id} style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",fontSize:13,color:doc.done?T.green:T.ts}}>
                  <input type="checkbox" checked={doc.done} onChange={()=>setForm(f=>({...f,docs:f.docs.map(d=>d.id!==doc.id?d:{...d,done:!d.done})}))} style={{accentColor:T.green}} />
                  <span style={{textDecoration:doc.done?"line-through":"none"}}>{doc.name}</span>
                </label>
              ))}
            </div>
            <div style={{display:"flex",gap:8,marginTop:8}}>
              <Input placeholder="Add custom document…" id="customDoc" style={{flex:1}} />
              <Btn size="sm" variant="ghost" onClick={()=>{
                const inp=document.getElementById("customDoc");
                if(inp.value.trim()){
                  setForm(f=>({...f,docs:[...f.docs,{id:Date.now(),name:inp.value.trim(),done:false}]}));
                  inp.value="";
                }
              }}>+ Add</Btn>
            </div>
          </div>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end",paddingTop:16}}>
            <Btn onClick={()=>setModal(null)} variant="ghost">Cancel</Btn>
            <Btn onClick={save} variant="primary">💾 Save Application</Btn>
          </div>
        </Modal>
      )}

      {/* Docs checklist modal */}
      {docsModal && (() => {
        const app = data.find(a=>a.id===docsModal.appId);
        if(!app) return null;
        const done=(app.docs||[]).filter(d=>d.done).length;
        const total=(app.docs||[]).length;
        return (
          <Modal title={`📋 ${app.university} — Documents`} onClose={()=>setDocsModal(null)} width={500}>
            <div style={{marginBottom:14,display:"flex",alignItems:"center",gap:10}}>
              <div style={{flex:1,height:8,background:T.bg3,borderRadius:4,overflow:"hidden"}}>
                <div style={{height:"100%",width:`${total>0?Math.round(done/total*100):0}%`,background:done===total?T.green:T.teal,transition:"width .3s"}} />
              </div>
              <span style={{fontSize:13,fontWeight:600,color:done===total?T.green:T.ts}}>{done}/{total} complete</span>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {(app.docs||[]).map(doc=>(
                <label key={doc.id} style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer",padding:"8px 12px",borderRadius:8,background:doc.done?"rgba(34,197,94,.08)":T.bg3,border:`1px solid ${doc.done?"rgba(34,197,94,.2)":T.border}`,transition:"all .2s"}}>
                  <input type="checkbox" checked={doc.done} onChange={()=>toggleDoc(app.id,doc.id)} style={{accentColor:T.green,width:16,height:16}} />
                  <span style={{fontSize:13,color:doc.done?T.green:T.tp,textDecoration:doc.done?"line-through":"none",flex:1}}>{doc.name}</span>
                  {doc.done && <span style={{fontSize:16}}>✅</span>}
                </label>
              ))}
            </div>
            <div style={{display:"flex",gap:8,marginTop:14}}>
              <Input placeholder="Add a document…" id="newDoc2" style={{flex:1}} />
              <Btn size="sm" variant="ghost" onClick={()=>{
                const inp=document.getElementById("newDoc2");
                if(inp.value.trim()){
                  setData(data.map(a=>a.id!==app.id?a:{...a,docs:[...(a.docs||[]),{id:Date.now(),name:inp.value.trim(),done:false}]}));
                  inp.value="";
                }
              }}>+ Add</Btn>
            </div>
          </Modal>
        );
      })()}

      {confirm && <Confirm message="Delete this application?" onConfirm={()=>del(confirm)} onCancel={()=>setConfirm(null)} />}
    </div>
  );
}

// ─── Email Tracker ────────────────────────────────────────────────────────────
const emptyEmail = () => ({
  id:uid(), professorName:"", university:"", email:"",
  department:"", researchArea:"",
  mailedDate:"", responseDate:"", followUpDate:"", followUpResponseDate:"",
  verdict:"pending", remarks:""
});

const EMAIL_COLS = [
  {key:"professorName",label:"Professor",sortable:true},
  {key:"university",label:"University",sortable:true},
  {key:"email",label:"Email",sortable:false},
  {key:"mailedDate",label:"Mailed",sortable:true},
  {key:"responseDate",label:"Response",sortable:true},
  {key:"followUpDate",label:"Follow-up",sortable:true},
  {key:"followUpResponseDate",label:"FU Response",sortable:true},
  {key:"verdict",label:"Verdict",sortable:true},
  {key:"remarks",label:"Remarks",sortable:false},
];

function EmailsView({data,setData}) {
  const [modal, setModal] = useState(null);
  const [confirm, setConfirm] = useState(null);
  const [form, setForm] = useState({});
  const [search, setSearch] = useState("");
  const [filterVerdict, setFilterVerdict] = useState("all");
  const [sortCol, setSortCol] = useState("mailedDate");
  const [sortDir, setSortDir] = useState(-1);
  const [page, setPage] = useState(1);
  const [perPage, setPerPage] = useState(25);
  const [showFilter, setShowFilter] = useState(false);
  const [filterDateFrom, setFilterDateFrom] = useState("");
  const [filterDateTo, setFilterDateTo] = useState("");
  const fileRef = useRef();

  const filtered = useMemo(()=>{
    const q = search.toLowerCase();
    return data.filter(e=>{
      const mQ = !q || [e.professorName,e.university,e.email,e.department,e.researchArea,e.remarks].some(f=>(f||"").toLowerCase().includes(q));
      const mV = filterVerdict==="all"||(e.verdict||"pending")===filterVerdict;
      const mFrom = !filterDateFrom||e.mailedDate>=filterDateFrom;
      const mTo = !filterDateTo||e.mailedDate<=filterDateTo;
      return mQ&&mV&&mFrom&&mTo;
    });
  },[data,search,filterVerdict,filterDateFrom,filterDateTo]);

  const sorted = useMemo(()=>{
    return [...filtered].sort((a,b)=>{
      const av=a[sortCol]||""; const bv=b[sortCol]||"";
      return av<bv?-sortDir:av>bv?sortDir:0;
    });
  },[filtered,sortCol,sortDir]);

  const totalPages = Math.max(1,Math.ceil(sorted.length/perPage));
  const paginated = useMemo(()=>sorted.slice((page-1)*perPage, page*perPage),[sorted,page,perPage]);

  const sort = (col) => { if(sortCol===col){setSortDir(d=>-d);}else{setSortCol(col);setSortDir(1);} setPage(1); };
  useEffect(()=>setPage(1),[search,filterVerdict,filterDateFrom,filterDateTo,perPage]);

  const openAdd = () => { setForm(emptyEmail()); setModal({mode:"add"}); };
  const openEdit = (e) => { setForm({...e}); setModal({mode:"edit"}); };
  const save = async () => {
    if(!form.professorName.trim()){alert("Professor name required");return;}
    if(modal.mode==="add"){
      await API.create('emails', form);
      setData([...data,{...form}]);
    } else {
      await API.update('emails', form.id, form);
      setData(data.map(e=>e.id===form.id?{...form}:e));
    }
    setModal(null);
  };
  const del = async (id) => { await API.delete('emails', id); setData(data.filter(e=>e.id!==id)); setConfirm(null); };
  const F = (k) => ({value:form[k]||"", onChange:e=>setForm(f=>({...f,[k]:e.target.value}))});

  const exportXLSX = () => {
    const rows = sorted.map(e=>({
      "Professor Name":e.professorName, University:e.university, Email:e.email,
      Department:e.department||"", "Research Area":e.researchArea||"",
      "Mailed Date":e.mailedDate||"", "Response Date":e.responseDate||"",
      "Follow-up Date":e.followUpDate||"", "Follow-up Response":e.followUpResponseDate||"",
      Verdict:e.verdict||"pending", Remarks:e.remarks||""
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Professor Emails");
    XLSX.writeFile(wb,"professor-emails.xlsx");
  };

  const importXLSX = (ev) => {
    const file=ev.target.files[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=(e)=>{
      try{
        const wb=XLSX.read(e.target.result,{type:"binary"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws);
        const imported=rows.map(r=>({
          id:uid(),
          professorName:r["Professor Name"]||r.professorName||r.name||"",
          university:r.University||r.university||"",
          email:r.Email||r.email||"",
          department:r.Department||r.department||"",
          researchArea:r["Research Area"]||r.researchArea||"",
          mailedDate:r["Mailed Date"]||r.mailedDate||"",
          responseDate:r["Response Date"]||r.responseDate||"",
          followUpDate:r["Follow-up Date"]||r.followUpDate||"",
          followUpResponseDate:r["Follow-up Response"]||r.followUpResponseDate||"",
          verdict:VERDICTS.includes(r.Verdict||r.verdict)?r.Verdict||r.verdict:"pending",
          remarks:r.Remarks||r.remarks||""
        })).filter(r=>r.professorName||r.email);
        if(window.confirm(`Import ${imported.length} email records?`)){
          setData(prev=>[...prev,...imported]);
        }
      }catch(err){alert("Error reading XLSX: "+err.message);}
    };
    reader.readAsBinaryString(file);
    ev.target.value="";
  };

  const needsFollowUp = (e) => e.mailedDate && !e.responseDate && daysLeft(e.mailedDate)!==null && daysLeft(e.mailedDate) < -14;

  return (
    <div>
      {/* Toolbar */}
      <div style={{display:"flex",gap:10,marginBottom:12,flexWrap:"wrap",alignItems:"center"}}>
        <div style={{flex:1,minWidth:220,position:"relative"}}>
          <span style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:T.ts,pointerEvents:"none",fontSize:14}}>🔍</span>
          <Input placeholder="Search professor, university, email, remarks…" value={search} onChange={e=>setSearch(e.target.value)} style={{paddingLeft:32}} />
        </div>
        <Select style={{width:140}} value={filterVerdict} onChange={e=>setFilterVerdict(e.target.value)}>
          <option value="all">All verdicts</option>
          {VERDICTS.map(v=><option key={v} value={v}>{v}</option>)}
        </Select>
        <Btn onClick={()=>setShowFilter(f=>!f)} variant={showFilter?"teal":"ghost"}>⚙ Filters</Btn>
        <Btn onClick={openAdd} variant="primary">+ Add Email</Btn>
        <Btn onClick={exportXLSX} variant="teal">⬇ Export</Btn>
        <Btn onClick={()=>fileRef.current.click()} variant="ghost">⬆ Import</Btn>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={importXLSX} />
      </div>

      {showFilter && (
        <div style={{background:T.bg2,border:`1px solid ${T.border}`,borderRadius:10,padding:16,marginBottom:12,display:"flex",gap:14,flexWrap:"wrap",alignItems:"flex-end"}}>
          <div><Label>Mailed From</Label><Input type="date" value={filterDateFrom} onChange={e=>setFilterDateFrom(e.target.value)} style={{width:160}} /></div>
          <div><Label>Mailed To</Label><Input type="date" value={filterDateTo} onChange={e=>setFilterDateTo(e.target.value)} style={{width:160}} /></div>
          <Btn size="sm" variant="ghost" onClick={()=>{setFilterDateFrom("");setFilterDateTo("");}}>Clear dates</Btn>
        </div>
      )}

      {/* Stats strip */}
      <div style={{display:"flex",gap:10,marginBottom:14,flexWrap:"wrap"}}>
        {[
          {label:"Total",val:data.length,color:T.ts},
          {label:"Filtered",val:filtered.length,color:T.accent},
          {label:"Pending",val:data.filter(e=>!e.responseDate).length,color:T.orange},
          {label:"Responded",val:data.filter(e=>e.responseDate).length,color:T.teal},
          {label:"Positive",val:data.filter(e=>["positive","interested"].includes(e.verdict)).length,color:T.green},
          {label:"Need Follow-up",val:data.filter(needsFollowUp).length,color:T.red},
        ].map(s=>(
          <div key={s.label} style={{background:T.bg2,border:`1px solid ${T.border}`,borderRadius:8,padding:"6px 14px",display:"flex",gap:8,alignItems:"center"}}>
            <span style={{fontWeight:700,color:s.color,fontSize:15}}>{s.val}</span>
            <span style={{fontSize:11,color:T.ts}}>{s.label}</span>
          </div>
        ))}
      </div>

      {/* Table */}
      <div style={{background:T.bg2,border:`1px solid ${T.border}`,borderRadius:12,overflow:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",minWidth:1100}}>
          <thead><tr>
            {EMAIL_COLS.map(c=>(
              <TH key={c.key} onClick={c.sortable?()=>sort(c.key):undefined} sorted={sortCol===c.key?sortDir:0}>{c.label}</TH>
            ))}
            <TH style={{width:80}}>Actions</TH>
          </tr></thead>
          <tbody>
            {paginated.length===0 && <tr><TD colSpan={10}><span style={{color:T.tm}}>No emails found. Import or add your first outreach!</span></TD></tr>}
            {paginated.map(e=>(
              <tr key={e.id}
                style={{background:needsFollowUp(e)?"rgba(239,68,68,.04)":""}}
                onMouseEnter={ev=>ev.currentTarget.style.background="rgba(255,255,255,.03)"}
                onMouseLeave={ev=>ev.currentTarget.style.background=needsFollowUp(e)?"rgba(239,68,68,.04)":""}>
                <TD>
                  <span style={{fontWeight:600,color:T.tp}}>{e.professorName}</span>
                  {e.department&&<div style={{fontSize:11,color:T.tm}}>{e.department}</div>}
                </TD>
                <TD><span style={{color:T.ts,fontSize:12}}>{e.university||"—"}</span></TD>
                <TD><a href={`mailto:${e.email}`} style={{color:T.blue,fontSize:12,textDecoration:"none"}}>{e.email||"—"}</a></TD>
                <TD style={{fontSize:12,color:T.ts}}>{fmtDate(e.mailedDate)}</TD>
                <TD>{e.responseDate?<span style={{fontSize:12,color:T.green}}>{fmtDate(e.responseDate)}</span>:<span style={{fontSize:11,color:T.tm}}>—</span>}</TD>
                <TD style={{fontSize:12,color:T.ts}}>{fmtDate(e.followUpDate)}</TD>
                <TD>{e.followUpResponseDate?<span style={{fontSize:12,color:T.green}}>{fmtDate(e.followUpResponseDate)}</span>:<span style={{fontSize:11,color:T.tm}}>—</span>}</TD>
                <TD><Badge v={e.verdict||"pending"}/>{needsFollowUp(e)&&<div style={{fontSize:10,color:T.red,marginTop:3}}>⚠ Follow-up!</div>}</TD>
                <TD><span style={{color:T.ts,fontSize:12,maxWidth:160,display:"block",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={e.remarks}>{e.remarks||"—"}</span></TD>
                <TD>
                  <div style={{display:"flex",gap:5}}>
                    <Btn size="sm" variant="ghost" onClick={()=>openEdit(e)}>✏️</Btn>
                    <Btn size="sm" variant="danger" onClick={()=>setConfirm(e.id)}>🗑</Btn>
                  </div>
                </TD>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Pagination */}
      <div style={{display:"flex",alignItems:"center",gap:12,marginTop:14,flexWrap:"wrap"}}>
        <span style={{fontSize:12,color:T.ts}}>
          {sorted.length===0?"No results":`${(page-1)*perPage+1}–${Math.min(page*perPage,sorted.length)} of ${sorted.length}`}
        </span>
        <div style={{flex:1}} />
        <label style={{fontSize:12,color:T.ts}}>Per page:</label>
        <Select style={{width:80}} value={perPage} onChange={e=>{setPerPage(Number(e.target.value));setPage(1);}}>
          {[10,25,50,100].map(n=><option key={n} value={n}>{n}</option>)}
        </Select>
        <div style={{display:"flex",gap:6}}>
          <Btn size="sm" variant="ghost" onClick={()=>setPage(1)} disabled={page===1}>«</Btn>
          <Btn size="sm" variant="ghost" onClick={()=>setPage(p=>Math.max(1,p-1))} disabled={page===1}>‹</Btn>
          {Array.from({length:Math.min(7,totalPages)},(_,i)=>{
            let p;
            if(totalPages<=7) p=i+1;
            else if(page<=4) p=i+1;
            else if(page>=totalPages-3) p=totalPages-6+i;
            else p=page-3+i;
            return p>=1&&p<=totalPages?(
              <Btn key={p} size="sm" variant={page===p?"primary":"ghost"} onClick={()=>setPage(p)}>{p}</Btn>
            ):null;
          })}
          <Btn size="sm" variant="ghost" onClick={()=>setPage(p=>Math.min(totalPages,p+1))} disabled={page===totalPages}>›</Btn>
          <Btn size="sm" variant="ghost" onClick={()=>setPage(totalPages)} disabled={page===totalPages}>»</Btn>
        </div>
      </div>

      {/* Email Modal */}
      {modal && (
        <Modal title={modal.mode==="add"?"Add Professor Email":"Edit Professor Email"} onClose={()=>setModal(null)} width={680}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            <FormGroup label="Professor Name *"><Input {...F("professorName")} placeholder="Dr. Jane Smith" /></FormGroup>
            <FormGroup label="University"><Input {...F("university")} placeholder="MIT" /></FormGroup>
            <FormGroup label="Email Address"><Input type="email" {...F("email")} placeholder="jsmith@mit.edu" /></FormGroup>
            <FormGroup label="Department"><Input {...F("department")} placeholder="Computer Science" /></FormGroup>
            <FormGroup label="Research Area"><Input {...F("researchArea")} placeholder="Machine Learning, NLP…" /></FormGroup>
            <FormGroup label="Verdict">
              <Select {...F("verdict")}>{VERDICTS.map(v=><option key={v} value={v}>{v}</option>)}</Select>
            </FormGroup>
            <FormGroup label="Mailed Date"><Input type="date" {...F("mailedDate")} /></FormGroup>
            <FormGroup label="Response Date"><Input type="date" {...F("responseDate")} /></FormGroup>
            <FormGroup label="Follow-up Mail Date"><Input type="date" {...F("followUpDate")} /></FormGroup>
            <FormGroup label="Follow-up Response Date"><Input type="date" {...F("followUpResponseDate")} /></FormGroup>
          </div>
          <FormGroup label="Remarks"><Textarea {...F("remarks")} placeholder="Notes about the response, interest level, next steps…" /></FormGroup>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end",paddingTop:8}}>
            <Btn onClick={()=>setModal(null)} variant="ghost">Cancel</Btn>
            <Btn onClick={save} variant="primary">💾 Save</Btn>
          </div>
        </Modal>
      )}
      {confirm && <Confirm message="Delete this email record?" onConfirm={()=>del(confirm)} onCancel={()=>setConfirm(null)} />}
    </div>
  );
}

// ─── App Root ─────────────────────────────────────────────────────────────────
const NAV = [
  {id:"dashboard",label:"Dashboard",icon:"🏠"},
  {id:"scholarships",label:"Scholarships",icon:"🎓"},
  {id:"applications",label:"Applications",icon:"📋"},
  {id:"emails",label:"Professor Emails",icon:"📧"},
];

export default function App() {
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("dashboard");
  const [scholarships, setScholarshipsRaw] = useState([]);
  const [applications, setApplicationsRaw] = useState([]);
  const [emails, setEmailsRaw] = useState([]);
  const [sideCollapsed, setSideCollapsed] = useState(false);
  const [theme, setTheme] = useState(() => localStorage.getItem("theme") || "system");

  useEffect(() => {
    localStorage.setItem("theme", theme);
    const root = document.documentElement;
    if (theme === "system") {
      const isDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
      root.setAttribute("data-theme", isDark ? "dark" : "light");
    } else {
      root.setAttribute("data-theme", theme);
    }
  }, [theme]);

  useEffect(()=>{
    (async()=>{
      const [s,a,e] = await Promise.all([API.get('scholarships'),API.get('applications'),API.get('emails')]);
      setScholarshipsRaw(s||[]); setApplicationsRaw(a||[]); setEmailsRaw(e||[]);
      setLoading(false);
    })();
  },[]);

  const setScholarships = useCallback((v)=>{ const next=typeof v==="function"?v(scholarships):v; setScholarshipsRaw(next); },[scholarships]);
  const setApplications = useCallback((v)=>{ const next=typeof v==="function"?v(applications):v; setApplicationsRaw(next); },[applications]);
  const setEmails = useCallback((v)=>{ const next=typeof v==="function"?v(emails):v; setEmailsRaw(next); },[emails]);

  const activeTab = NAV.find(n=>n.id===tab);

  if(loading) return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:T.bg0,color:T.ts,fontFamily:"system-ui",fontSize:14}}>
      <div style={{textAlign:"center"}}>
        <div style={{fontSize:40,marginBottom:16,animation:"spin 2s linear infinite"}}>🎓</div>
        <div>Loading your EduTracker…</div>
        <style>{`@keyframes spin{from{transform:rotate(0)}to{transform:rotate(360deg)}}`}</style>
      </div>
    </div>
  );

  return (
    <div style={{display:"flex",height:"100vh",fontFamily:"'Segoe UI',system-ui,sans-serif",background:T.bg0,color:T.tp,overflow:"hidden"}}>
      {/* Sidebar */}
      <div style={{width:sideCollapsed?60:220,background:T.bg1,borderRight:`1px solid ${T.border}`,display:"flex",flexDirection:"column",transition:"width .25s",flexShrink:0,overflow:"hidden"}}>
        <div style={{padding:"18px 16px",borderBottom:`1px solid ${T.border}`,display:"flex",alignItems:"center",gap:10,minWidth:220}}>
          <span style={{fontSize:22}}>🎓</span>
          {!sideCollapsed&&<div>
            <div style={{fontWeight:800,fontSize:14,color:T.tp,letterSpacing:.5}}>EduTracker</div>
            <div style={{fontSize:10,color:T.tm}}>My Dream Hub</div>
          </div>}
          <button onClick={()=>setSideCollapsed(s=>!s)} style={{marginLeft:"auto",background:"none",border:"none",color:T.ts,cursor:"pointer",fontSize:16,padding:4}}>☰</button>
        </div>
        <nav style={{flex:1,padding:"10px 8px"}}>
          {NAV.map(n=>{
            const active=tab===n.id;
            const cnt = n.id==="scholarships"?scholarships.length:n.id==="applications"?applications.length:n.id==="emails"?emails.length:null;
            return (
              <button key={n.id} onClick={()=>setTab(n.id)} style={{
                width:"100%",display:"flex",alignItems:"center",gap:10,padding:"10px 10px",
                background:active?T.accentDim:"none",
                border:"none",borderRadius:8,cursor:"pointer",textAlign:"left",
                color:active?T.accent:T.ts,fontWeight:active?700:400,fontSize:13,
                transition:"all .15s",marginBottom:2,minWidth:0,
              }}>
                <span style={{fontSize:17,flexShrink:0}}>{n.icon}</span>
                {!sideCollapsed&&<>
                  <span style={{flex:1,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{n.label}</span>
                  {cnt!==null&&<span style={{background:active?T.accentDim:T.bg3,color:active?T.accent:T.tm,fontSize:10,fontWeight:700,padding:"1px 7px",borderRadius:20,flexShrink:0}}>{cnt}</span>}
                </>}
              </button>
            );
          })}
        </nav>
        {!sideCollapsed&&<div style={{padding:"12px 16px",borderTop:`1px solid ${T.border}`,fontSize:10,color:T.tm}}>
          Data saved locally in your computer.
        </div>}
      </div>

      {/* Main */}
      <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"}}>
        <header style={{padding:"0 24px",borderBottom:`1px solid ${T.border}`,height:52,display:"flex",alignItems:"center",justifyContent:"space-between",background:T.bg1,flexShrink:0}}>
          <div>
            <span style={{fontWeight:700,fontSize:15,color:T.tp}}>{activeTab?.icon} {activeTab?.label}</span>
          </div>
          <div style={{display:"flex",gap:12,alignItems:"center"}}>
            <Select value={theme} onChange={e=>setTheme(e.target.value)} style={{width:110, padding:"4px 8px", fontSize:12}}>
              <option value="system">💻 System</option>
              <option value="light">☀️ Light</option>
              <option value="dark">🌙 Dark</option>
            </Select>
            <span style={{fontSize:11,color:T.tm,whiteSpace:"nowrap"}}>{new Date().toLocaleDateString("en-GB",{weekday:"long",day:"2-digit",month:"short",year:"numeric"})}</span>
          </div>
        </header>
        <main style={{flex:1,overflow:"auto",padding:24}}>
          {tab==="dashboard"&&<Dashboard scholarships={scholarships} applications={applications} emails={emails} setTab={setTab}/>}
          {tab==="scholarships"&&<ScholarshipsView data={scholarships} setData={setScholarships}/>}
          {tab==="applications"&&<ApplicationsView data={applications} setData={setApplications}/>}
          {tab==="emails"&&<EmailsView data={emails} setData={setEmails}/>}
        </main>
      </div>
    </div>
  );
}
