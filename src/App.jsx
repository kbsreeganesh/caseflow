import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// ═══════════════════════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════════════════════
const LIMIT = 120;

const SK = {
  cfg:    "cf9_cfg",
  cases:  "cf9_cases",
  users:  "cf9_users",
  log:    "cf9_log",
  uidreg: "cf9_uidreg",
  caseArch: ts => `cf9_cases_arch_${ts}`,
  userArch: ts => `cf9_users_arch_${ts}`,
  archIdx:  "cf9_arch_index",   // list of {key, ts, type, label}
};

const DEFAULT_FORMS = {
  analyser:      "https://forms.gle/PrAjh7xVL8HGGdq37",
  supervisor:    "https://forms.gle/6S9PzhsRHs9mnmYn8",
  preauthoriser: "",
};

const ROLES = {
  analyser:      { label:"Analyser",          color:"#0891B2", bg:"#E0F7FA", icon:"🔬" },
  supervisor:    { label:"Supervisor",         color:"#7C3AED", bg:"#EDE9FE", icon:"👨‍⚕️" },
  preauthoriser: { label:"Preauthoriser",      color:"#0369A1", bg:"#DBEAFE", icon:"📋" },
  upload:        { label:"Data Upload",        color:"#D97706", bg:"#FEF3C7", icon:"📤" },
  mis:           { label:"MIS",                color:"#065F46", bg:"#D1FAE5", icon:"📊" },
  state_manager: { label:"State Manager",      color:"#4C1D95", bg:"#EDE9FE", icon:"🗺️" },
  viewer:        { label:"Viewer",             color:"#475569", bg:"#F1F5F9", icon:"👁️"  },
  credential:    { label:"Credential Manager", color:"#059669", bg:"#D1FAE5", icon:"🔑" },
  admin:         { label:"Super Admin",        color:"#DC2626", bg:"#FEE2E2", icon:"⚙️"  },
};

// ═══════════════════════════════════════════════════════════
// PALETTE & STYLES
// ═══════════════════════════════════════════════════════════
const C = {
  bg:"#EDF2F7", surf:"#FFFFFF", navy:"#0F2030", navyM:"#1A3348",
  teal:"#0891B2", green:"#059669", amber:"#D97706", red:"#DC2626", purp:"#7C3AED",
  bdr:"#DDE5EE", txt:"#1A2535", muted:"#637187", lt:"#F0F5FA",
};

const G = `
  @import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700;800&family=DM+Sans:wght@400;500;600&display=swap');
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:'DM Sans',sans-serif;background:${C.bg};}
  @keyframes spin{to{transform:rotate(360deg);}}
  @keyframes fadeUp{from{opacity:0;transform:translateY(10px);}to{opacity:1;transform:none;}}
  .up{animation:fadeUp .28s ease forwards;}
  input:focus,select:focus{outline:2.5px solid ${C.teal};outline-offset:-1px;}
  ::-webkit-scrollbar{width:5px;height:5px;}
  ::-webkit-scrollbar-thumb{background:${C.bdr};border-radius:4px;}
  button{transition:opacity .15s,transform .1s;}
  button:hover{opacity:.86;}
  button:active{transform:scale(.97);}
  input[type=radio]{accent-color:${C.green};cursor:pointer;}
  table{border-collapse:collapse;}
`;

const $ = {
  btn:(bg=C.navy,fg="#fff")=>({background:bg,color:fg,border:"none",borderRadius:9,
    padding:"10px 20px",fontSize:13,fontWeight:600,cursor:"pointer",
    fontFamily:"'DM Sans',sans-serif",letterSpacing:"0.01em"}),
  inp:{width:"100%",border:`1.5px solid ${C.bdr}`,borderRadius:9,padding:"10px 13px",
    fontSize:14,fontFamily:"'DM Sans',sans-serif",color:C.txt,background:C.surf},
  card:{background:C.surf,borderRadius:14,padding:22,
    boxShadow:"0 1px 3px rgba(0,0,0,.05),0 6px 20px rgba(0,0,0,.06)",border:`1px solid ${C.bdr}`},
  lbl:{fontSize:11,fontWeight:700,color:C.muted,textTransform:"uppercase",
    letterSpacing:"0.08em",display:"block",marginBottom:6},
  tag:(bg,fg)=>({background:bg,color:fg,padding:"3px 10px",borderRadius:6,
    fontSize:11,fontWeight:700,letterSpacing:"0.06em",display:"inline-block"}),
};

// ═══════════════════════════════════════════════════════════
// UTILITY
// ═══════════════════════════════════════════════════════════
const norm = s => String(s||"").trim().toLowerCase();

function detectRole(val) {
  const v = norm(val);
  if (v.includes("supervisor")) return "supervisor";
  if (v.includes("analys"))     return "analyser";
  return null;
}

function parseXLSXBuf(ab) {
  const wb = XLSX.read(ab,{type:"array"});
  const out = {};
  wb.SheetNames.forEach(name=>{
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:""});
    if (rows.length<2) return;
    const headers = rows[0].map(h=>String(h).trim());
    const data    = rows.slice(1)
      .filter(r=>r.some(c=>String(c).trim()))
      .map(r=>Object.fromEntries(headers.map((h,i)=>[h,String(r[i]??"").trim()])));
    out[name]={headers,data};
  });
  return out;
}

function gsExportUrl(url) {
  const m=url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m?`https://docs.google.com/spreadsheets/d/${m[1]}/export?format=xlsx`:null;
}

function fmtEmbed(url) {
  if (!url) return "";
  if (url.includes("forms.gle")) return url;
  const base=url.split("?")[0].replace(/\/(edit|viewform)$/,"");
  return base+"/viewform?embedded=true";
}

function todayStr() { return new Date().toLocaleDateString("en-IN"); }

function randPin() { return String(1000+Math.floor(Math.random()*9000)); }

// Global UID registry: first UID where global count < LIMIT
function findActiveUID(userEntry, uidReg) {
  if (!userEntry?.userIds?.length) return null;
  return [...userEntry.userIds]
    .sort((a,b)=>(a.seq||0)-(b.seq||0))
    .find(u=>u.active!==false && (uidReg?.[u.uid]||0)<LIMIT) || null;
}

// Case-insensitive queue build
function buildQueue(cases, doctorName, role) {
  const dn = norm(doctorName);
  return cases
    .filter(c=>norm(c.doctorName)===dn && c.role===role && c.status==="pending")
    .sort((a,b)=>(a.seq||0)-(b.seq||0));
}

// Auto-detect column layout from headers/values
// Handles "Anlayser" misspelling and variants
function isAnalyserHeader(h) {
  const s=norm(h).replace(/\s+/g,"");
  return s==="analyser"||s==="analyzer"||s==="analysers"||s==="anlayser"||s==="anlaysers"||s.includes("anlays");
}
function isSupervisorHeader(h) {
  const s=norm(h).replace(/\s+/g,"");
  return s==="supervisor"||s==="supervisors";
}

const COMPLETED_STATUSES = ["query","approve","approved","already processed","reject","rejected","partially rejected","revert to analyser","revert to analyzer","hold","closed","done"];
function isCompletedStatus(val) {
  const v=norm(val||"").trim(); if (!v) return false;
  return COMPLETED_STATUSES.some(s=>v===s||v.startsWith(s));
}

// Auto-detect column layout — cases always imported as pending (status comes from form response sync)
function autoDetectCols(parsed) {
  for (const {headers} of Object.values(parsed)) {
    let ac=null, sc=null;
    for (const h of headers) {
      if (!ac && isAnalyserHeader(h))   ac=h;
      if (!sc && isSupervisorHeader(h)) sc=h;
    }
    if (ac||sc) return {mode:"separate",analyserCol:ac||"",supervisorCol:sc||""};
  }
  for (const {headers,data} of Object.values(parsed)) {
    for (const h of headers) {
      if (data.slice(0,25).some(r=>detectRole(r[h]||"")))
        return {mode:"single",roleCol:h};
    }
  }
  return {mode:"single",roleCol:""};
}

// Fetch a Google Sheet (response sheet) as CSV
function sheetToCsvUrl(url) {
  const m=url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (!m) return null;
  const gid=url.match(/[?&]gid=(\d+)/)?.[1]||"0";
  return `https://docs.google.com/spreadsheets/d/${m[1]}/export?format=csv&gid=${gid}`;
}

function parseCSV(text) {
  const rows=[]; let cur=[], field="", inQ=false;
  for (let i=0;i<text.length;i++){
    const c=text[i];
    if (inQ){
      if(c==='"'&&text[i+1]==='"'){field+='"';i++;}
      else if(c==='"'){inQ=false;}
      else{field+=c;}
    } else if(c==='"'){inQ=true;}
    else if(c===','){cur.push(field);field="";}
    else if(c==='\n'||c==='\r'){
      if(c==='\r'&&text[i+1]==='\n')i++;
      cur.push(field);rows.push(cur);cur=[];field="";
    } else{field+=c;}
  }
  if(field||cur.length){cur.push(field);rows.push(cur);}
  if(!rows.length)return[];
  const hdrs=rows[0].map(h=>h.trim());
  return rows.slice(1).filter(r=>r.some(c=>c.trim())).map(r=>Object.fromEntries(hdrs.map((h,i)=>[h,String(r[i]??"").trim()])));
}

// Parse User Master Excel:
// Columns: Name | Role | PIN | User ID 1 | Password 1 | User ID 2 | Password 2 | ...
function parseUserMaster(ab) {
  const wb  = XLSX.read(ab,{type:"array"});
  const ws  = wb.Sheets[wb.SheetNames[0]];
  const rows= XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
  if (rows.length<2) return [];
  const hdrs = rows[0].map(h=>String(h).trim().toLowerCase());

  const nameIdx = hdrs.findIndex(h=>h==="name"||h==="doctor name"||h==="doctor");
  const roleIdx = hdrs.findIndex(h=>h==="role");
  const pinIdx  = hdrs.findIndex(h=>h==="pin");

  const users=[];
  rows.slice(1).forEach(row=>{
    const name = String(row[nameIdx]||"").trim(); if (!name) return;
    const role = detectRole(row[roleIdx]||"") || norm(row[roleIdx]||"") || "analyser";
    const pin  = String(row[pinIdx]||"").trim() || randPin();

    // Collect User ID / Password pairs dynamically
    const userIds=[];
    hdrs.forEach((h,i)=>{
      const m=h.match(/^user\s*id\s*(\d+)$/);
      if (!m) return;
      const seq    = parseInt(m[1]);
      const uid    = String(row[i]||"").trim(); if (!uid) return;
      // find matching password column: "password N" or "pwd N" or "pass N"
      const pi = hdrs.findIndex((ph,pi2)=>pi2>i&&ph.match(new RegExp(`(password|pwd|pass)\\s*${seq}$`)));
      const pwd = pi>=0?String(row[pi]||"").trim():"";
      userIds.push({uid,password:pwd,active:true,seq});
    });

    users.push({name,pin,role,userIds});
  });
  return users;
}

// ═══════════════════════════════════════════════════════════
// SMALL UI COMPONENTS
// ═══════════════════════════════════════════════════════════
const Spinner=({size=36,color=C.teal})=>(
  <div style={{width:size,height:size,border:`3.5px solid ${C.bdr}`,borderTopColor:color,borderRadius:"50%",animation:"spin .8s linear infinite"}}/>
);

const StatCard=({icon,value,label,color=C.navy})=>(
  <div style={{...$.card,padding:"16px 20px"}}>
    <div style={{fontSize:20,marginBottom:8}}>{icon}</div>
    <div style={{fontSize:26,fontWeight:800,color,fontFamily:"'Sora',sans-serif",lineHeight:1}}>{value}</div>
    <div style={{fontSize:12,color:C.muted,marginTop:4}}>{label}</div>
  </div>
);

const PBar=({pct,color=C.teal,h=6})=>(
  <div style={{height:h,background:C.lt,borderRadius:4,overflow:"hidden"}}>
    <div style={{height:"100%",width:`${Math.min(100,pct)}%`,background:pct>=100?C.green:color,borderRadius:4,transition:"width .5s"}}/>
  </div>
);

const IBox=({children,type="info"})=>{
  const s={info:{bg:"#EFF6FF",bdr:"#BFDBFE",c:"#1E40AF"},warn:{bg:"#FFF7ED",bdr:"#FED7AA",c:"#92400E"},ok:{bg:"#F0FDF4",bdr:"#BBF7D0",c:"#166534"}}[type];
  return <div style={{background:s.bg,border:`1px solid ${s.bdr}`,borderRadius:10,padding:"12px 16px",marginTop:12}}><p style={{fontSize:13,color:s.c,margin:0,lineHeight:1.65}}>{children}</p></div>;
};

const TopBar=({subtitle,right})=>(
  <div style={{background:C.navy,padding:"0 20px",height:54,display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
    <div style={{display:"flex",alignItems:"center",gap:10}}>
      <span style={{color:"#fff",fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:18}}>⚕ CaseFlow</span>
      {subtitle&&<span style={{background:"rgba(255,255,255,.12)",color:"rgba(255,255,255,.75)",padding:"2px 10px",borderRadius:6,fontSize:11,fontWeight:700}}>{subtitle}</span>}
    </div>
    {right}
  </div>
);

const UploadBox=({label,icon,accept,onFile,inputRef,loading})=>(
  <div style={{border:`2px dashed ${C.bdr}`,borderRadius:12,padding:"32px 20px",textAlign:"center",background:C.lt}}>
    <div style={{fontSize:36,marginBottom:10}}>{icon}</div>
    <div style={{fontWeight:700,color:C.navy,fontSize:14,marginBottom:4}}>{label}</div>
    {loading
      ? <div style={{marginTop:14,display:"flex",justifyContent:"center"}}><Spinner size={26}/></div>
      : <>
          <input ref={inputRef} type="file" accept={accept} style={{display:"none"}}
            onChange={e=>{const f=e.target.files[0];if(f)onFile(f);e.target.value="";}}/>
          <button style={{...$.btn(C.teal),marginTop:12,padding:"10px 26px",fontSize:13}}
            onClick={()=>inputRef.current?.click()}>📂 Choose File</button>
        </>
    }
  </div>
);

// ═══════════════════════════════════════════════════════════
// USER FORM (Add / Edit)
// ═══════════════════════════════════════════════════════════
function UserForm({initial,onSave,onCancel,existingUsers}) {
  const [name,    setName]    = useState(initial?.name||"");
  const [pin,     setPin]     = useState(initial?.pin ||"");
  const [role,    setRole]    = useState(initial?.role||"analyser");
  const [userIds, setUserIds] = useState(
    initial?.userIds?.length ? initial.userIds
    : ["analyser","supervisor"].includes(initial?.role||"analyser")
      ? [{uid:"",password:"",active:true,seq:1}] : []
  );
  const [err,setErr]=useState("");

  const addUID=()=>setUserIds(ids=>[...ids,{uid:"",password:"",active:true,seq:ids.length+1}]);
  const updUID=(i,f,v)=>setUserIds(ids=>ids.map((x,j)=>j===i?{...x,[f]:v}:x));
  const remUID=i=>setUserIds(ids=>ids.filter((_,j)=>j!==i).map((x,j)=>({...x,seq:j+1})));
  const isDr=["analyser","supervisor","preauthoriser"].includes(role);

  const save=()=>{
    if (!name.trim()){setErr("Name is required.");return;}
    if (!pin.trim()) {setErr("PIN is required."); return;}
    const dup=existingUsers?.find(u=>norm(u.name)===norm(name)&&u.name!==initial?.name);
    if (dup){setErr("Name already exists.");return;}
    // Only require UID to be filled — password is optional
    if (isDr) for (const u of userIds) if (!u.uid.trim()){setErr("All User IDs must be filled.");return;}
    onSave({name:name.trim(),pin:pin.trim(),role,userIds:isDr?userIds:[]});
  };

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.6)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999,padding:20,overflowY:"auto"}}>
      <div style={{...$.card,maxWidth:560,width:"100%",maxHeight:"90vh",overflowY:"auto"}} className="up">
        <h3 style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:18,marginBottom:20}}>{initial?"Edit":"Add"} User</h3>

        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
          <div><label style={$.lbl}>Full Name</label>
            <input style={$.inp} value={name} onChange={e=>setName(e.target.value)} placeholder="Dr. Name"/></div>
          <div><label style={$.lbl}>PIN (4–6 digits)</label>
            <input style={{...$.inp,textAlign:"center",letterSpacing:4,fontWeight:700}} value={pin} maxLength={6}
              onChange={e=>setPin(e.target.value)} placeholder="1234"/></div>
        </div>

        <div style={{marginBottom:20}}>
          <label style={$.lbl}>Role</label>
          <select style={$.inp} value={role}
            onChange={e=>{setRole(e.target.value);if(!["analyser","supervisor"].includes(e.target.value))setUserIds([]);}}>
            {Object.entries(ROLES).map(([k,v])=><option key={k} value={k}>{v.icon} {v.label}</option>)}
          </select>
        </div>

        {isDr&&(
          <div style={{marginBottom:20}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
              <div>
                <label style={$.lbl}>User IDs (each = {LIMIT} cases globally)</label>
                <p style={{fontSize:12,color:C.muted,marginTop:-4}}>Passwords are optional — add only if you want CaseFlow to display them to the doctor.</p>
              </div>
              <button onClick={addUID} style={{...$.btn(C.teal),padding:"5px 12px",fontSize:12}}>+ Add ID</button>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {userIds.map((uid,i)=>(
                <div key={i} style={{display:"grid",gridTemplateColumns:"26px 1fr 1fr auto",gap:8,alignItems:"center",
                  padding:"10px 12px",background:C.lt,borderRadius:9,border:`1px solid ${C.bdr}`}}>
                  <div style={{width:24,height:24,background:C.navy,color:"#fff",borderRadius:"50%",
                    display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700}}>{i+1}</div>
                  <input style={{...$.inp,padding:"7px 10px",fontSize:13}} placeholder={`User ID e.g. ANLSR00${i+1}`}
                    value={uid.uid} onChange={e=>updUID(i,"uid",e.target.value)}/>
                  <input style={{...$.inp,padding:"7px 10px",fontSize:13}} placeholder="Password (optional)"
                    value={uid.password||""} onChange={e=>updUID(i,"password",e.target.value)}/>
                  <button onClick={()=>remUID(i)} style={{...$.btn("#FEE2E2",C.red),padding:"6px 10px",fontSize:12}}>✕</button>
                </div>
              ))}
              {!userIds.length&&<div style={{textAlign:"center",padding:16,color:C.muted,fontSize:13,background:C.lt,borderRadius:9}}>No IDs yet. Click + Add ID.</div>}
            </div>
          </div>
        )}

        {err&&<p style={{color:C.red,fontSize:13,marginBottom:12}}>⚠ {err}</p>}
        <div style={{display:"flex",gap:10}}>
          <button onClick={onCancel} style={{...$.btn(C.lt,C.muted),border:`1.5px solid ${C.bdr}`,flex:1}}>Cancel</button>
          <button onClick={save}     style={{...$.btn(C.green),flex:1}}>{initial?"Save":"Add User"}</button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════════
export default function App() {
  // ── Core data ───────────────────────────────────────────
  const [sc,     setSc]     = useState("boot");
  const [cfg,    setCfg]    = useState(null);
  const [cases,  setCases]  = useState([]);
  const [users,  setUsers]  = useState([]);
  const [log,    setLog]    = useState([]);
  const [uidReg, setUidReg] = useState({});

  // ── Session ─────────────────────────────────────────────
  const [session,   setSession]   = useState(null);
  const [showCreds, setShowCreds] = useState(false);
  const [confirmed, setConfirmed] = useState(false);

  // ── Login ────────────────────────────────────────────────
  const [lMode,setLMode]=useState("pin");
  const [pinV, setPinV] =useState("");
  const [uidV, setUidV] =useState("");
  const [pwV,  setPwV]  =useState("");
  const [lErr, setLErr] =useState("");

  // ── Cases ────────────────────────────────────────────────
  const [queue,setQueue]=useState([]);
  const [fKey, setFKey] =useState(0);

  // ── Admin ────────────────────────────────────────────────
  const [tab,      setTab]      =useState("dash");
  const [adminPw,  setAdminPw]  =useState("");
  const [adminErr, setAdminErr] =useState("");

  // ── Import wizard ────────────────────────────────────────
  const [wStep,       setWStep]       =useState(1);
  const [iSrc,        setISrc]        =useState("xlsx");
  const [gsUrl,       setGsUrl]       =useState("");
  const [sheets,      setSheets]      =useState(null);
  const [allH,        setAllH]        =useState([]);
  const [importMode,  setImportMode]  =useState("single");
  const [analyserCol, setAnalyserCol] =useState("");
  const [supvCol,     setSupvCol]     =useState("");
  const [roleCol,     setRoleCol]     =useState("");
  const [nameCol,     setNameCol]     =useState("");
  const [detCols,     setDetCols]     =useState([]);
  const [autoPins,    setAutoPins]    =useState({});  // doctorName → PIN
  const [iErr,        setIErr]        =useState("");
  const [iLoad,       setILoad]       =useState(false);

  // ── User master upload ───────────────────────────────────
  const [umLoad,  setUmLoad]  =useState(false);
  const [umMsg,   setUmMsg]   =useState("");
  const [umErr,   setUmErr]   =useState("");

  // ── User management ──────────────────────────────────────
  const [editUser,setEditUser]=useState(null);
  const [addUser, setAddUser] =useState(false);

  // ── Settings ─────────────────────────────────────────────
  const [newForms,      setNewForms]      =useState({analyser:"",supervisor:"",preauthoriser:""});
  const [newDataUrls,   setNewDataUrls]   =useState({analyser:"",supervisor:"",preauthoriser:""});
  const [newAdminPw,    setNewAdminPw]    =useState("");

  // Sync state
  const [syncLog,    setSyncLog]    =useState([]);
  const [syncing,    setSyncing]    =useState(false);
  const [syncMsg,    setSyncMsg]    =useState("");
  const [syncCfgOpen,setSyncCfgOpen]=useState(false);

  // Archive state
  const [archIdx,   setArchIdx]    =useState([]);
  const [archOpen,  setArchOpen]   =useState(false);

  // ── File input refs ──────────────────────────────────────
  const caseFileRef = useRef(null);
  const umFileRef   = useRef(null);

  // ═══════════════════════════════════════════════════════
  // BOOT
  // ═══════════════════════════════════════════════════════
  useEffect(()=>{
    (async()=>{
      try {
        const [cr,csr,ur,lr,rr,ar]=await Promise.all([
          window.storage.get(SK.cfg,     true).catch(()=>null),
          window.storage.get(SK.cases,   true).catch(()=>null),
          window.storage.get(SK.users,   true).catch(()=>null),
          window.storage.get(SK.log,     true).catch(()=>null),
          window.storage.get(SK.uidreg,  true).catch(()=>null),
          window.storage.get(SK.archIdx, true).catch(()=>null),
        ]);
        const lCfg    =cr?.value  ?JSON.parse(cr.value)  :null;
        const lCases  =csr?.value ?JSON.parse(csr.value) :[];
        const lUsers  =ur?.value  ?JSON.parse(ur.value)  :[];
        const lLog    =lr?.value  ?JSON.parse(lr.value)  :[];
        const lUidReg =rr?.value  ?JSON.parse(rr.value)  :{};
        const lArch   =ar?.value  ?JSON.parse(ar.value)  :[];
        setCfg(lCfg);setCases(lCases);setUsers(lUsers);setLog(lLog);setUidReg(lUidReg);setArchIdx(lArch);
        if (!lCfg){
          const dc={formUrls:DEFAULT_FORMS,formDataUrls:{analyser:"",supervisor:"",preauthoriser:""},syncCols:{caseKey:"",formCaseKey:"",formUidCol:""},adminPass:"admin123",detailCols:[],importedAt:null};
          await window.storage.set(SK.cfg,JSON.stringify(dc),true);
          setCfg(dc);setSc("admin");setTab("import");
        } else { setSc("login"); }
      } catch {
        setCfg({formUrls:DEFAULT_FORMS,adminPass:"admin123",detailCols:[]});setSc("login");
      }
    })();
  },[]);

  // ── Persist ──────────────────────────────────────────────
  const saveCfg   =async nc=>{await window.storage.set(SK.cfg,   JSON.stringify(nc),true);setCfg(nc);};
  const saveCases =async nc=>{await window.storage.set(SK.cases, JSON.stringify(nc),true);setCases(nc);};
  const saveUsers =async nu=>{await window.storage.set(SK.users, JSON.stringify(nu),true);setUsers(nu);};
  const saveLog   =async nl=>{await window.storage.set(SK.log,   JSON.stringify(nl),true);setLog(nl);};
  const saveUidReg=async nr=>{await window.storage.set(SK.uidreg,JSON.stringify(nr),true);setUidReg(nr);};

  // Archive old data then replace — keeps history for reference
  const archiveThenSaveCases=async(newData,label)=>{
    if (cases.length>0){
      const ts=Date.now();
      const key=SK.caseArch(ts);
      await window.storage.set(key,JSON.stringify(cases),true).catch(()=>{});
      const newIdx=[...archIdx,{key,ts,type:"cases",label:label||`Cases ${new Date(ts).toLocaleString("en-IN")}`,count:cases.length}];
      await window.storage.set(SK.archIdx,JSON.stringify(newIdx),true).catch(()=>{});
      setArchIdx(newIdx);
    }
    await saveCases(newData);
  };

  const archiveThenSaveUsers=async(newData,label)=>{
    if (users.length>0){
      const ts=Date.now();
      const key=SK.userArch(ts);
      await window.storage.set(key,JSON.stringify(users),true).catch(()=>{});
      const newIdx=[...archIdx,{key,ts,type:"users",label:label||`User Master ${new Date(ts).toLocaleString("en-IN")}`,count:users.length}];
      await window.storage.set(SK.archIdx,JSON.stringify(newIdx),true).catch(()=>{});
      setArchIdx(newIdx);
    }
    await saveUsers(newData);
  };

  // ═══════════════════════════════════════════════════════
  // LOGIN
  // ═══════════════════════════════════════════════════════
  const doLogin=async()=>{
    setLErr("");
    let userEntry=null,activeUID=null;
    if (lMode==="pin") {
      if (pinV.trim()===cfg?.adminPass){setSession({role:"admin",isAdmin:true});setSc("admin");setTab("dash");return;}
      userEntry=users.find(u=>u.pin===pinV.trim());
      if (!userEntry){setLErr("PIN not recognised.");return;}
      activeUID=findActiveUID(userEntry,uidReg);
    } else {
      for (const u of users){
        const uid=u.userIds?.find(x=>x.uid===uidV.trim()&&x.password===pwV.trim()&&x.active!==false);
        if (uid){userEntry=u;activeUID=uid;break;}
      }
      if (!userEntry){setLErr("Invalid User ID or password.");return;}
    }
    const role=userEntry.role;
    if (role==="admin")        {setSession({userEntry,role,isAdmin:true});setSc("admin");setTab("dash");   return;}
    if (role==="upload")       {setSession({userEntry,role});             setSc("admin");setTab("import"); return;}
    if (role==="mis")          {setSession({userEntry,role});             setSc("admin");setTab("import"); return;}
    if (role==="state_manager"){setSession({userEntry,role});             setSc("admin");setTab("import"); return;}
    if (role==="credential")   {setSession({userEntry,role});             setSc("admin");setTab("users");  return;}
    if (role==="viewer")       {setSession({userEntry,role});             setSc("admin");setTab("report"); return;}

    // Preauthoriser — just show the form, no case allocation
    if (role==="preauthoriser"){
      setSession({userEntry,activeUID:null,role,doctorName:userEntry.name});
      setSc("preauth"); return;
    }

    if (!activeUID){setLErr("All your User IDs have reached the 120-case limit. Contact your coordinator.");return;}

    const q=buildQueue(cases,userEntry.name,role);
    setSession({userEntry,activeUID,role,doctorName:userEntry.name});
    setQueue(q);setFKey(f=>f+1);setConfirmed(false);setShowCreds(true);setSc("cases");
  };

  // ═══════════════════════════════════════════════════════
  // COMPLETE CASE
  // ═══════════════════════════════════════════════════════
  const doComplete=async()=>{
    const cur=queue[0];if (!cur)return;
    const newCases=cases.map(c=>c.id===cur.id?{...c,status:"completed"}:c);
    await saveCases(newCases);
    const entry={doctorName:session.doctorName,uid:session.activeUID.uid,role:session.role,
      caseId:cur.id,ts:new Date().toISOString(),date:todayStr()};
    await saveLog([...log,entry]);
    const newReg={...uidReg,[session.activeUID.uid]:(uidReg[session.activeUID.uid]||0)+1};
    await saveUidReg(newReg);
    if (newReg[session.activeUID.uid]>=LIMIT){
      const nxt=findActiveUID(session.userEntry,newReg);
      setSession(s=>({...s,activeUID:nxt||null}));
      if (nxt){setConfirmed(false);setShowCreds(true);}
    }
    const newQ=newCases
      .filter(c=>norm(c.doctorName)===norm(session.doctorName)&&c.role===session.role&&c.status==="pending")
      .sort((a,b)=>(a.seq||0)-(b.seq||0));
    setQueue(newQ);setFKey(f=>f+1);
    if (!newQ.length) setSc("done");
  };

  // ═══════════════════════════════════════════════════════
  // ADMIN GATE
  // ═══════════════════════════════════════════════════════
  const checkAdmin=()=>{
    if (adminPw===cfg?.adminPass){setAdminPw("");setAdminErr("");setSession({role:"admin",isAdmin:true});setSc("admin");setTab("dash");}
    else setAdminErr("Incorrect password");
  };

  // ═══════════════════════════════════════════════════════
  // IMPORT: parse file
  // ═══════════════════════════════════════════════════════
  const afterParse=parsed=>{
    if (!Object.keys(parsed).length){setIErr("No usable data found.");return;}
    setSheets(parsed);
    const hset=new Set();
    Object.values(parsed).forEach(({headers})=>headers.forEach(h=>hset.add(h)));
    const hdrs=[...hset].filter(Boolean);
    setAllH(hdrs);
    const det=autoDetectCols(parsed);
    setImportMode(det.mode);
    if (det.mode==="separate"){
      setAnalyserCol(det.analyserCol||"");
      setSupvCol(det.supervisorCol||"");
      setDetCols(hdrs.filter(h=>h!==det.analyserCol&&h!==det.supervisorCol));
    } else {
      setRoleCol(det.roleCol||"");
      setDetCols(hdrs.filter(h=>h!==det.roleCol));
    }
    setWStep(2);
  };

  const onCaseFile=f=>{
    setIErr("");setILoad(true);
    const rd=new FileReader();
    rd.onload=ev=>{try{afterParse(parseXLSXBuf(ev.target.result));}catch(e){setIErr("Could not read: "+e.message);}setILoad(false);};
    rd.readAsArrayBuffer(f);
  };

  const fetchGS=async()=>{
    const url=gsExportUrl(gsUrl);
    if (!url){setIErr("Invalid URL.");return;}
    setILoad(true);setIErr("");
    try{const res=await fetch(url);if(!res.ok)throw new Error(`HTTP ${res.status}`);afterParse(parseXLSXBuf(await res.arrayBuffer()));}
    catch(e){setIErr(e.message);}
    setILoad(false);
  };

  // ── Build preview cases ────────────────────────────────
  const buildPreview=()=>{
    if (!sheets)return[];
    const prev=[];let seq=1;
    Object.entries(sheets).forEach(([sh,{data}])=>{
      data.forEach((row,ri)=>{
        const cdata={};detCols.forEach(col=>{cdata[col]=row[col]??"";});
        // All cases imported as PENDING. Status is driven by Google Form response sync.
        if (importMode==="separate"){
          if (analyserCol&&row[analyserCol]?.trim())
            prev.push({id:`CASE-${String(seq++).padStart(5,"0")}`,role:"analyser",doctorName:row[analyserCol].trim(),seq:ri+1,sheet:sh,status:"pending",data:cdata});
          if (supvCol&&row[supvCol]?.trim())
            prev.push({id:`CASE-${String(seq++).padStart(5,"0")}`,role:"supervisor",doctorName:row[supvCol].trim(),seq:ri+1,sheet:sh,status:"pending",data:cdata});
        } else {
          const role=detectRole(row[roleCol]||"");
          const dn=nameCol?(row[nameCol]||"").trim():"";
          if (role&&dn) prev.push({id:`CASE-${String(seq++).padStart(5,"0")}`,role,doctorName:dn,seq,sheet:sh,status:"pending",data:cdata});
        }
      });
    });
    return prev;
  };


  // Compute auto-PINs for doctors not yet in User Master
  const computeAutoPins=(preview,currentUsers)=>{
    const pins={};
    const docSet=[...new Set(preview.map(p=>`${p.doctorName}|${p.role}`))];
    docSet.forEach(dk=>{
      const [dn]=dk.split("|");
      const exists=currentUsers.find(u=>norm(u.name)===norm(dn));
      if (!exists) pins[dn]=autoPins[dn]||randPin();
    });
    return pins;
  };

  const activateImport=async(preview,pins,currentUsers,importLabel)=>{
    await archiveThenSaveCases(preview, importLabel||`Import ${new Date().toLocaleDateString("en-IN")} (${preview.length} cases)`);
    let newUsers=[...currentUsers];
    Object.entries(pins).forEach(([dn,pin])=>{
      if (!newUsers.find(u=>norm(u.name)===norm(dn))){
        const role=preview.find(p=>norm(p.doctorName)===norm(dn))?.role||"analyser";
        newUsers.push({name:dn,pin,role,userIds:[]});
      }
    });
    await saveUsers(newUsers);
    await saveCfg({...cfg,detailCols:detCols,importMode,analyserCol,supervisorCol:supvCol,roleCol,doctorNameCol:nameCol,importedAt:new Date().toISOString()});
    alert(`✅ Imported ${preview.length} cases. Previous data archived. ${Object.keys(pins).length} new doctor account(s) created.`);
    setWStep(1);setSheets(null);setAllH([]);setAutoPins({});setTab("dash");
  };

  // ═══════════════════════════════════════════════════════
  // SYNC FROM GOOGLE FORM RESPONSE SHEETS
  // ═══════════════════════════════════════════════════════
  // Architecture:
  //   - Google Form responses land in a Google Sheet (one per role)
  //   - Each response row has: User ID (what doctor typed), Case Key (e.g. HrnId/TID), Status/Action
  //   - We match response rows → cases by case key, mark completed, rebuild UID counts
  const doSync=async()=>{
    const sc = cfg?.syncCols||{};
    const fdUrls = cfg?.formDataUrls||{};
    if (!sc.caseKey)    { setSyncMsg("⚠ Set the Case Key column in Sync Settings first."); return; }
    if (!sc.formCaseKey){ setSyncMsg("⚠ Set the Form Response Case ID column in Sync Settings first."); return; }
    if (!sc.formUidCol) { setSyncMsg("⚠ Set the Form Response User ID column in Sync Settings first."); return; }
    if (!fdUrls.analyser&&!fdUrls.supervisor){ setSyncMsg("⚠ Add at least one Form Response Sheet URL in Settings first."); return; }

    setSyncing(true); setSyncMsg("Fetching form responses…");

    let allResponses=[]; // {caseKey, uid, statusVal, role}
    const errors=[];

    for (const role of ["analyser","supervisor"]) {
      const url=fdUrls[role]; if (!url) continue;
      const csvUrl=sheetToCsvUrl(url);
      if (!csvUrl){errors.push(`Invalid URL for ${role}`);continue;}
      try {
        const res=await fetch(csvUrl);
        if (!res.ok) throw new Error(`HTTP ${res.status} — share sheet as Anyone with link → Viewer`);
        const text=await res.text();
        const rows=parseCSV(text);
        rows.forEach(row=>{
          const ck  = String(row[sc.formCaseKey]||"").trim(); if (!ck) return;
          const uid = String(row[sc.formUidCol]||"").trim();
          const sv  = String(row[sc.formStatusCol]||"").trim(); // optional
          allResponses.push({caseKey:ck, uid, statusVal:sv, role});
        });
        setSyncMsg(`Fetched ${role} responses… (${rows.length} rows)`);
      } catch(e){ errors.push(`${role}: ${e.message}`); }
    }

    if (!allResponses.length && errors.length){
      setSyncMsg("❌ Could not fetch any responses. " + errors.join(" | ")); setSyncing(false); return;
    }

    // Rebuild UID registry from ALL responses
    const newReg={};
    allResponses.forEach(r=>{ if(r.uid) newReg[r.uid]=(newReg[r.uid]||0)+1; });

    // Mark matching cases completed
    let updated=0;
    const keyField=sc.caseKey;
    const newCases=cases.map(c=>{
      const caseKeyVal=String(c.data?.[keyField]||"").trim();
      if (!caseKeyVal) return c;
      // Find any response matching this case key
      const match=allResponses.find(r=>norm(r.caseKey)===norm(caseKeyVal)&&(!r.role||r.role===c.role));
      if (match && c.status!=="completed"){ updated++; return {...c,status:"completed",syncedUid:match.uid,syncedStatus:match.statusVal}; }
      return c;
    });

    await saveCases(newCases);
    await saveUidReg(newReg);
    const entry={ts:new Date().toISOString(),updated,uidCounts:Object.keys(newReg).length,errors};
    setSyncLog(l=>[entry,...l.slice(0,9)]);

    const msg=`✅ Sync complete — ${updated} case(s) marked completed, ${Object.values(newReg).reduce((a,b)=>a+b,0)} total form submissions counted across ${Object.keys(newReg).length} User ID(s).${errors.length?" Errors: "+errors.join(" | "):""}`;
    setSyncMsg(msg); setSyncing(false);
  };

  // ═══════════════════════════════════════════════════════
  // USER MASTER UPLOAD
  // ═══════════════════════════════════════════════════════
  const onUserMasterFile=async f=>{
    setUmLoad(true);setUmMsg("");setUmErr("");
    try {
      const ab=await f.arrayBuffer();
      const imported=parseUserMaster(ab);
      if (!imported.length){setUmErr("No users found. Check the format.");setUmLoad(false);return;}
      let merged=[...users];
      imported.forEach(nu=>{
        const idx=merged.findIndex(u=>norm(u.name)===norm(nu.name));
        if (idx>=0) merged[idx]={...merged[idx],...nu};
        else merged.push(nu);
      });
      await archiveThenSaveUsers(merged, `User Master ${new Date().toLocaleDateString("en-IN")} (${merged.length} users)`);
      setUmMsg(`✅ ${imported.length} user(s) imported/updated. Previous User Master archived.`);
    } catch(e){ setUmErr("Error reading file: "+e.message); }
    setUmLoad(false);
  };

  // ═══════════════════════════════════════════════════════
  // REPORT
  // ═══════════════════════════════════════════════════════
  const getReport=()=>{
    const td=todayStr();const docMap={};
    cases.forEach(c=>{
      const key=`${norm(c.doctorName)}|${c.role}`;
      if (!docMap[key]) docMap[key]={name:c.doctorName,role:c.role,total:0,comp:0,todayN:0};
      docMap[key].total++;
      if (c.status==="completed"){
        docMap[key].comp++;
        if (log.find(l=>l.caseId===c.id&&l.date===td)) docMap[key].todayN++;
      }
    });
    const rows=Object.values(docMap).map(r=>({...r,pending:r.total-r.comp}));
    const totals=rows.reduce((a,r)=>({total:a.total+r.total,comp:a.comp+r.comp,pending:a.pending+r.pending,todayN:a.todayN+r.todayN}),{total:0,comp:0,pending:0,todayN:0});
    return {rows:rows.sort((a,b)=>a.name.localeCompare(b.name)),totals};
  };

  const exportCSV=()=>{
    const {rows,totals}=getReport();const td=todayStr();
    const lines=[`CaseFlow Report — ${td}`,"","Doctor,Role,Total,Completed,Pending,Today",
      ...rows.map(r=>`${r.name},${ROLES[r.role]?.label||r.role},${r.total},${r.comp},${r.pending},${r.todayN}`),"",
      `TOTAL,,${totals.total},${totals.comp},${totals.pending},${totals.todayN}`];
    const a=document.createElement("a");
    a.href=URL.createObjectURL(new Blob([lines.join("\n")],{type:"text/csv"}));
    a.download=`caseflow_${td.replace(/\//g,"-")}.csv`;a.click();
  };

  const wrap={minHeight:"100vh",background:C.bg,color:C.txt,fontFamily:"'DM Sans',sans-serif"};

  // ══════════════════════════════════════════════════════
  // BOOT
  // ══════════════════════════════════════════════════════
  if (sc==="boot") return (
    <div style={{...wrap,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:14}}>
      <style>{G}</style><Spinner size={44}/><span style={{color:C.muted}}>Loading CaseFlow…</span>
    </div>
  );

  // ══════════════════════════════════════════════════════
  // LOGIN
  // ══════════════════════════════════════════════════════
  if (sc==="login") return (
    <div style={{...wrap,display:"flex",flexDirection:"column"}}>
      <style>{G}</style>
      <TopBar right={<button style={{...$.btn("rgba(255,255,255,.1)","rgba(255,255,255,.6)"),padding:"6px 14px",fontSize:12}}
        onClick={()=>{setAdminPw("");setAdminErr("");setSc("adminGate");}}>Admin →</button>}/>
      <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
        <div style={{width:"100%",maxWidth:420}} className="up">
          <div style={{textAlign:"center",marginBottom:24}}>
            <div style={{width:70,height:70,background:C.navyM,borderRadius:20,display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 14px",fontSize:32}}>🏥</div>
            <h1 style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:24,marginBottom:4}}>CaseFlow</h1>
            <p style={{color:C.muted,fontSize:14}}>Secure medical case management</p>
          </div>
          <div style={{display:"flex",background:C.lt,borderRadius:10,padding:4,gap:4,marginBottom:20}}>
            {[["pin","🔢 PIN Login"],["uid","🆔 User ID + Password"]].map(([m,lb])=>(
              <button key={m} onClick={()=>{setLMode(m);setLErr("");}}
                style={{...$.btn(lMode===m?C.navy:C.lt,lMode===m?"#fff":C.muted),flex:1,borderRadius:8,padding:"9px 8px",fontSize:12}}>{lb}</button>
            ))}
          </div>
          <div style={$.card}>
            {lMode==="pin"?(
              <>
                <label style={$.lbl}>Your PIN</label>
                <input style={{...$.inp,textAlign:"center",fontSize:28,letterSpacing:8,fontWeight:700,marginBottom:20}}
                  type="password" maxLength={6} placeholder="••••" value={pinV}
                  onChange={e=>{setPinV(e.target.value);setLErr("");}} onKeyDown={e=>e.key==="Enter"&&doLogin()} autoFocus/>
              </>
            ):(
              <>
                <div style={{marginBottom:14}}><label style={$.lbl}>User ID</label>
                  <input style={$.inp} placeholder="e.g. ANLSR001" value={uidV} onChange={e=>{setUidV(e.target.value);setLErr("");}}/></div>
                <div style={{marginBottom:20}}><label style={$.lbl}>Password</label>
                  <input style={$.inp} type="password" placeholder="••••••••" value={pwV}
                    onChange={e=>{setPwV(e.target.value);setLErr("");}} onKeyDown={e=>e.key==="Enter"&&doLogin()}/></div>
              </>
            )}
            {lErr&&<p style={{color:C.red,fontSize:13,marginBottom:12}}>⚠ {lErr}</p>}
            <button style={{...$.btn(),width:"100%",padding:"13px",fontSize:15}} onClick={doLogin}>Sign In →</button>
          </div>
        </div>
      </div>
    </div>
  );

  // ══════════════════════════════════════════════════════
  // ADMIN GATE
  // ══════════════════════════════════════════════════════
  if (sc==="adminGate") return (
    <div style={{...wrap,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <style>{G}</style>
      <div style={{...$.card,maxWidth:340,width:"100%",textAlign:"center"}} className="up">
        <div style={{fontSize:36,marginBottom:10}}>🔐</div>
        <h2 style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:20,marginBottom:6}}>Admin Access</h2>
        <input style={{...$.inp,textAlign:"center",fontSize:20,letterSpacing:4,marginBottom:10}}
          type="password" placeholder="••••••" value={adminPw}
          onChange={e=>{setAdminPw(e.target.value);setAdminErr("");}} onKeyDown={e=>e.key==="Enter"&&checkAdmin()} autoFocus/>
        {adminErr&&<p style={{color:C.red,fontSize:13,marginBottom:10}}>⚠ {adminErr}</p>}
        <div style={{display:"flex",gap:8}}>
          <button style={{...$.btn(C.lt,C.muted),border:`1.5px solid ${C.bdr}`,flex:1}} onClick={()=>setSc("login")}>← Back</button>
          <button style={{...$.btn(),flex:1}} onClick={checkAdmin}>Enter</button>
        </div>
      </div>
    </div>
  );

  // ══════════════════════════════════════════════════════
  // CASES SCREEN
  // ══════════════════════════════════════════════════════
  if (sc==="cases") {
    const curCase =queue[0];
    const ri      =ROLES[session?.role]||ROLES.analyser;
    const formUrl =cfg?.formUrls?.[session?.role]||"";
    const uidInfo =session?.activeUID;
    const totalAss=cases.filter(c=>norm(c.doctorName)===norm(session?.doctorName)&&c.role===session?.role).length;
    const doneN   =log.filter(l=>norm(l.doctorName)===norm(session?.doctorName)&&l.role===session?.role).length;
    const uidUsed =uidInfo?(uidReg[uidInfo.uid]||0):0;

    return (
      <div style={{...wrap,display:"flex",flexDirection:"column"}}>
        <style>{G}</style>

        {/* CREDENTIALS MODAL */}
        {showCreds&&uidInfo&&(
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.65)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999,padding:20}}>
            <div style={{...$.card,maxWidth:460,width:"100%"}} className="up">
              <div style={{fontSize:44,textAlign:"center",marginBottom:12}}>🔑</div>
              <h2 style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:20,textAlign:"center",marginBottom:4}}>Your Login Credentials</h2>
              <p style={{color:C.muted,fontSize:13,textAlign:"center",marginBottom:20,lineHeight:1.65}}>
                Enter these in the Google Form for every case.<br/>
                <strong>Next credentials shown only after {LIMIT} form submissions on this ID.</strong>
              </p>
              <label style={{display:"flex",alignItems:"center",gap:16,padding:"18px 20px",
                borderRadius:12,border:`2.5px solid ${confirmed?C.green:C.bdr}`,
                background:confirmed?"#F0FDF4":C.lt,cursor:"pointer",transition:"all .2s",marginBottom:16}}>
                <input type="radio" checked={confirmed} onChange={()=>setConfirmed(true)} style={{width:22,height:22,flexShrink:0}}/>
                <div style={{flex:1}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10,marginBottom:8}}>
                    <div>
                      <div style={{fontSize:10,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:"0.06em"}}>User ID</div>
                      <div style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:24,color:C.navy,letterSpacing:1}}>{uidInfo.uid}</div>
                    </div>
                    <div>
                      <div style={{fontSize:10,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:"0.06em"}}>Password</div>
                      <div style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:24,color:C.teal,letterSpacing:1}}>{uidInfo.password}</div>
                    </div>
                  </div>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                    <span style={$.tag(ri.bg,ri.color)}>{ri.icon} {ri.label}</span>
                    <span style={$.tag(uidUsed>=100?"#FEE2E2":uidUsed>=80?"#FFF7ED":"#F0FDF4",uidUsed>=100?C.red:uidUsed>=80?C.amber:C.green)}>
                      {uidUsed}/{LIMIT} used globally
                    </span>
                    <span style={$.tag(C.lt,C.muted)}>
                      ID #{(session.userEntry?.userIds?.findIndex(u=>u.uid===uidInfo.uid)??0)+1} of {session.userEntry?.userIds?.length||1}
                    </span>
                  </div>
                </div>
              </label>
              {!confirmed&&<p style={{color:C.amber,fontSize:13,textAlign:"center",fontWeight:600,marginBottom:12}}>⚠ Click the option above to confirm you have noted your credentials</p>}
              <button style={{...$.btn(C.green),width:"100%",padding:"13px",fontSize:15,opacity:confirmed?1:0.4}}
                disabled={!confirmed} onClick={()=>setShowCreds(false)}>✓ Noted — Start Cases →</button>
            </div>
          </div>
        )}

        <TopBar subtitle={ri.label.toUpperCase()} right={
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <span style={{color:"rgba(255,255,255,.6)",fontSize:13}}>{session?.doctorName} · {doneN}/{totalAss}</span>
            <button style={{...$.btn("rgba(255,255,255,.1)","rgba(255,255,255,.6)"),padding:"5px 12px",fontSize:12}}
              onClick={()=>{setSc("login");setSession(null);setPinV("");setPwV("");setUidV("");}}>Logout</button>
          </div>
        }/>
        <div style={{height:4,background:"rgba(0,0,0,.08)"}}>
          <div style={{height:"100%",width:`${totalAss?(doneN/totalAss)*100:0}%`,background:`linear-gradient(90deg,${ri.color},${C.green})`,transition:"width .5s"}}/>
        </div>

        {/* No cases assigned */}
        {totalAss===0&&(
          <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center"}}>
            <div style={{...$.card,maxWidth:420,textAlign:"center",padding:40}} className="up">
              <div style={{fontSize:52,marginBottom:14}}>⚠️</div>
              <h2 style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:22,marginBottom:10}}>No Cases Assigned</h2>
              <p style={{color:C.muted,fontSize:14,lineHeight:1.7,marginBottom:20}}>
                No cases found for <strong>{session?.doctorName}</strong> ({ri.label}).<br/>
                This usually means the name in the spreadsheet doesn't exactly match your account name.<br/>
                <strong>Contact your admin</strong> to verify the spelling matches exactly.
              </p>
              <button style={{...$.btn(C.lt,C.muted),border:`1.5px solid ${C.bdr}`}}
                onClick={()=>{setSc("login");setSession(null);setPinV("");}}>← Return to Login</button>
            </div>
          </div>
        )}

        {/* Cases exist but all done */}
        {totalAss>0&&!curCase&&(
          <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center"}}>
            <div style={{...$.card,maxWidth:440,textAlign:"center",padding:40}} className="up">
              <div style={{fontSize:60,marginBottom:14}}>✅</div>
              <h2 style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:24,marginBottom:8}}>All Cases Processed</h2>
              <p style={{color:C.muted,fontSize:14,lineHeight:1.7,marginBottom:10}}>
                You have stepped through all <strong>{totalAss}</strong> cases assigned to you in CaseFlow.
              </p>
              <div style={{background:"#FFF7ED",border:"1px solid #FED7AA",borderRadius:10,padding:"12px 16px",marginBottom:20,textAlign:"left"}}>
                <p style={{fontSize:13,color:"#92400E",lineHeight:1.7}}>
                  ⚠️ <strong>Note:</strong> CaseFlow tracks that you have <em>viewed and stepped through</em> each case. It does <em>not</em> confirm whether you have submitted the Google Form for each case. Please verify your form submissions independently in Google Forms.
                </p>
              </div>
              <button style={{...$.btn(C.lt,C.muted),border:`1.5px solid ${C.bdr}`}}
                onClick={()=>{setSc("login");setSession(null);setPinV("");}}>← Return to Login</button>
            </div>
          </div>
        )}

        {/* Active case */}
        {totalAss>0&&curCase&&(
          <div style={{maxWidth:840,margin:"0 auto",width:"100%",padding:"20px 16px"}}>
            {/* UID banner — shows usage count only; credentials available via View Creds */}
            {uidInfo&&(
              <div style={{background:C.navyM,borderRadius:10,padding:"10px 16px",marginBottom:14,display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <span style={{fontSize:12,color:"rgba(255,255,255,.5)"}}>Session active</span>
                  <span style={{...$.tag("rgba(255,255,255,.08)","rgba(255,255,255,.5)"),fontSize:11}}>{ri.icon} {ri.label}</span>
                </div>
                <div style={{display:"flex",alignItems:"center",gap:14}}>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontSize:10,color:"rgba(255,255,255,.45)",textTransform:"uppercase",letterSpacing:"0.08em"}}>This ID used</div>
                    <div style={{color:uidUsed>=100?"#FCA5A5":uidUsed>=80?"#FCD34D":"#6EE7B7",fontWeight:700,fontSize:14}}>{uidUsed} / {LIMIT}</div>
                  </div>
                  <button style={{...$.btn("rgba(255,255,255,.12)","rgba(255,255,255,.7)"),padding:"6px 14px",fontSize:12}}
                    onClick={()=>{setConfirmed(true);setShowCreds(true);}}>🔑 View Creds</button>
                </div>
              </div>
            )}

            {/* Case card */}
            <div style={{...$.card,marginBottom:14}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8,marginBottom:16}}>
                <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
                  <span style={$.tag(ri.bg,ri.color)}>{ri.icon} {ri.label.toUpperCase()}</span>
                  <span style={$.tag(C.navy,"#fff")}>CASE #{doneN+1}</span>
                  <span style={$.tag(C.lt,C.muted)}>{curCase.sheet}</span>
                </div>
                <span style={{fontSize:13,color:C.muted}}>{queue.length-1} more pending</span>
              </div>
              {/* Only show fields with actual data — blank/empty fields and other-sheet columns hidden automatically */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))",gap:10}}>
                {Object.entries(curCase.data||{})
                  .filter(([,v])=>v&&String(v).trim()&&String(v).trim()!=="—"&&String(v).trim()!=="-")
                  .map(([k,v])=>(
                  <div key={k} style={{background:C.lt,borderRadius:10,padding:"12px 14px",border:`1px solid ${C.bdr}`}}>
                    <div style={{fontSize:10,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:4}}>{k}</div>
                    <div style={{fontWeight:700,fontSize:15,wordBreak:"break-word",fontFamily:"'Sora',sans-serif"}}>{v}</div>
                  </div>
                ))}
              </div>
            </div>

            {/* Embedded form */}
            <div style={{...$.card,padding:0,overflow:"hidden",marginBottom:14}}>
              <div style={{background:ri.color,padding:"12px 18px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
                <span style={{color:"#fff",fontWeight:700,fontSize:14,fontFamily:"'Sora',sans-serif"}}>
                  {ri.icon} {ri.label} Form
                </span>
                <a href={formUrl} target="_blank" rel="noopener noreferrer"
                  style={{color:"rgba(255,255,255,.8)",fontSize:12,textDecoration:"none",background:"rgba(255,255,255,.15)",padding:"4px 10px",borderRadius:6}}>Open in tab ↗</a>
              </div>
              {formUrl
                ?<iframe key={fKey} src={fmtEmbed(formUrl)} style={{width:"100%",height:640,border:"none",display:"block"}} title="Review Form"/>
                :<div style={{padding:32,textAlign:"center",color:C.muted}}>Form URL not configured — Admin → Settings.</div>}
            </div>

            {/* CTA — honest language */}
            <div style={{...$.card,background:"#F0FDF4",border:"1px solid #A7F3D0",padding:"20px 24px"}}>
              <p style={{color:"#065F46",fontSize:14,marginBottom:6,fontWeight:700}}>
                Step {doneN+1} of {totalAss}
              </p>
              <p style={{color:"#065F46",fontSize:13,marginBottom:14,lineHeight:1.6}}>
                Fill and <strong>submit the form above</strong>, then click below to load the next case.
              </p>
              <button style={{...$.btn(C.green),padding:"13px 32px",fontSize:15}} onClick={doComplete}>
                {queue.length>1
                  ?`I have submitted the form — Load Case ${doneN+2} →`
                  :"I have submitted the form — Finish ✓"}
              </button>
              <p style={{fontSize:11,color:"#4B7A5E",marginTop:10,lineHeight:1.5}}>
                ⚠ Clicking this only advances the queue in CaseFlow. The Google Form submission is tracked separately by Google.
              </p>
            </div>
          </div>
        )}
      </div>
    );
  }

  // ══════════════════════════════════════════════════════
  // ADMIN PANEL
  // ══════════════════════════════════════════════════════
  if (sc==="admin") {
    const role    =session?.role||"admin";
    const isAdmin =role==="admin";
    const visTabs = isAdmin
      ? ["dash","import","users","report","settings"]
      : role==="upload"        ? ["import"]
      : role==="mis"           ? ["import","report"]
      : role==="state_manager" ? ["import","report"]
      : role==="credential"    ? ["users"]
      : role==="viewer"        ? ["dash","report"]
      : ["dash"];
    const TAB={
      dash:    {icon:"📊",label:"Dashboard"},
      import:  {icon:"📥",label:"Import Cases"},
      users:   {icon:"👥",label:"User Master"},
      report:  {icon:"📈",label:"Reports"},
      settings:{icon:"⚙️",label:"Settings"},
    };
    const {rows:repRows,totals}=getReport();
    const totalCases=cases.length;
    const totalComp=cases.filter(c=>c.status==="completed").length;

    return (
      <div style={{...wrap,display:"flex",flexDirection:"column",height:"100vh"}}>
        <style>{G}</style>
        <TopBar subtitle="ADMIN" right={isAdmin&&<button style={{...$.btn("rgba(255,255,255,.08)","rgba(255,255,255,.6)"),padding:"6px 14px",fontSize:12}} onClick={()=>setSc("login")}>← Login</button>}/>
        <div style={{display:"flex",flex:1,overflow:"hidden"}}>

          {/* Sidebar */}
          <div style={{width:185,background:C.navyM,padding:"12px 8px",display:"flex",flexDirection:"column",gap:2,flexShrink:0,overflowY:"auto"}}>
            {visTabs.map(t=>(
              <button key={t} onClick={()=>setTab(t)} style={{display:"flex",alignItems:"center",gap:9,padding:"10px 12px",borderRadius:9,
                border:"none",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:13,
                background:tab===t?"rgba(255,255,255,.14)":"transparent",color:tab===t?"#fff":"rgba(255,255,255,.45)",textAlign:"left"}}>
                <span style={{fontSize:16}}>{TAB[t].icon}</span>{TAB[t].label}
              </button>
            ))}
            <div style={{flex:1}}/>
            <div style={{padding:"8px 12px",fontSize:10,color:"rgba(255,255,255,.3)",lineHeight:1.6}}>
              {ROLES[role]?.label||role}<br/>
              {cfg?.importedAt?"Import: "+new Date(cfg.importedAt).toLocaleDateString("en-IN"):"No import yet"}
            </div>
          </div>

          <div style={{flex:1,overflow:"auto",padding:24}}>

            {/* ══ DASHBOARD ══ */}
            {tab==="dash"&&(
              <div className="up">
                <h2 style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:21,marginBottom:20}}>Dashboard</h2>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:14,marginBottom:20}}>
                  <StatCard icon="📋" value={totalCases} label="Total Cases" color={C.navy}/>
                  <StatCard icon="✅" value={totalComp}  label="Stepped Through" color={C.green}/>
                  <StatCard icon="⏳" value={totalCases-totalComp} label="Pending" color={C.amber}/>
                  <StatCard icon="👥" value={users.filter(u=>["analyser","supervisor"].includes(u.role)).length} label="Doctors" color={C.teal}/>
                </div>
                <div style={{...$.card,marginBottom:14}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
                    <span style={{fontWeight:700}}>Overall Workflow Progress</span>
                    <span style={{color:C.muted,fontSize:14}}>{totalCases?Math.round(totalComp/totalCases*100):0}%</span>
                  </div>
                  <PBar pct={totalCases?(totalComp/totalCases)*100:0}/>
                  <p style={{fontSize:11,color:C.muted,marginTop:8}}>⚠ Tracks cases stepped through in CaseFlow only — not actual Google Form submissions.</p>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
                  {["analyser","supervisor"].map(r=>{
                    const rc=ROLES[r];const arr=cases.filter(c=>c.role===r);const done=arr.filter(c=>c.status==="completed").length;
                    return(
                      <div key={r} style={{...$.card,background:rc.bg,border:`1px solid ${rc.color}30`}}>
                        <div style={{fontSize:13,fontWeight:700,color:rc.color,marginBottom:8}}>{rc.icon} {rc.label}</div>
                        <div style={{fontWeight:800,fontSize:22,fontFamily:"'Sora',sans-serif",marginBottom:8}}>{done}/{arr.length}</div>
                        <PBar pct={arr.length?(done/arr.length)*100:0} color={rc.color} h={5}/>
                      </div>
                    );
                  })}
                </div>

                {/* Sync shortcut */}
                <div style={{...$.card,border:`1.5px solid ${C.green}30`,background:"#F0FDF4"}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12}}>
                    <div>
                      <div style={{fontWeight:700,fontSize:14,color:C.green,marginBottom:4}}>🔄 Sync from Google Form Responses</div>
                      <div style={{fontSize:13,color:"#4B7A5E",lineHeight:1.6}}>
                        Pull latest form submissions to update case statuses and User ID usage counts.<br/>
                        {syncLog[0]?<span>Last sync: <strong>{new Date(syncLog[0].ts).toLocaleString("en-IN")}</strong> · {syncLog[0].updated} cases updated</span>:"Never synced"}
                      </div>
                    </div>
                    <div style={{display:"flex",alignItems:"center",gap:10,flexShrink:0}}>
                      {syncMsg&&<span style={{fontSize:12,color:syncMsg.startsWith("❌")||syncMsg.startsWith("⚠")?C.red:C.green,maxWidth:220,lineHeight:1.4}}>{syncMsg}</span>}
                      <button style={{...$.btn(C.green),padding:"10px 20px",opacity:syncing?0.6:1}}
                        disabled={syncing} onClick={doSync}>
                        {syncing?"⏳ Syncing…":"🔄 Sync Now"}
                      </button>
                      <button style={{...$.btn(C.lt,C.muted),border:`1px solid ${C.bdr}`,padding:"10px 14px",fontSize:12}}
                        onClick={()=>setTab("settings")}>Configure →</button>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* ══ IMPORT ══ */}
            {tab==="import"&&(
              <div className="up">
                <h2 style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:21,marginBottom:4}}>Import Case Data</h2>
                <p style={{color:C.muted,fontSize:14,marginBottom:20}}>Column headers "Analyser" / "Supervisor" are detected automatically — position doesn't matter.</p>

                {/* Step indicators */}
                <div style={{display:"flex",alignItems:"center",marginBottom:22}}>
                  {["Load File","Map Columns","Preview & Activate"].map((s,i)=>(
                    <div key={i} style={{display:"flex",alignItems:"center",flex:i<2?1:"none"}}>
                      <div style={{width:26,height:26,borderRadius:"50%",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,
                        background:wStep>i+1?C.green:wStep===i+1?C.navy:C.bdr,color:wStep>=i+1?"#fff":C.muted}}>
                        {wStep>i+1?"✓":i+1}
                      </div>
                      <span style={{margin:"0 8px",fontSize:12,fontWeight:600,whiteSpace:"nowrap",color:wStep===i+1?C.navy:C.muted}}>{s}</span>
                      {i<2&&<div style={{flex:1,height:2,background:wStep>i+1?C.green:C.bdr,minWidth:12}}/>}
                    </div>
                  ))}
                </div>

                {/* Step 1 */}
                {wStep===1&&(
                  <div style={$.card}>
                    <div style={{display:"flex",background:C.lt,borderRadius:9,padding:4,gap:4,marginBottom:20,width:"fit-content"}}>
                      {[["xlsx","📁 Upload XLSX"],["gsheet","🔗 Google Sheets"]].map(([id,lb])=>(
                        <button key={id} onClick={()=>{setISrc(id);setIErr("");}}
                          style={{...$.btn(iSrc===id?C.navy:C.lt,iSrc===id?"#fff":C.muted),padding:"8px 16px",fontSize:13,borderRadius:8}}>{lb}</button>
                      ))}
                    </div>
                    {iSrc==="xlsx"&&(
                      <UploadBox label="Upload your case data Excel file" icon="📊" accept=".xlsx,.xls"
                        onFile={onCaseFile} inputRef={caseFileRef} loading={iLoad}/>
                    )}
                    {iSrc==="gsheet"&&(
                      <div>
                        <p style={{fontSize:13,color:C.muted,marginBottom:12}}>Share as <strong>Anyone with the link → Viewer</strong> then paste URL.</p>
                        <div style={{display:"flex",gap:10}}>
                          <input style={{...$.inp,flex:1}} placeholder="https://docs.google.com/spreadsheets/d/…" value={gsUrl} onChange={e=>{setGsUrl(e.target.value);setIErr("");}}/>
                          <button style={{...$.btn(C.teal),flexShrink:0}} onClick={fetchGS} disabled={iLoad}>{iLoad?"Loading…":"Load →"}</button>
                        </div>
                      </div>
                    )}
                    {iErr&&<p style={{color:C.red,fontSize:13,marginTop:12}}>⚠ {iErr}</p>}
                    <IBox type="info">Supports two formats: (A) Separate "Analyser" and "Supervisor" column headers where cell values are doctor names, or (B) one role column with "Analyser"/"Supervisor" values plus a doctor-name column. Auto-detected.</IBox>
                  </div>
                )}

                {/* Step 2 */}
                {wStep===2&&sheets&&(
                  <div style={$.card}>
                    <div style={{marginBottom:18,padding:"12px 16px",background:C.lt,borderRadius:10,fontSize:13}}>
                      <strong>{Object.keys(sheets).length} sheet(s):</strong> {Object.keys(sheets).join(", ")}
                    </div>
                    <div style={{marginBottom:20}}>
                      <label style={$.lbl}>Column format in your spreadsheet</label>
                      <div style={{display:"flex",flexDirection:"column",gap:10}}>
                        {[["separate","Separate columns — 'Analyser' and 'Supervisor' are column headers, cell values = doctor names"],
                          ["single",  "Single role column — one column holds 'Analyser'/'Supervisor' values + a doctor-name column"]].map(([m,desc])=>(
                          <label key={m} style={{display:"flex",alignItems:"flex-start",gap:12,padding:"14px 16px",
                            borderRadius:10,border:`2px solid ${importMode===m?C.navy:C.bdr}`,
                            background:importMode===m?C.navyM+"18":"#fff",cursor:"pointer"}}>
                            <input type="radio" value={m} checked={importMode===m} onChange={()=>setImportMode(m)} style={{marginTop:2,width:18,height:18}}/>
                            <span style={{fontSize:13,color:importMode===m?C.navy:C.muted,fontWeight:importMode===m?600:400,lineHeight:1.5}}>{desc}</span>
                          </label>
                        ))}
                      </div>
                    </div>
                    {importMode==="separate"&&(
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:16}}>
                        <div><label style={$.lbl}>🔬 Analyser Column <span style={{color:C.teal}}>(auto-detected)</span></label>
                          <select style={$.inp} value={analyserCol} onChange={e=>setAnalyserCol(e.target.value)}>
                            <option value="">— None —</option>
                            {allH.filter(h=>h!==supvCol).map(h=><option key={h} value={h}>{h}</option>)}
                          </select></div>
                        <div><label style={$.lbl}>👨‍⚕️ Supervisor Column <span style={{color:C.purp}}>(auto-detected)</span></label>
                          <select style={$.inp} value={supvCol} onChange={e=>setSupvCol(e.target.value)}>
                            <option value="">— None —</option>
                            {allH.filter(h=>h!==analyserCol).map(h=><option key={h} value={h}>{h}</option>)}
                          </select></div>
                      </div>
                    )}
                    {importMode==="single"&&(
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:16}}>
                        <div><label style={$.lbl}>Role Column</label>
                          <select style={$.inp} value={roleCol} onChange={e=>setRoleCol(e.target.value)}>
                            <option value="">— Select —</option>
                            {allH.filter(h=>h!==nameCol).map(h=><option key={h} value={h}>{h}</option>)}
                          </select></div>
                        <div><label style={$.lbl}>Doctor Name Column</label>
                          <select style={$.inp} value={nameCol} onChange={e=>setNameCol(e.target.value)}>
                            <option value="">— Select —</option>
                            {allH.filter(h=>h!==roleCol).map(h=><option key={h} value={h}>{h}</option>)}
                          </select></div>
                      </div>
                    )}

                    <div style={{marginBottom:22}}>
                      <label style={$.lbl}>Detail columns to show on the doctor's case card</label>
                      <div style={{display:"flex",flexWrap:"wrap",gap:8,marginTop:8}}>
                        {allH.filter(h=>![analyserCol,supvCol,roleCol,nameCol].includes(h)).map(h=>{
                          const on=detCols.includes(h);
                          return <button key={h} onClick={()=>setDetCols(d=>on?d.filter(x=>x!==h):[...d,h])}
                            style={{padding:"7px 14px",borderRadius:8,border:`2px solid ${on?C.navy:C.bdr}`,background:on?C.navy:"#fff",color:on?"#fff":C.txt,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:13}}>
                            {on?"✓ ":""}{h}
                          </button>;
                        })}
                      </div>
                    </div>
                    <div style={{display:"flex",gap:10}}>
                      <button style={{...$.btn(C.lt,C.muted),border:`1.5px solid ${C.bdr}`}} onClick={()=>setWStep(1)}>← Back</button>
                      <button style={{...$.btn(),opacity:((importMode==="separate"&&(analyserCol||supvCol))||(importMode==="single"&&roleCol&&nameCol))&&detCols.length?1:0.4}}
                        onClick={()=>{const p=buildPreview();const pins=computeAutoPins(p,users);setAutoPins(pins);setWStep(3);}}
                        disabled={!((importMode==="separate"&&(analyserCol||supvCol))||(importMode==="single"&&roleCol&&nameCol))||!detCols.length}>
                        Preview →
                      </button>
                    </div>
                  </div>
                )}

                {/* Step 3 */}
                {wStep===3&&sheets&&(()=>{
                  const preview=buildPreview();
                  const docSet=[...new Set(preview.map(p=>`${p.doctorName}|${p.role}`))];
                  const newDocs=docSet.filter(dk=>{const [dn]=dk.split("|");return !users.find(u=>norm(u.name)===norm(dn));});

                  return (
                    <div style={$.card}>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:20}}>
                        {[["📋",preview.length,"Total Cases",C.navy],["🔬",preview.filter(p=>p.role==="analyser").length,"Analyser",C.teal],["👨‍⚕️",preview.filter(p=>p.role==="supervisor").length,"Supervisor",C.purp]].map(([ic,v,lb,c])=>(
                          <div key={lb} style={{background:C.lt,borderRadius:10,padding:16,textAlign:"center"}}>
                            <div style={{fontSize:20,marginBottom:4}}>{ic}</div>
                            <div style={{fontSize:24,fontWeight:800,fontFamily:"'Sora',sans-serif",color:c}}>{v}</div>
                            <div style={{fontSize:12,color:C.muted}}>{lb}</div>
                          </div>
                        ))}
                      </div>

                      <label style={$.lbl}>{docSet.length} doctor assignments</label>
                      <div style={{maxHeight:220,overflowY:"auto",display:"flex",flexDirection:"column",gap:6,marginBottom:16}}>
                        {docSet.map(dk=>{
                          const [dn,role]=dk.split("|");
                          const cnt=preview.filter(p=>p.doctorName===dn&&p.role===role).length;
                          const ri=ROLES[role]||{};
                          const isNew=!!autoPins[dn];
                          return (
                            <div key={dk} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 14px",background:C.lt,borderRadius:9,flexWrap:"wrap",gap:8}}>
                              <div style={{display:"flex",alignItems:"center",gap:8}}>
                                <span style={$.tag(ri.bg||C.lt,ri.color||C.muted)}>{ri.icon} {ri.label||role}</span>
                                <span style={{fontWeight:600,fontSize:14}}>{dn}</span>
                                {isNew&&<span style={$.tag("#FEF3C7",C.amber)}>NEW ACCOUNT</span>}
                              </div>
                              <div style={{display:"flex",alignItems:"center",gap:12}}>
                                <span style={{color:C.muted,fontSize:13}}>{cnt} cases</span>
                                {isNew&&(
                                  <div style={{display:"flex",alignItems:"center",gap:6}}>
                                    <span style={{fontSize:11,color:C.muted}}>Auto-PIN:</span>
                                    <input style={{...$.inp,width:80,padding:"5px 8px",textAlign:"center",fontSize:14,fontWeight:700,letterSpacing:2,margin:0}}
                                      value={autoPins[dn]||""} maxLength={6}
                                      onChange={e=>setAutoPins(p=>({...p,[dn]:e.target.value}))}/>
                                  </div>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>

                      {newDocs.length>0&&(
                        <IBox type="warn">
                          <strong>{newDocs.length} new doctor account(s)</strong> will be created with the PINs shown above. You can edit PINs before activating. User IDs must be added later via User Master.
                        </IBox>
                      )}
                      <IBox type="info">CaseFlow tracks case workflow steps only — not actual Google Form submissions. Remind doctors to submit the form before clicking "Next".</IBox>

                      <div style={{display:"flex",gap:10,marginTop:16}}>
                        <button style={{...$.btn(C.lt,C.muted),border:`1.5px solid ${C.bdr}`}} onClick={()=>setWStep(2)}>← Back</button>
                        <button style={$.btn(C.green)} onClick={()=>activateImport(preview,autoPins,users)}>✓ Activate Import ({preview.length} cases)</button>
                      </div>
                    </div>
                  );
                })()}
              </div>
            )}

            {/* ══ USER MASTER ══ */}
            {tab==="users"&&(
              <div className="up">
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:12}}>
                  <div>
                    <h2 style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:21,margin:0}}>User Master</h2>
                    <p style={{color:C.muted,fontSize:13,marginTop:4}}>{users.length} users · PINs, User IDs, roles</p>
                  </div>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                    {/* Download current user master as XLSX */}
                    <button style={{...$.btn(C.green),padding:"9px 16px",fontSize:13}} onClick={()=>{
                      const XLSX_=window.XLSX||require&&require('xlsx');
                      // Build rows
                      const maxIds=Math.max(0,...users.map(u=>u.userIds?.length||0));
                      const idCols=Array.from({length:maxIds},(_, i)=>[`User ID ${i+1}`,`Password ${i+1}`]).flat();
                      const headers=["Name","Role","PIN",...idCols];
                      const rows=users.map(u=>{
                        const base=[u.name,ROLES[u.role]?.label||u.role,u.pin];
                        const ids=(u.userIds||[]).flatMap(x=>[x.uid||"",x.password||""]);
                        while(ids.length<maxIds*2) ids.push("");
                        return [...base,...ids];
                      });
                      const wb=XLSX.utils.book_new();
                      const ws=XLSX.utils.aoa_to_sheet([headers,...rows]);
                      XLSX.utils.book_append_sheet(wb,ws,"User Master");
                      XLSX.writeFile(wb,`UserMaster_${new Date().toLocaleDateString("en-IN").replace(/\//g,"-")}.xlsx`);
                    }}>⬇ Download XLSX</button>
                    {/* Only admin and credential manager can add users */}
                    {(isAdmin||role==="credential")&&(
                      <button style={$.btn(C.teal)} onClick={()=>setAddUser(true)}>+ Add User</button>
                    )}
                  </div>
                </div>

                {addUser &&<UserForm onSave={async d=>{await saveUsers([...users,d]);setAddUser(false);}} onCancel={()=>setAddUser(false)} existingUsers={users}/>}
                {editUser&&<UserForm initial={editUser} onSave={async d=>{await saveUsers(users.map(u=>u.name===editUser.name?d:u));setEditUser(null);}} onCancel={()=>setEditUser(null)} existingUsers={users.filter(u=>u.name!==editUser?.name)}/>}

                {/* Bulk upload from Excel */}
                <div style={{...$.card,marginBottom:20,background:C.lt,border:`1.5px dashed ${C.teal}`}}>
                  <h3 style={{fontFamily:"'Sora',sans-serif",fontWeight:700,fontSize:15,marginBottom:6,color:C.teal}}>📋 Bulk Upload User Master via Excel</h3>
                  <p style={{fontSize:13,color:C.muted,marginBottom:12,lineHeight:1.6}}>
                    Upload an Excel file with these exact column headers:<br/>
                    <code style={{background:"#fff",padding:"2px 6px",borderRadius:4,fontSize:12}}>
                      Name | Role | PIN | User ID 1 | Password 1 | User ID 2 | Password 2 | User ID 3 | Password 3 | ...
                    </code>
                  </p>
                  <div style={{display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
                    <input ref={umFileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}}
                      onChange={e=>{const f=e.target.files[0];if(f)onUserMasterFile(f);e.target.value="";}}/>
                    <button style={{...$.btn(C.teal),padding:"10px 22px"}} onClick={()=>umFileRef.current?.click()} disabled={umLoad}>
                      {umLoad?"Importing…":"📂 Upload User Master Excel"}
                    </button>
                    {umMsg&&<span style={{color:C.green,fontWeight:600,fontSize:13}}>{umMsg}</span>}
                    {umErr&&<span style={{color:C.red,fontSize:13}}>⚠ {umErr}</span>}
                  </div>
                  <div style={{marginTop:10,fontSize:12,color:C.muted,lineHeight:1.7}}>
                    Role values accepted: <strong>Analyser, Supervisor, Upload, Credential, Admin</strong><br/>
                    Existing users will be updated; new names will be added. PINs auto-generated if blank.
                  </div>
                </div>

                {!users.length&&(
                  <div style={{...$.card,textAlign:"center",padding:48,color:C.muted}}>
                    <div style={{fontSize:40,marginBottom:12}}>👥</div>
                    <p style={{fontWeight:700,marginBottom:8}}>No users yet</p>
                    <button style={$.btn(C.teal)} onClick={()=>setAddUser(true)}>+ Add First User</button>
                  </div>
                )}

                <div style={{display:"flex",flexDirection:"column",gap:10}}>
                  {users.map(u=>{
                    const ri=ROLES[u.role]||{};
                    const activeUID=findActiveUID(u,uidReg);
                    return(
                      <div key={u.name} style={{...$.card,padding:"16px 20px"}}>
                        <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
                          <div style={{flex:1,minWidth:220}}>
                            <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginBottom:10}}>
                              <span style={{fontWeight:700,fontSize:16,fontFamily:"'Sora',sans-serif"}}>{u.name}</span>
                              <span style={$.tag(ri.bg||C.lt,ri.color||C.muted)}>{ri.icon} {ri.label||u.role}</span>
                              <span style={{...$.tag(C.lt,C.muted),letterSpacing:2}}>PIN: {u.pin}</span>
                            </div>
                            {u.userIds?.length>0&&(
                              <div style={{display:"flex",flexWrap:"wrap",gap:7,marginBottom:8}}>
                                {u.userIds.map((uid,idx)=>{
                                  const cnt=uidReg[uid.uid]||0;
                                  const isCurrent=uid.uid===activeUID?.uid;
                                  const isFull=cnt>=LIMIT;
                                  const prevFull=u.userIds.slice(0,idx).every(p=>(uidReg[p.uid]||0)>=LIMIT);
                                  const visible=idx===0||prevFull;
                                  return(
                                    <div key={uid.uid} style={{padding:"5px 11px",borderRadius:8,
                                      border:`2px solid ${isFull?C.red:isCurrent?C.green:visible?C.bdr:"transparent"}`,
                                      background:isFull?"#FEF2F2":isCurrent?"#F0FDF4":visible?C.lt:"transparent",
                                      opacity:visible?1:0.3,fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:7}}>
                                      <span style={{color:C.muted,fontSize:10}}>#{idx+1}</span>
                                      <span style={{fontFamily:"'Sora',sans-serif"}}>{uid.uid}</span>
                                      <span style={{color:isFull?C.red:isCurrent?C.green:C.muted,fontWeight:700}}>{cnt}/{LIMIT}</span>
                                      {isCurrent&&!isFull&&<span style={$.tag("#F0FDF4",C.green)}>● Active</span>}
                                      {isFull&&<span style={$.tag("#FEE2E2",C.red)}>FULL</span>}
                                      {!visible&&<span>🔒</span>}
                                    </div>
                                  );
                                })}
                              </div>
                            )}
                            {!u.userIds?.length&&["analyser","supervisor"].includes(u.role)&&(
                              <p style={{fontSize:12,color:C.amber,marginBottom:4}}>⚠ No User IDs assigned — doctor cannot use forms.</p>
                            )}
                            <div style={{fontSize:12,color:C.muted}}>Active: {activeUID?.uid||"None"}</div>
                          </div>
                          <div style={{display:"flex",gap:8,flexShrink:0}}>
                            <button style={{...$.btn("#EFF6FF","#1E40AF"),padding:"6px 12px",fontSize:12}} onClick={()=>setEditUser(u)}>Edit</button>
                            <button style={{...$.btn("#FEE2E2",C.red),padding:"6px 12px",fontSize:12}}
                              onClick={async()=>{if(!confirm(`Delete ${u.name}?`))return;await saveUsers(users.filter(x=>x.name!==u.name));}}>Delete</button>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>

                <IBox type="info">
                  🔒 = UID locked until previous is full. Each User ID has a global {LIMIT}-case limit across all doctors sharing it.
                  Doctors with no User IDs cannot proceed past the credentials screen.
                </IBox>
              </div>
            )}

            {/* ══ REPORTS ══ */}
            {tab==="report"&&(
              <div className="up">
                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:20,flexWrap:"wrap",gap:12}}>
                  <div>
                    <h2 style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:21,margin:0}}>Reports</h2>
                    <p style={{color:C.muted,fontSize:13,marginTop:4}}>{new Date().toLocaleDateString("en-IN",{weekday:"long",year:"numeric",month:"long",day:"numeric"})}</p>
                  </div>
                  <button style={$.btn(C.green)} onClick={exportCSV}>⬇ Export CSV</button>
                </div>
                <div style={{...$.card,background:"#FFF7ED",border:"1px solid #FED7AA",marginBottom:18}}>
                  <p style={{fontSize:13,color:"#92400E",lineHeight:1.6,margin:0}}>
                    ⚠️ <strong>Important:</strong> These counts reflect cases <em>stepped through</em> in CaseFlow — not verified Google Form submissions. For actual submission data, check your Google Forms response sheet.
                  </p>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:12,marginBottom:22}}>
                  <StatCard icon="📋" value={totals.total}   label="Total Assigned"   color={C.navy}/>
                  <StatCard icon="✅" value={totals.comp}    label="Stepped Through"  color={C.green}/>
                  <StatCard icon="⏳" value={totals.pending} label="Pending"          color={C.amber}/>
                  <StatCard icon="📅" value={totals.todayN}  label="Done Today"       color={C.teal}/>
                </div>
                <div style={{...$.card,padding:0,overflow:"hidden",marginBottom:20}}>
                  <div style={{overflowX:"auto"}}>
                    <table style={{width:"100%",fontSize:14}}>
                      <thead><tr style={{background:C.lt}}>
                        {["Doctor","Role","Total","Stepped Through","Pending","Today","Progress"].map(h=>(
                          <th key={h} style={{padding:"11px 14px",textAlign:"left",fontSize:10,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:"0.08em",borderBottom:`1px solid ${C.bdr}`,whiteSpace:"nowrap"}}>{h}</th>
                        ))}
                      </tr></thead>
                      <tbody>
                        {repRows.map((r,i)=>{
                          const pct=r.total?Math.round(r.comp/r.total*100):0;const ri=ROLES[r.role]||{};
                          return(
                            <tr key={`${r.name}-${r.role}`} style={{borderBottom:`1px solid ${C.bdr}`,background:i%2===0?C.surf:C.lt}}>
                              <td style={{padding:"11px 14px",fontWeight:700,whiteSpace:"nowrap"}}>{r.name}</td>
                              <td style={{padding:"11px 14px"}}><span style={$.tag(ri.bg||C.lt,ri.color||C.muted)}>{ri.icon} {ri.label||r.role}</span></td>
                              <td style={{padding:"11px 14px"}}>{r.total}</td>
                              <td style={{padding:"11px 14px",color:C.green,fontWeight:700}}>{r.comp}</td>
                              <td style={{padding:"11px 14px",color:r.pending>0?C.amber:C.muted,fontWeight:r.pending>0?700:400}}>{r.pending}</td>
                              <td style={{padding:"11px 14px",color:C.teal,fontWeight:600}}>{r.todayN}</td>
                              <td style={{padding:"11px 14px",minWidth:100}}>
                                <PBar pct={pct} color={pct>=100?C.green:C.teal}/>
                                <span style={{fontSize:11,color:C.muted,marginTop:3,display:"block"}}>{pct}%</span>
                              </td>
                            </tr>
                          );
                        })}
                        <tr style={{background:C.navyM}}>
                          {["TOTAL","",`${totals.total}`,`${totals.comp}`,`${totals.pending}`,`${totals.todayN}`,`${totals.total?Math.round(totals.comp/totals.total*100):0}%`].map((v,i)=>(
                            <td key={i} style={{padding:"12px 14px",color:i===0?"#fff":i===3?"#6EE7B7":i===4?"#FCD34D":i===5?"#7DD3FC":"#fff",fontWeight:700,fontFamily:i===0?"'Sora',sans-serif":"inherit"}}>{v}</td>
                          ))}
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>
                {isAdmin&&Object.keys(uidReg).length>0&&(
                  <div style={$.card}>
                    <h3 style={{fontFamily:"'Sora',sans-serif",fontWeight:700,fontSize:16,marginBottom:14}}>🆔 Global User ID Utilisation</h3>
                    <div style={{display:"flex",flexDirection:"column",gap:10}}>
                      {users.filter(u=>["analyser","supervisor"].includes(u.role)).flatMap(u=>(u.userIds||[]).map(uid=>({name:u.name,role:u.role,uid}))).map(({name,role,uid})=>{
                        const cnt=uidReg[uid.uid]||0;const pct=Math.min(100,Math.round(cnt/LIMIT*100));const ri=ROLES[role]||{};
                        return(
                          <div key={uid.uid} style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                            <span style={{minWidth:100,fontSize:13,fontFamily:"'Sora',sans-serif",fontWeight:700}}>{uid.uid}</span>
                            <span style={{minWidth:120,fontSize:12,color:C.muted,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{name}</span>
                            <span style={$.tag(ri.bg||C.lt,ri.color||C.muted)}>{ri.icon}</span>
                            <div style={{flex:1,minWidth:80}}><PBar pct={pct} color={pct>=100?C.red:pct>=80?C.amber:C.green}/></div>
                            <span style={{minWidth:60,fontSize:12,textAlign:"right",color:pct>=100?C.red:C.muted,fontWeight:700}}>{cnt}/{LIMIT}</span>
                            {pct>=100&&<span style={$.tag("#FEE2E2",C.red)}>FULL</span>}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* ══ SETTINGS ══ */}
            {tab==="settings"&&(
              <div className="up">
                <h2 style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:21,marginBottom:20}}>Settings</h2>

                {/* A. Form submission URLs (embedded in doctor screen) */}
                <div style={{...$.card,marginBottom:16}}>
                  <h3 style={{fontFamily:"'Sora',sans-serif",fontWeight:700,fontSize:16,marginBottom:4}}>🔗 Google Form URLs <span style={{fontSize:12,color:C.muted,fontWeight:400}}>(embedded in doctor's case screen)</span></h3>
                  <p style={{fontSize:13,color:C.muted,marginBottom:16}}>The form doctors fill in for each case. Shown as an embedded iframe alongside the case data.</p>
                  {["analyser","supervisor","preauthoriser"].map(r=>{
                    const ri=ROLES[r];
                    return(
                      <div key={r} style={{marginBottom:16}}>
                        <label style={$.lbl}>{ri.icon} {ri.label} Form URL</label>
                        <div style={{fontSize:12,color:C.muted,padding:"8px 12px",background:C.lt,borderRadius:8,marginBottom:8,wordBreak:"break-all"}}>
                          Current: {cfg?.formUrls?.[r]||"Not set"}
                        </div>
                        <input style={$.inp} placeholder="https://forms.gle/… or full /viewform URL" value={newForms[r]||""}
                          onChange={e=>setNewForms(f=>({...f,[r]:e.target.value}))}/>
                      </div>
                    );
                  })}
                  <button style={$.btn(C.teal)} onClick={async()=>{
                    const fu={
                      analyser:      newForms.analyser.trim()      ||cfg?.formUrls?.analyser||"",
                      supervisor:    newForms.supervisor.trim()    ||cfg?.formUrls?.supervisor||"",
                      preauthoriser: newForms.preauthoriser.trim() ||cfg?.formUrls?.preauthoriser||"",
                    };
                    await saveCfg({...cfg,formUrls:fu});alert("✅ Updated!");setNewForms({analyser:"",supervisor:"",preauthoriser:""});
                  }}>Update Form URLs</button>
                  <IBox type="warn">If the form does not embed, use the full browser URL ending in /viewform — not the short forms.gle link.</IBox>
                  <button style={$.btn(C.teal)} onClick={async()=>{
                    const fu={analyser:newForms.analyser.trim()||cfg?.formUrls?.analyser,supervisor:newForms.supervisor.trim()||cfg?.formUrls?.supervisor};
                    await saveCfg({...cfg,formUrls:fu});alert("✅ Updated!");setNewForms({analyser:"",supervisor:""});
                  }}>Update Form URLs</button>
                  <IBox type="warn">If the form does not embed in the iframe, use the full browser address bar URL ending in /viewform — not the short forms.gle link.</IBox>
                </div>

                {/* B. Form Response Sheet URLs + sync column mapping */}
                <div style={{...$.card,marginBottom:16,border:`1.5px solid ${C.teal}`}}>
                  <h3 style={{fontFamily:"'Sora',sans-serif",fontWeight:700,fontSize:16,marginBottom:4,color:C.teal}}>📊 Google Form Response Sheet URLs <span style={{fontSize:12,color:C.muted,fontWeight:400}}>(source of truth for status &amp; 120-case count)</span></h3>
                  <p style={{fontSize:13,color:C.muted,marginBottom:16,lineHeight:1.6}}>
                    These are the Google Sheets where your Google Form responses are recorded automatically.<br/>
                    Open your Form → Responses tab → click the Google Sheets icon → copy the sheet URL.<br/>
                    The sheet must be shared as <strong>Anyone with the link → Viewer</strong>.
                  </p>
                  {["analyser","supervisor","preauthoriser"].map(r=>{
                    const ri=ROLES[r];
                    return(
                      <div key={r} style={{marginBottom:16}}>
                        <label style={$.lbl}>{ri.icon} {ri.label} Form Response Sheet URL</label>
                        <div style={{fontSize:12,color:C.muted,padding:"8px 12px",background:C.lt,borderRadius:8,marginBottom:8,wordBreak:"break-all"}}>
                          Current: {cfg?.formDataUrls?.[r]||"Not set"}
                        </div>
                        <input style={$.inp} placeholder="https://docs.google.com/spreadsheets/d/…" value={newDataUrls[r]||""}
                          onChange={e=>setNewDataUrls(f=>({...f,[r]:e.target.value}))}/>
                      </div>
                    );
                  })}
                  <button style={{...$.btn(C.teal),marginBottom:16}} onClick={async()=>{
                    const fu={
                      analyser:      newDataUrls.analyser.trim()      ||cfg?.formDataUrls?.analyser||"",
                      supervisor:    newDataUrls.supervisor.trim()    ||cfg?.formDataUrls?.supervisor||"",
                      preauthoriser: newDataUrls.preauthoriser.trim() ||cfg?.formDataUrls?.preauthoriser||"",
                    };
                    await saveCfg({...cfg,formDataUrls:fu});alert("✅ Form Response URLs updated!");setNewDataUrls({analyser:"",supervisor:"",preauthoriser:""});
                  }}>Save Response Sheet URLs</button>

                  {/* Sync Column Mapping */}
                  <div style={{background:C.lt,borderRadius:10,padding:"16px",border:`1px solid ${C.bdr}`}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                      <span style={{fontWeight:700,fontSize:14}}>🔧 Column Mapping for Sync</span>
                      <button style={{...$.btn(C.lt,C.muted),border:`1px solid ${C.bdr}`,padding:"4px 10px",fontSize:12}}
                        onClick={()=>setSyncCfgOpen(o=>!o)}>{syncCfgOpen?"▲ Hide":"▼ Show"}</button>
                    </div>
                    {syncCfgOpen&&(
                      <div>
                        <p style={{fontSize:12,color:C.muted,marginBottom:12,lineHeight:1.6}}>
                          Tell CaseFlow which columns to use when matching form responses to cases.<br/>
                          The <strong>Case Key</strong> must appear in both the case data AND the form response (e.g. TID, HrnId, Claim No).
                        </p>
                        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:12}}>
                          <div>
                            <label style={$.lbl}>Case Key Column <span style={{color:C.red}}>*</span><br/><span style={{fontSize:10,fontWeight:400,textTransform:"none",letterSpacing:0,color:C.muted}}>Column in your imported cases (e.g. HrnId, TID)</span></label>
                            <input style={$.inp} placeholder="e.g. HrnId" value={cfg?.syncCols?.caseKey||""}
                              onChange={async e=>await saveCfg({...cfg,syncCols:{...(cfg?.syncCols||{}),caseKey:e.target.value}})}/>
                          </div>
                          <div>
                            <label style={$.lbl}>Form Response: Case ID Column <span style={{color:C.red}}>*</span><br/><span style={{fontSize:10,fontWeight:400,textTransform:"none",letterSpacing:0,color:C.muted}}>Column in form response sheet matching the case</span></label>
                            <input style={$.inp} placeholder="e.g. HrnId or Claim Number" value={cfg?.syncCols?.formCaseKey||""}
                              onChange={async e=>await saveCfg({...cfg,syncCols:{...(cfg?.syncCols||{}),formCaseKey:e.target.value}})}/>
                          </div>
                          <div>
                            <label style={$.lbl}>Form Response: User ID Column <span style={{color:C.red}}>*</span><br/><span style={{fontSize:10,fontWeight:400,textTransform:"none",letterSpacing:0,color:C.muted}}>Column where doctor entered their User ID</span></label>
                            <input style={$.inp} placeholder="e.g. User ID" value={cfg?.syncCols?.formUidCol||""}
                              onChange={async e=>await saveCfg({...cfg,syncCols:{...(cfg?.syncCols||{}),formUidCol:e.target.value}})}/>
                          </div>
                          <div>
                            <label style={$.lbl}>Form Response: Status/Action Column <span style={{fontSize:10,fontWeight:400,color:C.muted}}>(optional)</span><br/><span style={{fontSize:10,fontWeight:400,textTransform:"none",letterSpacing:0,color:C.muted}}>If your form captures the action taken</span></label>
                            <input style={$.inp} placeholder="e.g. Action Taken (leave blank if not in form)" value={cfg?.syncCols?.formStatusCol||""}
                              onChange={async e=>await saveCfg({...cfg,syncCols:{...(cfg?.syncCols||{}),formStatusCol:e.target.value}})}/>
                          </div>
                        </div>
                        <IBox type="info">
                          Every form submission = 1 count toward the 120-case User ID limit. Cases are matched by the Case Key column. A case is marked <strong>completed</strong> when any matching form response is found — regardless of what action the doctor took.
                        </IBox>
                      </div>
                    )}
                  </div>

                  {/* Sync button */}
                  <div style={{marginTop:16,display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
                    <button style={{...$.btn(C.green),padding:"11px 24px",fontSize:14,opacity:syncing?0.6:1}}
                      disabled={syncing} onClick={doSync}>
                      {syncing?"⏳ Syncing…":"🔄 Sync Now from Form Responses"}
                    </button>
                    {syncMsg&&(
                      <span style={{fontSize:13,color:syncMsg.startsWith("❌")||syncMsg.startsWith("⚠")?C.red:C.green,fontWeight:600,maxWidth:480,lineHeight:1.5}}>{syncMsg}</span>
                    )}
                  </div>

                  {syncLog.length>0&&(
                    <div style={{marginTop:14}}>
                      <label style={$.lbl}>Recent Sync History</label>
                      <div style={{display:"flex",flexDirection:"column",gap:6}}>
                        {syncLog.map((s,i)=>(
                          <div key={i} style={{padding:"8px 12px",background:"#fff",borderRadius:8,border:`1px solid ${C.bdr}`,fontSize:13,display:"flex",gap:16,flexWrap:"wrap"}}>
                            <span style={{color:C.muted}}>{new Date(s.ts).toLocaleString("en-IN")}</span>
                            <span style={{color:C.green,fontWeight:600}}>{s.updated} cases updated</span>
                            <span style={{color:C.muted}}>{s.uidCounts} UIDs tracked</span>
                            {s.errors?.length>0&&<span style={{color:C.red}}>⚠ {s.errors.join(" | ")}</span>}
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

                {/* C. Admin password */}
                <div style={{...$.card,marginBottom:16}}>
                  <h3 style={{fontFamily:"'Sora',sans-serif",fontWeight:700,fontSize:16,marginBottom:14}}>🔐 Admin Password</h3>
                  <div style={{display:"flex",gap:10}}>
                    <input style={{...$.inp,flex:1}} type="password" placeholder="New admin password" value={newAdminPw} onChange={e=>setNewAdminPw(e.target.value)}/>
                    <button style={{...$.btn(C.amber),flexShrink:0}} onClick={async()=>{if(!newAdminPw.trim())return;await saveCfg({...cfg,adminPass:newAdminPw.trim()});alert("✅ Updated!");setNewAdminPw("");}}>Update</button>
                  </div>
                </div>

                {/* D. Archive — download/delete old uploads */}
                <div style={{...$.card,marginBottom:16}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14,flexWrap:"wrap",gap:8}}>
                    <h3 style={{fontFamily:"'Sora',sans-serif",fontWeight:700,fontSize:16,margin:0}}>🗄️ Archive — Previous Uploads</h3>
                    <button style={{...$.btn(C.lt,C.muted),border:`1px solid ${C.bdr}`,padding:"5px 12px",fontSize:12}}
                      onClick={()=>setArchOpen(o=>!o)}>{archOpen?"▲ Hide":"▼ Show"} ({archIdx.length} archived)</button>
                  </div>
                  {archOpen&&(
                    archIdx.length===0
                    ? <p style={{color:C.muted,fontSize:13}}>No archived files yet. Previous uploads are saved here when you import new data.</p>
                    : <div style={{display:"flex",flexDirection:"column",gap:8}}>
                        {[...archIdx].reverse().map((a,i)=>(
                          <div key={a.key} style={{display:"flex",alignItems:"center",justifyContent:"space-between",
                            padding:"10px 14px",background:C.lt,borderRadius:9,border:`1px solid ${C.bdr}`,flexWrap:"wrap",gap:8}}>
                            <div>
                              <div style={{fontWeight:600,fontSize:13,marginBottom:2}}>
                                {a.type==="cases"?"📋 Case Data":"👥 User Master"} — {a.label}
                              </div>
                              <div style={{fontSize:12,color:C.muted}}>{a.count} records · {new Date(a.ts).toLocaleString("en-IN")}</div>
                            </div>
                            <div style={{display:"flex",gap:8,flexShrink:0}}>
                              <button style={{...$.btn("#EFF6FF","#1E40AF"),padding:"5px 12px",fontSize:12}} onClick={async()=>{
                                try {
                                  const r=await window.storage.get(a.key,true).catch(()=>null);
                                  if (!r?.value){alert("Archive file not found.");return;}
                                  const data=JSON.parse(r.value);
                                  const wb=XLSX.utils.book_new();
                                  if (a.type==="cases"){
                                    const rows=data.map(c=>({...c.data,role:c.role,doctor:c.doctorName,seq:c.seq,status:c.status,sheet:c.sheet}));
                                    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Cases");
                                  } else {
                                    const maxIds=Math.max(0,...data.map(u=>u.userIds?.length||0));
                                    const hdrs=["Name","Role","PIN",...Array.from({length:maxIds},(_,j)=>[`User ID ${j+1}`,`Password ${j+1}`]).flat()];
                                    const rows=data.map(u=>{const b=[u.name,ROLES[u.role]?.label||u.role,u.pin];const ids=(u.userIds||[]).flatMap(x=>[x.uid||"",x.password||""]);while(ids.length<maxIds*2)ids.push("");return [...b,...ids];});
                                    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([hdrs,...rows]),"User Master");
                                  }
                                  XLSX.writeFile(wb,`Archive_${a.type}_${new Date(a.ts).toLocaleDateString("en-IN").replace(/\//g,"-")}.xlsx`);
                                } catch(e){alert("Download failed: "+e.message);}
                              }}>⬇ Download</button>
                              {isAdmin&&(
                                <button style={{...$.btn("#FEE2E2",C.red),padding:"5px 12px",fontSize:12}} onClick={async()=>{
                                  if(!confirm(`Delete this archive entry? This cannot be undone.`))return;
                                  await window.storage.delete(a.key,true).catch(()=>{});
                                  const newIdx=archIdx.filter(x=>x.key!==a.key);
                                  await window.storage.set(SK.archIdx,JSON.stringify(newIdx),true).catch(()=>{});
                                  setArchIdx(newIdx);
                                }}>🗑 Delete</button>
                              )}
                            </div>
                          </div>
                        ))}
                      </div>
                  )}
                </div>

                {/* E. Danger zone */}
                <div style={$.card}>
                  <h3 style={{fontFamily:"'Sora',sans-serif",fontWeight:700,fontSize:16,marginBottom:8}}>⚠️ Danger Zone</h3>
                  <p style={{fontSize:13,color:C.muted,marginBottom:12}}>Admin only. These actions cannot be undone.</p>
                  <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                    <button style={$.btn("#FEE2E2",C.red)} onClick={async()=>{if(!confirm("Reset ALL progress and UID counts?"))return;await saveCases(cases.map(c=>({...c,status:"pending"})));await saveLog([]);await saveUidReg({});setSyncLog([]);setSyncMsg("");alert("✅ Reset.");}}>Reset All Progress</button>
                    <button style={$.btn("#FEE2E2",C.red)} onClick={async()=>{if(!confirm("Delete ALL cases from active database?"))return;await saveCases([]);await saveLog([]);await saveUidReg({});setSyncLog([]);setSyncMsg("");alert("✅ Deleted. Archive preserved.");}}>Clear Active Cases</button>
                    <button style={$.btn("#FEE2E2",C.red)} onClick={async()=>{if(!confirm("Delete ALL archived files permanently?"))return;for(const a of archIdx){await window.storage.delete(a.key,true).catch(()=>{});}await window.storage.delete(SK.archIdx,true).catch(()=>{});setArchIdx([]);alert("✅ All archives deleted.");}}>🗑 Clear All Archives</button>
                  </div>
                </div>
              </div>
            )}

          </div>
        </div>
      </div>
    );
  }

  return null;
}
