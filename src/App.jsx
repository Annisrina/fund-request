import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── GANTI DENGAN URL APPS SCRIPT KAMU ───────────────────────────────────────
const API_URL = "https://script.google.com/macros/s/AKfycbz3_FHxZpkBUNr-GMStg0DDAX7-pL8D6EAlcwEh_bZYWOprGYivpktXTvw87HpDYMKiLQ/exec";

const SPV_PASSWORD   = "SPV";
const ADMIN_PASSWORD = "ADM";
const DEPARTMENTS    = ["Finance","SCM","HC","Marketing","Operational"];

const C = {
  bg:"#0F1117", card:"#1A1D27", border:"#2A2D3A",
  accent:"#4F8EF7", accent2:"#38D9A9", warn:"#F7A94F",
  danger:"#F76F6F", text:"#E8EAF0", muted:"#7B7F96",
  green:"#38D9A9", yellow:"#F7A94F", red:"#F76F6F",
};

const css = `
  @import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700&family=Space+Mono:wght@400;700&display=swap');
  *{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}
  html,body{background:${C.bg};color:${C.text};font-family:'Sora',sans-serif;min-height:100vh;-webkit-text-size-adjust:100%}
  ::-webkit-scrollbar{width:4px;height:4px}
  ::-webkit-scrollbar-track{background:${C.bg}}
  ::-webkit-scrollbar-thumb{background:${C.border};border-radius:3px}

  input,select,textarea{
    background:${C.bg};color:${C.text};border:1.5px solid ${C.border};
    border-radius:12px;padding:14px 16px;font-family:'Sora',sans-serif;
    font-size:16px;outline:none;width:100%;transition:border .2s;
    -webkit-appearance:none;appearance:none;
  }
  input:focus,select:focus,textarea:focus{border-color:${C.accent};box-shadow:0 0 0 3px rgba(79,142,247,.15)}
  input[type=file]{padding:14px;cursor:pointer;font-size:14px}
  select{background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%237B7F96' d='M6 8L1 3h10z'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 14px center;padding-right:36px}

  button{
    cursor:pointer;font-family:'Sora',sans-serif;font-size:15px;
    border:none;border-radius:12px;padding:13px 20px;transition:all .15s;
    -webkit-tap-highlight-color:transparent;user-select:none;
    active:opacity:.7;
  }
  button:active{opacity:.75;transform:scale(.97)}

  table{width:100%;border-collapse:collapse;font-size:13px}
  th{background:${C.border};color:${C.muted};font-weight:600;padding:11px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.5px;white-space:nowrap}
  td{padding:11px 12px;border-bottom:1px solid ${C.border};vertical-align:middle}

  .badge{display:inline-block;padding:4px 11px;border-radius:20px;font-size:11px;font-weight:700;letter-spacing:.3px;white-space:nowrap}
  .b-pending{background:rgba(247,169,79,.18);color:${C.yellow}}
  .b-spv{background:rgba(79,142,247,.2);color:${C.accent}}
  .b-approved{background:rgba(56,217,169,.18);color:${C.green}}
  .b-rejected{background:rgba(247,111,111,.18);color:${C.red}}
  .b-dept{background:rgba(79,142,247,.13);color:${C.accent}}
  .b-rmb{background:rgba(56,217,169,.12);color:${C.accent2}}
  .b-real{background:rgba(247,169,79,.13);color:${C.warn}}

  .card{background:${C.card};border:1px solid ${C.border};border-radius:16px;padding:18px}
  .label{font-size:13px;color:${C.muted};margin-bottom:8px;font-weight:600}
  .sec{font-size:11px;text-transform:uppercase;letter-spacing:1.4px;color:${C.muted};margin-bottom:14px;font-weight:700}
  .flow{background:rgba(79,142,247,.07);border-left:3px solid ${C.accent};border-radius:0 10px 10px 0;padding:12px 16px;margin-bottom:16px;font-size:13px;line-height:1.9}
  .flow li{list-style:none}
  .flow li::before{content:"▸ ";color:${C.accent}}

  .step-track{display:flex;align-items:center;gap:0;font-size:11px}
  .step{padding:5px 10px;border-radius:20px;font-weight:700;white-space:nowrap;font-size:11px}
  .step-line{flex:1;height:2px;background:${C.border};min-width:8px}
  .step-line.done{background:${C.accent}}

  .spinner{display:inline-block;width:16px;height:16px;border:2.5px solid ${C.border};border-top-color:${C.accent};border-radius:50%;animation:spin .7s linear infinite;vertical-align:middle}
  @keyframes spin{to{transform:rotate(360deg)}}

  .toast{position:fixed;bottom:90px;left:50%;transform:translateX(-50%);padding:13px 22px;border-radius:12px;font-size:14px;font-weight:600;z-index:3000;animation:fadeUp .3s;white-space:nowrap;box-shadow:0 4px 20px rgba(0,0,0,.4)}
  @keyframes fadeUp{from{opacity:0;transform:translateX(-50%) translateY(10px)}to{opacity:1;transform:translateX(-50%) translateY(0)}}

  /* bottom nav */
  .bot-nav{position:fixed;bottom:0;left:0;right:0;background:${C.card};border-top:1px solid ${C.border};display:flex;z-index:200;padding-bottom:env(safe-area-inset-bottom)}
  .bot-nav-item{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:10px 4px 8px;gap:3px;cursor:pointer;border:none;background:none;font-family:'Sora',sans-serif;transition:all .15s}
  .bot-nav-item span.icon{font-size:22px;line-height:1}
  .bot-nav-item span.lbl{font-size:10px;font-weight:600;letter-spacing:.2px}

  /* card list item for mobile */
  .list-card{background:${C.card};border:1px solid ${C.border};border-radius:14px;padding:14px 16px;margin-bottom:10px}
  .list-card-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px}
  .list-card-id{font-family:'Space Mono',monospace;font-size:11px;color:${C.accent}}
  .list-card-body{display:grid;grid-template-columns:1fr 1fr;gap:6px 12px;font-size:13px;margin-bottom:10px}
  .list-card-body .k{color:${C.muted};font-size:11px;font-weight:600}
  .list-card-body .v{color:${C.text};font-weight:500}
  .list-card-actions{display:flex;gap:8px;margin-top:10px;flex-wrap:wrap}

  @media(max-width:640px){
    .card{padding:14px}
    .hide-mobile{display:none!important}
  }
  @media(min-width:641px){
    .bot-nav{display:none}
    .desktop-tabs{display:flex!important}
    .main-content{padding-bottom:20px!important}
  }
  .desktop-tabs{display:none}
`;

const fmtRp = v => "Rp " + Number(v||0).toLocaleString("id-ID");
const today = () => new Date().toISOString().slice(0,10);
const uid   = () => Math.random().toString(36).slice(2,8).toUpperCase();

async function apiGet(){
  const r=await fetch(API_URL);return r.json();
}
async function apiPost(body){
  const r=await fetch(API_URL,{method:"POST",body:JSON.stringify(body)});return r.json();
}
async function fileToBase64(file){
  return new Promise((res,rej)=>{const rd=new FileReader();rd.onload=()=>res(rd.result);rd.onerror=rej;rd.readAsDataURL(file);});
}

// ── MICRO COMPONENTS ──────────────────────────────────────────────────────────
function Btn({children,variant="primary",loading,disabled,...p}){
  const s={
    primary:{background:C.accent,color:C.bg},
    success:{background:C.green,color:C.bg},
    danger:{background:C.red,color:"#fff"},
    warn:{background:C.warn,color:C.bg},
    ghost:{background:"rgba(255,255,255,.05)",color:C.muted,border:`1px solid ${C.border}`},
    tab:{background:"transparent",color:C.muted},
  }[variant]||{background:C.accent,color:C.bg};
  return(
    <button {...p} disabled={loading||disabled} style={{fontWeight:700,opacity:(loading||disabled)?.55:1,...s,...(p.style||{})}}>
      {loading&&<><span className="spinner"/> </>}{children}
    </button>
  );
}
function Field({label,children,hint}){
  return(
    <div style={{marginBottom:16}}>
      <div className="label">{label}{hint&&<span style={{color:C.muted,fontWeight:400,fontSize:11,marginLeft:6}}>{hint}</span>}</div>
      {children}
    </div>
  );
}
function Empty({text="Belum ada data"}){return <div style={{color:C.muted,fontSize:14,padding:"28px 0",textAlign:"center"}}>{text}</div>;}
function Spinner(){return <div style={{textAlign:"center",padding:40}}><span className="spinner" style={{width:32,height:32,borderWidth:3}}/></div>;}
function Toast({toast}){
  if(!toast)return null;
  return <div className="toast" style={{background:toast.type==="error"?C.red:C.green,color:toast.type==="error"?"#fff":C.bg}}>{toast.msg}</div>;
}

function StepTrack({spvStatus,financeAcc}){
  const spvOk=spvStatus==="Approved",spvFail=spvStatus==="Rejected";
  const finOk=financeAcc==="Approved",finFail=financeAcc==="Rejected";
  return(
    <div className="step-track">
      <div className="step" style={{background:"rgba(56,217,169,.15)",color:C.green}}>Staff</div>
      <div className={`step-line${spvOk||spvFail?" done":""}`}/>
      <div className="step" style={{background:spvFail?"rgba(247,111,111,.15)":spvOk?"rgba(79,142,247,.2)":"rgba(247,169,79,.1)",color:spvFail?C.red:spvOk?C.accent:C.yellow}}>
        {spvFail?"✗SPV":spvOk?"✓SPV":"SPV"}
      </div>
      <div className={`step-line${finOk||finFail?" done":""}`}/>
      <div className="step" style={{background:finFail?"rgba(247,111,111,.15)":finOk?"rgba(56,217,169,.15)":"rgba(123,127,150,.08)",color:finFail?C.red:finOk?C.green:C.muted,opacity:spvOk?1:.4}}>
        {finFail?"✗Fin":finOk?"✓Fin":"Fin"}
      </div>
    </div>
  );
}

function ImageViewer({src,onClose}){
  if(!src)return null;
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.95)",zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={onClose}>
      <img src={src} alt="bukti" style={{maxWidth:"100%",maxHeight:"85vh",borderRadius:14,border:`1px solid ${C.border}`}}/>
      <button onClick={onClose} style={{position:"absolute",top:20,right:20,background:"rgba(255,255,255,.1)",color:"#fff",fontSize:22,border:"none",cursor:"pointer",borderRadius:"50%",width:44,height:44,display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
    </div>
  );
}

// ── CARD ITEM (mobile-friendly list item) ─────────────────────────────────────
function AdvCard({x,isSpv,isAdmin,onSpvAcc,onFinanceAcc}){
  return(
    <div className="list-card">
      <div className="list-card-header">
        <div>
          <div className="list-card-id">{x.ID}</div>
          <div style={{fontWeight:700,fontSize:15,marginTop:2}}>{x.Item}</div>
        </div>
        <StepTrack spvStatus={x.SpvStatus} financeAcc={x.FinanceAcc}/>
      </div>
      <div className="list-card-body">
        <div><div className="k">Nama</div><div className="v">{x.Nama}</div></div>
        <div><div className="k">Tanggal</div><div className="v">{x.Tanggal}</div></div>
        <div><div className="k">Qty</div><div className="v">{x.Qty}</div></div>
        <div><div className="k">Est.</div><div className="v">{fmtRp(x.TotalEst)}</div></div>
        <div><div className="k">Status SPV</div><div className="v">
          {x.SpvStatus==="Approved"?<span className="badge b-spv">✓ SPV Acc</span>:x.SpvStatus==="Rejected"?<span className="badge b-rejected">✗ Ditolak</span>:<span className="badge b-pending">Pending</span>}
        </div></div>
        <div><div className="k">Finance</div><div className="v">
          {x.FinanceAcc==="Approved"?<span className="badge b-approved">✓ Approved</span>:x.FinanceAcc==="Rejected"?<span className="badge b-rejected">✗ Ditolak</span>:x.SpvStatus==="Approved"?<span className="badge b-pending">Menunggu</span>:<span style={{color:C.muted,fontSize:12}}>—</span>}
        </div></div>
      </div>
      {x.PhotoUrl&&<a href={x.PhotoUrl} target="_blank" rel="noreferrer" style={{color:C.accent,fontSize:13,fontWeight:600}}>🖼 Lihat Foto Nota</a>}
      <div className="list-card-actions">
        {isSpv&&x.SpvStatus==="Pending"&&(
          <>
            <Btn variant="success" style={{flex:1,padding:"11px"}} onClick={()=>onSpvAcc(x.ID,"Approved")}>✓ SPV Approve</Btn>
            <Btn variant="danger"  style={{flex:1,padding:"11px"}} onClick={()=>onSpvAcc(x.ID,"Rejected")}>✗ Tolak</Btn>
          </>
        )}
        {isAdmin&&x.SpvStatus==="Approved"&&x.FinanceAcc==="—"&&(
          <>
            <Btn variant="success" style={{flex:1,padding:"11px"}} onClick={()=>onFinanceAcc(x.ID,"Approved")}>✓ Finance Acc</Btn>
            <Btn variant="danger"  style={{flex:1,padding:"11px"}} onClick={()=>onFinanceAcc(x.ID,"Rejected")}>✗ Tolak</Btn>
          </>
        )}
        {isAdmin&&x.SpvStatus!=="Approved"&&x.FinanceAcc==="—"&&(
          <div style={{color:C.muted,fontSize:12,padding:"4px 0"}}>⏳ Tunggu SPV terlebih dahulu</div>
        )}
      </div>
    </div>
  );
}

function RmbCard({x,isSpv,isAdmin,onSpvAcc,onFinanceAcc}){
  return(
    <div className="list-card">
      <div className="list-card-header">
        <div>
          <div className="list-card-id">{x.ID} · <span className={`badge ${x.Tipe==="Reimburse"?"b-rmb":"b-real"}`}>{x.Tipe==="Reimburse"?"Reimburse":"Realisasi"}</span></div>
          <div style={{fontWeight:700,fontSize:15,marginTop:4}}>{x.Item}</div>
        </div>
        <StepTrack spvStatus={x.SpvStatus} financeAcc={x.FinanceAcc}/>
      </div>
      <div className="list-card-body">
        <div><div className="k">Nama</div><div className="v">{x.Nama}</div></div>
        <div><div className="k">Tanggal</div><div className="v">{x.Tanggal}</div></div>
        <div><div className="k">Qty</div><div className="v">{x.Qty}</div></div>
        <div><div className="k">Total Bayar</div><div className="v">{fmtRp(x.TotalPay)}</div></div>
        <div><div className="k">Status SPV</div><div className="v">
          {x.SpvStatus==="Approved"?<span className="badge b-spv">✓ SPV Acc</span>:x.SpvStatus==="Rejected"?<span className="badge b-rejected">✗ Ditolak</span>:<span className="badge b-pending">Pending</span>}
        </div></div>
        <div><div className="k">Finance</div><div className="v">
          {x.FinanceAcc==="Approved"?<span className="badge b-approved">✓ Approved</span>:x.FinanceAcc==="Rejected"?<span className="badge b-rejected">✗ Ditolak</span>:x.SpvStatus==="Approved"?<span className="badge b-pending">Menunggu</span>:<span style={{color:C.muted,fontSize:12}}>—</span>}
        </div></div>
      </div>
      {x.PhotoUrl&&<a href={x.PhotoUrl} target="_blank" rel="noreferrer" style={{color:C.accent,fontSize:13,fontWeight:600}}>🖼 Lihat Bukti</a>}
      <div className="list-card-actions">
        {isSpv&&x.SpvStatus==="Pending"&&(
          <>
            <Btn variant="success" style={{flex:1,padding:"11px"}} onClick={()=>onSpvAcc(x.ID,"Approved")}>✓ SPV Approve</Btn>
            <Btn variant="danger"  style={{flex:1,padding:"11px"}} onClick={()=>onSpvAcc(x.ID,"Rejected")}>✗ Tolak</Btn>
          </>
        )}
        {isAdmin&&x.SpvStatus==="Approved"&&x.FinanceAcc==="—"&&(
          <>
            <Btn variant="success" style={{flex:1,padding:"11px"}} onClick={()=>onFinanceAcc(x.ID,"Approved")}>✓ Finance Acc</Btn>
            <Btn variant="danger"  style={{flex:1,padding:"11px"}} onClick={()=>onFinanceAcc(x.ID,"Rejected")}>✗ Tolak</Btn>
          </>
        )}
        {isAdmin&&x.SpvStatus!=="Approved"&&x.FinanceAcc==="—"&&(
          <div style={{color:C.muted,fontSize:12,padding:"4px 0"}}>⏳ Tunggu SPV terlebih dahulu</div>
        )}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
//  LOGIN SCREEN
// ══════════════════════════════════════════════════════════════════
function LoginScreen({deptSel,setDeptSel,roleSel,setRoleSel,step,pwInput,setPwInput,err,onLogin,loading}){
  return(
    <>
      <style>{css}</style>
      <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:C.bg,padding:20}}>
        <div style={{position:"fixed",inset:0,backgroundImage:`linear-gradient(${C.border} 1px,transparent 1px),linear-gradient(90deg,${C.border} 1px,transparent 1px)`,backgroundSize:"40px 40px",opacity:.2,pointerEvents:"none"}}/>
        <div style={{position:"fixed",inset:0,background:`radial-gradient(ellipse 70% 50% at 50% 40%,rgba(79,142,247,.1) 0%,transparent 70%)`,pointerEvents:"none"}}/>
        <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:22,padding:28,width:"100%",maxWidth:440,position:"relative"}}>
          <div style={{textAlign:"center",marginBottom:28}}>
            <div style={{width:64,height:64,borderRadius:18,background:`linear-gradient(135deg,${C.accent},${C.accent2})`,display:"inline-flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:24,color:C.bg,marginBottom:12}}>FR</div>
            <div style={{fontWeight:800,fontSize:22}}>Fund Request</div>
            <div style={{color:C.muted,fontSize:13,marginTop:4}}>Sistem Pengajuan Dana Internal</div>
          </div>

          {step==="dept"&&(
            <>
              <Field label="Departemen">
                <select value={deptSel} onChange={e=>setDeptSel(e.target.value)}>
                  {DEPARTMENTS.map(d=><option key={d}>{d}</option>)}
                </select>
              </Field>
              <Field label="Masuk sebagai">
                <select value={roleSel} onChange={e=>setRoleSel(e.target.value)}>
                  <option>Staff</option>
                  <option>SPV</option>
                  {deptSel==="Finance"&&<option>Admin Finance</option>}
                </select>
              </Field>
            </>
          )}
          {step==="role"&&(
            <>
              <div style={{marginBottom:12,fontSize:14,color:C.muted,background:"rgba(79,142,247,.08)",borderRadius:10,padding:"10px 14px"}}>
                Login sebagai <b style={{color:C.text}}>{roleSel}</b> · Dept <b style={{color:C.accent}}>{deptSel}</b>
              </div>
              <Field label="Password" hint="(3 digit)">
                <input type="password" maxLength={10} value={pwInput} onChange={e=>setPwInput(e.target.value)} onKeyDown={e=>e.key==="Enter"&&onLogin()} placeholder="···" autoFocus inputMode="text" autoComplete="off"/>
              </Field>
              {err&&<div style={{color:C.red,fontSize:13,marginBottom:12,padding:"8px 12px",background:"rgba(247,111,111,.1)",borderRadius:8}}>{err}</div>}
            </>
          )}

          <Btn loading={loading} onClick={onLogin} style={{width:"100%",padding:"15px",fontSize:16}}>
            {step==="dept"&&roleSel==="Staff"?"Masuk sebagai Staff →":"Lanjut →"}
          </Btn>

          <div style={{marginTop:18,padding:14,background:"rgba(79,142,247,.07)",borderRadius:12,fontSize:13,color:C.muted,lineHeight:1.8}}>
            <b style={{color:C.accent}}>Alur Persetujuan:</b><br/>
            📝 Staff kirim → 🔑 <b style={{color:C.accent}}>SPV acc</b> → 👑 <b style={{color:C.accent2}}>Finance acc final</b>
          </div>
        </div>
      </div>
    </>
  );
}

// ══════════════════════════════════════════════════════════════════
//  APP ROOT
// ══════════════════════════════════════════════════════════════════
export default function App(){
  const [dept,setDept]           = useState(null);
  const [role,setRole]           = useState(null);
  const [loginStep,setLoginStep] = useState("dept");
  const [pwInput,setPwInput]     = useState("");
  const [loginErr,setLoginErr]   = useState("");
  const [deptSel,setDeptSel]     = useState(DEPARTMENTS[0]);
  const [roleSel,setRoleSel]     = useState("Staff");
  const [sheet,setSheet]         = useState("advance");
  const [imgView,setImgView]     = useState(null);
  const [toast,setToast]         = useState(null);
  const [advances,setAdvances]   = useState([]);
  const [reimburse,setReimburse] = useState([]);
  const [dataLoading,setDataLoading] = useState(false);

  const showToast = useCallback((msg,type="success")=>{
    setToast({msg,type}); setTimeout(()=>setToast(null),3000);
  },[]);

  const loadData = useCallback(async()=>{
    if(API_URL==="PASTE_APPS_SCRIPT_URL_DISINI") return;
    setDataLoading(true);
    try{
      const res=await apiGet();
      if(res.ok){ setAdvances(res.advances||[]); setReimburse(res.reimburse||[]); }
    }catch(e){ showToast("Gagal load data","error"); }
    setDataLoading(false);
  },[showToast]);

  useEffect(()=>{ if(role){ loadData(); const t=setInterval(loadData,30000); return()=>clearInterval(t); } },[role,loadData]);

  function handleLogin(){
    setLoginErr("");
    if(loginStep==="dept"){
      if(roleSel==="Staff"){ setDept(deptSel);setRole("staff");setLoginStep("done");return; }
      setLoginStep("role"); return;
    }
    if(roleSel==="Admin Finance"){
      if(pwInput!==ADMIN_PASSWORD){setLoginErr("Password salah");return;}
      setDept("Finance");setRole("finance_admin");setLoginStep("done");
    } else {
      if(pwInput!==SPV_PASSWORD){setLoginErr("Password SPV salah");return;}
      setDept(deptSel);setRole("spv");setLoginStep("done");
    }
  }
  function handleLogout(){ setDept(null);setRole(null);setLoginStep("dept");setPwInput("");setLoginErr("");setRoleSel("Staff");setAdvances([]);setReimburse([]); }

  async function submitAdvance(data){
    const id="ADV-"+uid();
    const photoBase64=data.photo?await fileToBase64(data.photo):"";
    const res=await apiPost({action:"submitAdvance",id,dept,...data,photoBase64,photoName:data.photo?.name||""});
    if(res.ok){ await loadData(); showToast("✓ Pengajuan advance terkirim!"); return true; }
    showToast(res.error||"Gagal kirim","error"); return false;
  }
  async function spvAdvance(id,status){
    const res=await apiPost({action:"spvAdvance",id,status});
    if(res.ok){ await loadData(); showToast(`✓ SPV ${status}`); }
    else showToast(res.error||"Gagal","error");
  }
  async function financeAdvance(id,val){
    const res=await apiPost({action:"financeAdvance",id,val});
    if(res.ok){ await loadData(); showToast(`✓ Finance: ${val}`); }
    else showToast(res.error||"Gagal","error");
  }
  async function submitReimburse(data){
    const id="RMB-"+uid();
    const photoBase64=data.photo?await fileToBase64(data.photo):"";
    const res=await apiPost({action:"submitReimburse",id,dept,...data,photoBase64,photoName:data.photo?.name||""});
    if(res.ok){ await loadData(); showToast("✓ Pengajuan terkirim!"); return true; }
    showToast(res.error||"Gagal kirim","error"); return false;
  }
  async function spvReimburse(id,status){
    const res=await apiPost({action:"spvReimburse",id,status});
    if(res.ok){ await loadData(); showToast(`✓ SPV ${status}`); }
    else showToast(res.error||"Gagal","error");
  }
  async function financeReimburse(id,val){
    const res=await apiPost({action:"financeReimburse",id,val});
    if(res.ok){ await loadData(); showToast(`✓ Finance: ${val}`); }
    else showToast(res.error||"Gagal","error");
  }

  function exportXLSX(){
    const wb=XLSX.utils.book_new();
    const aR=advances.map(x=>({ID:x.ID,Dept:x.Dept,Nama:x.Nama,Tanggal:x.Tanggal,Item:x.Item,Qty:x.Qty,"Est(Rp)":x.TotalEst,"Status SPV":x.SpvStatus,"Acc Finance":x.FinanceAcc}));
    const rR=reimburse.map(x=>({ID:x.ID,Dept:x.Dept,Tipe:x.Tipe,Nama:x.Nama,Tanggal:x.Tanggal,Item:x.Item,Qty:x.Qty,"Bayar(Rp)":x.TotalPay,"Ref":x.AdvanceRef||"-","Status SPV":x.SpvStatus,"Acc Finance":x.FinanceAcc}));
    const sum=DEPARTMENTS.map(d=>({Dept:d,"Adv Total":advances.filter(x=>x.Dept===d).length,"Adv Fin OK":advances.filter(x=>x.Dept===d&&x.FinanceAcc==="Approved").length,"Est Advance":advances.filter(x=>x.Dept===d).reduce((s,x)=>s+(+x.TotalEst||0),0),"Rmb Total":reimburse.filter(x=>x.Dept===d).length,"Rmb Fin OK":reimburse.filter(x=>x.Dept===d&&x.FinanceAcc==="Approved").length,"Total Bayar":reimburse.filter(x=>x.Dept===d).reduce((s,x)=>s+(+x.TotalPay||0),0)}));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(aR.length?aR:[{note:"kosong"}]),"Advance");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rR.length?rR:[{note:"kosong"}]),"Reimburse_Realisasi");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(sum),"Dashboard");
    XLSX.writeFile(wb,`FundRequest_${today()}.xlsx`);
  }

  if(loginStep!=="done") return <LoginScreen deptSel={deptSel} setDeptSel={setDeptSel} roleSel={roleSel} setRoleSel={setRoleSel} step={loginStep} pwInput={pwInput} setPwInput={setPwInput} err={loginErr} onLogin={handleLogin} loading={false}/>;

  const isSpv=role==="spv", isAdmin=role==="finance_admin";
  const roleLabel=isAdmin?"👑 Admin Finance":isSpv?"🔑 SPV":"Staff";

  const TABS=[
    {id:"advance",  icon:"📋", label:"Advance"},
    {id:"reimburse",icon:"💸", label:"Reimburse"},
    {id:"dashboard",icon:"📊", label:"Dashboard"},
  ];

  return(
    <>
      <style>{css}</style>
      <ImageViewer src={imgView} onClose={()=>setImgView(null)}/>
      <Toast toast={toast}/>

      {API_URL==="PASTE_APPS_SCRIPT_URL_DISINI"&&(
        <div style={{background:C.warn,color:C.bg,textAlign:"center",padding:"11px 16px",fontSize:13,fontWeight:700,lineHeight:1.5}}>
          ⚠️ Belum terhubung ke Google Sheet<br/>
          <span style={{fontWeight:400,fontSize:12}}>Ganti API_URL di baris 5 dengan URL Apps Script kamu</span>
        </div>
      )}

      {/* TOPBAR */}
      <div style={{background:C.card,borderBottom:`1px solid ${C.border}`,padding:"12px 18px",display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:36,height:36,borderRadius:10,background:`linear-gradient(135deg,${C.accent},${C.accent2})`,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:15,color:C.bg,flexShrink:0}}>FR</div>
          <div>
            <div style={{fontWeight:700,fontSize:14,lineHeight:1.2}}>Fund Request</div>
            <div style={{fontSize:11,color:C.muted}}>{dept} · {roleLabel}</div>
          </div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          {dataLoading&&<span className="spinner"/>}
          <button onClick={loadData} style={{background:"none",border:`1px solid ${C.border}`,color:C.muted,borderRadius:8,padding:"7px 10px",fontSize:15}}>↻</button>
          <button onClick={handleLogout} style={{background:"none",border:`1px solid ${C.border}`,color:C.muted,borderRadius:8,padding:"7px 12px",fontSize:13,fontWeight:600}}>Keluar</button>
        </div>
      </div>

      {/* DESKTOP TABS */}
      <div className="desktop-tabs" style={{background:C.card,borderBottom:`1px solid ${C.border}`,padding:"0 20px"}}>
        {TABS.map(t=>(
          <button key={t.id} onClick={()=>setSheet(t.id)} style={{background:"none",border:"none",borderBottom:sheet===t.id?`2px solid ${C.accent}`:"2px solid transparent",color:sheet===t.id?C.accent:C.muted,padding:"13px 20px",fontWeight:sheet===t.id?700:400,fontSize:14,cursor:"pointer",fontFamily:"'Sora',sans-serif"}}>
            {t.icon} {t.label}
          </button>
        ))}
      </div>

      {/* MAIN CONTENT */}
      <div className="main-content" style={{flex:1,padding:"16px",paddingBottom:90,maxWidth:900,margin:"0 auto",width:"100%"}}>
        {sheet==="advance"   && <AdvanceSheet   dept={dept} isSpv={isSpv} isAdmin={isAdmin} advances={advances} onSubmit={submitAdvance} onSpvAcc={spvAdvance} onFinanceAcc={financeAdvance} onViewImg={setImgView}/>}
        {sheet==="reimburse" && <ReimburseSheet dept={dept} isSpv={isSpv} isAdmin={isAdmin} advances={advances} reimburse={reimburse} onSubmit={submitReimburse} onSpvAcc={spvReimburse} onFinanceAcc={financeReimburse} onViewImg={setImgView}/>}
        {sheet==="dashboard" && <DashboardSheet isAdmin={isAdmin} advances={advances} reimburse={reimburse} onExport={exportXLSX} onFinanceAdvance={financeAdvance} onFinanceReimburse={financeReimburse}/>}
      </div>

      {/* BOTTOM NAV (mobile/iPad) */}
      <nav className="bot-nav">
        {TABS.map(t=>(
          <button key={t.id} className="bot-nav-item" onClick={()=>setSheet(t.id)}
            style={{color:sheet===t.id?C.accent:C.muted,borderTop:sheet===t.id?`2px solid ${C.accent}`:"2px solid transparent"}}>
            <span className="icon">{t.icon}</span>
            <span className="lbl">{t.label}</span>
          </button>
        ))}
      </nav>
    </>
  );
}

// ══════════════════════════════════════════════════════════════════
//  ADVANCE SHEET
// ══════════════════════════════════════════════════════════════════
function AdvanceSheet({dept,isSpv,isAdmin,advances,onSubmit,onSpvAcc,onFinanceAcc,onViewImg}){
  const [viewMode,setViewMode]=useState((isSpv||isAdmin)?"list":"form");
  const [form,setForm]=useState({name:"",date:today(),item:"",qty:1,totalEst:"",photo:null,photoUrl:null});
  const [err,setErr]=useState("");const [loading,setLoading]=useState(false);
  const myList=advances.filter(x=>x.Dept===dept);

  function handleFile(e){const f=e.target.files[0];if(!f)return;setForm(p=>({...p,photo:f,photoUrl:URL.createObjectURL(f)}));}
  async function handleSubmit(){
    const{name,date,item,qty,totalEst,photo}=form;
    if(!name||!date||!item||!qty||!totalEst||!photo){setErr("Semua field wajib diisi termasuk foto.");return;}
    setLoading(true);setErr("");
    const ok=await onSubmit({name,date,item,qty,totalEst,photo});
    if(ok) setForm({name:"",date:today(),item:"",qty:1,totalEst:"",photo:null,photoUrl:null});
    setLoading(false);
  }

  return(
    <div>
      <div className="flow"><ul>
        <li>Ajukan uang muka <b>sebelum</b> pembelian + foto nota perkiraan</li>
        <li><b>Tahap 1 — SPV:</b> Setujui atau tolak</li>
        <li><b>Tahap 2 — Finance:</b> Acc final setelah SPV approve</li>
      </ul></div>

      {(isSpv||isAdmin)&&(
        <div style={{display:"flex",gap:8,marginBottom:14}}>
          <Btn variant={viewMode==="form"?"primary":"ghost"} onClick={()=>setViewMode("form")} style={{flex:1,padding:"12px"}}>📝 Form Baru</Btn>
          <Btn variant={viewMode==="list"?"primary":"ghost"} onClick={()=>setViewMode("list")} style={{flex:1,padding:"12px"}}>📋 Daftar ({myList.length})</Btn>
        </div>
      )}

      {viewMode==="form"&&(
        <div className="card">
          <div className="sec">Form Pengajuan Advance — {dept}</div>
          <Field label="Nama Pengaju">
            <input value={form.name} onChange={e=>setForm(p=>({...p,name:e.target.value}))} placeholder="Nama lengkap" autoComplete="name"/>
          </Field>
          <Field label="Tanggal">
            <input type="date" value={form.date} onChange={e=>setForm(p=>({...p,date:e.target.value}))}/>
          </Field>
          <Field label="Nama Pengajuan / Keperluan">
            <input value={form.item} onChange={e=>setForm(p=>({...p,item:e.target.value}))} placeholder="Contoh: Pembelian ATK Mei 2025"/>
          </Field>
          <Field label="Qty">
            <input type="number" min={1} value={form.qty} onChange={e=>setForm(p=>({...p,qty:e.target.value}))} inputMode="numeric"/>
          </Field>
          <Field label="Total Perkiraan (Rp)">
            <input type="number" value={form.totalEst} onChange={e=>setForm(p=>({...p,totalEst:e.target.value}))} placeholder="0" inputMode="numeric"/>
          </Field>
          <Field label="Foto Nota / Referensi" hint="wajib">
            <input type="file" accept="image/*" capture="environment" onChange={handleFile}/>
          </Field>
          {form.photoUrl&&(
            <div style={{marginBottom:14}}>
              <img src={form.photoUrl} alt="preview" style={{width:"100%",maxWidth:260,borderRadius:10,border:`1px solid ${C.border}`,cursor:"pointer"}} onClick={()=>onViewImg&&onViewImg(form.photoUrl)}/>
            </div>
          )}
          {err&&<div style={{color:C.red,fontSize:13,marginBottom:12,padding:"10px 14px",background:"rgba(247,111,111,.1)",borderRadius:10}}>{err}</div>}
          <Btn loading={loading} onClick={handleSubmit} style={{width:"100%",padding:"15px",fontSize:15}}>
            {loading?"Mengirim ke Google Sheet...":"Kirim Pengajuan"}
          </Btn>
        </div>
      )}

      {viewMode==="list"&&(
        <div>
          <div className="sec" style={{marginBottom:10}}>{myList.length} Pengajuan — {dept}</div>
          {myList.length===0?<Empty text="Belum ada pengajuan advance"/>:myList.map(x=>(
            <AdvCard key={x.ID} x={x} isSpv={isSpv} isAdmin={isAdmin} onSpvAcc={onSpvAcc} onFinanceAcc={onFinanceAcc}/>
          ))}
        </div>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
//  REIMBURSE & REALISASI SHEET
// ══════════════════════════════════════════════════════════════════
function ReimburseSheet({dept,isSpv,isAdmin,advances,reimburse,onSubmit,onSpvAcc,onFinanceAcc,onViewImg}){
  const [tab,setTab]=useState("reimburse");
  const [viewMode,setViewMode]=useState((isSpv||isAdmin)?"list":"form");
  const [form,setForm]=useState({name:"",date:today(),item:"",qty:1,totalPay:"",photo:null,photoUrl:null});
  const [advRef,setAdvRef]=useState("");
  const [err,setErr]=useState("");const [loading,setLoading]=useState(false);
  const approvedAdv=advances.filter(x=>x.Dept===dept&&x.SpvStatus==="Approved");
  const myList=reimburse.filter(x=>x.Dept===dept&&(tab==="reimburse"?x.Tipe==="Reimburse":x.Tipe==="Realisasi Advance"));

  function handleFile(e){const f=e.target.files[0];if(!f)return;setForm(p=>({...p,photo:f,photoUrl:URL.createObjectURL(f)}));}
  async function handleSubmit(){
    setLoading(true);setErr("");
    let data;
    if(tab==="realisasi"){
      if(!advRef){setErr("Pilih referensi advance.");setLoading(false);return;}
      const adv=advances.find(x=>x.ID===advRef);
      if(!form.totalPay||!form.photo){setErr("Isi total bayar dan upload bukti.");setLoading(false);return;}
      data={name:adv.Nama,date:form.date,item:adv.Item,qty:adv.Qty,totalPay:form.totalPay,photo:form.photo,type:"Realisasi Advance",advanceRef:advRef};
    } else {
      const{name,date,item,qty,totalPay,photo}=form;
      if(!name||!date||!item||!qty||!totalPay||!photo){setErr("Semua field wajib diisi.");setLoading(false);return;}
      data={name,date,item,qty,totalPay,photo,type:"Reimburse",advanceRef:""};
    }
    const ok=await onSubmit(data);
    if(ok){setForm({name:"",date:today(),item:"",qty:1,totalPay:"",photo:null,photoUrl:null});setAdvRef("");}
    setLoading(false);
  }

  return(
    <div>
      <div className="flow"><ul>
        <li><b>Reimburse:</b> Sudah beli pakai uang pribadi, minta penggantian</li>
        <li><b>Realisasi:</b> Pertanggungjawaban dana advance — data otomatis dari advance SPV Approved</li>
        <li>SPV acc dulu → Finance acc final</li>
      </ul></div>

      <div style={{display:"flex",gap:8,marginBottom:14}}>
        <Btn variant={tab==="reimburse"?"primary":"ghost"} onClick={()=>setTab("reimburse")} style={{flex:1,padding:"12px"}}>💳 Reimburse</Btn>
        <Btn variant={tab==="realisasi"?"warn":"ghost"} onClick={()=>setTab("realisasi")} style={{flex:1,padding:"12px"}}>🔄 Realisasi</Btn>
      </div>

      {(isSpv||isAdmin)&&(
        <div style={{display:"flex",gap:8,marginBottom:14}}>
          <Btn variant={viewMode==="form"?"primary":"ghost"} onClick={()=>setViewMode("form")} style={{flex:1,padding:"11px"}}>📝 Form Baru</Btn>
          <Btn variant={viewMode==="list"?"primary":"ghost"} onClick={()=>setViewMode("list")} style={{flex:1,padding:"11px"}}>📋 Daftar ({myList.length})</Btn>
        </div>
      )}

      {viewMode==="form"&&(
        <div className="card">
          <div className="sec">Form {tab==="reimburse"?"Reimburse":"Realisasi Advance"}</div>
          {tab==="realisasi"?(
            <>
              <Field label="Pilih Advance (SPV Approved)">
                <select value={advRef} onChange={e=>{setAdvRef(e.target.value);const a=advances.find(x=>x.ID===e.target.value);if(a)setForm(p=>({...p,name:a.Nama,item:a.Item,qty:a.Qty}));}}>
                  <option value="">— Pilih ID Advance —</option>
                  {approvedAdv.map(a=><option key={a.ID} value={a.ID}>{a.ID} · {a.Item}</option>)}
                </select>
              </Field>
              {advRef&&<div style={{background:"rgba(79,142,247,.08)",borderRadius:10,padding:"10px 14px",fontSize:13,marginBottom:14,color:C.muted}}>
                📎 Data dari: <b style={{color:C.text}}>{advances.find(x=>x.ID===advRef)?.Item}</b>
              </div>}
              <Field label="Tanggal Realisasi">
                <input type="date" value={form.date} onChange={e=>setForm(p=>({...p,date:e.target.value}))}/>
              </Field>
              <Field label="Total Dibayar (Rp)">
                <input type="number" value={form.totalPay} onChange={e=>setForm(p=>({...p,totalPay:e.target.value}))} placeholder="0" inputMode="numeric"/>
              </Field>
              <Field label="Upload Bukti / Nota" hint="wajib">
                <input type="file" accept="image/*" capture="environment" onChange={handleFile}/>
              </Field>
            </>
          ):(
            <>
              <Field label="Nama Pengaju">
                <input value={form.name} onChange={e=>setForm(p=>({...p,name:e.target.value}))} placeholder="Nama lengkap"/>
              </Field>
              <Field label="Tanggal">
                <input type="date" value={form.date} onChange={e=>setForm(p=>({...p,date:e.target.value}))}/>
              </Field>
              <Field label="Nama Pengajuan / Item">
                <input value={form.item} onChange={e=>setForm(p=>({...p,item:e.target.value}))} placeholder="Deskripsi item"/>
              </Field>
              <Field label="Qty">
                <input type="number" min={1} value={form.qty} onChange={e=>setForm(p=>({...p,qty:e.target.value}))} inputMode="numeric"/>
              </Field>
              <Field label="Total Bayar (Rp)">
                <input type="number" value={form.totalPay} onChange={e=>setForm(p=>({...p,totalPay:e.target.value}))} placeholder="0" inputMode="numeric"/>
              </Field>
              <Field label="Upload Bukti Pembayaran" hint="wajib">
                <input type="file" accept="image/*" capture="environment" onChange={handleFile}/>
              </Field>
            </>
          )}
          {form.photoUrl&&<div style={{marginBottom:14}}><img src={form.photoUrl} alt="preview" style={{width:"100%",maxWidth:260,borderRadius:10,border:`1px solid ${C.border}`}}/></div>}
          {err&&<div style={{color:C.red,fontSize:13,marginBottom:12,padding:"10px 14px",background:"rgba(247,111,111,.1)",borderRadius:10}}>{err}</div>}
          <Btn loading={loading} onClick={handleSubmit} style={{width:"100%",padding:"15px",fontSize:15}}>
            {loading?"Mengirim...":"Kirim Pengajuan"}
          </Btn>
        </div>
      )}

      {viewMode==="list"&&(
        <div>
          <div className="sec" style={{marginBottom:10}}>{myList.length} Pengajuan</div>
          {myList.length===0?<Empty text={`Belum ada ${tab==="reimburse"?"reimburse":"realisasi"}`}/>:myList.map(x=>(
            <RmbCard key={x.ID} x={x} isSpv={isSpv} isAdmin={isAdmin} onSpvAcc={onSpvAcc} onFinanceAcc={onFinanceAcc}/>
          ))}
        </div>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
//  DASHBOARD SHEET
// ══════════════════════════════════════════════════════════════════
function DashboardSheet({isAdmin,advances,reimburse,onExport,onFinanceAdvance,onFinanceReimburse}){
  const [filterDept,setFilterDept]=useState("All");
  const [filterStatus,setFilterStatus]=useState("All");

  function sm(x){
    if(filterStatus==="All") return true;
    if(filterStatus==="Pending") return x.SpvStatus==="Pending";
    if(filterStatus==="SPV Approved") return x.SpvStatus==="Approved"&&x.FinanceAcc==="—";
    if(filterStatus==="Finance Approved") return x.FinanceAcc==="Approved";
    if(filterStatus==="Rejected") return x.SpvStatus==="Rejected"||x.FinanceAcc==="Rejected";
    return true;
  }
  const af=advances.filter(x=>(filterDept==="All"||x.Dept===filterDept)&&sm(x));
  const rf=reimburse.filter(x=>(filterDept==="All"||x.Dept===filterDept)&&sm(x));
  const all=[...advances,...reimburse];
  const totalEst=advances.reduce((s,x)=>s+(+x.TotalEst||0),0);
  const totalBayar=reimburse.reduce((s,x)=>s+(+x.TotalPay||0),0);
  const pendingSPV=all.filter(x=>x.SpvStatus==="Pending").length;
  const pendingFin=all.filter(x=>x.SpvStatus==="Approved"&&x.FinanceAcc==="—").length;
  const finDone=all.filter(x=>x.FinanceAcc==="Approved").length;

  return(
    <div>
      <div className="flow"><ul>
        <li>Rangkuman semua pengajuan — data realtime dari Google Sheet</li>
        <li>Auto-refresh setiap 30 detik · tekan ↻ untuk refresh manual</li>
        {isAdmin&&<li>Admin Finance bisa Acc Final dan Download XLSX</li>}
      </ul></div>

      {/* STAT CARDS */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16}}>
        {[
          [fmtRp(totalEst),"Est. Advance",C.accent],
          [fmtRp(totalBayar),"Total Reimburse",C.accent2],
          [pendingSPV,"Pending SPV",C.yellow],
          [pendingFin,"Tunggu Finance",C.warn],
          [finDone,"Finance Approved",C.green],
          [advances.length+reimburse.length,"Total Pengajuan",C.muted],
        ].map(([v,l,c])=>(
          <div key={l} className="card" style={{textAlign:"center",padding:14}}>
            <div style={{fontSize:17,fontWeight:800,color:c,fontFamily:"'Space Mono',monospace",lineHeight:1.2}}>{v}</div>
            <div style={{fontSize:11,color:C.muted,marginTop:4,lineHeight:1.3}}>{l}</div>
          </div>
        ))}
      </div>

      {/* FILTER DEPT */}
      <div style={{overflowX:"auto",whiteSpace:"nowrap",marginBottom:10,paddingBottom:4}}>
        {["All",...DEPARTMENTS].map(d=>(
          <button key={d} onClick={()=>setFilterDept(d)} style={{background:filterDept===d?C.accent:"rgba(255,255,255,.04)",color:filterDept===d?C.bg:C.muted,border:`1px solid ${filterDept===d?C.accent:C.border}`,borderRadius:20,padding:"7px 14px",fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"'Sora',sans-serif",marginRight:6,display:"inline-block"}}>
            {d}
          </button>
        ))}
      </div>

      {/* FILTER STATUS + EXPORT */}
      <div style={{display:"flex",gap:8,marginBottom:16,alignItems:"center"}}>
        <select value={filterStatus} onChange={e=>setFilterStatus(e.target.value)} style={{flex:1,fontSize:14,padding:"11px 14px"}}>
          <option>All</option>
          <option>Pending</option>
          <option>SPV Approved</option>
          <option>Finance Approved</option>
          <option>Rejected</option>
        </select>
        {isAdmin&&<Btn variant="success" onClick={onExport} style={{whiteSpace:"nowrap",padding:"11px 14px",fontSize:13}}>⬇ XLSX</Btn>}
      </div>

      {/* DEPT SUMMARY */}
      <div className="card" style={{marginBottom:14}}>
        <div className="sec">Ringkasan per Departemen</div>
        {DEPARTMENTS.filter(d=>filterDept==="All"||d===filterDept).map(d=>{
          const da=advances.filter(x=>x.Dept===d), dr=reimburse.filter(x=>x.Dept===d);
          return(
            <div key={d} style={{borderBottom:`1px solid ${C.border}`,paddingBottom:12,marginBottom:12}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                <span className="badge b-dept" style={{fontSize:13,padding:"5px 14px"}}>{d}</span>
                <span style={{fontSize:12,color:C.muted}}>{da.length+dr.length} pengajuan</span>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,fontSize:12}}>
                <div style={{color:C.muted}}>Advance SPV ✓ <b style={{color:C.accent}}>{da.filter(x=>x.SpvStatus==="Approved").length}</b></div>
                <div style={{color:C.muted}}>Advance Fin ✓ <b style={{color:C.green}}>{da.filter(x=>x.FinanceAcc==="Approved").length}</b></div>
                <div style={{color:C.muted}}>Rmb SPV ✓ <b style={{color:C.accent}}>{dr.filter(x=>x.SpvStatus==="Approved").length}</b></div>
                <div style={{color:C.muted}}>Rmb Fin ✓ <b style={{color:C.green}}>{dr.filter(x=>x.FinanceAcc==="Approved").length}</b></div>
                <div style={{color:C.muted}}>Est. Adv <b style={{color:C.text}}>{fmtRp(da.reduce((s,x)=>s+(+x.TotalEst||0),0))}</b></div>
                <div style={{color:C.muted}}>Total Bayar <b style={{color:C.text}}>{fmtRp(dr.reduce((s,x)=>s+(+x.TotalPay||0),0))}</b></div>
              </div>
            </div>
          );
        })}
      </div>

      {/* DETAIL ADVANCE */}
      <div style={{marginBottom:6}}>
        <div className="sec">Detail Advance ({af.length})</div>
        {af.length===0?<Empty text="Tidak ada data advance"/>:af.map(x=>(
          <AdvCard key={x.ID} x={x} isSpv={false} isAdmin={isAdmin} onSpvAcc={()=>{}} onFinanceAcc={onFinanceAdvance}/>
        ))}
      </div>

      {/* DETAIL REIMBURSE */}
      <div>
        <div className="sec" style={{marginTop:16}}>Detail Reimburse & Realisasi ({rf.length})</div>
        {rf.length===0?<Empty text="Tidak ada data reimburse"/>:rf.map(x=>(
          <RmbCard key={x.ID} x={x} isSpv={false} isAdmin={isAdmin} onSpvAcc={()=>{}} onFinanceAcc={onFinanceReimburse}/>
        ))}
      </div>
    </div>
  );
}
