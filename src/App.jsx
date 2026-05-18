import { useMemo, useState, useEffect } from "react";
import { AreaChart, Area, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid, ReferenceArea, LineChart, Line, ReferenceLine, Legend, ComposedChart } from "recharts";
import * as XLSX from "xlsx";

// ─── Icons ────────────────────────────────────────────────────────────────────
const Icon = {
  Download:()=><span style={{fontSize:14}}>↓</span>,
  Lock:()=><span style={{fontSize:14}}>🔒</span>,
  Check:({color,size=14})=><span style={{color,fontSize:size}}>✓</span>,
  Alert:({color})=><span style={{color,fontSize:14}}>⚠</span>,
  File:({color})=><span style={{color,fontSize:14}}>📄</span>,
  Layers:({color})=><span style={{color,fontSize:14}}>📋</span>,
  ZoomOut:()=><span style={{fontSize:12}}>🔍</span>,
  Info:({color})=><span style={{color,fontSize:13}}>ℹ</span>,
  Circle:({fill,size=8})=><span style={{color:fill,fontSize:size}}>●</span>,
  Edit:()=><span style={{fontSize:13}}>✏️</span>,
};

// ─── Brand colors ─────────────────────────────────────────────────────────────
const WB = {
  bg:"#0d1b2a", card:"#152535", border:"#1e3448", borderLight:"#2a4560",
  primary:"#4a7c8e", primaryLight:"#8aabb8",
  text:"#f8fafc", textMuted:"#8aabb8", textDim:"#475569",
  green:"#22c55e", greenBg:"#052e16", greenBorder:"#14532d",
  red:"#ef4444", redBg:"#1a0505", redBorder:"#7f1d1d",
  yellow:"#eab308", yellowBg:"#1a1200", yellowBorder:"#854d0e",
  orange:"#f97316",
};
const PAL=["#4a7c8e","#8aabb8","#c8a96e","#5b8a6d","#7b9eb8","#c87d5b","#6b8a9e","#a8c4b8","#9e7b6e","#5b7a8e","#b8956e","#6e8e7b","#8e6e5b","#7ab8c4","#c4a87a","#e07b54","#54a0e0","#a0e054","#e054a0","#54e0a0"];

// ─── Helpers ──────────────────────────────────────────────────────────────────
function toMoney(v){if(!v&&v!==0)return 0;const n=Number(String(v).replace(/[$,\s]/g,""));return isFinite(n)?n:0;}
function toPercent(v){if(!v&&v!==0)return 0;const n=Number(String(v).replace(/[%\s]/g,""));if(!isFinite(n))return 0;return n>1?n/100:n;}
function toDate(v){
  if(!v&&v!==0)return null;
  if(v instanceof Date&&!isNaN(v))return v;
  if(typeof v==="number"){try{const p=XLSX.SSF.parse_date_code(v);return p?new Date(p.y,p.m-1,p.d):null;}catch{return null;}}
  const d=new Date(v);return isNaN(d.getTime())?null:d;
}
function fmtDate(d){return d?d.toISOString().slice(0,10):"—";}
function fmtMonth(s){if(!s||s==="—")return"";try{const d=new Date(s+"-01");return d.toLocaleDateString("en-US",{month:"short",year:"numeric"});}catch{return s;}}
function fmtShort(s){if(!s||s==="—")return"";try{const d=new Date(s);return d.toLocaleDateString("en-US",{month:"short",year:"2-digit"});}catch{return s;}}
function currency(v){return new Intl.NumberFormat("en-US",{style:"currency",currency:"USD",maximumFractionDigits:0}).format(v||0);}
function today(){return fmtDate(new Date());}
function monthKey(d){const dt=d instanceof Date?d:new Date(d);return`${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,"0")}`;}
function getText(node,tag){const el=node?.getElementsByTagName(tag)?.[0];return el?.textContent?.trim()||"";}

// ─── SOV parser ───────────────────────────────────────────────────────────────
const SUBTOTAL_RE=/grand|total|subtotal|sub-total|\bsum\b/i;
const GCGR_RE=/general\s*req|gen\s*req|\bGR\b|\bGC\b|general\s*condition|mobilization/i;
const OHP_RE=/oh\s*[&+]\s*p|overhead.*profit|markup|\bprofit\b|\bfee\b/i;
function classifyRow(desc,val){
  if(!desc&&!val)return"skip";if(val===0)return"skip";
  const d=String(desc||"");
  if(SUBTOTAL_RE.test(d))return"subtotal";
  if(OHP_RE.test(d))return"ohp";
  if(GCGR_RE.test(d))return"gcgr";
  return"line_item";
}
function parseSOV(arrayBuffer){
  const wb=XLSX.read(new Uint8Array(arrayBuffer),{type:"array"});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const raw=XLSX.utils.sheet_to_json(ws,{defval:"",header:1});
  const lineItems=[];let gcgrTotal=0,ohpTotal=0;
  raw.forEach(row=>{
    let desc="",val=0;
    row.forEach(cell=>{const s=String(cell||"").trim();const n=toMoney(s);if(!desc&&s.length>1&&isNaN(Number(s.replace(/[$,]/g,""))))desc=s;if(!val&&n>=100&&s!==desc)val=n;});
    if(!desc&&!val)return;
    const kind=classifyRow(desc,val);
    if(kind==="skip"||kind==="subtotal")return;
    if(kind==="gcgr"){gcgrTotal+=val;return;}
    if(kind==="ohp"){ohpTotal+=val;return;}
    if(val>0&&desc)lineItems.push({id:`sov-${lineItems.length}`,description:desc,baseValue:val});
  });
  if(!lineItems.length)return{items:[],gcgrTotal,ohpTotal,contractTotal:0,baseSum:0,parseWarning:"No line items found."};
  const baseSum=lineItems.reduce((s,i)=>s+i.baseValue,0);
  const items=lineItems.map(item=>{
    const share=baseSum>0?item.baseValue/baseSum:0;
    const gcgrShare=gcgrTotal*share,ohpShare=ohpTotal*share;
    return{...item,gcgrShare,ohpShare,scheduledValue:item.baseValue+gcgrShare+ohpShare,percentComplete:0};
  });
  return{items,gcgrTotal,ohpTotal,contractTotal:items.reduce((s,i)=>s+i.scheduledValue,0),baseSum,parseWarning:null};
}

// ─── MS Project XML parser (resource-aware) ───────────────────────────────────
function parseMSP(xmlText){
  const parser=new DOMParser();
  const doc=parser.parseFromString(xmlText,"application/xml");
  if(doc.querySelector("parsererror"))throw new Error("Invalid XML.");
  const resources=new Map();
  doc.querySelectorAll("Resource").forEach(r=>{
    const uid=getText(r,"UID");
    const name=getText(r,"Name");
    if(uid&&name&&uid!=="0")resources.set(uid,{uid,name});
  });
  const taskResources=new Map();
  doc.querySelectorAll("Assignment").forEach(a=>{
    const taskUID=getText(a,"TaskUID");
    const resourceUID=getText(a,"ResourceUID");
    if(!taskUID||!resourceUID||resourceUID==="-65535"||resourceUID==="0")return;
    if(!resources.has(resourceUID))return;
    if(!taskResources.has(taskUID))taskResources.set(taskUID,[]);
    taskResources.get(taskUID).push({resourceUID});
  });
  const rows=[],errors=[];
  doc.querySelectorAll("Task").forEach((task,i)=>{
    const uid=getText(task,"UID");if(uid==="0")return;
    const name=getText(task,"Name");if(!name)return;
    const isSummary=getText(task,"Summary")==="1";
    const startDate=toDate(getText(task,"Start"));
    const finishDate=toDate(getText(task,"Finish"));
    const wbs=getText(task,"WBS")||uid;
    const assignedResources=taskResources.get(uid)||[];
    if(!isSummary){
      if(!startDate)errors.push({row:i+1,issue:`"${name}" — missing Start`});
      if(!finishDate)errors.push({row:i+1,issue:`"${name}" — missing Finish`});
    }
    rows.push({rowNum:i+1,id:wbs,uid,name,startDate,finishDate,
      percentComplete:toPercent(getText(task,"PercentComplete")||"0"),
      isSummary,assignedResources});
  });
  return{rows,errors,resources};
}

// ─── Fuzzy matching ───────────────────────────────────────────────────────────
const STOP=new Set(["the","a","an","of","in","at","to","and","or","for","with","on","is","are","be","by","as","from","this","it","work","scope","item","all","per","each","existing","new","provide","install","furnish","allowance"]);
const ALIASES={"demo":["demolition","demolish","remove","removal"],"demolition":["demo","remove","removal"],"conc":["concrete"],"concrete":["conc","slab","footing","foundation"],"elec":["electrical","electric","power"],"electrical":["elec","electric","power","wiring"],"mech":["mechanical","hvac","plumbing"],"hvac":["mechanical","heating","cooling","mech"],"plumb":["plumbing","pipe"],"plumbing":["plumb","pipe","piping","sanitary"],"framing":["structural","frame","steel","wood","lumber"],"masonry":["brick","block","cmu","stone"],"drywall":["gypsum","gwb","partition","wall"],"flooring":["floor","tile","carpet","lvt"],"roofing":["roof","membrane","waterproof"],"sitework":["site","earthwork","grading","excavation","utilities"],"painting":["paint","coating","finish"],"insulation":["insulate","thermal","acoustic"],"siding":["cladding","facade","exterior"],"doors":["door","hardware","frame"],"windows":["window","glazing","storefront"]};
function tokenize(s){return String(s||"").toLowerCase().replace(/[^a-z0-9\s]/g," ").split(/\s+/).filter(t=>t.length>1&&!STOP.has(t));}
function expand(tokens){const s=new Set(tokens);tokens.forEach(t=>{if(ALIASES[t])ALIASES[t].forEach(a=>s.add(a));});return Array.from(s);}
function matchScore(a,b){const ta=expand(tokenize(a)),tb=expand(tokenize(b));if(!ta.length||!tb.length)return 0;const sa=new Set(ta),sb=new Set(tb);let ov=0;sa.forEach(t=>{if(sb.has(t))ov++;});const union=new Set([...sa,...sb]).size;const j=union>0?ov/union:0;const oa=new Set(tokenize(a)),ob=new Set(tokenize(b));let orig=0;oa.forEach(t=>{if(ob.has(t))orig++;});return Math.min(1,j+orig*0.15);}
function autoLink(tasks,sovItems){
  const map={};
  tasks.forEach(t=>{
    if(t.isSummary)return;
    let best=null,bestScore=0;
    sovItems.forEach(s=>{const sc=matchScore(t.name,s.description);if(sc>bestScore){bestScore=sc;best=s;}});
    if(best&&bestScore>=0.08){if(!map[best.id])map[best.id]=[];map[best.id].push(t.uid);}
  });
  return map;
}
function autoMatchSubsToSOV(resources,sovItems){
  const map={};
  resources.forEach((res,uid)=>{
    let best=null,bestScore=0;
    sovItems.forEach(s=>{const sc=matchScore(res.name,s.description);if(sc>bestScore){bestScore=sc;best=s;}});
    map[uid]=best&&bestScore>=0.08?best.id:null;
  });
  return map;
}

// ─── Bell curve ───────────────────────────────────────────────────────────────
function bellCurve(startDate,finishDate,totalValue){
  if(!startDate||!finishDate||totalValue<=0)return[];
  const dur=Math.max(1,Math.round((finishDate-startDate)/86400000)+1);
  const sigma=dur/6,mid=dur/2;let sum=0;
  const w=Array.from({length:dur},(_,i)=>{const g=Math.exp(-0.5*Math.pow((i-mid)/sigma,2));sum+=g;return g;});
  let cum=0;
  return w.map((wt,i)=>{const d=new Date(startDate);d.setDate(d.getDate()+i);const daily=sum>0?(wt/sum)*totalValue:0;cum+=daily;return{date:fmtDate(d),daily,cumulative:cum};});
}

// ─── Build scope curves ───────────────────────────────────────────────────────
function buildScopeCurves(tasks,sovItems,linksMap){
  if(!sovItems.length||!tasks.length)return{byScope:{},allDates:[],scopeKeys:[]};
  const taskByUid=new Map(tasks.map(t=>[t.uid,t]));
  const byScope={};
  let globalMin=Infinity,globalMax=-Infinity;
  sovItems.forEach(sov=>{
    const linked=(linksMap[sov.id]||[]).map(u=>taskByUid.get(u)).filter(t=>t&&t.startDate&&t.finishDate);
    if(!linked.length)return;
    const minD=new Date(Math.min(...linked.map(t=>t.startDate.getTime())));
    const maxD=new Date(Math.max(...linked.map(t=>t.finishDate.getTime())));
    const full=bellCurve(minD,maxD,sov.scheduledValue);
    const lookup=new Map(full.map(p=>[p.date,p.daily]));
    byScope[sov.id]={sov,tasks:linked,startDate:minD,finishDate:maxD,full,lookup};
    if(minD.getTime()<globalMin)globalMin=minD.getTime();
    if(maxD.getTime()>globalMax)globalMax=maxD.getTime();
  });
  if(!Object.keys(byScope).length)return{byScope:{},allDates:[],scopeKeys:[]};
  const scopeIds=Object.keys(byScope);
  const totalDays=Math.round((globalMax-globalMin)/86400000)+1;
  const step=Math.max(1,Math.floor(totalDays/150));
  const allDates=[];let cumTotal=0;
  for(let i=0;i<totalDays;i+=step){
    const d=new Date(globalMin+i*86400000);const dateStr=fmtDate(d);
    const row={date:dateStr,total:0};
    scopeIds.forEach(sid=>{const val=byScope[sid].lookup.get(dateStr)||0;const key=byScope[sid].sov.description.slice(0,18);row[key]=(row[key]||0)+Math.round(val);row.total+=val;});
    row.total=Math.round(row.total);cumTotal+=row.total;row.cumTotal=Math.round(cumTotal);
    allDates.push(row);
  }
  scopeIds.forEach(sid=>{const{full}=byScope[sid];const s=Math.max(1,Math.floor(full.length/120));byScope[sid].weekly=full.filter((_,i)=>i%s===0||(i===full.length-1));});
  return{byScope,allDates,scopeKeys:scopeIds};
}

// ─── Build sub curves ─────────────────────────────────────────────────────────
function buildSubCurves(tasks,resources,subContractValues){
  if(!resources.size||!tasks.length)return{bySubUID:{},subAllDates:[],subKeys:[]};
  const workTasks=tasks.filter(t=>!t.isSummary&&t.startDate&&t.finishDate);
  const bySubUID={};
  let globalMin=Infinity,globalMax=-Infinity;
  resources.forEach((res,uid)=>{
    const value=subContractValues[uid]||0;
    if(value<=0)return;
    const subTasks=workTasks.filter(t=>t.assignedResources.some(r=>r.resourceUID===uid));
    if(!subTasks.length)return;
    const minD=new Date(Math.min(...subTasks.map(t=>t.startDate.getTime())));
    const maxD=new Date(Math.max(...subTasks.map(t=>t.finishDate.getTime())));
    const full=bellCurve(minD,maxD,value);
    const lookup=new Map(full.map(p=>[p.date,p.daily]));
    bySubUID[uid]={res,tasks:subTasks,startDate:minD,finishDate:maxD,full,lookup,value};
    if(minD.getTime()<globalMin)globalMin=minD.getTime();
    if(maxD.getTime()>globalMax)globalMax=maxD.getTime();
  });
  if(!Object.keys(bySubUID).length)return{bySubUID:{},subAllDates:[],subKeys:[]};
  const subIds=Object.keys(bySubUID);
  const totalDays=Math.round((globalMax-globalMin)/86400000)+1;
  const step=Math.max(1,Math.floor(totalDays/150));
  const subAllDates=[];let cumTotal=0;
  for(let i=0;i<totalDays;i+=step){
    const d=new Date(globalMin+i*86400000);const dateStr=fmtDate(d);
    const row={date:dateStr,total:0};
    subIds.forEach(uid=>{const val=bySubUID[uid].lookup.get(dateStr)||0;const key=bySubUID[uid].res.name.slice(0,16);row[key]=(row[key]||0)+Math.round(val);row.total+=val;});
    row.total=Math.round(row.total);cumTotal+=row.total;row.cumTotal=Math.round(cumTotal);
    subAllDates.push(row);
  }
  subIds.forEach(uid=>{const{full}=bySubUID[uid];const s=Math.max(1,Math.floor(full.length/120));bySubUID[uid].weekly=full.filter((_,i)=>i%s===0||(i===full.length-1));});
  return{bySubUID,subAllDates,subKeys:subIds};
}

// ─── Zoom hook ────────────────────────────────────────────────────────────────
function useZoom(data){
  const [range,setRange]=useState(null);
  const [sel,setSel]=useState(null);
  const sliced=useMemo(()=>(!range||!data.length)?data:data.slice(range[0],range[1]+1),[data,range]);
  function onDown(e){if(e?.activeLabel)setSel({start:e.activeLabel,end:null,active:true});}
  function onMove(e){if(sel?.active&&e?.activeLabel)setSel(s=>({...s,end:e.activeLabel}));}
  function onUp(){
    if(sel?.active&&sel.start&&sel.end&&sel.start!==sel.end){
      const a=data.findIndex(d=>d.date===sel.start||d.month===sel.start),b=data.findIndex(d=>d.date===sel.end||d.month===sel.end);
      if(a>=0&&b>=0)setRange([Math.min(a,b),Math.max(a,b)]);
    }
    setSel(null);
  }
  return{sliced,onDown,onMove,onUp,reset:()=>setRange(null),isZoomed:!!range,sel};
}

// ─── UI Atoms ─────────────────────────────────────────────────────────────────
function Metric({title,value,sub,color}){
  const c=color||WB.primary;
  return(
    <div style={{background:WB.card,borderRadius:12,padding:"13px 15px",border:`1px solid ${WB.border}`,position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:0,right:0,width:50,height:50,background:c,opacity:.07,borderRadius:"0 12px 0 50px"}}/>
      <div style={{fontSize:9,fontFamily:"monospace",color:WB.textDim,textTransform:"uppercase",letterSpacing:1,marginBottom:5}}>{title}</div>
      <div style={{fontSize:18,fontWeight:700,color:WB.text,lineHeight:1}}>{value}</div>
      {sub&&<div style={{fontSize:11,color:c,marginTop:3}}>{sub}</div>}
    </div>
  );
}

function DropZone({label,accept,onFile,fileName,icon,hint}){
  const [drag,setDrag]=useState(false);
  return(
    <div>
      <label onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
        onDrop={e=>{e.preventDefault();setDrag(false);const f=e.dataTransfer.files?.[0];if(f)onFile(f);}}
        style={{display:"flex",flexDirection:"column",alignItems:"center",gap:8,padding:"16px",border:`2px dashed ${drag?WB.primary:fileName?"#1e4d2b":WB.border}`,borderRadius:12,cursor:"pointer",background:drag?"#0c1a2e":fileName?"#020d06":WB.bg,transition:"all .2s",textAlign:"center"}}>
        <div style={{width:34,height:34,borderRadius:8,background:WB.card,display:"flex",alignItems:"center",justifyContent:"center",border:`1px solid ${WB.border}`}}>{icon}</div>
        <div>
          <div style={{fontWeight:600,color:fileName?WB.green:WB.text,fontSize:13,marginBottom:2}}>{fileName||label}</div>
          <div style={{fontSize:11,color:WB.textDim}}>{fileName?"Loaded — click to replace":"Drag & drop or click"}</div>
        </div>
        <input type="file" accept={accept} style={{display:"none"}} onChange={e=>e.target.files?.[0]&&onFile(e.target.files[0])}/>
      </label>
      {hint&&<div style={{fontSize:11,color:WB.textDim,marginTop:3}}>{hint}</div>}
    </div>
  );
}

function ZoomBar({isZoomed,onReset}){
  return isZoomed
    ?<button onClick={onReset} style={{display:"flex",alignItems:"center",gap:5,padding:"4px 9px",background:WB.card,border:`1px solid ${WB.borderLight}`,borderRadius:6,color:WB.textMuted,cursor:"pointer",fontSize:11}}><Icon.ZoomOut/> Reset zoom</button>
    :<span style={{fontSize:11,color:WB.textDim}}>Drag to zoom</span>;
}

function Empty({msg}){
  return(
    <div style={{minHeight:160,display:"flex",alignItems:"center",justifyContent:"center",border:`1px dashed ${WB.border}`,borderRadius:12}}>
      <div style={{fontSize:13,color:WB.textDim,textAlign:"center",padding:20,maxWidth:340}}>{msg}</div>
    </div>
  );
}

function CTip({active,payload,label}){
  if(!active||!payload?.length)return null;
  const items=payload.filter(p=>p.value>0&&p.name!=="Project Total").sort((a,b)=>b.value-a.value);
  const totalVal=payload.find(p=>p.name==="Project Total")?.value||items.reduce((s,p)=>s+p.value,0);
  return(
    <div style={{background:WB.card,border:`1px solid ${WB.border}`,borderRadius:10,padding:"10px 13px",fontSize:12,maxWidth:300,boxShadow:"0 8px 32px rgba(0,0,0,.6)"}}>
      <div style={{color:WB.textDim,marginBottom:5,fontFamily:"monospace",fontSize:10}}>{label}</div>
      {items.slice(0,10).map(p=>(
        <div key={p.name||p.dataKey} style={{display:"flex",gap:7,alignItems:"center",marginBottom:2}}>
          <span style={{width:7,height:7,borderRadius:"50%",background:p.color,flexShrink:0,display:"inline-block"}}/>
          <span style={{color:WB.textMuted,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontSize:11}}>{p.name||p.dataKey}:</span>
          <span style={{fontWeight:600,flexShrink:0}}>{currency(p.value)}</span>
        </div>
      ))}
      {totalVal!=null&&<div style={{borderTop:`1px solid ${WB.borderLight}`,marginTop:7,paddingTop:7,display:"flex",justifyContent:"space-between",gap:8}}><span style={{color:WB.text,fontWeight:700}}>Total:</span><span style={{color:WB.text,fontWeight:800,fontSize:14}}>{currency(totalVal)}</span></div>}
    </div>
  );
}

// ─── Sub Dollar Override Panel ────────────────────────────────────────────────
function SubDollarOverride({resources,sovItems,subContractValues,onUpdate,subSOVLinks,onSOVLinkChange}){
  const [editingUID,setEditingUID]=useState(null);
  const [editVal,setEditVal]=useState("");
  const totalAssigned=Object.values(subContractValues).reduce((s,v)=>s+v,0);
  const contractTotal=sovItems.reduce((s,i)=>s+i.scheduledValue,0);

  function startEdit(uid){setEditingUID(uid);setEditVal(String(subContractValues[uid]||0));}
  function saveEdit(uid){onUpdate(prev=>({...prev,[uid]:toMoney(editVal)}));setEditingUID(null);}

  return(
    <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:18}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16,flexWrap:"wrap",gap:10}}>
        <div>
          <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 3px",color:WB.text}}>Subcontractor Values & SOV Links</h2>
          <p style={{color:WB.textDim,fontSize:12,margin:0}}>Auto-distributed from SOV match. Click ✏️ to override any amount. Link each sub to their SOV scope.</p>
        </div>
        <div style={{background:WB.bg,borderRadius:9,padding:"8px 14px",border:`1px solid ${WB.border}`,fontSize:12}}>
          <span style={{color:WB.textDim}}>Assigned: </span>
          <span style={{color:totalAssigned<=contractTotal?WB.green:WB.red,fontWeight:700,fontFamily:"monospace"}}>{currency(totalAssigned)}</span>
          <span style={{color:WB.textDim}}> / {currency(contractTotal)}</span>
        </div>
      </div>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",minWidth:680,borderCollapse:"collapse",fontSize:12}}>
          <thead>
            <tr style={{background:WB.bg}}>
              {["Sub / Resource","Linked SOV Scope","Contract Value","Tasks"].map(h=>(
                <th key={h} style={{padding:"7px 10px",textAlign:"left",color:WB.textDim,fontFamily:"monospace",fontSize:9,borderBottom:`1px solid ${WB.border}`,textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {Array.from(resources.entries()).map(([uid,res],i)=>{
              const col=PAL[i%PAL.length];
              const val=subContractValues[uid]||0;
              const linkedSOV=subSOVLinks[uid]||"none";
              return(
                <tr key={uid} style={{borderBottom:`1px solid ${WB.border}`}}>
                  <td style={{padding:"8px 10px",color:WB.text}}>
                    <div style={{display:"flex",alignItems:"center",gap:7}}>
                      <div style={{width:8,height:8,borderRadius:"50%",background:col,flexShrink:0}}/>
                      <span style={{fontWeight:600}}>{res.name}</span>
                    </div>
                  </td>
                  <td style={{padding:"8px 10px",minWidth:220}}>
                    <select value={linkedSOV} onChange={e=>onSOVLinkChange(uid,e.target.value)}
                      style={{width:"100%",padding:"5px 8px",background:WB.bg,border:`1px solid ${linkedSOV!=="none"?col:WB.border}`,borderRadius:6,color:linkedSOV!=="none"?col:WB.textDim,fontSize:11,cursor:"pointer",outline:"none"}}>
                      <option value="none">— Unlinked —</option>
                      {sovItems.map(s=><option key={s.id} value={s.id}>{s.description} ({currency(s.scheduledValue)})</option>)}
                    </select>
                  </td>
                  <td style={{padding:"8px 10px",minWidth:180}}>
                    {editingUID===uid?(
                      <div style={{display:"flex",gap:6,alignItems:"center"}}>
                        <input type="text" value={editVal} onChange={e=>setEditVal(e.target.value)}
                          onKeyDown={e=>{if(e.key==="Enter")saveEdit(uid);if(e.key==="Escape")setEditingUID(null);}}
                          autoFocus
                          style={{width:110,padding:"4px 7px",background:WB.bg,border:`1px solid ${col}`,borderRadius:5,color:col,fontSize:12,fontFamily:"monospace",outline:"none"}}/>
                        <button onClick={()=>saveEdit(uid)} style={{padding:"4px 8px",background:WB.greenBg,border:`1px solid ${WB.greenBorder}`,borderRadius:5,color:WB.green,cursor:"pointer",fontSize:11}}>✓</button>
                        <button onClick={()=>setEditingUID(null)} style={{padding:"4px 8px",background:"transparent",border:`1px solid ${WB.border}`,borderRadius:5,color:WB.textDim,cursor:"pointer",fontSize:11}}>✕</button>
                      </div>
                    ):(
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <span style={{fontFamily:"monospace",fontWeight:700,color:val>0?col:WB.textDim}}>{currency(val)}</span>
                        <button onClick={()=>startEdit(uid)} style={{background:"none",border:"none",cursor:"pointer",color:WB.textDim,fontSize:12,padding:0}}><Icon.Edit/></button>
                        {val===0&&<span style={{fontSize:10,color:WB.yellow}}>⚠ no value</span>}
                      </div>
                    )}
                  </td>
                  <td style={{padding:"8px 10px",color:WB.textDim,fontFamily:"monospace",fontSize:11}}>
                    —
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── Link Tasks Panel ─────────────────────────────────────────────────────────
function LinkManager({tasks,sovItems,linksMap,onLinksChange,resources}){
  const [taskFilter,setTaskFilter]=useState("");
  const [unlinkedOnly,setUnlinkedOnly]=useState(false);
  const sovColorMap=useMemo(()=>{const m={};sovItems.forEach((s,i)=>{m[s.id]=PAL[i%PAL.length];});return m;},[sovItems]);
  const taskToSov=useMemo(()=>{const m={};Object.entries(linksMap).forEach(([sid,uids])=>uids.forEach(uid=>{m[uid]=sid;}));return m;},[linksMap]);
  function setTaskLink(uid,sovId){
    const m={...linksMap};Object.keys(m).forEach(sid=>{m[sid]=(m[sid]||[]).filter(u=>u!==uid);});
    if(sovId!=="none"){if(!m[sovId])m[sovId]=[];m[sovId]=[...m[sovId],uid];}
    onLinksChange(m);
  }
  const workTasks=tasks.filter(t=>!t.isSummary);
  const filtered=workTasks.filter(t=>t.name.toLowerCase().includes(taskFilter.toLowerCase())&&(!unlinkedOnly||!taskToSov[t.uid]));
  const linked=workTasks.filter(t=>taskToSov[t.uid]).length;
  return(
    <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:18}}>
      <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 3px",color:WB.text}}>Link Tasks → SOV Scopes</h2>
      <p style={{color:WB.textDim,fontSize:12,margin:"0 0 12px"}}>Assign each task to a SOV scope for the scope-based curves. Resource assignments come from the schedule automatically.</p>
      <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap",marginBottom:12}}>
        <input value={taskFilter} onChange={e=>setTaskFilter(e.target.value)} placeholder="Filter tasks…"
          style={{padding:"6px 10px",background:WB.bg,border:`1px solid ${WB.border}`,borderRadius:7,color:WB.textMuted,fontSize:12,outline:"none",width:200}}/>
        <label style={{display:"flex",alignItems:"center",gap:5,fontSize:12,color:WB.textDim,cursor:"pointer"}}>
          <input type="checkbox" checked={unlinkedOnly} onChange={e=>setUnlinkedOnly(e.target.checked)} style={{accentColor:WB.primary}}/>
          Unlinked only
        </label>
        <span style={{fontSize:12,color:WB.textDim,marginLeft:"auto"}}><span style={{color:WB.green,fontWeight:600}}>{linked}</span> linked · <span style={{color:linked<workTasks.length?WB.yellow:WB.green,fontWeight:600}}>{workTasks.length-linked}</span> unlinked</span>
      </div>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",minWidth:680,borderCollapse:"collapse",fontSize:12}}>
          <thead><tr style={{background:WB.bg}}>{["WBS","Task Name","Resource(s)","Dates","SOV Scope"].map(h=><th key={h} style={{padding:"7px 10px",textAlign:"left",color:WB.textDim,fontFamily:"monospace",fontSize:9,borderBottom:`1px solid ${WB.border}`,textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
          <tbody>
            {filtered.map((t,i)=>{
              const sid=taskToSov[t.uid];const col=sid?sovColorMap[sid]:null;
              const resNames=t.assignedResources?.map(r=>resources.get(r.resourceUID)?.name||r.resourceUID).join(", ")||"—";
              return(
                <tr key={t.uid} style={{borderBottom:`1px solid ${WB.border}`,background:i%2===0?"transparent":WB.bg}}>
                  <td style={{padding:"7px 10px",fontFamily:"monospace",color:WB.primary,fontSize:10,whiteSpace:"nowrap"}}>{t.id}</td>
                  <td style={{padding:"7px 10px",color:WB.text,maxWidth:200}}>
                    <div style={{display:"flex",alignItems:"center",gap:6}}>{col&&<div style={{width:3,height:14,borderRadius:2,background:col,flexShrink:0}}/>}<span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{t.name}</span></div>
                  </td>
                  <td style={{padding:"7px 10px",color:WB.orange,fontSize:10,maxWidth:130,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{resNames}</td>
                  <td style={{padding:"7px 10px",fontFamily:"monospace",color:WB.textDim,fontSize:10,whiteSpace:"nowrap"}}>{fmtDate(t.startDate)} → {fmtDate(t.finishDate)}</td>
                  <td style={{padding:"7px 10px",minWidth:220}}>
                    <select value={sid||"none"} onChange={e=>setTaskLink(t.uid,e.target.value)}
                      style={{width:"100%",padding:"5px 9px",background:col?`${col}18`:WB.bg,border:`1px solid ${col||WB.border}`,borderRadius:6,color:col||WB.textDim,fontSize:12,cursor:"pointer",outline:"none",fontWeight:col?600:400}}>
                      <option value="none">— Unassigned —</option>
                      {sovItems.map(s=><option key={s.id} value={s.id}>{s.description} ({currency(s.scheduledValue)})</option>)}
                    </select>
                  </td>
                </tr>
              );
            })}
            {!filtered.length&&<tr><td colSpan={5} style={{padding:24,textAlign:"center",color:WB.textDim}}>No tasks match filter.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── Scope Curves Tab ─────────────────────────────────────────────────────────
function ScopeCurvesTab({byScope,scopeKeys,allDates,sovColorMap,sovItems}){
  const zoom=useZoom(allDates);
  const [selected,setSelected]=useState("all");
  if(!scopeKeys.length)return<Empty msg="Upload both files and link tasks to SOV scopes to see scope curves."/>;
  return(
    <div style={{display:"grid",gap:14}}>
      <div style={{display:"flex",gap:5,flexWrap:"wrap",alignItems:"center"}}>
        <button onClick={()=>setSelected("all")} style={{padding:"4px 10px",borderRadius:99,border:`1px solid ${selected==="all"?WB.primary:WB.border}`,background:selected==="all"?"#1e3448":"transparent",color:selected==="all"?WB.primaryLight:WB.textDim,cursor:"pointer",fontSize:12}}>All scopes</button>
        {scopeKeys.map((sid,i)=>{const s=byScope[sid];const col=PAL[i%PAL.length];return(<button key={sid} onClick={()=>setSelected(sid)} style={{padding:"4px 10px",borderRadius:99,border:`1px solid ${selected===sid?col:WB.border}`,background:selected===sid?`${col}20`:"transparent",color:selected===sid?col:WB.textDim,cursor:"pointer",fontSize:12,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.sov.description.slice(0,22)}</button>);})}
      </div>
      {selected==="all"&&(
        <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:20}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
            <div><h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 2px",color:WB.text}}>All Scopes — Stacked Bell Curves</h2><p style={{color:WB.textDim,fontSize:12,margin:0}}>By SOV scope · task dates drive timing · drag to zoom</p></div>
            <ZoomBar isZoomed={zoom.isZoomed} onReset={zoom.reset}/>
          </div>
          <div style={{height:440,cursor:"crosshair"}}>
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={zoom.sliced} margin={{top:8,right:8,left:8,bottom:8}} onMouseDown={zoom.onDown} onMouseMove={zoom.onMove} onMouseUp={zoom.onUp}>
                <defs>{scopeKeys.map((sid,i)=>(<linearGradient key={sid} id={`sg${i}`} x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={PAL[i%PAL.length]} stopOpacity={0.55}/><stop offset="95%" stopColor={PAL[i%PAL.length]} stopOpacity={0.04}/></linearGradient>))}</defs>
                <CartesianGrid strokeDasharray="3 3" stroke={WB.border}/>
                <XAxis dataKey="date" tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} minTickGap={55} tickFormatter={fmtShort}/>
                <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} width={52}/>
                <Tooltip content={<CTip/>}/>
                <Legend wrapperStyle={{fontSize:10,color:WB.textDim}}/>
                {scopeKeys.map((sid,i)=>{const s=byScope[sid];const col=PAL[i%PAL.length];const key=s.sov.description.slice(0,18);return(<Area key={sid} type="monotone" dataKey={key} stroke={col} strokeWidth={1.5} fill={`url(#sg${i})`} stackId="s" name={key}/>);})}
                <Line type="monotone" dataKey="total" stroke="#ffffff" strokeWidth={2.5} dot={false} name="Project Total" strokeOpacity={0.9}/>
                {zoom.sel?.active&&zoom.sel.start&&zoom.sel.end&&<ReferenceArea x1={zoom.sel.start} x2={zoom.sel.end} fill={WB.primary} fillOpacity={0.08}/>}
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}
      {selected!=="all"&&byScope[selected]&&(()=>{
        const{sov,startDate,finishDate,weekly}=byScope[selected];
        const col=sovColorMap[selected];
        return(
          <div style={{display:"grid",gap:12}}>
            <div style={{background:WB.card,borderRadius:13,border:`1px solid ${col}40`,padding:16,display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:10}}>
              <div><div style={{fontSize:9,fontFamily:"monospace",color:WB.textDim,textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Scope</div><div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:13,color:WB.text}}>{sov.description}</div></div>
              <div><div style={{fontSize:9,fontFamily:"monospace",color:WB.textDim,textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Contract Value</div><div style={{fontWeight:700,fontSize:15,color:col}}>{currency(sov.scheduledValue)}</div></div>
              <div><div style={{fontSize:9,fontFamily:"monospace",color:WB.textDim,textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Duration</div><div style={{color:WB.textMuted,fontSize:11}}>{fmtDate(startDate)} → {fmtDate(finishDate)}</div></div>
            </div>
            <div style={{background:WB.card,borderRadius:13,border:`1px solid ${WB.border}`,padding:20}}>
              <h3 style={{fontFamily:"'Syne',sans-serif",fontSize:15,fontWeight:700,margin:"0 0 14px",color:WB.text}}>{sov.description} — Bell Curve</h3>
              <div style={{height:280}}>
                <ResponsiveContainer width="100%" height="100%">
                  <AreaChart data={weekly} margin={{top:8,right:8,left:8,bottom:8}}>
                    <defs><linearGradient id="sg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={col} stopOpacity={0.5}/><stop offset="95%" stopColor={col} stopOpacity={0.02}/></linearGradient></defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={WB.border}/>
                    <XAxis dataKey="date" tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} minTickGap={40} tickFormatter={fmtShort}/>
                    <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} width={52}/>
                    <Tooltip content={<CTip/>}/>
                    <Area type="monotone" dataKey="daily" stroke={col} strokeWidth={2.5} fill="url(#sg)" name="Daily Spend"/>
                  </AreaChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}

// ─── Sub Curves Tab ───────────────────────────────────────────────────────────
function SubCurvesTab({bySubUID,subKeys,subAllDates,resources}){
  const zoom=useZoom(subAllDates);
  const [selected,setSelected]=useState("all");
  if(!subKeys.length){
    if(!resources.size)return<Empty msg="No resources found in your schedule. Upload a resource-loaded XML (with subcontractors assigned to tasks) to see sub curves. Non-resource-loaded schedules still work — use the Scope Curves tab instead."/>;
    return<Empty msg="Resources detected but no sub curves yet. Go to the Sub Values tab and link each sub to a SOV scope to generate their curves."/>;
  }
  return(
    <div style={{display:"grid",gap:14}}>
      <div style={{display:"flex",gap:5,flexWrap:"wrap",alignItems:"center"}}>
        <button onClick={()=>setSelected("all")} style={{padding:"4px 10px",borderRadius:99,border:`1px solid ${selected==="all"?WB.primary:WB.border}`,background:selected==="all"?"#1e3448":"transparent",color:selected==="all"?WB.primaryLight:WB.textDim,cursor:"pointer",fontSize:12}}>All subs</button>
        {subKeys.map((uid,i)=>{const res=bySubUID[uid].res;const col=PAL[i%PAL.length];return(<button key={uid} onClick={()=>setSelected(uid)} style={{padding:"4px 10px",borderRadius:99,border:`1px solid ${selected===uid?col:WB.border}`,background:selected===uid?`${col}20`:"transparent",color:selected===uid?col:WB.textDim,cursor:"pointer",fontSize:12,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{res.name.slice(0,20)}</button>);})}
      </div>
      {selected==="all"&&(
        <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:20}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
            <div><h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 2px",color:WB.text}}>All Subs — Stacked Bell Curves</h2><p style={{color:WB.textDim,fontSize:12,margin:0}}>By subcontractor resource · task assignments drive timing · drag to zoom</p></div>
            <ZoomBar isZoomed={zoom.isZoomed} onReset={zoom.reset}/>
          </div>
          <div style={{height:440,cursor:"crosshair"}}>
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={zoom.sliced} margin={{top:8,right:8,left:8,bottom:8}} onMouseDown={zoom.onDown} onMouseMove={zoom.onMove} onMouseUp={zoom.onUp}>
                <defs>{subKeys.map((uid,i)=>(<linearGradient key={uid} id={`sub${i}`} x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={PAL[i%PAL.length]} stopOpacity={0.55}/><stop offset="95%" stopColor={PAL[i%PAL.length]} stopOpacity={0.04}/></linearGradient>))}</defs>
                <CartesianGrid strokeDasharray="3 3" stroke={WB.border}/>
                <XAxis dataKey="date" tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} minTickGap={55} tickFormatter={fmtShort}/>
                <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} width={52}/>
                <Tooltip content={<CTip/>}/>
                <Legend wrapperStyle={{fontSize:10,color:WB.textDim}}/>
                {subKeys.map((uid,i)=>{const s=bySubUID[uid];const col=PAL[i%PAL.length];const key=s.res.name.slice(0,16);return(<Area key={uid} type="monotone" dataKey={key} stroke={col} strokeWidth={1.5} fill={`url(#sub${i})`} stackId="s" name={key}/>);})}
                <Line type="monotone" dataKey="total" stroke="#ffffff" strokeWidth={2.5} dot={false} name="Project Total" strokeOpacity={0.9}/>
                {zoom.sel?.active&&zoom.sel.start&&zoom.sel.end&&<ReferenceArea x1={zoom.sel.start} x2={zoom.sel.end} fill={WB.primary} fillOpacity={0.08}/>}
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}
      {selected!=="all"&&bySubUID[selected]&&(()=>{
        const{res,startDate,finishDate,weekly,value,tasks:subTasks}=bySubUID[selected];
        const idx=subKeys.indexOf(selected);
        const col=PAL[idx%PAL.length];
        return(
          <div style={{display:"grid",gap:12}}>
            <div style={{background:WB.card,borderRadius:13,border:`1px solid ${col}40`,padding:16,display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:10}}>
              <div><div style={{fontSize:9,fontFamily:"monospace",color:WB.textDim,textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Subcontractor</div><div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:13,color:WB.text}}>{res.name}</div></div>
              <div><div style={{fontSize:9,fontFamily:"monospace",color:WB.textDim,textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Contract Value</div><div style={{fontWeight:700,fontSize:15,color:col}}>{currency(value)}</div></div>
              <div><div style={{fontSize:9,fontFamily:"monospace",color:WB.textDim,textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Tasks</div><div style={{fontWeight:700,color:WB.textMuted}}>{subTasks.length}</div></div>
              <div><div style={{fontSize:9,fontFamily:"monospace",color:WB.textDim,textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Duration</div><div style={{color:WB.textMuted,fontSize:11}}>{fmtDate(startDate)} → {fmtDate(finishDate)}</div></div>
            </div>
            <div style={{background:WB.card,borderRadius:13,border:`1px solid ${WB.border}`,padding:20}}>
              <h3 style={{fontFamily:"'Syne',sans-serif",fontSize:15,fontWeight:700,margin:"0 0 14px",color:WB.text}}>{res.name} — Spend Bell Curve</h3>
              <div style={{height:280}}>
                <ResponsiveContainer width="100%" height="100%">
                  <AreaChart data={weekly} margin={{top:8,right:8,left:8,bottom:8}}>
                    <defs><linearGradient id="subg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={col} stopOpacity={0.5}/><stop offset="95%" stopColor={col} stopOpacity={0.02}/></linearGradient></defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={WB.border}/>
                    <XAxis dataKey="date" tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} minTickGap={40} tickFormatter={fmtShort}/>
                    <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} width={52}/>
                    <Tooltip content={<CTip/>}/>
                    <Area type="monotone" dataKey="daily" stroke={col} strokeWidth={2.5} fill="url(#subg)" name="Daily Spend"/>
                  </AreaChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div style={{background:WB.card,borderRadius:13,border:`1px solid ${WB.border}`,padding:18,overflowX:"auto"}}>
              <h3 style={{fontFamily:"'Syne',sans-serif",fontSize:15,fontWeight:700,margin:"0 0 12px",color:WB.text}}>Assigned Tasks</h3>
              <table style={{width:"100%",minWidth:500,borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{background:WB.bg}}>{["WBS","Task","Start","Finish","% Done"].map(h=><th key={h} style={{padding:"6px 10px",textAlign:"left",color:WB.textDim,fontFamily:"monospace",fontSize:9,borderBottom:`1px solid ${WB.border}`,textTransform:"uppercase",letterSpacing:.5}}>{h}</th>)}</tr></thead>
                <tbody>
                  {subTasks.map((t,i)=>(
                    <tr key={t.uid} style={{borderBottom:`1px solid ${WB.border}`,background:i%2===0?"transparent":WB.bg}}>
                      <td style={{padding:"6px 10px",fontFamily:"monospace",color:WB.primary,fontSize:10}}>{t.id}</td>
                      <td style={{padding:"6px 10px",color:WB.text,maxWidth:220,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{t.name}</td>
                      <td style={{padding:"6px 10px",fontFamily:"monospace",color:WB.textDim,fontSize:10}}>{fmtDate(t.startDate)}</td>
                      <td style={{padding:"6px 10px",fontFamily:"monospace",color:WB.textDim,fontSize:10}}>{fmtDate(t.finishDate)}</td>
                      <td style={{padding:"6px 10px",fontFamily:"monospace",color:t.percentComplete>0?WB.green:WB.textDim,fontSize:10}}>{Math.round(t.percentComplete*100)}%</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        );
      })()}
    </div>
  );
}

// ─── Dashboard ────────────────────────────────────────────────────────────────
function Dashboard({allDates,subAllDates,scopeKeys,subKeys,contractTotal,subContractValues,resources}){
  const scopeZoom=useZoom(allDates);
  const subZoom=useZoom(subAllDates);
  const totalSubValue=Object.values(subContractValues).reduce((s,v)=>s+v,0);

  return(
    <div style={{display:"grid",gap:16}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:10}}>
        <Metric title="Contract Total" value={currency(contractTotal)} color={WB.primary}/>
        <Metric title="SOV Scopes" value={scopeKeys.length} color={WB.green} sub="plotted"/>
        <Metric title="Subs w/ Curves" value={resources.size>0?`${subKeys.length} / ${resources.size}`:"N/A"} color={WB.orange} sub={resources.size===0?"non-resource schedule":undefined}/>
        <Metric title="Sub Value Assigned" value={currency(totalSubValue)} color={totalSubValue>0&&totalSubValue<=contractTotal?WB.green:WB.yellow}/>
      </div>

      {/* Curve 1: By Scope */}
      <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:20}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
          <div>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:16,fontWeight:700,margin:"0 0 2px",color:WB.text}}>① Project Curve — By SOV Scope</h2>
            <p style={{color:WB.textDim,fontSize:12,margin:0}}>Cumulative S-curve built from task dates × SOV scope values</p>
          </div>
          <ZoomBar isZoomed={scopeZoom.isZoomed} onReset={scopeZoom.reset}/>
        </div>
        <div style={{height:300,cursor:"crosshair"}}>
          {allDates.length?(
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={scopeZoom.sliced} margin={{top:8,right:8,left:8,bottom:8}} onMouseDown={scopeZoom.onDown} onMouseMove={scopeZoom.onMove} onMouseUp={scopeZoom.onUp}>
                <CartesianGrid strokeDasharray="3 3" stroke={WB.border}/>
                <XAxis dataKey="date" tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} minTickGap={55} tickFormatter={fmtShort}/>
                <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} width={52}/>
                <Tooltip content={<CTip/>}/>
                <ReferenceLine x={today()} stroke={WB.borderLight} strokeDasharray="3 3" label={{value:"Today",fill:WB.textDim,fontSize:10}}/>
                <Line type="monotone" dataKey="cumTotal" stroke={WB.primary} strokeWidth={2.5} dot={false} name="Scope Baseline"/>
                {scopeZoom.sel?.active&&scopeZoom.sel.start&&scopeZoom.sel.end&&<ReferenceArea x1={scopeZoom.sel.start} x2={scopeZoom.sel.end} fill={WB.primary} fillOpacity={0.08}/>}
              </LineChart>
            </ResponsiveContainer>
          ):<Empty msg="Link tasks to SOV scopes to see the scope curve."/>}
        </div>
      </div>

      {/* Curve 2: By Sub */}
      <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:20}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
          <div>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:16,fontWeight:700,margin:"0 0 2px",color:WB.text}}>② Project Curve — By Subcontractor</h2>
            <p style={{color:WB.textDim,fontSize:12,margin:0}}>Cumulative S-curve built from each sub's assigned task dates × their contract value</p>
          </div>
          <ZoomBar isZoomed={subZoom.isZoomed} onReset={subZoom.reset}/>
        </div>
        <div style={{height:300,cursor:"crosshair"}}>
          {subAllDates.length?(
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={subZoom.sliced} margin={{top:8,right:8,left:8,bottom:8}} onMouseDown={subZoom.onDown} onMouseMove={subZoom.onMove} onMouseUp={subZoom.onUp}>
                <CartesianGrid strokeDasharray="3 3" stroke={WB.border}/>
                <XAxis dataKey="date" tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} minTickGap={55} tickFormatter={fmtShort}/>
                <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} width={52}/>
                <Tooltip content={<CTip/>}/>
                <ReferenceLine x={today()} stroke={WB.borderLight} strokeDasharray="3 3" label={{value:"Today",fill:WB.textDim,fontSize:10}}/>
                <Line type="monotone" dataKey="cumTotal" stroke={WB.orange} strokeWidth={2.5} dot={false} name="Sub Baseline"/>
                {subZoom.sel?.active&&subZoom.sel.start&&subZoom.sel.end&&<ReferenceArea x1={subZoom.sel.start} x2={subZoom.sel.end} fill={WB.orange} fillOpacity={0.08}/>}
              </LineChart>
            </ResponsiveContainer>
          ):<Empty msg="Assign values to subs in the Sub Values tab to see this curve."/>}
        </div>
      </div>

      {/* Overlay comparison */}
      {allDates.length>0&&subAllDates.length>0&&(
        <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:20}}>
          <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:16,fontWeight:700,margin:"0 0 4px",color:WB.text}}>① vs ② — Overlay Comparison</h2>
          <p style={{color:WB.textDim,fontSize:12,marginBottom:14}}>Blue = scope-based · Orange = sub-based · Divergence shows allocation or timing differences</p>
          <div style={{height:300}}>
            <ResponsiveContainer width="100%" height="100%">
              <LineChart margin={{top:8,right:8,left:8,bottom:8}}>
                <CartesianGrid strokeDasharray="3 3" stroke={WB.border}/>
                <XAxis dataKey="date" type="category" allowDuplicatedCategory={false} tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} minTickGap={55} tickFormatter={fmtShort}/>
                <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} width={52}/>
                <Tooltip content={<CTip/>}/>
                <Legend wrapperStyle={{fontSize:11,color:WB.textDim}}/>
                <ReferenceLine x={today()} stroke={WB.borderLight} strokeDasharray="3 3"/>
                <Line data={allDates} type="monotone" dataKey="cumTotal" stroke={WB.primary} strokeWidth={2.5} dot={false} name="By Scope" strokeDasharray="6 3"/>
                <Line data={subAllDates} type="monotone" dataKey="cumTotal" stroke={WB.orange} strokeWidth={2.5} dot={false} name="By Sub"/>
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
const STORAGE_KEY="wb_scope_tracker_v2";

export default function App({user=null,project=null,onBackToProjects=null,onSignOut=null,supabase=null}={}){
  const [rawTasks,setRawTasks]=useState([]);
  const [resources,setResources]=useState(new Map());
  const [sovResult,setSovResult]=useState(null);
  const [parseErrors,setParseErrors]=useState([]);
  const [schedFile,setSchedFile]=useState("");
  const [sovFile,setSovFile]=useState("");
  const [linksMap,setLinksMap]=useState({});
  const [subSOVLinks,setSubSOVLinks]=useState({});
  const [subContractValues,setSubContractValues]=useState({});
  const [activeTab,setActiveTab]=useState("dashboard");
  const [baselineLocked,setBaselineLocked]=useState(false);
  const [saveStatus,setSaveStatus]=useState("");

  const workTasks=useMemo(()=>rawTasks.filter(t=>!t.isSummary),[rawTasks]);
  const sovItems=sovResult?.items||[];
  const contractTotal=sovResult?.contractTotal||0;

  // Auto-set sub values when SOV links change
  useEffect(()=>{
    if(!sovItems.length||!resources.size)return;
    const scopeSubCount={};
    resources.forEach((_,uid)=>{const sovId=subSOVLinks[uid];if(sovId&&sovId!=="none")scopeSubCount[sovId]=(scopeSubCount[sovId]||0)+1;});
    setSubContractValues(prev=>{
      const next={...prev};
      resources.forEach((_,uid)=>{
        const sovId=subSOVLinks[uid];
        if(!sovId||sovId==="none"){if(!next[uid])next[uid]=0;return;}
        const sov=sovItems.find(s=>s.id===sovId);
        if(!sov){next[uid]=0;return;}
        next[uid]=Math.round(sov.scheduledValue/(scopeSubCount[sovId]||1));
      });
      return next;
    });
  },[subSOVLinks,sovItems.length,resources.size]);

  // Auto-link tasks to SOV
  useEffect(()=>{
    if(!workTasks.length||!sovItems.length||Object.keys(linksMap).length>0)return;
    setLinksMap(autoLink(workTasks,sovItems));
  },[workTasks.length,sovItems.length]);

  // Auto-match subs to SOV
  useEffect(()=>{
    if(!resources.size||!sovItems.length||Object.keys(subSOVLinks).length>0)return;
    const matched=autoMatchSubsToSOV(resources,sovItems);
    const filtered={};
    Object.entries(matched).forEach(([uid,sovId])=>{if(sovId)filtered[uid]=sovId;});
    if(Object.keys(filtered).length>0)setSubSOVLinks(filtered);
  },[resources.size,sovItems.length]);

  // Save
  useEffect(()=>{
    if(!rawTasks.length&&!sovItems.length)return;
    const state={
      rawTasks:rawTasks.map(t=>({...t,startDate:t.startDate?t.startDate.toISOString():null,finishDate:t.finishDate?t.finishDate.toISOString():null})),
      resources:Array.from(resources.entries()),
      sovResult,schedFile,sovFile,linksMap,subSOVLinks,subContractValues,baselineLocked
    };
    if(supabase&&project){
      const t=setTimeout(async()=>{const{error}=await supabase.from("projects").update({data:state,updated_at:new Date().toISOString()}).eq("id",project.id);if(!error){setSaveStatus("saved");setTimeout(()=>setSaveStatus(""),2000);}},1500);
      return()=>clearTimeout(t);
    }else{
      try{localStorage.setItem(STORAGE_KEY,JSON.stringify(state));setSaveStatus("saved");setTimeout(()=>setSaveStatus(""),2000);}catch(e){}
    }
  },[rawTasks,sovResult,linksMap,subSOVLinks,subContractValues,baselineLocked]);

  // Restore
  useEffect(()=>{
    async function restore(){
      let state=null;
      if(supabase&&project){const{data}=await supabase.from("projects").select("data").eq("id",project.id).single();if(data?.data&&Object.keys(data.data).length>0)state=data.data;}
      else{try{const s=localStorage.getItem(STORAGE_KEY);if(s)state=JSON.parse(s);}catch(e){}}
      if(!state)return;
      if(state.rawTasks)setRawTasks(state.rawTasks.map(t=>({...t,startDate:t.startDate?new Date(t.startDate):null,finishDate:t.finishDate?new Date(t.finishDate):null})));
      if(state.resources)setResources(new Map(state.resources));
      if(state.sovResult)setSovResult(state.sovResult);
      if(state.schedFile)setSchedFile(state.schedFile);
      if(state.sovFile)setSovFile(state.sovFile);
      if(state.linksMap)setLinksMap(state.linksMap);
      if(state.subSOVLinks)setSubSOVLinks(state.subSOVLinks);
      if(state.subContractValues)setSubContractValues(state.subContractValues);
      if(state.baselineLocked)setBaselineLocked(state.baselineLocked);
      setSaveStatus("restored");setTimeout(()=>setSaveStatus(""),3000);
    }
    restore();
  },[project?.id]);

  function clearSave(){
    if(supabase&&project)supabase.from("projects").update({data:{}}).eq("id",project.id);
    localStorage.removeItem(STORAGE_KEY);window.location.reload();
  }

  function handleScheduleFile(file){
    setSchedFile(file.name);setLinksMap({});setSubSOVLinks({});setSubContractValues({});
    const reader=new FileReader();
    reader.onload=e=>{
      try{const{rows,errors,resources:res}=parseMSP(e.target.result);setRawTasks(rows);setParseErrors(errors);setResources(res);}
      catch(err){setParseErrors([{row:0,issue:`Parse error: ${err.message}`}]);}
    };
    reader.readAsText(file);
  }

  function handleSOVFile(file){
    setSovFile(file.name);
    const reader=new FileReader();
    reader.onload=e=>{
      try{setSovResult(parseSOV(e.target.result));}
      catch(err){setParseErrors(p=>[...p,{row:0,issue:`SOV error: ${err.message}`}]);}
    };
    reader.readAsArrayBuffer(file);
  }

  const sovColorMap=useMemo(()=>{const m={};sovItems.forEach((s,i)=>{m[s.id]=PAL[i%PAL.length];});return m;},[sovItems]);
  const{byScope,allDates,scopeKeys}=useMemo(()=>buildScopeCurves(workTasks,sovItems,linksMap),[workTasks,sovItems,linksMap]);
  const{bySubUID,subAllDates,subKeys}=useMemo(()=>buildSubCurves(workTasks,resources,subContractValues),[workTasks,resources,subContractValues]);

  const totals=useMemo(()=>({
    contract:contractTotal,base:sovResult?.baseSum||0,gcgr:sovResult?.gcgrTotal||0,ohp:sovResult?.ohpTotal||0,
    scopes:sovItems.length,plotted:scopeKeys.length,linked:new Set(Object.values(linksMap).flat()).size,tasks:workTasks.length,
    resources:resources.size,subsWithCurves:subKeys.length,
  }),[sovResult,sovItems,scopeKeys,linksMap,workTasks,contractTotal,resources,subKeys]);

  const subsWithNoValue=Array.from(resources.keys()).filter(uid=>!subContractValues[uid]||subContractValues[uid]===0).length;
  const hasData=scopeKeys.length>0||subKeys.length>0;

  function exportReport(){
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet([{Metric:"Contract Total",Value:contractTotal},{Metric:"Base",Value:totals.base},{Metric:"Gen Req",Value:totals.gcgr},{Metric:"OH&P",Value:totals.ohp}]),"Summary");
    if(allDates.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(allDates),"Scope Curve");
    if(subAllDates.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(subAllDates),"Sub Curve");
    const subRows=Array.from(resources.entries()).map(([uid,res])=>({Subcontractor:res.name,ContractValue:subContractValues[uid]||0,LinkedSOV:subSOVLinks[uid]||"—"}));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(subRows),"Sub Values");
    XLSX.writeFile(wb,"WB_Scope_Curves.xlsx");
  }

  const TABS=[
    {id:"dashboard",label:"Dashboard",icon:"📊"},
    {id:"scope-curves",label:"Scope Curves",icon:"📈"},
    {id:"sub-curves",label:"Sub Curves",icon:"👷"},
    {id:"subs",label:"Sub Values",icon:"💰",badge:subsWithNoValue||null,badgeColor:WB.yellow},
    {id:"links",label:"Link Tasks",icon:"🔗",badge:totals.tasks-totals.linked>0?totals.tasks-totals.linked:null,badgeColor:WB.yellow},
    {id:"sov",label:"SOV Detail",icon:"📋"},
    {id:"validation",label:"Validation",icon:"⚠",badge:parseErrors.length||null,badgeColor:WB.red},
  ];

  return(
    <div style={{minHeight:"100vh",background:WB.bg,color:WB.text,fontFamily:"'Inter',system-ui,sans-serif",padding:"24px"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Syne:wght@700;800&display=swap');
        *{box-sizing:border-box;}
        ::-webkit-scrollbar{width:4px;height:4px;}
        ::-webkit-scrollbar-track{background:${WB.card};}
        ::-webkit-scrollbar-thumb{background:${WB.border};border-radius:3px;}
        .tab-btn{transition:all .12s;} .tab-btn:hover{color:${WB.text}!important;}
        input::placeholder{color:${WB.textDim}!important;}
        select{appearance:auto;}
        input[type=range]{height:4px;}
      `}</style>
      <div style={{maxWidth:1400,margin:"0 auto"}}>

        {/* Header */}
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:18,flexWrap:"wrap",gap:12}}>
          <div style={{display:"flex",alignItems:"center",gap:16,flexWrap:"wrap"}}>
            <img src="/Wright Brothers Logo.png" alt="Wright Brothers" style={{height:56,objectFit:"contain"}} onError={e=>{e.target.style.display="none";}}/>
            <div>
              {onBackToProjects&&<button onClick={onBackToProjects} style={{background:"none",border:"none",color:WB.primaryLight,cursor:"pointer",fontSize:11,padding:0,marginBottom:4,display:"block"}}>← All Projects</button>}
              <h1 style={{fontFamily:"'Syne',sans-serif",fontSize:22,fontWeight:800,margin:0,lineHeight:1.1,color:WB.text}}>{project?.name||"Project Spend Baseline Tracker"}</h1>
              {user&&<div style={{fontSize:11,color:WB.textDim,marginTop:2}}>{user.email}</div>}
            </div>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",alignSelf:"flex-end"}}>
            {saveStatus==="saved"&&<span style={{fontSize:11,color:WB.green,display:"flex",alignItems:"center",gap:4}}><Icon.Check color={WB.green}/> Saved</span>}
            {saveStatus==="restored"&&<span style={{fontSize:11,color:WB.primaryLight,display:"flex",alignItems:"center",gap:4}}><Icon.Check color={WB.primaryLight}/> Restored</span>}
            <button onClick={clearSave} style={{padding:"6px 12px",background:"transparent",border:`1px solid ${WB.border}`,borderRadius:8,color:WB.textDim,cursor:"pointer",fontSize:11}}>Clear Data</button>
            {onSignOut&&<button onClick={onSignOut} style={{padding:"6px 12px",background:"transparent",border:`1px solid ${WB.border}`,borderRadius:8,color:WB.textDim,cursor:"pointer",fontSize:11}}>Sign Out</button>}
            {hasData&&!baselineLocked&&(
              <button onClick={()=>setBaselineLocked(true)}
                style={{display:"flex",alignItems:"center",gap:6,padding:"8px 14px",background:"#1e3448",border:`1px solid ${WB.primary}`,borderRadius:8,color:WB.primaryLight,cursor:"pointer",fontSize:12,fontWeight:600}}>
                <Icon.Lock/> Lock Baseline
              </button>
            )}
            {baselineLocked&&(
              <span style={{display:"flex",alignItems:"center",gap:5,padding:"7px 12px",background:WB.greenBg,border:`1px solid ${WB.greenBorder}`,borderRadius:8,color:WB.green,fontSize:12,fontWeight:600}}>
                <Icon.Lock/> Baseline Locked
              </span>
            )}
            <button onClick={exportReport} disabled={!hasData}
              style={{display:"flex",alignItems:"center",gap:6,padding:"8px 14px",background:"transparent",border:`1px solid ${WB.border}`,borderRadius:8,color:hasData?WB.textMuted:WB.textDim,cursor:hasData?"pointer":"not-allowed",fontSize:12}}>
              <Icon.Download/> Export
            </button>
          </div>
        </div>

        {/* Uploads */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
          <DropZone label="Upload MS Project XML (Resource-Loaded)" accept=".xml" onFile={handleScheduleFile}
            fileName={schedFile?`${schedFile} · ${workTasks.length} tasks${resources.size>0?` · ${resources.size} subs`:""}`:""}
            icon={<Icon.File color={WB.primary}/>} hint="File → Save As → XML Format · Resource-loaded = sub curves enabled · Non-resource-loaded = scope curves only"/>
          <DropZone label="Upload SOV Excel" accept=".xlsx,.xls,.csv" onFile={handleSOVFile}
            fileName={sovFile?`${sovFile} · ${sovItems.length} scopes`:""}
            icon={<Icon.Layers color={WB.green}/>} hint="GC/GR and OH&P auto-detected and distributed"/>
        </div>

        {/* SOV info */}
        {sovResult&&(
          <div style={{display:"flex",gap:8,alignItems:"center",padding:"8px 12px",background:WB.card,border:`1px solid ${WB.border}`,borderRadius:8,marginBottom:12,fontSize:12,flexWrap:"wrap"}}>
            <Icon.Info color={WB.primary}/>
            <span style={{color:WB.textDim}}>SOV:</span>
            <span style={{color:WB.text,fontFamily:"monospace"}}>{currency(totals.base)}</span><span style={{color:WB.textDim}}>base +</span>
            <span style={{color:WB.orange,fontFamily:"monospace"}}>{currency(totals.gcgr)}</span><span style={{color:WB.textDim}}>Gen Req +</span>
            <span style={{color:"#a855f7",fontFamily:"monospace"}}>{currency(totals.ohp)}</span><span style={{color:WB.textDim}}>OH&P =</span>
            <span style={{color:WB.green,fontWeight:700,fontFamily:"monospace"}}>{currency(totals.contract)}</span>
            {sovResult.parseWarning&&<span style={{color:WB.yellow,marginLeft:6}}>⚠ {sovResult.parseWarning}</span>}
          </div>
        )}

        {/* KPIs */}
        {(workTasks.length>0||sovItems.length>0)&&(
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:9,marginBottom:14}}>
            <Metric title="Contract Total" value={currency(totals.contract)} color={WB.green}/>
            <Metric title="Scopes Plotted" value={`${totals.plotted} / ${totals.scopes}`} color={totals.scopes-totals.plotted>0?WB.yellow:WB.green}/>
            <Metric title="Tasks Linked" value={`${totals.linked} / ${totals.tasks}`} color={WB.primary}/>
            <Metric title="Resources Found" value={totals.resources} color={WB.orange} sub="from schedule"/>
            <Metric title="Subs w/ Curves" value={totals.subsWithCurves} color={WB.primaryLight}/>
          </div>
        )}

        {/* Tabs */}
        <div style={{display:"flex",gap:2,marginBottom:14,background:WB.card,padding:3,borderRadius:10,border:`1px solid ${WB.border}`,width:"fit-content",flexWrap:"wrap"}}>
          {TABS.map(t=>(
            <button key={t.id} className="tab-btn" onClick={()=>setActiveTab(t.id)}
              style={{display:"flex",alignItems:"center",gap:5,padding:"6px 12px",borderRadius:7,border:"none",cursor:"pointer",fontSize:12,fontWeight:500,background:activeTab===t.id?WB.border:"transparent",color:activeTab===t.id?WB.text:WB.textDim}}>
              <span>{t.icon}</span>{t.label}
              {t.badge?<span style={{background:t.badgeColor,color:"#fff",borderRadius:99,padding:"1px 5px",fontSize:9,fontWeight:700}}>{t.badge}</span>:null}
            </button>
          ))}
        </div>

        {activeTab==="dashboard"&&(
          hasData
            ?<Dashboard allDates={allDates} subAllDates={subAllDates} scopeKeys={scopeKeys} subKeys={subKeys} contractTotal={contractTotal} subContractValues={subContractValues} resources={resources}/>
            :<Empty msg="Upload a resource-loaded XML schedule and SOV Excel to get started."/>
        )}

        {activeTab==="scope-curves"&&(
          <ScopeCurvesTab byScope={byScope} scopeKeys={scopeKeys} allDates={allDates} sovColorMap={sovColorMap} sovItems={sovItems}/>
        )}

        {activeTab==="sub-curves"&&(
          <SubCurvesTab bySubUID={bySubUID} subKeys={subKeys} subAllDates={subAllDates} resources={resources}/>
        )}

        {activeTab==="subs"&&(
          resources.size>0&&sovItems.length>0
            ?<SubDollarOverride resources={resources} sovItems={sovItems} subContractValues={subContractValues} onUpdate={setSubContractValues} subSOVLinks={subSOVLinks} onSOVLinkChange={(uid,sovId)=>setSubSOVLinks(prev=>({...prev,[uid]:sovId}))}/>
            :resources.size===0
              ?<Empty msg="No resources found in your schedule — this tab is only available for resource-loaded XML files. You can still use Scope Curves and Link Tasks with any schedule."/>
              :<Empty msg="Upload your SOV Excel to link subcontractors to scope values."/>
        )}

        {activeTab==="links"&&(
          workTasks.length&&sovItems.length
            ?<LinkManager tasks={workTasks} sovItems={sovItems} linksMap={linksMap} onLinksChange={setLinksMap} resources={resources}/>
            :<Empty msg="Upload both files to use the link manager."/>
        )}

        {activeTab==="sov"&&(
          <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:18,overflowX:"auto"}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 4px",color:WB.text}}>SOV Line Items</h2>
            <p style={{color:WB.textDim,fontSize:12,marginBottom:14}}>Base + Gen Req + OH&P = contract value per scope</p>
            {sovItems.length?(
              <table style={{width:"100%",minWidth:620,borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{background:WB.bg}}>{["Description","Base","+ Gen Req","+ OH&P","Contract Value","Tasks Linked"].map(h=><th key={h} style={{padding:"7px 9px",textAlign:"left",color:WB.textDim,fontFamily:"monospace",fontSize:9,borderBottom:`1px solid ${WB.border}`,textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                <tbody>
                  {sovItems.map((s,i)=>{const col=PAL[i%PAL.length];const count=(linksMap[s.id]||[]).length;return(
                    <tr key={s.id} style={{borderBottom:`1px solid ${WB.border}`,background:i%2===0?"transparent":WB.bg}}>
                      <td style={{padding:"7px 9px",color:WB.text,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.description}</td>
                      <td style={{padding:"7px 9px",color:WB.textMuted,fontFamily:"monospace"}}>{currency(s.baseValue)}</td>
                      <td style={{padding:"7px 9px",color:WB.orange,fontFamily:"monospace",fontSize:10}}>{currency(s.gcgrShare)}</td>
                      <td style={{padding:"7px 9px",color:"#a855f7",fontFamily:"monospace",fontSize:10}}>{currency(s.ohpShare)}</td>
                      <td style={{padding:"7px 9px",fontWeight:700,color:col,fontFamily:"monospace"}}>{currency(s.scheduledValue)}</td>
                      <td style={{padding:"7px 9px"}}><span style={{fontSize:10,color:count>0?col:WB.textDim,background:count>0?`${col}20`:WB.bg,padding:"1px 7px",borderRadius:99,border:`1px solid ${count>0?`${col}40`:WB.border}`}}>{count}</span></td>
                    </tr>
                  );})}
                </tbody>
                <tfoot><tr style={{borderTop:`2px solid ${WB.border}`}}>
                  <td style={{padding:"8px 9px",fontWeight:700,color:WB.text}}>TOTAL</td>
                  <td style={{padding:"8px 9px",fontWeight:700,color:WB.textMuted,fontFamily:"monospace"}}>{currency(totals.base)}</td>
                  <td style={{padding:"8px 9px",fontWeight:700,color:WB.orange,fontFamily:"monospace"}}>{currency(totals.gcgr)}</td>
                  <td style={{padding:"8px 9px",fontWeight:700,color:"#a855f7",fontFamily:"monospace"}}>{currency(totals.ohp)}</td>
                  <td style={{padding:"8px 9px",fontWeight:700,color:WB.green,fontFamily:"monospace",fontSize:14}}>{currency(totals.contract)}</td>
                  <td/>
                </tr></tfoot>
              </table>
            ):<Empty msg="Upload your SOV to see the breakdown."/>}
          </div>
        )}

        {activeTab==="validation"&&(
          <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:18}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 4px",color:WB.text}}>Validation Log</h2>
            <p style={{color:WB.textDim,fontSize:12,marginBottom:14}}>Parse issues from uploaded files.</p>
            {parseErrors.length
              ?(<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(270px,1fr))",gap:7}}>{parseErrors.map((e,i)=>(<div key={i} style={{background:WB.redBg,border:`1px solid ${WB.redBorder}`,borderRadius:8,padding:"9px 12px"}}><div style={{fontFamily:"monospace",fontSize:9,color:WB.red,marginBottom:2}}>ROW {e.row||"—"}</div><div style={{color:WB.text,fontSize:12}}>{e.issue}</div></div>))}</div>)
              :(<div style={{display:"flex",alignItems:"center",gap:9,padding:"11px 14px",background:WB.greenBg,border:`1px solid ${WB.greenBorder}`,borderRadius:8}}><Icon.Check color={WB.green}/><span style={{color:"#86efac",fontSize:12}}>{rawTasks.length?"No issues found.":"Upload a schedule to validate."}</span></div>)
            }
          </div>
        )}

      </div>
    </div>
  );
}
