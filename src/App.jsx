import { useMemo, useState, useEffect, useCallback } from "react";
// No lucide-react needed — using inline icon components
const Upload=()=><span>↑</span>;
const AlertTriangle=({size,color,style})=><span style={{color,...(style||{})}}>⚠</span>;
const CheckCircle2=({size,color,style})=><span style={{color,...(style||{})}}>✓</span>;
const Download=()=><span>↓</span>;
const BarChart3=({size,color})=><span style={{color}}>📊</span>;
const TrendingUp=({size})=><span>📈</span>;
const FileText=({size,color})=><span style={{color}}>📄</span>;
const Layers=({size,color})=><span style={{color}}>📋</span>;
const Link2=({size,color})=><span style={{color}}>🔗</span>;
const ZoomOut=({size})=><span>🔍</span>;
const Info=({size,color})=><span style={{color}}>ℹ</span>;
const Lock=({size})=><span>🔒</span>;
const Activity=({size})=><span>📉</span>;
const RefreshCw=({size})=><span>↻</span>;
const Circle=({size,fill,stroke})=><span style={{color:fill,fontSize:size}}>●</span>;
import { AreaChart, Area, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid, ReferenceArea, LineChart, Line, ReferenceLine, Legend } from "recharts";
import * as XLSX from "xlsx";

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
function fmtShort(s){if(!s||s==="—")return"";try{const d=new Date(s);return d.toLocaleDateString("en-US",{month:"short",year:"2-digit"});}catch{return s;}}
function currency(v){return new Intl.NumberFormat("en-US",{style:"currency",currency:"USD",maximumFractionDigits:0}).format(v||0);}
function pct(v){return`${Math.round((v||0)*100)}%`;}
function getText(node,tag){const el=node?.getElementsByTagName(tag)?.[0];return el?.textContent?.trim()||"";}
function today(){return fmtDate(new Date());}

// ─── SOV parser ───────────────────────────────────────────────────────────────
const SUBTOTAL_RE=/grand|total|subtotal|sub-total|\bsum\b/i;
const GCGR_RE=/general\s*req|gen\s*req|\bGR\b|\bGC\b|general\s*condition|mobilization/i;
const OHP_RE=/oh\s*[&+]\s*p|overhead.*profit|markup|\bprofit\b|\bfee\b/i;
function classifyRow(desc,val){
  if(!desc&&!val)return"skip"; if(val===0)return"skip";
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
    if(val>0&&desc)lineItems.push({id:`sov-${lineItems.length}`,description:desc,baseValue:val,phase:row.find(c=>typeof c==="number"&&c>0&&c<100&&Number.isInteger(c))||null});
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

// ─── MS Project XML parser ────────────────────────────────────────────────────
function parseMSP(xmlText){
  const parser=new DOMParser();
  const doc=parser.parseFromString(xmlText,"application/xml");
  if(doc.querySelector("parsererror"))throw new Error("Invalid XML.");
  const aMap=new Map();
  doc.querySelectorAll("Assignment").forEach(a=>{
    const uid=getText(a,"TaskUID");if(!uid)return;
    const prev=aMap.get(uid)||{cost:0,actual:0};
    aMap.set(uid,{cost:prev.cost+toMoney(getText(a,"Cost")),actual:prev.actual+toMoney(getText(a,"ActualCost"))});
  });
  const rows=[],errors=[];
  doc.querySelectorAll("Task").forEach((task,i)=>{
    const uid=getText(task,"UID");if(uid==="0")return;
    const name=getText(task,"Name");if(!name)return;
    const isSummary=getText(task,"Summary")==="1";
    const startDate=toDate(getText(task,"Start")),finishDate=toDate(getText(task,"Finish"));
    const wbs=getText(task,"WBS")||uid;
    const a=aMap.get(uid)||{cost:0,actual:0};
    if(!isSummary){if(!startDate)errors.push({row:i+1,issue:`"${name}" — missing Start`});if(!finishDate)errors.push({row:i+1,issue:`"${name}" — missing Finish`});}
    rows.push({rowNum:i+1,id:wbs,uid,name,startDate,finishDate,budgetedCost:a.cost>0?a.cost:toMoney(getText(task,"Cost")),actualCost:a.actual>0?a.actual:toMoney(getText(task,"ActualCost")),percentComplete:toPercent(getText(task,"PercentComplete")||"0"),isSummary});
  });
  return{rows,errors};
}

// ─── Fuzzy matching ───────────────────────────────────────────────────────────
const STOP=new Set(["the","a","an","of","in","at","to","and","or","for","with","on","is","are","be","by","as","from","this","it","work","scope","item","all","per","each","existing","new","provide","install","furnish","allowance"]);
const ALIASES={"demo":["demolition","demolish","remove","removal"],"demolition":["demo","remove","removal"],"conc":["concrete"],"concrete":["conc","slab","footing","foundation"],"elec":["electrical","electric","power"],"electrical":["elec","electric","power","wiring"],"mech":["mechanical","hvac","plumbing"],"hvac":["mechanical","heating","cooling","mech"],"plumb":["plumbing","pipe"],"plumbing":["plumb","pipe","piping","sanitary"],"struct":["structural","steel","framing"],"structural":["struct","steel","frame","framing"],"framing":["structural","frame","steel","wood","lumber"],"masonry":["brick","block","cmu","stone"],"drywall":["gypsum","gwb","partition","wall"],"flooring":["floor","tile","carpet","lvt"],"roofing":["roof","membrane","waterproof"],"sitework":["site","earthwork","grading","excavation","utilities"],"painting":["paint","coating","finish"],"insulation":["insulate","thermal","acoustic"],"ceiling":["ceilings","acoustic","tile","grid"],"millwork":["casework","cabinetry","woodwork"],"landscaping":["landscape","planting","irrigation","lawn"],"permit":["permits","permitting"],"siding":["cladding","facade","exterior"],"doors":["door","hardware","frame"],"windows":["window","glazing","storefront"],"appliances":["appliance","equipment"],"countertops":["countertop","stone","quartz","granite"],"specialties":["specialty","signage","accessories"]};
function tokenize(s){return String(s||"").toLowerCase().replace(/[^a-z0-9\s]/g," ").split(/\s+/).filter(t=>t.length>1&&!STOP.has(t));}
function expand(tokens){const s=new Set(tokens);tokens.forEach(t=>{if(ALIASES[t])ALIASES[t].forEach(a=>s.add(a));});return Array.from(s);}
function matchScore(a,b){const ta=expand(tokenize(a)),tb=expand(tokenize(b));if(!ta.length||!tb.length)return 0;const sa=new Set(ta),sb=new Set(tb);let ov=0;sa.forEach(t=>{if(sb.has(t))ov++;});const union=new Set([...sa,...sb]).size;const j=union>0?ov/union:0;const oa=new Set(tokenize(a)),ob=new Set(tokenize(b));let orig=0;oa.forEach(t=>{if(ob.has(t))orig++;});return Math.min(1,j+orig*0.15);}
function autoLink(tasks,sovItems){
  const map={};
  tasks.forEach(t=>{if(t.isSummary)return;let best=null,bestScore=0;sovItems.forEach(s=>{const sc=matchScore(t.name,s.description);if(sc>bestScore){bestScore=sc;best=s;}});if(best&&bestScore>=0.08){if(!map[best.id])map[best.id]=[];map[best.id].push(t.uid);}});
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

// ─── Build curves (unified spine) ────────────────────────────────────────────
function buildCurves(tasks,sovItems,linksMap){
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
  const lastDate=fmtDate(new Date(globalMax));
  if(allDates.length&&allDates[allDates.length-1].date!==lastDate){
    const row={date:lastDate,total:0};
    scopeIds.forEach(sid=>{const val=byScope[sid].lookup.get(lastDate)||0;const key=byScope[sid].sov.description.slice(0,18);row[key]=(row[key]||0)+Math.round(val);row.total+=val;});
    row.total=Math.round(row.total);cumTotal+=row.total;row.cumTotal=Math.round(cumTotal);allDates.push(row);
  }
  scopeIds.forEach(sid=>{const{full}=byScope[sid];const s=Math.max(1,Math.floor(full.length/120));byScope[sid].weekly=full.filter((_,i)=>i%s===0||(i===full.length-1));});
  return{byScope,allDates,scopeKeys:scopeIds};
}

// ─── Build actual earned value curve from progress updates ────────────────────
function buildActualCurve(sovItems,linksMap,progressUpdates,byScope,allDates){
  // progressUpdates: { sovId: { percentComplete: 0-1, asOfDate: 'YYYY-MM-DD' } }
  if(!allDates.length)return[];
  return allDates.map(d=>{
    let actualCum=0;
    sovItems.forEach(sov=>{
      const prog=progressUpdates[sov.id];
      if(!prog||!prog.asOfDate||d.date>prog.asOfDate)return;
      // Earned = scheduledValue * percentComplete, distributed up to asOfDate
      actualCum+=sov.scheduledValue*(prog.percentComplete||0);
    });
    return{...d,actualCum:Math.round(actualCum)};
  });
}

// ─── Colors ───────────────────────────────────────────────────────────────────
const PAL=["#3b82f6","#22c55e","#f97316","#a855f7","#06b6d4","#f43f5e","#eab308","#10b981","#8b5cf6","#ec4899","#14b8a6","#84cc16","#ef4444","#6366f1","#f59e0b"];

// ─── Zoom hook ────────────────────────────────────────────────────────────────
function useZoom(data){
  const [range,setRange]=useState(null);
  const [sel,setSel]=useState(null);
  const sliced=useMemo(()=>(!range||!data.length)?data:data.slice(range[0],range[1]+1),[data,range]);
  function onDown(e){if(e?.activeLabel)setSel({start:e.activeLabel,end:null,active:true});}
  function onMove(e){if(sel?.active&&e?.activeLabel)setSel(s=>({...s,end:e.activeLabel}));}
  function onUp(){
    if(sel?.active&&sel.start&&sel.end&&sel.start!==sel.end){
      const a=data.findIndex(d=>d.date===sel.start),b=data.findIndex(d=>d.date===sel.end);
      if(a>=0&&b>=0)setRange([Math.min(a,b),Math.max(a,b)]);
    }
    setSel(null);
  }
  return{sliced,onDown,onMove,onUp,reset:()=>setRange(null),isZoomed:!!range,sel};
}

// ─── UI atoms ────────────────────────────────────────────────────────────────
function CTip({active,payload,label}){
  if(!active||!payload?.length)return null;
  const items=payload.filter(p=>p.value>0).slice(0,8);
  return(
    <div style={{background:"#0f172a",border:"1px solid #1e293b",borderRadius:10,padding:"10px 13px",fontSize:12,maxWidth:280,boxShadow:"0 8px 32px rgba(0,0,0,.6)"}}>
      <div style={{color:"#64748b",marginBottom:5,fontFamily:"monospace",fontSize:10}}>{fmtShort(label)}</div>
      {items.map(p=>(
        <div key={p.dataKey} style={{display:"flex",gap:7,alignItems:"center",marginBottom:2}}>
          <span style={{width:7,height:7,borderRadius:"50%",background:p.color,flexShrink:0,display:"inline-block"}}/>
          <span style={{color:"#94a3b8",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontSize:11}}>{p.dataKey}:</span>
          <span style={{fontWeight:600,flexShrink:0}}>{currency(p.value)}</span>
        </div>
      ))}
    </div>
  );
}

function Metric({title,value,sub,color="#3b82f6",large}){
  return(
    <div style={{background:"#0f172a",borderRadius:12,padding:large?"18px 20px":"13px 15px",border:"1px solid #1e293b",position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:0,right:0,width:55,height:55,background:color,opacity:.07,borderRadius:"0 12px 0 55px"}}/>
      <div style={{fontSize:9,fontFamily:"monospace",color:"#475569",textTransform:"uppercase",letterSpacing:1,marginBottom:5}}>{title}</div>
      <div style={{fontSize:large?24:18,fontWeight:700,color:"#f1f5f9",lineHeight:1}}>{value}</div>
      {sub&&<div style={{fontSize:11,color,marginTop:3}}>{sub}</div>}
    </div>
  );
}

function DropZone({label,accept,onFile,fileName,icon,hint}){
  const [drag,setDrag]=useState(false);
  return(
    <div>
      <label onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
        onDrop={e=>{e.preventDefault();setDrag(false);const f=e.dataTransfer.files?.[0];if(f)onFile(f);}}
        style={{display:"flex",flexDirection:"column",alignItems:"center",gap:8,padding:"16px",border:`2px dashed ${drag?"#3b82f6":fileName?"#1e4d2b":"#1e293b"}`,borderRadius:12,cursor:"pointer",background:drag?"#0c1a2e":fileName?"#020d06":"#020817",transition:"all .2s",textAlign:"center"}}>
        <div style={{width:34,height:34,borderRadius:8,background:"#0f172a",display:"flex",alignItems:"center",justifyContent:"center",border:"1px solid #1e293b"}}>{icon}</div>
        <div>
          <div style={{fontWeight:600,color:fileName?"#22c55e":"#cbd5e1",fontSize:13,marginBottom:2}}>{fileName||label}</div>
          <div style={{fontSize:11,color:"#334155"}}>{fileName?"Loaded — click to replace":"Drag & drop or click"}</div>
        </div>
        <input type="file" accept={accept} style={{display:"none"}} onChange={e=>e.target.files?.[0]&&onFile(e.target.files[0])}/>
      </label>
      {hint&&<div style={{fontSize:11,color:"#334155",marginTop:3}}>{hint}</div>}
    </div>
  );
}

function ZoomBar({isZoomed,onReset}){
  return isZoomed
    ?<button onClick={onReset} style={{display:"flex",alignItems:"center",gap:5,padding:"4px 9px",background:"#1e293b",border:"1px solid #334155",borderRadius:6,color:"#94a3b8",cursor:"pointer",fontSize:11}}><ZoomOut size={10}/> Reset zoom</button>
    :<span style={{fontSize:11,color:"#334155"}}>Drag on chart to zoom</span>;
}

function Empty({msg}){
  return(<div style={{minHeight:160,display:"flex",alignItems:"center",justifyContent:"center",border:"1px dashed #1e293b",borderRadius:12}}><div style={{fontSize:13,color:"#334155",textAlign:"center",padding:20,maxWidth:300}}>{msg}</div></div>);
}

function TrafficLight({status}){
  const col=status==="green"?"#22c55e":status==="yellow"?"#eab308":"#ef4444";
  const bg=status==="green"?"#052e16":status==="yellow"?"#1a1200":"#1a0505";
  const label=status==="green"?"On Track":status==="yellow"?"Watch":"Behind";
  return(<span style={{display:"inline-flex",alignItems:"center",gap:4,fontSize:11,color:col,background:bg,padding:"2px 8px",borderRadius:99,border:`1px solid ${col}40`}}><Circle size={7} fill={col} stroke="none"/>{label}</span>);
}

// ─── Link Manager ─────────────────────────────────────────────────────────────
function LinkManager({tasks,sovItems,linksMap,onLinksChange}){
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
    <div style={{background:"#0f172a",borderRadius:14,border:"1px solid #1e293b",padding:18}}>
      <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 3px",color:"#f1f5f9"}}>Link Tasks → SOV Scopes</h2>
      <p style={{color:"#475569",fontSize:12,margin:"0 0 12px"}}>Assign each task to a SOV scope. Fuzzy matches are pre-filled.</p>
      <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap",marginBottom:12}}>
        <input value={taskFilter} onChange={e=>setTaskFilter(e.target.value)} placeholder="Filter tasks…"
          style={{padding:"6px 10px",background:"#020817",border:"1px solid #1e293b",borderRadius:7,color:"#94a3b8",fontSize:12,outline:"none",width:200}}/>
        <label style={{display:"flex",alignItems:"center",gap:5,fontSize:12,color:"#64748b",cursor:"pointer"}}>
          <input type="checkbox" checked={unlinkedOnly} onChange={e=>setUnlinkedOnly(e.target.checked)} style={{accentColor:"#3b82f6"}}/>
          Unlinked only
        </label>
        <span style={{fontSize:12,color:"#475569",marginLeft:"auto"}}><span style={{color:"#22c55e",fontWeight:600}}>{linked}</span> linked · <span style={{color:linked<workTasks.length?"#f59e0b":"#22c55e",fontWeight:600}}>{workTasks.length-linked}</span> unlinked</span>
      </div>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",minWidth:640,borderCollapse:"collapse",fontSize:12}}>
          <thead><tr style={{background:"#020817"}}>{["WBS","Task Name","Dates","SOV Scope"].map(h=><th key={h} style={{padding:"7px 10px",textAlign:"left",color:"#475569",fontFamily:"monospace",fontSize:9,borderBottom:"1px solid #1e293b",textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
          <tbody>
            {filtered.map((t,i)=>{
              const sid=taskToSov[t.uid];const col=sid?sovColorMap[sid]:null;
              return(
                <tr key={t.uid} style={{borderBottom:"1px solid #1e293b",background:i%2===0?"transparent":"#020817"}}>
                  <td style={{padding:"7px 10px",fontFamily:"monospace",color:"#3b82f6",fontSize:10,whiteSpace:"nowrap"}}>{t.id}</td>
                  <td style={{padding:"7px 10px",color:"#cbd5e1",maxWidth:240}}>
                    <div style={{display:"flex",alignItems:"center",gap:6}}>{col&&<div style={{width:3,height:14,borderRadius:2,background:col,flexShrink:0}}/>}<span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{t.name}</span></div>
                  </td>
                  <td style={{padding:"7px 10px",fontFamily:"monospace",color:"#475569",fontSize:10,whiteSpace:"nowrap"}}>{fmtDate(t.startDate)} → {fmtDate(t.finishDate)}</td>
                  <td style={{padding:"7px 10px",minWidth:240}}>
                    <select value={sid||"none"} onChange={e=>setTaskLink(t.uid,e.target.value)}
                      style={{width:"100%",padding:"5px 9px",background:col?`${col}18`:"#020817",border:`1px solid ${col||"#1e293b"}`,borderRadius:6,color:col||"#64748b",fontSize:12,cursor:"pointer",outline:"none",fontWeight:col?600:400}}>
                      <option value="none" style={{background:"#020817",color:"#475569"}}>— Unassigned —</option>
                      {sovItems.map(s=><option key={s.id} value={s.id} style={{background:"#020817",color:"#cbd5e1"}}>{s.description} ({currency(s.scheduledValue)})</option>)}
                    </select>
                  </td>
                </tr>
              );
            })}
            {!filtered.length&&<tr><td colSpan={4} style={{padding:24,textAlign:"center",color:"#334155"}}>No tasks match filter.</td></tr>}
          </tbody>
        </table>
      </div>
      <div style={{marginTop:16,paddingTop:14,borderTop:"1px solid #1e293b"}}>
        <div style={{fontSize:11,fontFamily:"monospace",color:"#475569",textTransform:"uppercase",letterSpacing:1,marginBottom:8}}>Scope Task Count</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:7}}>
          {sovItems.map((s,i)=>{const col=PAL[i%PAL.length];const count=(linksMap[s.id]||[]).length;return(<div key={s.id} style={{padding:"8px 10px",borderRadius:8,background:`${col}10`,border:`1px solid ${count>0?col:"#1e293b"}`,display:"flex",justifyContent:"space-between",alignItems:"center",gap:6}}><div style={{overflow:"hidden"}}><div style={{fontSize:11,color:"#cbd5e1",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.description}</div><div style={{fontSize:10,color:col,fontFamily:"monospace"}}>{currency(s.scheduledValue)}</div></div><span style={{flexShrink:0,fontSize:10,fontFamily:"monospace",color:count>0?col:"#334155",background:count>0?`${col}20`:"#0f172a",padding:"1px 7px",borderRadius:99,border:`1px solid ${count>0?`${col}40`:"#1e293b"}`}}>{count}</span></div>);})}
        </div>
      </div>
    </div>
  );
}

// ─── Progress Update Panel ────────────────────────────────────────────────────
function ProgressPanel({sovItems,progressUpdates,onUpdate,baseline}){
  const [updateDate,setUpdateDate]=useState(today());
  function setProgress(sovId,val){onUpdate({...progressUpdates,[sovId]:{percentComplete:Math.min(1,Math.max(0,val/100)),asOfDate:updateDate}});}
  const totalEarned=sovItems.reduce((s,sov)=>{const p=progressUpdates[sov.id];return s+(p?sov.scheduledValue*p.percentComplete:0);},0);
  const totalContract=sovItems.reduce((s,sov)=>s+sov.scheduledValue,0);
  return(
    <div style={{background:"#0f172a",borderRadius:14,border:"1px solid #1e293b",padding:18}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16,flexWrap:"wrap",gap:10}}>
        <div>
          <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 3px",color:"#f1f5f9"}}>Progress Update</h2>
          <p style={{color:"#475569",fontSize:12,margin:0}}>Enter % complete per scope as of a given date to track earned value vs. baseline.</p>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:12,color:"#64748b"}}>As of date:</span>
          <input type="date" value={updateDate} onChange={e=>setUpdateDate(e.target.value)}
            style={{padding:"5px 9px",background:"#020817",border:"1px solid #1e293b",borderRadius:7,color:"#94a3b8",fontSize:12,outline:"none"}}/>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:8,marginBottom:16}}>
        <Metric title="Total Earned" value={currency(totalEarned)} color="#22c55e" sub={`${Math.round(totalContract>0?totalEarned/totalContract*100:0)}% of contract`}/>
        <Metric title="Remaining" value={currency(totalContract-totalEarned)} color="#f97316"/>
        {baseline&&<Metric title="Baseline at Date" value={currency(baseline)} color="#3b82f6" sub="planned spend"/>}
        {baseline&&<Metric title="Variance" value={currency(totalEarned-baseline)} color={totalEarned>=baseline?"#22c55e":"#ef4444"} sub={totalEarned>=baseline?"Ahead":"Behind"}/>}
      </div>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
        <thead><tr style={{background:"#020817"}}>{["Scope","Contract Value","% Complete","Earned Value","Status"].map(h=><th key={h} style={{padding:"7px 10px",textAlign:"left",color:"#475569",fontFamily:"monospace",fontSize:9,borderBottom:"1px solid #1e293b",textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
        <tbody>
          {sovItems.map((sov,i)=>{
            const col=PAL[i%PAL.length];
            const prog=progressUpdates[sov.id];
            const pctVal=prog?Math.round(prog.percentComplete*100):0;
            const earned=sov.scheduledValue*(prog?.percentComplete||0);
            const status=pctVal>=80?"green":pctVal>=40?"yellow":"red";
            return(
              <tr key={sov.id} style={{borderBottom:"1px solid #1e293b"}}>
                <td style={{padding:"8px 10px",color:"#cbd5e1",maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                  <div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:3,height:14,borderRadius:2,background:col,flexShrink:0}}/>{sov.description}</div>
                </td>
                <td style={{padding:"8px 10px",color:"#94a3b8",fontFamily:"monospace"}}>{currency(sov.scheduledValue)}</td>
                <td style={{padding:"8px 10px",minWidth:160}}>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <input type="range" min={0} max={100} value={pctVal} onChange={e=>setProgress(sov.id,Number(e.target.value))}
                      style={{flex:1,accentColor:col,cursor:"pointer"}}/>
                    <input type="number" min={0} max={100} value={pctVal} onChange={e=>setProgress(sov.id,Number(e.target.value))}
                      style={{width:46,padding:"3px 5px",background:"#020817",border:`1px solid ${col}`,borderRadius:5,color:col,fontSize:12,fontFamily:"monospace",outline:"none",textAlign:"center"}}/>
                    <span style={{color:"#475569",fontSize:11}}>%</span>
                  </div>
                </td>
                <td style={{padding:"8px 10px",color:col,fontFamily:"monospace",fontWeight:600}}>{currency(earned)}</td>
                <td style={{padding:"8px 10px"}}>{prog?<TrafficLight status={status}/>:<span style={{fontSize:11,color:"#334155"}}>—</span>}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

// ─── Health Dashboard ─────────────────────────────────────────────────────────
function Dashboard({sovItems,allDates,byScope,sovKeys,progressUpdates,baselineLocked,contractTotal,combinedZoom,PAL}){
  const sovColorMap=useMemo(()=>{const m={};sovItems.forEach((s,i)=>{m[s.id]=PAL[i%PAL.length];});return m;},[sovItems]);
  const totalEarned=sovItems.reduce((s,sov)=>{const p=progressUpdates[sov.id];return s+(p?sov.scheduledValue*p.percentComplete:0);},0);
  const cpi=contractTotal>0?totalEarned/contractTotal:0;
  // Find planned value at today
  const todayStr=today();
  const todayPoint=allDates.find(d=>d.date>=todayStr)||allDates[allDates.length-1];
  const plannedToDate=todayPoint?.cumTotal||0;
  const spi=plannedToDate>0?totalEarned/plannedToDate:0;
  const variance=totalEarned-plannedToDate;
  // Scope health
  const scopeHealth=sovItems.map((sov,i)=>{
    const prog=progressUpdates[sov.id];
    const pct=prog?.percentComplete||0;
    const scope=byScope[sov.id];
    // Find where in baseline this scope should be today
    let baselineToday=0;
    if(scope){
      const todayPt=scope.full.find(p=>p.date>=todayStr)||scope.full[scope.full.length-1];
      baselineToday=todayPt?todayPt.cumulative/sov.scheduledValue:0;
    }
    const diff=pct-baselineToday;
    const status=diff>=-0.05?"green":diff>=-0.15?"yellow":"red";
    return{sov,pct,baselineToday,diff,status,col:PAL[i%PAL.length]};
  });
  // S-curve data with actual overlay
  const sCurveData=useMemo(()=>combinedZoom.sliced.map(d=>{
    let actualCum=0;
    sovItems.forEach(sov=>{const p=progressUpdates[sov.id];if(p&&p.asOfDate&&d.date<=p.asOfDate)actualCum+=sov.scheduledValue*p.percentComplete;});
    return{...d,actualCum:d.date<=todayStr?Math.round(actualCum):null};
  }),[combinedZoom.sliced,progressUpdates,sovItems]);

  const green=scopeHealth.filter(s=>s.status==="green").length;
  const yellow=scopeHealth.filter(s=>s.status==="yellow").length;
  const red=scopeHealth.filter(s=>s.status==="red").length;

  return(
    <div style={{display:"grid",gap:16}}>
      {/* KPI row */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:10}}>
        <Metric title="Contract Value" value={currency(contractTotal)} color="#3b82f6" large/>
        <Metric title="Earned Value" value={currency(totalEarned)} color="#22c55e" sub={`${Math.round(contractTotal>0?totalEarned/contractTotal*100:0)}% complete`} large/>
        <Metric title="Planned to Date" value={currency(plannedToDate)} color="#64748b" sub="baseline" large/>
        <Metric title="Schedule Variance" value={currency(variance)} color={variance>=0?"#22c55e":"#ef4444"} sub={variance>=0?"Ahead of plan":"Behind plan"} large/>
        <Metric title="CPI" value={cpi?cpi.toFixed(2):"—"} color={cpi>=1?"#22c55e":cpi>=0.9?"#eab308":"#ef4444"} sub={cpi>=1?"Efficient":"Overspend"} large/>
        <Metric title="SPI" value={spi?spi.toFixed(2):"—"} color={spi>=1?"#22c55e":spi>=0.9?"#eab308":"#ef4444"} sub={spi>=1?"On schedule":"Behind"} large/>
      </div>

      {/* Status summary */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
        {[{label:"On Track",count:green,color:"#22c55e",bg:"#052e16"},{label:"Watch",count:yellow,color:"#eab308",bg:"#1a1200"},{label:"Behind",count:red,color:"#ef4444",bg:"#1a0505"}].map(s=>(
          <div key={s.label} style={{background:s.bg,borderRadius:12,padding:"14px 16px",border:`1px solid ${s.color}30`,textAlign:"center"}}>
            <div style={{fontSize:32,fontWeight:800,fontFamily:"'Syne',sans-serif",color:s.color,lineHeight:1}}>{s.count}</div>
            <div style={{fontSize:12,color:s.color,marginTop:4}}>{s.label}</div>
          </div>
        ))}
      </div>

      {/* S-Curve with actual overlay */}
      <div style={{background:"#0f172a",borderRadius:14,border:"1px solid #1e293b",padding:20}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
          <div>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:16,fontWeight:700,margin:"0 0 2px",color:"#f1f5f9"}}>Baseline vs. Actual Earned Value</h2>
            <p style={{color:"#475569",fontSize:12,margin:0}}>Planned cumulative spend (dashed) vs. actual earned value (solid)</p>
          </div>
          <ZoomBar isZoomed={combinedZoom.isZoomed} onReset={combinedZoom.reset}/>
        </div>
        <div style={{height:320,cursor:"crosshair"}}>
          {allDates.length?(
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={sCurveData} margin={{top:8,right:8,left:8,bottom:8}}
                onMouseDown={combinedZoom.onDown} onMouseMove={combinedZoom.onMove} onMouseUp={combinedZoom.onUp}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1e293b"/>
                <XAxis dataKey="date" tick={{fontSize:9,fill:"#475569",fontFamily:"monospace"}} minTickGap={55} tickFormatter={fmtShort}/>
                <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:"#475569",fontFamily:"monospace"}} width={52}/>
                <Tooltip content={<CTip/>}/>
                <Legend wrapperStyle={{fontSize:11,color:"#64748b"}}/>
                <ReferenceLine x={todayStr} stroke="#334155" strokeDasharray="3 3" label={{value:"Today",fill:"#475569",fontSize:10}}/>
                <Line type="monotone" dataKey="cumTotal" stroke="#3b82f6" strokeWidth={2} dot={false} strokeDasharray="6 3" name="Baseline (Planned)"/>
                <Line type="monotone" dataKey="actualCum" stroke="#22c55e" strokeWidth={2.5} dot={false} name="Actual Earned"/>
                {combinedZoom.sel?.active&&combinedZoom.sel.start&&combinedZoom.sel.end&&<ReferenceArea x1={combinedZoom.sel.start} x2={combinedZoom.sel.end} fill="#3b82f6" fillOpacity={0.08}/>}
              </LineChart>
            </ResponsiveContainer>
          ):<Empty msg="Upload files and link scopes to see the dashboard."/>}
        </div>
      </div>

      {/* Scope health table */}
      <div style={{background:"#0f172a",borderRadius:14,border:"1px solid #1e293b",padding:18,overflowX:"auto"}}>
        <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:16,fontWeight:700,margin:"0 0 14px",color:"#f1f5f9"}}>Scope Health</h2>
        <table style={{width:"100%",minWidth:600,borderCollapse:"collapse",fontSize:12}}>
          <thead><tr style={{background:"#020817"}}>{["Scope","Contract Value","Baseline % Today","Actual %","Variance","Status"].map(h=><th key={h} style={{padding:"7px 10px",textAlign:"left",color:"#475569",fontFamily:"monospace",fontSize:9,borderBottom:"1px solid #1e293b",textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
          <tbody>
            {scopeHealth.map(({sov,pct,baselineToday,diff,status,col},i)=>(
              <tr key={sov.id} style={{borderBottom:"1px solid #1e293b",background:i%2===0?"transparent":"#020817"}}>
                <td style={{padding:"8px 10px",color:"#cbd5e1",maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                  <div style={{display:"flex",alignItems:"center",gap:5}}><div style={{width:3,height:14,borderRadius:2,background:col,flexShrink:0}}/>{sov.description}</div>
                </td>
                <td style={{padding:"8px 10px",color:"#94a3b8",fontFamily:"monospace"}}>{currency(sov.scheduledValue)}</td>
                <td style={{padding:"8px 10px",fontFamily:"monospace",color:"#64748b"}}>{Math.round(baselineToday*100)}%</td>
                <td style={{padding:"8px 10px",minWidth:120}}>
                  <div style={{display:"flex",alignItems:"center",gap:6}}>
                    <div style={{flex:1,height:4,background:"#1e293b",borderRadius:2}}><div style={{width:`${Math.min(100,Math.round(pct*100))}%`,height:"100%",background:col,borderRadius:2}}/></div>
                    <span style={{fontFamily:"monospace",fontSize:10,color:col,minWidth:28}}>{Math.round(pct*100)}%</span>
                  </div>
                </td>
                <td style={{padding:"8px 10px",fontFamily:"monospace",color:diff>=0?"#22c55e":"#ef4444",fontSize:11}}>{diff>=0?"+":""}{Math.round(diff*100)}pp</td>
                <td style={{padding:"8px 10px"}}><TrafficLight status={status}/></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App(){
  const [rawTasks,setRawTasks]=useState([]);
  const [sovResult,setSovResult]=useState(null);
  const [parseErrors,setParseErrors]=useState([]);
  const [schedFile,setSchedFile]=useState("");
  const [sovFile,setSovFile]=useState("");
  const [linksMap,setLinksMap]=useState({});
  const [activeTab,setActiveTab]=useState("dashboard");
  const [selectedScope,setSelectedScope]=useState("all");
  const [scopeZooms,setScopeZooms]=useState({});
  const [scopeSels,setScopeSels]=useState({});
  const [baselineLocked,setBaselineLocked]=useState(false);
  const [progressUpdates,setProgressUpdates]=useState({}); // sovId -> {percentComplete, asOfDate}

  const workTasks=useMemo(()=>rawTasks.filter(t=>!t.isSummary),[rawTasks]);
  const sovItems=sovResult?.items||[];

  useEffect(()=>{if(!workTasks.length||!sovItems.length)return;setLinksMap(autoLink(workTasks,sovItems));},[workTasks.length,sovItems.length]);

  function handleScheduleFile(file){
    setSchedFile(file.name);
    const reader=new FileReader();
    reader.onload=e=>{try{const{rows,errors}=parseMSP(e.target.result);setRawTasks(rows);setParseErrors(errors);}catch(err){setParseErrors([{row:0,issue:`Parse error: ${err.message}`}]);}};
    reader.readAsText(file);
  }
  function handleSOVFile(file){
    setSovFile(file.name);
    const reader=new FileReader();
    reader.onload=e=>{try{setSovResult(parseSOV(e.target.result));}catch(err){setParseErrors(p=>[...p,{row:0,issue:`SOV error: ${err.message}`}]);}};
    reader.readAsArrayBuffer(file);
  }

  const sovColorMap=useMemo(()=>{const m={};sovItems.forEach((s,i)=>{m[s.id]=PAL[i%PAL.length];});return m;},[sovItems]);
  const{byScope,allDates,scopeKeys}=useMemo(()=>buildCurves(workTasks,sovItems,linksMap),[workTasks,sovItems,linksMap]);
  const sovKeys=scopeKeys||Object.keys(byScope);
  const combinedZoom=useZoom(allDates);

  function getScopeSlice(sid){const data=byScope[sid]?.weekly||[];const z=scopeZooms[sid];return z?data.slice(z[0],z[1]+1):data;}

  const contractTotal=sovResult?.contractTotal||0;
  const totals=useMemo(()=>({
    contract:contractTotal,base:sovResult?.baseSum||0,gcgr:sovResult?.gcgrTotal||0,ohp:sovResult?.ohpTotal||0,
    scopes:sovItems.length,plotted:sovKeys.length,linked:new Set(Object.values(linksMap).flat()).size,tasks:workTasks.length,
  }),[sovResult,sovItems,sovKeys,linksMap,workTasks,contractTotal]);

  // Baseline planned value at today for progress panel
  const baselineAtToday=useMemo(()=>{
    const todayStr=today();
    const pt=allDates.find(d=>d.date>=todayStr)||allDates[allDates.length-1];
    return pt?.cumTotal||0;
  },[allDates]);

  function exportReport(){
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet([{Metric:"Contract Total",Value:contractTotal},{Metric:"Base",Value:totals.base},{Metric:"Gen Req",Value:totals.gcgr},{Metric:"OH&P",Value:totals.ohp}]),"Summary");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(allDates),"Combined Curve");
    sovItems.forEach(s=>{if(byScope[s.id])XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(byScope[s.id].full),s.description.slice(0,28).replace(/[^a-zA-Z0-9 ]/g,""));});
    XLSX.writeFile(wb,"Scope_Curves.xlsx");
  }

  const hasData=sovKeys.length>0;
  const TABS=[
    {id:"dashboard",label:"Dashboard",icon:<Activity size={12}/>},
    {id:"curves",label:"Bell Curves",icon:<TrendingUp size={12}/>},
    {id:"progress",label:"Progress Update",icon:<RefreshCw size={12}/>},
    {id:"links",label:"Link Tasks",icon:<Link2 size={12}/>,badge:totals.tasks-totals.linked>0?totals.tasks-totals.linked:null,badgeColor:"#f59e0b"},
    {id:"sov",label:"SOV Detail",icon:<Layers size={12}/>},
    {id:"validation",label:"Validation",icon:<AlertTriangle size={12}/>,badge:parseErrors.length||null,badgeColor:"#ef4444"},
  ];

  return(
    <div style={{minHeight:"100vh",background:"#020817",color:"#e2e8f0",fontFamily:"'Inter',system-ui,sans-serif",padding:"24px"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Syne:wght@700;800&display=swap');
        *{box-sizing:border-box;}
        ::-webkit-scrollbar{width:4px;height:4px;} ::-webkit-scrollbar-track{background:#0f172a;} ::-webkit-scrollbar-thumb{background:#1e293b;border-radius:3px;}
        .tab-btn{transition:all .12s;} .tab-btn:hover{color:#f1f5f9!important;}
        input::placeholder{color:#334155!important;} select{appearance:auto;}
        input[type=range]{height:4px;}
      `}</style>
      <div style={{maxWidth:1400,margin:"0 auto"}}>

        {/* Header */}
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:18,flexWrap:"wrap",gap:12}}>
          <div>
            <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:6}}>
              <BarChart3 size={12} color="#3b82f6"/>
              <span style={{fontFamily:"monospace",fontSize:9,color:"#3b82f6",letterSpacing:2,textTransform:"uppercase"}}>Scope Spend · Baseline · Progress</span>
            </div>
            <h1 style={{fontFamily:"'Syne',sans-serif",fontSize:26,fontWeight:800,margin:0,lineHeight:1.1,color:"#f8fafc"}}>
              Project Spend<br/><span style={{color:"#3b82f6"}}>Baseline Tracker</span>
            </h1>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center",flexShrink:0}}>
            {hasData&&!baselineLocked&&(
              <button onClick={()=>setBaselineLocked(true)}
                style={{display:"flex",alignItems:"center",gap:6,padding:"8px 14px",background:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:8,color:"#93c5fd",cursor:"pointer",fontSize:12,fontWeight:600}}>
                <Lock size={12}/> Lock Baseline
              </button>
            )}
            {baselineLocked&&(
              <span style={{display:"flex",alignItems:"center",gap:5,padding:"7px 12px",background:"#052e16",border:"1px solid #14532d",borderRadius:8,color:"#22c55e",fontSize:12,fontWeight:600}}>
                <Lock size={11}/> Baseline Locked
              </span>
            )}
            <button onClick={exportReport} disabled={!hasData}
              style={{display:"flex",alignItems:"center",gap:6,padding:"8px 14px",background:"transparent",border:"1px solid #1e293b",borderRadius:8,color:hasData?"#94a3b8":"#334155",cursor:hasData?"pointer":"not-allowed",fontSize:12}}>
              <Download size={12}/> Export
            </button>
          </div>
        </div>

        {/* Uploads */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
          <DropZone label="Upload MS Project XML" accept=".xml" onFile={handleScheduleFile} fileName={schedFile?`${schedFile} · ${workTasks.length} tasks`:""} icon={<FileText size={15} color="#3b82f6"/>} hint="File → Save As → XML Format in MS Project"/>
          <DropZone label="Upload SOV Excel" accept=".xlsx,.xls,.csv" onFile={handleSOVFile} fileName={sovFile?`${sovFile} · ${sovItems.length} scopes`:""} icon={<Layers size={15} color="#22c55e"/>} hint="GC/GR and OH&P auto-detected and distributed"/>
        </div>

        {/* SOV info bar */}
        {sovResult&&(
          <div style={{display:"flex",gap:8,alignItems:"center",padding:"8px 12px",background:"#0f172a",border:"1px solid #1e293b",borderRadius:8,marginBottom:12,fontSize:12,flexWrap:"wrap"}}>
            <Info size={11} color="#3b82f6"/>
            <span style={{color:"#64748b"}}>SOV:</span>
            <span style={{color:"#f1f5f9",fontFamily:"monospace"}}>{currency(totals.base)}</span><span style={{color:"#334155"}}>base +</span>
            <span style={{color:"#f97316",fontFamily:"monospace"}}>{currency(totals.gcgr)}</span><span style={{color:"#334155"}}>Gen Req +</span>
            <span style={{color:"#a855f7",fontFamily:"monospace"}}>{currency(totals.ohp)}</span><span style={{color:"#334155"}}>OH&P =</span>
            <span style={{color:"#22c55e",fontWeight:700,fontFamily:"monospace"}}>{currency(totals.contract)}</span>
            {sovResult.parseWarning&&<span style={{color:"#f59e0b",marginLeft:6}}>⚠ {sovResult.parseWarning}</span>}
          </div>
        )}

        {/* Tabs */}
        <div style={{display:"flex",gap:2,marginBottom:14,background:"#0f172a",padding:3,borderRadius:10,border:"1px solid #1e293b",width:"fit-content",flexWrap:"wrap"}}>
          {TABS.map(t=>(
            <button key={t.id} className="tab-btn" onClick={()=>setActiveTab(t.id)}
              style={{display:"flex",alignItems:"center",gap:5,padding:"6px 12px",borderRadius:7,border:"none",cursor:"pointer",fontSize:12,fontWeight:500,background:activeTab===t.id?"#1e293b":"transparent",color:activeTab===t.id?"#f1f5f9":"#64748b"}}>
              {t.icon}{t.label}
              {t.badge?<span style={{background:t.badgeColor,color:"#fff",borderRadius:99,padding:"1px 5px",fontSize:9,fontWeight:700}}>{t.badge}</span>:null}
            </button>
          ))}
        </div>

        {/* ── Dashboard ── */}
        {activeTab==="dashboard"&&(
          hasData
            ?<Dashboard sovItems={sovItems} allDates={allDates} byScope={byScope} sovKeys={sovKeys} progressUpdates={progressUpdates} baselineLocked={baselineLocked} contractTotal={contractTotal} combinedZoom={combinedZoom} PAL={PAL}/>
            :<Empty msg="Upload both files and link tasks to scopes to see the dashboard."/>
        )}

        {/* ── Bell Curves ── */}
        {activeTab==="curves"&&(
          <div style={{display:"grid",gap:14}}>
            {!hasData&&<Empty msg="Upload both files and link tasks to generate bell curves."/>}
            {hasData&&(
              <div style={{display:"flex",gap:5,flexWrap:"wrap",alignItems:"center"}}>
                <button onClick={()=>setSelectedScope("all")} style={{padding:"4px 10px",borderRadius:99,border:`1px solid ${selectedScope==="all"?"#3b82f6":"#1e293b"}`,background:selectedScope==="all"?"#1e3a5f":"transparent",color:selectedScope==="all"?"#93c5fd":"#64748b",cursor:"pointer",fontSize:12}}>All scopes</button>
                {sovKeys.map((sid,i)=>{const s=byScope[sid];const col=PAL[i%PAL.length];return(<button key={sid} onClick={()=>setSelectedScope(sid)} style={{padding:"4px 10px",borderRadius:99,border:`1px solid ${selectedScope===sid?col:"#1e293b"}`,background:selectedScope===sid?`${col}20`:"transparent",color:selectedScope===sid?col:"#64748b",cursor:"pointer",fontSize:12,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.sov.description.slice(0,22)}</button>);})}
              </div>
            )}

            {/* All scopes stacked + total overlay */}
            {hasData&&selectedScope==="all"&&(
              <div style={{background:"#0f172a",borderRadius:14,border:"1px solid #1e293b",padding:20}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
                  <div>
                    <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 2px",color:"#f1f5f9"}}>All Scope Bell Curves — Stacked</h2>
                    <p style={{color:"#475569",fontSize:12,margin:0}}>Stacked scope spend with total project envelope (white line) · drag to zoom</p>
                  </div>
                  <ZoomBar isZoomed={combinedZoom.isZoomed} onReset={combinedZoom.reset}/>
                </div>
                <div style={{height:440,cursor:"crosshair"}}>
                  <ResponsiveContainer width="100%" height="100%">
                    <AreaChart data={combinedZoom.sliced} margin={{top:8,right:8,left:8,bottom:8}}
                      onMouseDown={combinedZoom.onDown} onMouseMove={combinedZoom.onMove} onMouseUp={combinedZoom.onUp}>
                      <defs>
                        {sovKeys.map((sid,i)=>(
                          <linearGradient key={sid} id={`cg${i}`} x1="0" y1="0" x2="0" y2="1">
                            <stop offset="5%" stopColor={PAL[i%PAL.length]} stopOpacity={0.55}/>
                            <stop offset="95%" stopColor={PAL[i%PAL.length]} stopOpacity={0.04}/>
                          </linearGradient>
                        ))}
                      </defs>
                      <CartesianGrid strokeDasharray="3 3" stroke="#1e293b"/>
                      <XAxis dataKey="date" tick={{fontSize:9,fill:"#475569",fontFamily:"monospace"}} minTickGap={55} tickFormatter={fmtShort}/>
                      <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:"#475569",fontFamily:"monospace"}} width={52}/>
                      <Tooltip content={<CTip/>}/>
                      <Legend wrapperStyle={{fontSize:10,color:"#64748b"}}/>
                      {/* Stacked scope areas */}
                      {sovKeys.map((sid,i)=>{
                        const s=byScope[sid];const col=PAL[i%PAL.length];const key=s.sov.description.slice(0,18);
                        return(<Area key={sid} type="monotone" dataKey={key} stroke={col} strokeWidth={1.5} fill={`url(#cg${i})`} stackId="s" name={key}/>);
                      })}
                      {/* Total project envelope line on top */}
                      <Line type="monotone" dataKey="total" stroke="#ffffff" strokeWidth={2.5} dot={false} name="Project Total" strokeOpacity={0.9}/>
                      {combinedZoom.sel?.active&&combinedZoom.sel.start&&combinedZoom.sel.end&&(<ReferenceArea x1={combinedZoom.sel.start} x2={combinedZoom.sel.end} fill="#3b82f6" fillOpacity={0.08}/>)}
                    </AreaChart>
                  </ResponsiveContainer>
                </div>
              </div>
            )}

            {/* Individual scope */}
            {hasData&&selectedScope!=="all"&&byScope[selectedScope]&&(()=>{
              const{sov,tasks:lt,startDate,finishDate}=byScope[selectedScope];
              const col=sovColorMap[selectedScope];const data=getScopeSlice(selectedScope);
              const zoomed=!!scopeZooms[selectedScope];const sel=scopeSels[selectedScope]||{};
              return(
                <div style={{display:"grid",gap:12}}>
                  <div style={{background:"#0f172a",borderRadius:13,border:`1px solid ${col}40`,padding:16,display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:10}}>
                    <div><div style={{fontSize:9,fontFamily:"monospace",color:"#475569",textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Scope</div><div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:13,color:"#f1f5f9"}}>{sov.description}</div></div>
                    <div><div style={{fontSize:9,fontFamily:"monospace",color:"#475569",textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Contract Value</div><div style={{fontWeight:700,fontSize:15,color:col}}>{currency(sov.scheduledValue)}</div></div>
                    <div><div style={{fontSize:9,fontFamily:"monospace",color:"#475569",textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Base Value</div><div style={{fontWeight:600,color:"#cbd5e1",fontSize:12}}>{currency(sov.baseValue)}</div></div>
                    <div><div style={{fontSize:9,fontFamily:"monospace",color:"#475569",textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Duration</div><div style={{color:"#cbd5e1",fontSize:11}}>{fmtDate(startDate)} → {fmtDate(finishDate)}</div></div>
                    <div><div style={{fontSize:9,fontFamily:"monospace",color:"#475569",textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Tasks</div><div style={{fontWeight:600,color:"#cbd5e1"}}>{lt.length}</div></div>
                  </div>
                  <div style={{background:"#0f172a",borderRadius:13,border:"1px solid #1e293b",padding:20}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
                      <div><h2 style={{fontFamily:"'Syne',sans-serif",fontSize:15,fontWeight:700,margin:"0 0 2px",color:"#f1f5f9"}}>{sov.description} — Bell Curve</h2><p style={{color:"#475569",fontSize:12,margin:0}}>Drag to zoom</p></div>
                      <ZoomBar isZoomed={zoomed} onReset={()=>setScopeZooms(p=>({...p,[selectedScope]:null}))}/>
                    </div>
                    <div style={{height:280,cursor:"crosshair"}}>
                      <ResponsiveContainer width="100%" height="100%">
                        <AreaChart data={data} margin={{top:8,right:8,left:8,bottom:8}}
                          onMouseDown={e=>{if(e?.activeLabel)setScopeSels(p=>({...p,[selectedScope]:{start:e.activeLabel,end:null,active:true}}));}}
                          onMouseMove={e=>{if(sel?.active&&e?.activeLabel)setScopeSels(p=>({...p,[selectedScope]:{...p[selectedScope],end:e.activeLabel}}));}}
                          onMouseUp={()=>{if(sel?.active&&sel.start&&sel.end&&sel.start!==sel.end){const d=byScope[selectedScope]?.weekly||[];const a=d.findIndex(x=>x.date===sel.start),b=d.findIndex(x=>x.date===sel.end);if(a>=0&&b>=0)setScopeZooms(p=>({...p,[selectedScope]:[Math.min(a,b),Math.max(a,b)]}));}setScopeSels(p=>({...p,[selectedScope]:{active:false}}));}}>
                          <defs><linearGradient id="sg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={col} stopOpacity={0.5}/><stop offset="95%" stopColor={col} stopOpacity={0.02}/></linearGradient></defs>
                          <CartesianGrid strokeDasharray="3 3" stroke="#1e293b"/>
                          <XAxis dataKey="date" tick={{fontSize:9,fill:"#475569",fontFamily:"monospace"}} minTickGap={40} tickFormatter={fmtShort}/>
                          <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:"#475569",fontFamily:"monospace"}} width={52}/>
                          <Tooltip content={<CTip/>}/>
                          <Area type="monotone" dataKey="daily" stroke={col} strokeWidth={2.5} fill="url(#sg)" name="Weekly Spend"/>
                          {sel?.active&&sel.start&&sel.end&&<ReferenceArea x1={sel.start} x2={sel.end} fill={col} fillOpacity={0.08}/>}
                        </AreaChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                  <div style={{background:"#0f172a",borderRadius:13,border:"1px solid #1e293b",padding:18}}>
                    <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:14,fontWeight:700,margin:"0 0 10px",color:"#f1f5f9"}}>Linked Tasks ({lt.length})</h2>
                    <table style={{width:"100%",minWidth:480,borderCollapse:"collapse",fontSize:12}}>
                      <thead><tr style={{background:"#020817"}}>{["WBS","Task","Start","Finish","%"].map(h=><th key={h} style={{padding:"6px 9px",textAlign:"left",color:"#475569",fontFamily:"monospace",fontSize:9,borderBottom:"1px solid #1e293b",textTransform:"uppercase",letterSpacing:.5}}>{h}</th>)}</tr></thead>
                      <tbody>{lt.map((t,i)=>(<tr key={i} style={{borderBottom:"1px solid #0f172a"}}><td style={{padding:"6px 9px",fontFamily:"monospace",color:"#3b82f6",fontSize:10}}>{t.id}</td><td style={{padding:"6px 9px",color:"#cbd5e1"}}>{t.name}</td><td style={{padding:"6px 9px",fontFamily:"monospace",color:"#64748b",fontSize:10}}>{fmtDate(t.startDate)}</td><td style={{padding:"6px 9px",fontFamily:"monospace",color:"#64748b",fontSize:10}}>{fmtDate(t.finishDate)}</td><td style={{padding:"6px 9px",minWidth:80}}><div style={{display:"flex",alignItems:"center",gap:5}}><div style={{flex:1,height:3,background:"#1e293b",borderRadius:2}}><div style={{width:`${Math.min(100,Math.round(t.percentComplete*100))}%`,height:"100%",background:col,borderRadius:2}}/></div><span style={{fontFamily:"monospace",fontSize:9,color:"#64748b",minWidth:22}}>{Math.round(t.percentComplete*100)}%</span></div></td></tr>))}</tbody>
                    </table>
                  </div>
                </div>
              );
            })()}
          </div>
        )}

        {/* ── Progress Update ── */}
        {activeTab==="progress"&&(
          hasData
            ?<ProgressPanel sovItems={sovItems} progressUpdates={progressUpdates} onUpdate={setProgressUpdates} baseline={baselineAtToday}/>
            :<Empty msg="Upload both files and link scopes to enter progress updates."/>
        )}

        {/* ── Link Tasks ── */}
        {activeTab==="links"&&(
          workTasks.length&&sovItems.length
            ?<LinkManager tasks={workTasks} sovItems={sovItems} linksMap={linksMap} onLinksChange={setLinksMap}/>
            :<div style={{background:"#0f172a",borderRadius:14,border:"1px solid #1e293b",padding:40,textAlign:"center"}}><Link2 size={24} color="#1e293b" style={{marginBottom:8}}/><div style={{fontSize:13,color:"#334155"}}>Upload both files to use the link manager.</div></div>
        )}

        {/* ── SOV Detail ── */}
        {activeTab==="sov"&&(
          <div style={{background:"#0f172a",borderRadius:14,border:"1px solid #1e293b",padding:18,overflowX:"auto"}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 4px",color:"#f1f5f9"}}>SOV Line Items</h2>
            <p style={{color:"#475569",fontSize:12,marginBottom:14}}>Base + distributed Gen Req + OH&P = contract value per scope</p>
            {sovItems.length?(
              <table style={{width:"100%",minWidth:660,borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{background:"#020817"}}>{["Ph","Description","Base","+ Gen Req","+ OH&P","Contract Value","Tasks"].map(h=><th key={h} style={{padding:"7px 9px",textAlign:"left",color:"#475569",fontFamily:"monospace",fontSize:9,borderBottom:"1px solid #1e293b",textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                <tbody>
                  {sovItems.map((s,i)=>{const col=PAL[i%PAL.length];const count=(linksMap[s.id]||[]).length;return(<tr key={s.id} style={{borderBottom:"1px solid #1e293b",background:i%2===0?"transparent":"#020817"}}><td style={{padding:"7px 9px",fontFamily:"monospace",color:"#475569",fontSize:10}}>{s.phase||"—"}</td><td style={{padding:"7px 9px",color:"#cbd5e1",maxWidth:170,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.description}</td><td style={{padding:"7px 9px",color:"#94a3b8",fontFamily:"monospace"}}>{currency(s.baseValue)}</td><td style={{padding:"7px 9px",color:"#f97316",fontFamily:"monospace",fontSize:10}}>{currency(s.gcgrShare)}</td><td style={{padding:"7px 9px",color:"#a855f7",fontFamily:"monospace",fontSize:10}}>{currency(s.ohpShare)}</td><td style={{padding:"7px 9px",fontWeight:700,color:col,fontFamily:"monospace"}}>{currency(s.scheduledValue)}</td><td style={{padding:"7px 9px"}}><span style={{fontSize:10,color:count>0?col:"#334155",background:count>0?`${col}20`:"#0f172a",padding:"1px 7px",borderRadius:99,border:`1px solid ${count>0?`${col}40`:"#1e293b"}`}}>{count}</span></td></tr>);})}
                </tbody>
                <tfoot><tr style={{borderTop:"2px solid #1e293b"}}><td colSpan={2} style={{padding:"8px 9px",fontWeight:700,color:"#f1f5f9"}}>TOTAL</td><td style={{padding:"8px 9px",fontWeight:700,color:"#94a3b8",fontFamily:"monospace"}}>{currency(totals.base)}</td><td style={{padding:"8px 9px",fontWeight:700,color:"#f97316",fontFamily:"monospace"}}>{currency(totals.gcgr)}</td><td style={{padding:"8px 9px",fontWeight:700,color:"#a855f7",fontFamily:"monospace"}}>{currency(totals.ohp)}</td><td style={{padding:"8px 9px",fontWeight:700,color:"#22c55e",fontFamily:"monospace",fontSize:14}}>{currency(totals.contract)}</td><td/></tr></tfoot>
              </table>
            ):<Empty msg="Upload your SOV to see the breakdown."/>}
          </div>
        )}

        {/* ── Validation ── */}
        {activeTab==="validation"&&(
          <div style={{background:"#0f172a",borderRadius:14,border:"1px solid #1e293b",padding:18}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 4px",color:"#f1f5f9"}}>Validation Log</h2>
            <p style={{color:"#475569",fontSize:12,marginBottom:14}}>Parse issues from uploaded files.</p>
            {parseErrors.length?(<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(270px,1fr))",gap:7}}>{parseErrors.map((e,i)=>(<div key={i} style={{background:"#1a0f00",border:"1px solid #431407",borderRadius:8,padding:"9px 12px"}}><div style={{fontFamily:"monospace",fontSize:9,color:"#f97316",marginBottom:2}}>ROW {e.row||"—"}</div><div style={{color:"#cbd5e1",fontSize:12}}>{e.issue}</div></div>))}</div>)
            :(<div style={{display:"flex",alignItems:"center",gap:9,padding:"11px 14px",background:"#052e16",border:"1px solid #14532d",borderRadius:8}}><CheckCircle2 size={14} color="#22c55e"/><span style={{color:"#86efac",fontSize:12}}>{rawTasks.length?"No issues found.":"Upload a schedule to validate."}</span></div>)}
          </div>
        )}

      </div>
    </div>
  );
}
