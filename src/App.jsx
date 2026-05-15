import { useMemo, useState, useEffect, useCallback } from "react";
import { AreaChart, Area, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid, ReferenceArea, LineChart, Line, ReferenceLine, Legend, ComposedChart, BarChart, Bar } from "recharts";
import * as XLSX from "xlsx";

// ─── Inline icons (no lucide-react needed) ────────────────────────────────────
const Icon = {
  Upload:()=><span style={{fontSize:14}}>↑</span>,
  Download:()=><span style={{fontSize:14}}>↓</span>,
  Lock:()=><span style={{fontSize:14}}>🔒</span>,
  Unlock:()=><span style={{fontSize:14}}>🔓</span>,
  Check:({color,size=14})=><span style={{color,fontSize:size}}>✓</span>,
  Alert:({color})=><span style={{color,fontSize:14}}>⚠</span>,
  Chart:()=><span style={{fontSize:14}}>📊</span>,
  Trend:()=><span style={{fontSize:14}}>📈</span>,
  File:({color})=><span style={{color,fontSize:14}}>📄</span>,
  Layers:({color})=><span style={{color,fontSize:14}}>📋</span>,
  Link:()=><span style={{fontSize:14}}>🔗</span>,
  ZoomOut:()=><span style={{fontSize:12}}>🔍</span>,
  Info:({color})=><span style={{color,fontSize:13}}>ℹ</span>,
  Activity:()=><span style={{fontSize:14}}>📉</span>,
  Refresh:()=><span style={{fontSize:14}}>↻</span>,
  Circle:({fill,size=8})=><span style={{color:fill,fontSize:size}}>●</span>,
  Save:()=><span style={{fontSize:14}}>💾</span>,
  Forecast:()=><span style={{fontSize:14}}>🔮</span>,
  Calendar:()=><span style={{fontSize:14}}>📅</span>,
};

// ─── Brand colors ─────────────────────────────────────────────────────────────
const WB = {
  bg:"#0d1b2a", card:"#152535", border:"#1e3448", borderLight:"#2a4560",
  primary:"#4a7c8e", primaryLight:"#8aabb8", primaryDark:"#3a6c7e",
  text:"#f8fafc", textMuted:"#8aabb8", textDim:"#475569",
  green:"#22c55e", greenBg:"#052e16", greenBorder:"#14532d",
  red:"#ef4444", redBg:"#1a0505", redBorder:"#7f1d1d",
  yellow:"#eab308", yellowBg:"#1a1200", yellowBorder:"#854d0e",
  orange:"#f97316",
};
const PAL=["#4a7c8e","#8aabb8","#c8a96e","#5b8a6d","#7b9eb8","#c87d5b","#6b8a9e","#a8c4b8","#9e7b6e","#5b7a8e","#b8956e","#6e8e7b","#8e6e5b","#7ab8c4","#c4a87a"];

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
function addMonths(dateStr,n){const d=new Date(dateStr+"-01");d.setMonth(d.getMonth()+n);return monthKey(d);}

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
    if(!isSummary){
      if(!startDate)errors.push({row:i+1,issue:`"${name}" — missing Start`});
      if(!finishDate)errors.push({row:i+1,issue:`"${name}" — missing Finish`});
    }
    rows.push({rowNum:i+1,id:wbs,uid,name,startDate,finishDate,
      budgetedCost:a.cost>0?a.cost:toMoney(getText(task,"Cost")),
      actualCost:a.actual>0?a.actual:toMoney(getText(task,"ActualCost")),
      percentComplete:toPercent(getText(task,"PercentComplete")||"0"),isSummary});
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

// ─── Build scope curves ───────────────────────────────────────────────────────
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
  scopeIds.forEach(sid=>{const{full}=byScope[sid];const s=Math.max(1,Math.floor(full.length/120));byScope[sid].weekly=full.filter((_,i)=>i%s===0||(i===full.length-1));});
  return{byScope,allDates,scopeKeys:scopeIds};
}

// ─── Build monthly data ───────────────────────────────────────────────────────
function buildMonthlyData(byScope,sovKeys){
  if(!sovKeys.length)return[];
  let globalMin=null,globalMax=null;
  sovKeys.forEach(sid=>{
    const{startDate,finishDate}=byScope[sid];
    if(!globalMin||startDate<globalMin)globalMin=startDate;
    if(!globalMax||finishDate>globalMax)globalMax=finishDate;
  });
  if(!globalMin||!globalMax)return[];
  const monthMap=new Map();
  const start=new Date(globalMin.getFullYear(),globalMin.getMonth(),1);
  const end=new Date(globalMax.getFullYear(),globalMax.getMonth(),1);
  for(let d=new Date(start);d<=end;d.setMonth(d.getMonth()+1)){
    const k=monthKey(d);
    monthMap.set(k,{month:k,total:0});
  }
  sovKeys.forEach(sid=>{
    const{sov,full}=byScope[sid];
    const key=sov.description.slice(0,18);
    monthMap.forEach(row=>{row[key]=row[key]||0;});
    full.forEach(p=>{
      const mk=p.date.slice(0,7);
      if(monthMap.has(mk)){const row=monthMap.get(mk);row[key]=(row[key]||0)+p.daily;row.total+=p.daily;}
    });
  });
  const sorted=Array.from(monthMap.values()).sort((a,b)=>a.month.localeCompare(b.month));
  sorted.forEach(row=>{Object.keys(row).forEach(k=>{if(k!=="month")row[k]=Math.round(row[k]);});});
  return sorted;
}

// ─── Monthly progress: compute earned value per scope per month ───────────────
// monthlyProgress: { sovId: { "2025-05": 45, "2025-06": 70, ... } } (cumulative % per month)
function buildActualMonthly(sovItems,monthlyProgress,byScope,allMonths){
  // For each month, find the latest cumulative % entered up to that month
  return allMonths.map(mk=>{
    let actualTotal=0;
    sovItems.forEach(sov=>{
      const prog=monthlyProgress[sov.id]||{};
      // Find the latest month entry <= mk
      const entries=Object.entries(prog).filter(([m])=>m<=mk).sort((a,b)=>b[0].localeCompare(a[0]));
      const latestPct=entries.length>0?entries[0][1]/100:0;
      actualTotal+=sov.scheduledValue*latestPct;
    });
    return{month:mk,actualCum:Math.round(actualTotal)};
  });
}

// ─── Variance detection ───────────────────────────────────────────────────────
function computeVariances(sovItems,monthlyProgress,byScope,threshold=0.10){
  const alerts=[];
  sovItems.forEach(sov=>{
    if(!byScope[sov.id])return;
    const prog=monthlyProgress[sov.id]||{};
    Object.entries(prog).forEach(([mk,pctVal])=>{
      const actualEarned=sov.scheduledValue*(pctVal/100);
      // Find baseline planned % at this month
      const scopeFull=byScope[sov.id].full;
      const lastDayOfMonth=new Date(mk+"-01");lastDayOfMonth.setMonth(lastDayOfMonth.getMonth()+1);lastDayOfMonth.setDate(0);
      const lastDay=fmtDate(lastDayOfMonth);
      const baselinePoint=scopeFull.filter(p=>p.date<=lastDay);
      const baselinePct=baselinePoint.length>0?baselinePoint[baselinePoint.length-1].cumulative/sov.scheduledValue:0;
      const baselineEarned=sov.scheduledValue*baselinePct;
      const variance=(actualEarned-baselineEarned)/sov.scheduledValue;
      if(Math.abs(variance)>threshold){
        alerts.push({sovId:sov.id,description:sov.description,month:mk,variance,actualPct:pctVal,baselinePct:Math.round(baselinePct*100),actualEarned,baselineEarned,status:variance<0?"behind":"ahead"});
      }
    });
  });
  return alerts;
}

// ─── Forecast to complete ─────────────────────────────────────────────────────
function computeForecasts(sovItems,monthlyProgress,byScope){
  return sovItems.map(sov=>{
    const prog=monthlyProgress[sov.id]||{};
    const entries=Object.entries(prog).sort((a,b)=>a[0].localeCompare(b[0]));
    if(entries.length<2)return{sov,hasData:false};
    const latestEntry=entries[entries.length-1];
    const prevEntry=entries[entries.length-2];
    const latestPct=latestEntry[1]/100;
    const prevPct=prevEntry[1]/100;
    const monthsDiff=1; // assuming monthly entries
    const burnRatePerMonth=(latestPct-prevPct)/monthsDiff;
    if(burnRatePerMonth<=0)return{sov,hasData:false,latestPct,note:"No progress this period"};
    const remainingPct=1-latestPct;
    const monthsToComplete=Math.ceil(remainingPct/burnRatePerMonth);
    const forecastCompleteMonth=addMonths(latestEntry[0],monthsToComplete);
    // EAC: if current burn rate continues
    const plannedEnd=byScope[sov.id]?monthKey(byScope[sov.id].finishDate):null;
    const isLate=plannedEnd&&forecastCompleteMonth>plannedEnd;
    const isEarly=plannedEnd&&forecastCompleteMonth<plannedEnd;
    return{sov,hasData:true,latestPct,burnRatePerMonth,monthsToComplete,forecastCompleteMonth,plannedEnd,isLate,isEarly,scheduledValue:sov.scheduledValue};
  });
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
      const a=data.findIndex(d=>d.date===sel.start),b=data.findIndex(d=>d.date===sel.end);
      if(a>=0&&b>=0)setRange([Math.min(a,b),Math.max(a,b)]);
    }
    setSel(null);
  }
  return{sliced,onDown,onMove,onUp,reset:()=>setRange(null),isZoomed:!!range,sel};
}

// ─── Tooltip ─────────────────────────────────────────────────────────────────
function CTip({active,payload,label,monthlyTotals}){
  if(!active||!payload?.length)return null;
  const items=payload.filter(p=>p.value>0&&p.name!=="▸ Project Total").sort((a,b)=>b.value-a.value);
  const totalVal=monthlyTotals?monthlyTotals.find(d=>d.month===label)?.total:payload.find(p=>p.name==="▸ Project Total")?.value;
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
      {totalVal!=null&&<div style={{borderTop:`1px solid ${WB.borderLight}`,marginTop:7,paddingTop:7,display:"flex",justifyContent:"space-between",gap:8}}><span style={{color:WB.text,fontWeight:700}}>Monthly Total:</span><span style={{color:WB.text,fontWeight:800,fontSize:14}}>{currency(totalVal)}</span></div>}
    </div>
  );
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
      <div style={{fontSize:13,color:WB.textDim,textAlign:"center",padding:20,maxWidth:300}}>{msg}</div>
    </div>
  );
}

function StatusBadge({status}){
  const col=status==="green"?WB.green:status==="yellow"?WB.yellow:WB.red;
  const bg=status==="green"?WB.greenBg:status==="yellow"?WB.yellowBg:WB.redBg;
  const label=status==="green"?"On Track":status==="yellow"?"Watch":"Behind";
  return(<span style={{display:"inline-flex",alignItems:"center",gap:4,fontSize:11,color:col,background:bg,padding:"2px 8px",borderRadius:99,border:`1px solid ${col}40`}}><Icon.Circle fill={col} size={7}/>{label}</span>);
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
    <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:18}}>
      <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 3px",color:WB.text}}>Link Tasks → SOV Scopes</h2>
      <p style={{color:WB.textDim,fontSize:12,margin:"0 0 12px"}}>Assign each schedule task to a SOV scope. Fuzzy matches are pre-filled.</p>
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
        <table style={{width:"100%",minWidth:640,borderCollapse:"collapse",fontSize:12}}>
          <thead><tr style={{background:WB.bg}}>{["WBS","Task Name","Dates","SOV Scope"].map(h=><th key={h} style={{padding:"7px 10px",textAlign:"left",color:WB.textDim,fontFamily:"monospace",fontSize:9,borderBottom:`1px solid ${WB.border}`,textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
          <tbody>
            {filtered.map((t,i)=>{
              const sid=taskToSov[t.uid];const col=sid?sovColorMap[sid]:null;
              return(
                <tr key={t.uid} style={{borderBottom:`1px solid ${WB.border}`,background:i%2===0?"transparent":WB.bg}}>
                  <td style={{padding:"7px 10px",fontFamily:"monospace",color:WB.primary,fontSize:10,whiteSpace:"nowrap"}}>{t.id}</td>
                  <td style={{padding:"7px 10px",color:WB.text,maxWidth:240}}>
                    <div style={{display:"flex",alignItems:"center",gap:6}}>{col&&<div style={{width:3,height:14,borderRadius:2,background:col,flexShrink:0}}/>}<span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{t.name}</span></div>
                  </td>
                  <td style={{padding:"7px 10px",fontFamily:"monospace",color:WB.textDim,fontSize:10,whiteSpace:"nowrap"}}>{fmtDate(t.startDate)} → {fmtDate(t.finishDate)}</td>
                  <td style={{padding:"7px 10px",minWidth:240}}>
                    <select value={sid||"none"} onChange={e=>setTaskLink(t.uid,e.target.value)}
                      style={{width:"100%",padding:"5px 9px",background:col?`${col}18`:WB.bg,border:`1px solid ${col||WB.border}`,borderRadius:6,color:col||WB.textDim,fontSize:12,cursor:"pointer",outline:"none",fontWeight:col?600:400}}>
                      <option value="none" style={{background:WB.bg,color:WB.textDim}}>— Unassigned —</option>
                      {sovItems.map(s=><option key={s.id} value={s.id} style={{background:WB.bg,color:WB.text}}>{s.description} ({currency(s.scheduledValue)})</option>)}
                    </select>
                  </td>
                </tr>
              );
            })}
            {!filtered.length&&<tr><td colSpan={4} style={{padding:24,textAlign:"center",color:WB.textDim}}>No tasks match filter.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── Monthly Progress Panel ───────────────────────────────────────────────────
function MonthlyProgressPanel({sovItems,monthlyProgress,onUpdate,byScope,sovKeys}){
  const [selectedMonth,setSelectedMonth]=useState(today().slice(0,7));
  // Get all available months from scope curves
  const availableMonths=useMemo(()=>{
    if(!sovKeys.length)return[];
    let min=null,max=null;
    sovKeys.forEach(sid=>{
      const{startDate,finishDate}=byScope[sid];
      if(!min||startDate<min)min=startDate;
      if(!max||finishDate>max)max=finishDate;
    });
    if(!min||!max)return[];
    const months=[];
    for(let d=new Date(min.getFullYear(),min.getMonth(),1);d<=max;d.setMonth(d.getMonth()+1)){
      months.push(monthKey(new Date(d)));
    }
    return months;
  },[byScope,sovKeys]);

  function setPct(sovId,pct){
    const val=Math.min(100,Math.max(0,Number(pct)||0));
    onUpdate(prev=>({...prev,[sovId]:{...(prev[sovId]||{}),[selectedMonth]:val}}));
  }

  function getPct(sovId){return monthlyProgress[sovId]?.[selectedMonth]??"";}

  // Compute baseline % for selected month
  function getBaselinePct(sovId){
    if(!byScope[sovId])return 0;
    const full=byScope[sovId].full;
    const lastDayOfMonth=new Date(selectedMonth+"-01");lastDayOfMonth.setMonth(lastDayOfMonth.getMonth()+1);lastDayOfMonth.setDate(0);
    const lastDay=fmtDate(lastDayOfMonth);
    const pts=full.filter(p=>p.date<=lastDay);
    const sov=sovItems.find(s=>s.id===sovId);
    if(!pts.length||!sov)return 0;
    return Math.round(pts[pts.length-1].cumulative/sov.scheduledValue*100);
  }

  const totalEarned=sovItems.reduce((s,sov)=>{
    const prog=monthlyProgress[sov.id]||{};
    const entries=Object.entries(prog).filter(([m])=>m<=selectedMonth).sort((a,b)=>b[0].localeCompare(a[0]));
    const pct=entries.length>0?entries[0][1]/100:0;
    return s+sov.scheduledValue*pct;
  },0);
  const contractTotal=sovItems.reduce((s,sov)=>s+sov.scheduledValue,0);

  return(
    <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:18}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16,flexWrap:"wrap",gap:10}}>
        <div>
          <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 3px",color:WB.text}}>Monthly Progress Update</h2>
          <p style={{color:WB.textDim,fontSize:12,margin:0}}>Enter cumulative % complete per scope for each month.</p>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:12,color:WB.textDim}}>Month:</span>
          <select value={selectedMonth} onChange={e=>setSelectedMonth(e.target.value)}
            style={{padding:"6px 10px",background:WB.bg,border:`1px solid ${WB.border}`,borderRadius:7,color:WB.textMuted,fontSize:12,outline:"none",cursor:"pointer"}}>
            {availableMonths.map(m=><option key={m} value={m}>{fmtMonth(m)}</option>)}
          </select>
        </div>
      </div>

      {/* Summary KPIs */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:9,marginBottom:16}}>
        <Metric title="Earned Value" value={currency(totalEarned)} color={WB.green} sub={`${Math.round(contractTotal>0?totalEarned/contractTotal*100:0)}% complete`}/>
        <Metric title="Remaining" value={currency(contractTotal-totalEarned)} color={WB.orange}/>
        <Metric title="Month" value={fmtMonth(selectedMonth)} color={WB.primary}/>
      </div>

      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
        <thead>
          <tr style={{background:WB.bg}}>
            {["Scope","Contract Value","Baseline %","Actual % (cumulative)","Variance","Status"].map(h=>(
              <th key={h} style={{padding:"7px 10px",textAlign:"left",color:WB.textDim,fontFamily:"monospace",fontSize:9,borderBottom:`1px solid ${WB.border}`,textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {sovItems.map((sov,i)=>{
            const col=PAL[i%PAL.length];
            const pctVal=getPct(sov.id);
            const baselinePct=getBaselinePct(sov.id);
            const actualPct=pctVal!==""?Number(pctVal):null;
            const variance=actualPct!==null?actualPct-baselinePct:null;
            const status=variance===null?"none":variance>=-5?"green":variance>=-15?"yellow":"red";
            return(
              <tr key={sov.id} style={{borderBottom:`1px solid ${WB.border}`}}>
                <td style={{padding:"8px 10px",color:WB.text,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                  <div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:3,height:14,borderRadius:2,background:col,flexShrink:0}}/>{sov.description}</div>
                </td>
                <td style={{padding:"8px 10px",color:WB.textMuted,fontFamily:"monospace"}}>{currency(sov.scheduledValue)}</td>
                <td style={{padding:"8px 10px",fontFamily:"monospace",color:WB.textDim}}>{baselinePct}%</td>
                <td style={{padding:"8px 10px",minWidth:200}}>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <input type="range" min={0} max={100} value={pctVal!==""?Number(pctVal):0} onChange={e=>setPct(sov.id,e.target.value)}
                      style={{flex:1,accentColor:col,cursor:"pointer"}}/>
                    <input type="number" min={0} max={100} value={pctVal} onChange={e=>setPct(sov.id,e.target.value)} placeholder="—"
                      style={{width:50,padding:"3px 5px",background:WB.bg,border:`1px solid ${col}`,borderRadius:5,color:col,fontSize:12,fontFamily:"monospace",outline:"none",textAlign:"center"}}/>
                    <span style={{color:WB.textDim,fontSize:11}}>%</span>
                  </div>
                </td>
                <td style={{padding:"8px 10px",fontFamily:"monospace",fontSize:11,color:variance===null?WB.textDim:variance>=0?WB.green:WB.red}}>
                  {variance!==null?`${variance>=0?"+":""}${variance.toFixed(1)}pp`:"—"}
                </td>
                <td style={{padding:"8px 10px"}}>{status!=="none"?<StatusBadge status={status}/>:<span style={{fontSize:11,color:WB.textDim}}>—</span>}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

// ─── Forecast Panel ───────────────────────────────────────────────────────────
function ForecastPanel({forecasts,variances}){
  return(
    <div style={{display:"grid",gap:16}}>
      {/* Variance alerts */}
      {variances.length>0&&(
        <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.yellowBorder}`,padding:18}}>
          <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 4px",color:WB.text}}>
            <Icon.Alert color={WB.yellow}/> Variance Alerts
          </h2>
          <p style={{color:WB.textDim,fontSize:12,marginBottom:14}}>Scopes where actual % complete deviates more than 10% from baseline plan.</p>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:8}}>
            {variances.map((v,i)=>{
              const col=v.status==="behind"?WB.red:WB.green;
              const bg=v.status==="behind"?WB.redBg:WB.greenBg;
              const border=v.status==="behind"?WB.redBorder:WB.greenBorder;
              return(
                <div key={i} style={{background:bg,border:`1px solid ${border}`,borderRadius:10,padding:"12px 14px"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                    <div style={{fontWeight:600,color:WB.text,fontSize:13}}>{v.description}</div>
                    <span style={{fontSize:11,color:col,fontFamily:"monospace",fontWeight:700}}>{v.status==="behind"?"BEHIND":"AHEAD"}</span>
                  </div>
                  <div style={{fontSize:11,color:WB.textDim,fontFamily:"monospace",marginBottom:2}}>{fmtMonth(v.month)}</div>
                  <div style={{display:"flex",gap:16,fontSize:12}}>
                    <div><span style={{color:WB.textDim}}>Baseline: </span><span style={{color:WB.textMuted,fontWeight:600}}>{v.baselinePct}%</span></div>
                    <div><span style={{color:WB.textDim}}>Actual: </span><span style={{color:col,fontWeight:600}}>{v.actualPct}%</span></div>
                    <div><span style={{color:WB.textDim}}>Δ </span><span style={{color:col,fontWeight:700}}>{v.variance>=0?"+":""}{(v.variance*100).toFixed(1)}pp</span></div>
                  </div>
                  <div style={{fontSize:11,color:WB.textDim,marginTop:4}}>
                    {v.status==="behind"?`${currency(v.baselineEarned-v.actualEarned)} behind planned earned value`:`${currency(v.actualEarned-v.baselineEarned)} ahead of planned earned value`}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Forecast to complete */}
      <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:18,overflowX:"auto"}}>
        <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 4px",color:WB.text}}>
          <Icon.Forecast/> Forecast to Complete
        </h2>
        <p style={{color:WB.textDim,fontSize:12,marginBottom:14}}>Projected completion month per scope based on current burn rate.</p>
        <table style={{width:"100%",minWidth:700,borderCollapse:"collapse",fontSize:12}}>
          <thead>
            <tr style={{background:WB.bg}}>
              {["Scope","Contract Value","Current %","Burn Rate/Mo","Planned Finish","Forecast Finish","Status"].map(h=>(
                <th key={h} style={{padding:"7px 10px",textAlign:"left",color:WB.textDim,fontFamily:"monospace",fontSize:9,borderBottom:`1px solid ${WB.border}`,textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {forecasts.map((f,i)=>{
              const col=PAL[i%PAL.length];
              if(!f.hasData){
                return(
                  <tr key={f.sov.id} style={{borderBottom:`1px solid ${WB.border}`,background:i%2===0?"transparent":WB.bg}}>
                    <td style={{padding:"8px 10px",color:WB.text}}>
                      <div style={{display:"flex",alignItems:"center",gap:5}}><div style={{width:3,height:14,borderRadius:2,background:col,flexShrink:0}}/>{f.sov.description}</div>
                    </td>
                    <td style={{padding:"8px 10px",color:WB.textMuted,fontFamily:"monospace"}}>{currency(f.sov.scheduledValue)}</td>
                    <td colSpan={5} style={{padding:"8px 10px",color:WB.textDim,fontSize:11}}>{f.note||"Need at least 2 months of data to forecast"}</td>
                  </tr>
                );
              }
              const statusColor=f.isLate?WB.red:f.isEarly?WB.green:WB.green;
              const statusLabel=f.isLate?"Late":f.isEarly?"Early":"On Track";
              return(
                <tr key={f.sov.id} style={{borderBottom:`1px solid ${WB.border}`,background:i%2===0?"transparent":WB.bg}}>
                  <td style={{padding:"8px 10px",color:WB.text,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                    <div style={{display:"flex",alignItems:"center",gap:5}}><div style={{width:3,height:14,borderRadius:2,background:col,flexShrink:0}}/>{f.sov.description}</div>
                  </td>
                  <td style={{padding:"8px 10px",color:WB.textMuted,fontFamily:"monospace"}}>{currency(f.sov.scheduledValue)}</td>
                  <td style={{padding:"8px 10px",color:col,fontFamily:"monospace",fontWeight:600}}>{Math.round(f.latestPct*100)}%</td>
                  <td style={{padding:"8px 10px",color:WB.textDim,fontFamily:"monospace"}}>{(f.burnRatePerMonth*100).toFixed(1)}%/mo</td>
                  <td style={{padding:"8px 10px",color:WB.textDim,fontFamily:"monospace"}}>{f.plannedEnd?fmtMonth(f.plannedEnd):"—"}</td>
                  <td style={{padding:"8px 10px",color:statusColor,fontFamily:"monospace",fontWeight:600}}>{fmtMonth(f.forecastCompleteMonth)}</td>
                  <td style={{padding:"8px 10px"}}>
                    <span style={{fontSize:11,color:statusColor,background:`${statusColor}18`,padding:"2px 8px",borderRadius:99,border:`1px solid ${statusColor}40`}}>{statusLabel}</span>
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

// ─── Dashboard ────────────────────────────────────────────────────────────────
function Dashboard({sovItems,allDates,byScope,sovKeys,monthlyProgress,contractTotal,combinedZoom,variances}){
  const totalEarned=useMemo(()=>sovItems.reduce((s,sov)=>{
    const prog=monthlyProgress[sov.id]||{};
    const todayMk=today().slice(0,7);
    const entries=Object.entries(prog).filter(([m])=>m<=todayMk).sort((a,b)=>b[0].localeCompare(a[0]));
    const pct=entries.length>0?entries[0][1]/100:0;
    return s+sov.scheduledValue*pct;
  },0),[sovItems,monthlyProgress]);

  const todayMk=today().slice(0,7);
  const todayPoint=allDates.find(d=>d.date>=today())||allDates[allDates.length-1];
  const plannedToDate=todayPoint?.cumTotal||0;
  const cpi=contractTotal>0?totalEarned/contractTotal:0;
  const spi=plannedToDate>0?totalEarned/plannedToDate:0;
  const variance=totalEarned-plannedToDate;

  // Build actual monthly cumulative for S-curve overlay
  const allMonths=useMemo(()=>{
    if(!sovKeys.length)return[];
    let min=null,max=null;
    sovKeys.forEach(sid=>{const{startDate,finishDate}=byScope[sid];if(!min||startDate<min)min=startDate;if(!max||finishDate>max)max=finishDate;});
    if(!min||!max)return[];
    const months=[];for(let d=new Date(min.getFullYear(),min.getMonth(),1);d<=max;d.setMonth(d.getMonth()+1))months.push(monthKey(new Date(d)));
    return months;
  },[byScope,sovKeys]);

  const actualMonthly=useMemo(()=>buildActualMonthly(sovItems,monthlyProgress,byScope,allMonths),[sovItems,monthlyProgress,byScope,allMonths]);

  // Merge with allDates for S-curve chart
  const sCurveData=useMemo(()=>{
    return combinedZoom.sliced.map(d=>{
      const mk=d.date.slice(0,7);
      const actual=actualMonthly.find(a=>a.month===mk);
      return{...d,actualCum:actual&&d.date<=today()?actual.actualCum:null};
    });
  },[combinedZoom.sliced,actualMonthly]);

  const scopeHealth=sovItems.map((sov,i)=>{
    const prog=monthlyProgress[sov.id]||{};
    const entries=Object.entries(prog).filter(([m])=>m<=todayMk).sort((a,b)=>b[0].localeCompare(a[0]));
    const pct=entries.length>0?entries[0][1]/100:0;
    const scope=byScope[sov.id];
    let baselineToday=0;
    if(scope){const pt=scope.full.filter(p=>p.date<=today());baselineToday=pt.length>0?pt[pt.length-1].cumulative/sov.scheduledValue:0;}
    const diff=pct-baselineToday;
    const status=diff>=-0.05?"green":diff>=-0.15?"yellow":"red";
    return{sov,pct,baselineToday,diff,status,col:PAL[i%PAL.length]};
  });

  const green=scopeHealth.filter(s=>s.status==="green").length;
  const yellow=scopeHealth.filter(s=>s.status==="yellow").length;
  const red=scopeHealth.filter(s=>s.status==="red").length;

  return(
    <div style={{display:"grid",gap:16}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:10}}>
        <Metric title="Contract Value" value={currency(contractTotal)} color={WB.primary}/>
        <Metric title="Earned Value" value={currency(totalEarned)} color={WB.green} sub={`${Math.round(contractTotal>0?totalEarned/contractTotal*100:0)}% complete`}/>
        <Metric title="Planned to Date" value={currency(plannedToDate)} color={WB.textDim} sub="baseline"/>
        <Metric title="Schedule Variance" value={currency(variance)} color={variance>=0?WB.green:WB.red} sub={variance>=0?"Ahead":"Behind"}/>
        <Metric title="CPI" value={cpi?cpi.toFixed(2):"—"} color={cpi>=1?WB.green:cpi>=0.9?WB.yellow:WB.red} sub={cpi>=1?"Efficient":"Overspend"}/>
        <Metric title="SPI" value={spi?spi.toFixed(2):"—"} color={spi>=1?WB.green:spi>=0.9?WB.yellow:WB.red} sub={spi>=1?"On schedule":"Behind"}/>
      </div>

      {/* Variance alert summary */}
      {variances.length>0&&(
        <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",background:WB.yellowBg,border:`1px solid ${WB.yellowBorder}`,borderRadius:9,fontSize:13}}>
          <Icon.Alert color={WB.yellow}/>
          <span style={{color:WB.yellow,fontWeight:600}}>{variances.length} variance alert{variances.length!==1?"s":""}</span>
          <span style={{color:WB.textDim}}>— go to the Forecast tab to review</span>
        </div>
      )}

      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
        {[{label:"On Track",count:green,color:WB.green,bg:WB.greenBg},{label:"Watch",count:yellow,color:WB.yellow,bg:WB.yellowBg},{label:"Behind",count:red,color:WB.red,bg:WB.redBg}].map(s=>(
          <div key={s.label} style={{background:s.bg,borderRadius:12,padding:"14px 16px",border:`1px solid ${s.color}30`,textAlign:"center"}}>
            <div style={{fontSize:32,fontWeight:800,fontFamily:"'Syne',sans-serif",color:s.color,lineHeight:1}}>{s.count}</div>
            <div style={{fontSize:12,color:s.color,marginTop:4}}>{s.label}</div>
          </div>
        ))}
      </div>

      <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:20}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
          <div>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:16,fontWeight:700,margin:"0 0 2px",color:WB.text}}>Baseline vs. Actual Earned Value</h2>
            <p style={{color:WB.textDim,fontSize:12,margin:0}}>Planned cumulative (dashed) vs. actual earned value</p>
          </div>
          <ZoomBar isZoomed={combinedZoom.isZoomed} onReset={combinedZoom.reset}/>
        </div>
        <div style={{height:320,cursor:"crosshair"}}>
          {allDates.length?(
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={sCurveData} margin={{top:8,right:8,left:8,bottom:8}}
                onMouseDown={combinedZoom.onDown} onMouseMove={combinedZoom.onMove} onMouseUp={combinedZoom.onUp}>
                <CartesianGrid strokeDasharray="3 3" stroke={WB.border}/>
                <XAxis dataKey="date" tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} minTickGap={55} tickFormatter={fmtShort}/>
                <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} width={52}/>
                <Tooltip content={<CTip/>}/>
                <Legend wrapperStyle={{fontSize:11,color:WB.textDim}}/>
                <ReferenceLine x={today()} stroke={WB.borderLight} strokeDasharray="3 3" label={{value:"Today",fill:WB.textDim,fontSize:10}}/>
                <Line type="monotone" dataKey="cumTotal" stroke={WB.primary} strokeWidth={2} dot={false} strokeDasharray="6 3" name="Baseline (Planned)"/>
                <Line type="monotone" dataKey="actualCum" stroke={WB.green} strokeWidth={2.5} dot={false} name="Actual Earned"/>
                {combinedZoom.sel?.active&&combinedZoom.sel.start&&combinedZoom.sel.end&&<ReferenceArea x1={combinedZoom.sel.start} x2={combinedZoom.sel.end} fill={WB.primary} fillOpacity={0.08}/>}
              </LineChart>
            </ResponsiveContainer>
          ):<Empty msg="Upload files to see the S-curve."/>}
        </div>
      </div>

      <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:18,overflowX:"auto"}}>
        <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:16,fontWeight:700,margin:"0 0 14px",color:WB.text}}>Scope Health</h2>
        <table style={{width:"100%",minWidth:600,borderCollapse:"collapse",fontSize:12}}>
          <thead><tr style={{background:WB.bg}}>{["Scope","Contract","Baseline %","Actual %","Variance","Status"].map(h=><th key={h} style={{padding:"7px 10px",textAlign:"left",color:WB.textDim,fontFamily:"monospace",fontSize:9,borderBottom:`1px solid ${WB.border}`,textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
          <tbody>
            {scopeHealth.map(({sov,pct,baselineToday,diff,status,col},i)=>(
              <tr key={sov.id} style={{borderBottom:`1px solid ${WB.border}`,background:i%2===0?"transparent":WB.bg}}>
                <td style={{padding:"8px 10px",color:WB.text,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                  <div style={{display:"flex",alignItems:"center",gap:5}}><div style={{width:3,height:14,borderRadius:2,background:col,flexShrink:0}}/>{sov.description}</div>
                </td>
                <td style={{padding:"8px 10px",color:WB.textMuted,fontFamily:"monospace"}}>{currency(sov.scheduledValue)}</td>
                <td style={{padding:"8px 10px",fontFamily:"monospace",color:WB.textDim}}>{Math.round(baselineToday*100)}%</td>
                <td style={{padding:"8px 10px",minWidth:120}}>
                  <div style={{display:"flex",alignItems:"center",gap:6}}>
                    <div style={{flex:1,height:4,background:WB.border,borderRadius:2}}><div style={{width:`${Math.min(100,Math.round(pct*100))}%`,height:"100%",background:col,borderRadius:2}}/></div>
                    <span style={{fontFamily:"monospace",fontSize:10,color:col,minWidth:28}}>{Math.round(pct*100)}%</span>
                  </div>
                </td>
                <td style={{padding:"8px 10px",fontFamily:"monospace",color:diff>=0?WB.green:WB.red,fontSize:11}}>{diff>=0?"+":""}{Math.round(diff*100)}pp</td>
                <td style={{padding:"8px 10px"}}><StatusBadge status={status}/></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}


// ─── Subcontractor Module ─────────────────────────────────────────────────────
function SubcontractorModule({subcontractors,onUpdate,sovItems,byScope,sovKeys,monthlyData,sovColorMap,selectedSub,onSelectSub}){
  const [activeView,setActiveView]=useState("list"); // list | detail | add
  const [newSub,setNewSub]=useState({name:"",trade:"",contractAmount:""});
  const [subFile,setSubFile]=useState("");
  const [importError,setImportError]=useState("");

  // Parse subcontractor import Excel
  function handleSubFile(file){
    setSubFile(file.name);setImportError("");
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const raw=XLSX.utils.sheet_to_json(ws,{defval:""});
        const imported=raw.map((row,i)=>{
          const name=String(row["Subcontractor"]||row["Name"]||row["Sub"]||"").trim();
          const trade=String(row["Trade"]||row["Division"]||"").trim();
          const amt=toMoney(row["Contract Amount"]||row["Amount"]||row["Value"]||0);
          if(!name)return null;
          return{id:`sub-import-${Date.now()}-${i}`,name,trade,contractAmount:amt,sovIds:[],baselineLocked:false,monthlyProgress:{}};
        }).filter(Boolean);
        if(!imported.length){setImportError("No subcontractors found. Expected columns: Subcontractor, Trade, Contract Amount");return;}
        onUpdate(prev=>[...prev,...imported.filter(s=>!prev.find(p=>p.name===s.name))]);
        setActiveView("list");
      }catch(err){setImportError(`Import error: ${err.message}`);}
    };
    reader.readAsArrayBuffer(file);
  }

  function addSub(){
    if(!newSub.name.trim())return;
    const sub={id:`sub-${Date.now()}`,name:newSub.name.trim(),trade:newSub.trade.trim(),contractAmount:toMoney(newSub.contractAmount),sovIds:[],baselineLocked:false,monthlyProgress:{}};
    onUpdate(prev=>[...prev,sub]);
    setNewSub({name:"",trade:"",contractAmount:""});
    setActiveView("list");
  }

  function deleteSub(id){
    if(!confirm("Delete this subcontractor?"))return;
    onUpdate(prev=>prev.filter(s=>s.id!==id));
    if(selectedSub?.id===id)onSelectSub(null);
  }

  function toggleSovForSub(subId,sovId){
    onUpdate(prev=>prev.map(s=>{
      if(s.id!==subId)return s;
      const has=s.sovIds.includes(sovId);
      return{...s,sovIds:has?s.sovIds.filter(id=>id!==sovId):[...s.sovIds,sovId]};
    }));
  }

  function lockSubBaseline(subId){
    onUpdate(prev=>prev.map(s=>s.id===subId?{...s,baselineLocked:true}:s));
  }

  function setSubProgress(subId,month,pct){
    onUpdate(prev=>prev.map(s=>{
      if(s.id!==subId)return s;
      return{...s,monthlyProgress:{...s.monthlyProgress,[month]:Math.min(100,Math.max(0,Number(pct)||0))}};
    }));
  }

  // Build sub bell curve data from its SOV scopes
  function buildSubCurve(sub){
    if(!sub.sovIds.length)return{monthly:[],contractTotal:0};
    const subSovItems=sovItems.filter(s=>sub.sovIds.includes(s.id));
    const contractTotal=subSovItems.reduce((s,i)=>s+i.scheduledValue,0);
    // Aggregate monthly data for this sub's scopes only
    const monthly=monthlyData.map(row=>{
      let total=0;
      subSovItems.forEach(sov=>{const key=sov.description.slice(0,18);total+=(row[key]||0);});
      return{month:row.month,total:Math.round(total)};
    }).filter(r=>r.total>0||monthly?.length<3);
    return{monthly,contractTotal};
  }

  // Sub health: compare progress vs baseline
  function getSubHealth(sub){
    if(!sub.monthlyProgress||!Object.keys(sub.monthlyProgress).length)return"none";
    const todayMk=today().slice(0,7);
    const entries=Object.entries(sub.monthlyProgress).filter(([m])=>m<=todayMk).sort((a,b)=>b[0].localeCompare(a[0]));
    if(!entries.length)return"none";
    const latestPct=entries[0][1]/100;
    // Rough baseline: how far through project duration are we
    const subSovItems=sovItems.filter(s=>sub.sovIds.includes(s.id));
    if(!subSovItems.length)return"none";
    const allScopes=subSovItems.map(s=>byScope[s.id]).filter(Boolean);
    if(!allScopes.length)return"none";
    const minStart=new Date(Math.min(...allScopes.map(s=>s.startDate.getTime())));
    const maxEnd=new Date(Math.max(...allScopes.map(s=>s.finishDate.getTime())));
    const totalDuration=maxEnd-minStart;
    const elapsed=new Date()-minStart;
    const baselinePct=Math.min(1,Math.max(0,elapsed/totalDuration));
    const diff=latestPct-baselinePct;
    return diff>=-0.05?"green":diff>=-0.15?"yellow":"red";
  }

  const SUB_COLORS=["#4a7c8e","#c8a96e","#5b8a6d","#c87d5b","#7b9eb8","#8e6e5b","#a8c4b8","#6b8a9e"];

  // ── List view ──
  if(activeView==="list"&&!selectedSub){
    return(
      <div style={{display:"grid",gap:14}}>
        {/* Header actions */}
        <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
          <button onClick={()=>setActiveView("add")}
            style={{padding:"8px 16px",background:WB.primary,border:"none",borderRadius:8,color:WB.text,fontSize:13,fontWeight:600,cursor:"pointer"}}>
            + Add Subcontractor
          </button>
          <label style={{display:"flex",alignItems:"center",gap:7,padding:"8px 16px",background:WB.card,border:`1px solid ${WB.border}`,borderRadius:8,cursor:"pointer",fontSize:13,color:WB.textMuted}}>
            ↑ Import Excel
            <input type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>e.target.files?.[0]&&handleSubFile(e.target.files[0])}/>
          </label>
          {subFile&&<span style={{fontSize:12,color:WB.green}}>✓ {subFile}</span>}
          {importError&&<span style={{fontSize:12,color:WB.red}}>⚠ {importError}</span>}
        </div>

        {/* Import template hint */}
        <div style={{padding:"10px 14px",background:WB.card,borderRadius:9,border:`1px solid ${WB.border}`,fontSize:12,color:WB.textDim}}>
          📋 Excel import format: columns <strong style={{color:WB.textMuted}}>Subcontractor</strong>, <strong style={{color:WB.textMuted}}>Trade</strong>, <strong style={{color:WB.textMuted}}>Contract Amount</strong>
        </div>

        {/* Sub cards */}
        {subcontractors.length===0?(
          <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:40,textAlign:"center"}}>
            <div style={{fontSize:32,marginBottom:8}}>👷</div>
            <div style={{fontSize:14,color:WB.textDim}}>No subcontractors yet. Add one above or import from Excel.</div>
          </div>
        ):(
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:12}}>
            {subcontractors.map((sub,i)=>{
              const col=SUB_COLORS[i%SUB_COLORS.length];
              const health=getSubHealth(sub);
              const {contractTotal}=buildSubCurve(sub);
              const healthColor=health==="green"?WB.green:health==="yellow"?WB.yellow:health==="red"?WB.red:WB.textDim;
              return(
                <div key={sub.id} style={{background:WB.card,borderRadius:13,border:`1px solid ${col}40`,padding:16,position:"relative"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
                    <div>
                      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:3}}>
                        <div style={{width:10,height:10,borderRadius:"50%",background:col}}/>
                        <span style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:15,color:WB.text}}>{sub.name}</span>
                        {sub.baselineLocked&&<span style={{fontSize:10,color:WB.green,background:WB.greenBg,padding:"1px 6px",borderRadius:99,border:`1px solid ${WB.greenBorder}`}}>🔒 Locked</span>}
                      </div>
                      {sub.trade&&<div style={{fontSize:12,color:WB.textDim}}>{sub.trade}</div>}
                    </div>
                    <button onClick={()=>deleteSub(sub.id)} style={{background:"none",border:"none",cursor:"pointer",color:WB.textDim,fontSize:16,padding:0,lineHeight:1}}>×</button>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
                    <div style={{background:WB.bg,borderRadius:8,padding:"8px 10px"}}>
                      <div style={{fontSize:9,fontFamily:"monospace",color:WB.textDim,textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Contract</div>
                      <div style={{fontWeight:700,color:col,fontSize:14}}>{currency(sub.contractAmount||contractTotal)}</div>
                    </div>
                    <div style={{background:WB.bg,borderRadius:8,padding:"8px 10px"}}>
                      <div style={{fontSize:9,fontFamily:"monospace",color:WB.textDim,textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Scopes</div>
                      <div style={{fontWeight:700,color:WB.textMuted,fontSize:14}}>{sub.sovIds.length}</div>
                    </div>
                  </div>
                  <div style={{display:"flex",gap:8}}>
                    <button onClick={()=>{onSelectSub(sub);setActiveView("detail");}}
                      style={{flex:1,padding:"7px",background:`${col}20`,border:`1px solid ${col}50`,borderRadius:7,color:col,cursor:"pointer",fontSize:12,fontWeight:600}}>
                      Open →
                    </button>
                    {!sub.baselineLocked&&<button onClick={()=>lockSubBaseline(sub.id)}
                      style={{padding:"7px 10px",background:"transparent",border:`1px solid ${WB.border}`,borderRadius:7,color:WB.textDim,cursor:"pointer",fontSize:12}}>
                      🔒 Lock
                    </button>}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    );
  }

  // ── Add sub form ──
  if(activeView==="add"){
    return(
      <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:24,maxWidth:500}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:20}}>
          <button onClick={()=>setActiveView("list")} style={{background:"none",border:"none",color:WB.primaryLight,cursor:"pointer",fontSize:13}}>← Back</button>
          <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:18,fontWeight:700,margin:0,color:WB.text}}>Add Subcontractor</h2>
        </div>
        {[{label:"Subcontractor Name",key:"name",placeholder:"e.g. Acme Framing LLC"},{label:"Trade / Division",key:"trade",placeholder:"e.g. Framing, Concrete, MEP"},{label:"Contract Amount",key:"contractAmount",placeholder:"e.g. 125000"}].map(f=>(
          <div key={f.key} style={{marginBottom:14}}>
            <label style={{fontSize:12,color:WB.textDim,display:"block",marginBottom:5}}>{f.label}</label>
            <input value={newSub[f.key]} onChange={e=>setNewSub(p=>({...p,[f.key]:e.target.value}))} placeholder={f.placeholder}
              onKeyDown={e=>e.key==="Enter"&&addSub()}
              style={{width:"100%",padding:"9px 12px",background:WB.bg,border:`1px solid ${WB.border}`,borderRadius:8,color:WB.text,fontSize:13,outline:"none"}}/>
          </div>
        ))}
        <button onClick={addSub} disabled={!newSub.name.trim()}
          style={{width:"100%",padding:"10px",background:WB.primary,border:"none",borderRadius:8,color:WB.text,fontSize:13,fontWeight:600,cursor:"pointer",opacity:newSub.name.trim()?1:0.5}}>
          Add Subcontractor
        </button>
      </div>
    );
  }

  // ── Detail view ──
  if(activeView==="detail"&&selectedSub){
    const sub=subcontractors.find(s=>s.id===selectedSub.id)||selectedSub;
    const subIdx=subcontractors.findIndex(s=>s.id===sub.id);
    const col=SUB_COLORS[subIdx%SUB_COLORS.length];
    const {monthly,contractTotal}=buildSubCurve(sub);
    const subSovItems=sovItems.filter(s=>sub.sovIds.includes(s.id));
    const totalEarned=subSovItems.reduce((s,sov)=>{
      const todayMk=today().slice(0,7);
      const entries=Object.entries(sub.monthlyProgress||{}).filter(([m])=>m<=todayMk).sort((a,b)=>b[0].localeCompare(a[0]));
      const pct=entries.length>0?entries[0][1]/100:0;
      return s+sov.scheduledValue*pct;
    },0);
    const availableMonths=monthly.map(m=>m.month);
    const [progMonth,setProgMonth]=useState(today().slice(0,7));

    return(
      <div style={{display:"grid",gap:14}}>
        {/* Back + header */}
        <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
          <button onClick={()=>{setActiveView("list");onSelectSub(null);}} style={{background:"none",border:"none",color:WB.primaryLight,cursor:"pointer",fontSize:13}}>← All Subs</button>
          <div style={{flex:1}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:12,height:12,borderRadius:"50%",background:col}}/>
              <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:20,fontWeight:700,margin:0,color:WB.text}}>{sub.name}</h2>
              {sub.baselineLocked&&<span style={{fontSize:11,color:WB.green,background:WB.greenBg,padding:"2px 8px",borderRadius:99,border:`1px solid ${WB.greenBorder}`}}>🔒 Baseline Locked</span>}
              {!sub.baselineLocked&&<button onClick={()=>lockSubBaseline(sub.id)} style={{padding:"4px 10px",background:"transparent",border:`1px solid ${WB.border}`,borderRadius:7,color:WB.textDim,cursor:"pointer",fontSize:11}}>🔒 Lock Baseline</button>}
            </div>
            {sub.trade&&<div style={{fontSize:12,color:WB.textDim,marginTop:2}}>{sub.trade}</div>}
          </div>
        </div>

        {/* KPIs */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:9}}>
          <Metric title="Contract Amount" value={currency(sub.contractAmount||contractTotal)} color={col}/>
          <Metric title="SOV Scopes" value={sub.sovIds.length} color={col}/>
          <Metric title="Earned to Date" value={currency(totalEarned)} color={WB.green} sub={`${Math.round(contractTotal>0?totalEarned/contractTotal*100:0)}%`}/>
          <Metric title="Remaining" value={currency((sub.contractAmount||contractTotal)-totalEarned)} color={WB.orange}/>
        </div>

        {/* Bell curve */}
        {monthly.length>0&&(
          <div style={{background:WB.card,borderRadius:13,border:`1px solid ${WB.border}`,padding:20}}>
            <h3 style={{fontFamily:"'Syne',sans-serif",fontSize:16,fontWeight:700,margin:"0 0 14px",color:WB.text}}>{sub.name} — Monthly Spend Bell Curve</h3>
            <div style={{height:280}}>
              <ResponsiveContainer width="100%" height="100%">
                <AreaChart data={monthly} margin={{top:8,right:8,left:8,bottom:8}}>
                  <defs><linearGradient id="subGrad" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={col} stopOpacity={0.5}/><stop offset="95%" stopColor={col} stopOpacity={0.02}/></linearGradient></defs>
                  <CartesianGrid strokeDasharray="3 3" stroke={WB.border}/>
                  <XAxis dataKey="month" tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} minTickGap={20}/>
                  <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} width={52}/>
                  <Tooltip content={({active,payload,label})=>{if(!active||!payload?.length)return null;return(<div style={{background:WB.card,border:`1px solid ${WB.border}`,borderRadius:9,padding:"9px 12px"}}><div style={{color:WB.textDim,fontSize:10,fontFamily:"monospace",marginBottom:4}}>{label}</div><div style={{color:col,fontWeight:700}}>{currency(payload[0]?.value)}</div></div>);}}/>
                  <Area type="monotone" dataKey="total" stroke={col} strokeWidth={2.5} fill="url(#subGrad)" name="Monthly Spend"/>
                </AreaChart>
              </ResponsiveContainer>
            </div>
          </div>
        )}

        {/* Assign SOV scopes */}
        <div style={{background:WB.card,borderRadius:13,border:`1px solid ${WB.border}`,padding:20}}>
          <h3 style={{fontFamily:"'Syne',sans-serif",fontSize:15,fontWeight:700,margin:"0 0 14px",color:WB.text}}>Assign SOV Scopes</h3>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:8}}>
            {sovItems.map((sov,i)=>{
              const assigned=sub.sovIds.includes(sov.id);
              const scopeCol=PAL[i%PAL.length];
              return(
                <div key={sov.id} onClick={()=>toggleSovForSub(sub.id,sov.id)}
                  style={{padding:"10px 12px",borderRadius:9,background:assigned?`${col}18`:WB.bg,border:`1px solid ${assigned?col:WB.border}`,cursor:"pointer",transition:"all .15s",display:"flex",justifyContent:"space-between",alignItems:"center",gap:8}}>
                  <div style={{overflow:"hidden"}}>
                    <div style={{fontSize:12,color:WB.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontWeight:assigned?600:400}}>{sov.description}</div>
                    <div style={{fontSize:11,color:assigned?col:WB.textDim,fontFamily:"monospace"}}>{currency(sov.scheduledValue)}</div>
                  </div>
                  <div style={{width:18,height:18,borderRadius:"50%",border:`2px solid ${assigned?col:WB.borderLight}`,background:assigned?col:"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontSize:11,color:"#fff"}}>
                    {assigned?"✓":""}
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Monthly progress */}
        <div style={{background:WB.card,borderRadius:13,border:`1px solid ${WB.border}`,padding:20}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap",gap:10}}>
            <h3 style={{fontFamily:"'Syne',sans-serif",fontSize:15,fontWeight:700,margin:0,color:WB.text}}>Monthly Progress</h3>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <span style={{fontSize:12,color:WB.textDim}}>Month:</span>
              <select value={progMonth} onChange={e=>setProgMonth(e.target.value)}
                style={{padding:"5px 9px",background:WB.bg,border:`1px solid ${WB.border}`,borderRadius:7,color:WB.textMuted,fontSize:12,outline:"none",cursor:"pointer"}}>
                {availableMonths.map(m=><option key={m} value={m}>{fmtMonth(m)}</option>)}
              </select>
            </div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <span style={{fontSize:13,color:WB.textMuted,minWidth:120}}>% Complete:</span>
            <input type="range" min={0} max={100} value={sub.monthlyProgress?.[progMonth]||0} onChange={e=>setSubProgress(sub.id,progMonth,e.target.value)}
              style={{flex:1,accentColor:col,cursor:"pointer"}}/>
            <input type="number" min={0} max={100} value={sub.monthlyProgress?.[progMonth]||""} onChange={e=>setSubProgress(sub.id,progMonth,e.target.value)} placeholder="0"
              style={{width:50,padding:"4px 6px",background:WB.bg,border:`1px solid ${col}`,borderRadius:6,color:col,fontSize:12,fontFamily:"monospace",outline:"none",textAlign:"center"}}/>
            <span style={{color:WB.textDim,fontSize:12}}>%</span>
          </div>
          {Object.keys(sub.monthlyProgress||{}).length>0&&(
            <div style={{marginTop:14}}>
              <div style={{fontSize:11,color:WB.textDim,fontFamily:"monospace",textTransform:"uppercase",letterSpacing:1,marginBottom:8}}>Progress History</div>
              <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                {Object.entries(sub.monthlyProgress).sort((a,b)=>a[0].localeCompare(b[0])).map(([m,pct])=>(
                  <div key={m} style={{padding:"4px 10px",background:`${col}18`,borderRadius:99,border:`1px solid ${col}40`,fontSize:11,fontFamily:"monospace",color:col}}>
                    {fmtMonth(m)}: {pct}%
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    );
  }

  return null;
}

// ─── Main App ─────────────────────────────────────────────────────────────────
const STORAGE_KEY="wb_scope_tracker_v1";

export default function App({user=null,project=null,onBackToProjects=null,onSignOut=null,supabase=null}={}){
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
  const [monthlyProgress,setMonthlyProgress]=useState({}); // {sovId: {month: pct}}
  const [mLeft,setMLeft]=useState(0);
  const [mRight,setMRight]=useState(null);
  const [mSel,setMSel]=useState(null);
  const [c1Left,setC1Left]=useState(0);
  const [c1Right,setC1Right]=useState(null);
  const [c1Sel,setC1Sel]=useState(null);
  const [saveStatus,setSaveStatus]=useState(""); // "saved" | "restored" | ""
  const [subcontractors,setSubcontractors]=useState([]); // [{id,name,trade,contractAmount,sovIds:[],baselineLocked,monthlyProgress:{}}]
  const [selectedSub,setSelectedSub]=useState(null);

  const workTasks=useMemo(()=>rawTasks.filter(t=>!t.isSummary),[rawTasks]);
  const sovItems=sovResult?.items||[];

  // ── Cloud save (Supabase) with localStorage fallback ─────────────────────
  useEffect(()=>{
    if(!rawTasks.length&&!sovItems.length)return;
    const state={rawTasks:rawTasks.map(t=>({...t,startDate:t.startDate?t.startDate.toISOString():null,finishDate:t.finishDate?t.finishDate.toISOString():null})),sovResult,schedFile,sovFile,linksMap,baselineLocked,monthlyProgress,subcontractors};
    if(supabase&&project){
      const t=setTimeout(async()=>{const{error}=await supabase.from("projects").update({data:state,updated_at:new Date().toISOString()}).eq("id",project.id);if(!error){setSaveStatus("saved");setTimeout(()=>setSaveStatus(""),2000);}},1500);
      return()=>clearTimeout(t);
    }else{
      try{localStorage.setItem(STORAGE_KEY,JSON.stringify(state));setSaveStatus("saved");setTimeout(()=>setSaveStatus(""),2000);}catch(e){}
    }
  },[rawTasks,sovResult,linksMap,baselineLocked,monthlyProgress]);

  // ── Restore from cloud or localStorage ───────────────────────────────────
  useEffect(()=>{
    async function restore(){
      let state=null;
      if(supabase&&project){const{data}=await supabase.from("projects").select("data").eq("id",project.id).single();if(data?.data&&Object.keys(data.data).length>0)state=data.data;}
      else{try{const s=localStorage.getItem(STORAGE_KEY);if(s)state=JSON.parse(s);}catch(e){}}
      if(!state)return;
      if(state.rawTasks)setRawTasks(state.rawTasks.map(t=>({...t,startDate:t.startDate?new Date(t.startDate):null,finishDate:t.finishDate?new Date(t.finishDate):null})));
      if(state.sovResult)setSovResult(state.sovResult);
      if(state.schedFile)setSchedFile(state.schedFile);
      if(state.sovFile)setSovFile(state.sovFile);
      if(state.linksMap)setLinksMap(state.linksMap);
      if(state.baselineLocked)setBaselineLocked(state.baselineLocked);
      if(state.monthlyProgress)setMonthlyProgress(state.monthlyProgress);
      if(state.subcontractors)setSubcontractors(state.subcontractors);
      setSaveStatus("restored");setTimeout(()=>setSaveStatus(""),3000);
    }
    restore();
  },[project?.id]);

  function clearSave(){
    if(supabase&&project)supabase.from("projects").update({data:{}}).eq("id",project.id);
    localStorage.removeItem(STORAGE_KEY);window.location.reload();
  }

  // Auto fuzzy link
  useEffect(()=>{
    if(!workTasks.length||!sovItems.length||Object.keys(linksMap).length>0)return;
    setLinksMap(autoLink(workTasks,sovItems));
  },[workTasks.length,sovItems.length]);

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
  const monthlyData=useMemo(()=>buildMonthlyData(byScope,sovKeys),[byScope,sovKeys]);
  const combinedZoom=useZoom(allDates);

  function getScopeSlice(sid){const data=byScope[sid]?.weekly||[];const z=scopeZooms[sid];return z?data.slice(z[0],z[1]+1):data;}

  const contractTotal=sovResult?.contractTotal||0;
  const totals=useMemo(()=>({
    contract:contractTotal,base:sovResult?.baseSum||0,gcgr:sovResult?.gcgrTotal||0,ohp:sovResult?.ohpTotal||0,
    scopes:sovItems.length,plotted:sovKeys.length,linked:new Set(Object.values(linksMap).flat()).size,tasks:workTasks.length,
  }),[sovResult,sovItems,sovKeys,linksMap,workTasks,contractTotal]);

  // Variance alerts and forecasts
  const variances=useMemo(()=>computeVariances(sovItems,monthlyProgress,byScope),[sovItems,monthlyProgress,byScope]);
  const forecasts=useMemo(()=>computeForecasts(sovItems,monthlyProgress,byScope),[sovItems,monthlyProgress,byScope]);

  // All months for monthly progress
  const allMonths=useMemo(()=>{
    if(!sovKeys.length)return[];
    let min=null,max=null;
    sovKeys.forEach(sid=>{const{startDate,finishDate}=byScope[sid];if(!min||startDate<min)min=startDate;if(!max||finishDate>max)max=finishDate;});
    if(!min||!max)return[];
    const months=[];for(let d=new Date(min.getFullYear(),min.getMonth(),1);d<=max;d.setMonth(d.getMonth()+1))months.push(monthKey(new Date(d)));
    return months;
  },[byScope,sovKeys]);

  function exportReport(){
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet([{Metric:"Contract Total",Value:contractTotal},{Metric:"Base",Value:totals.base},{Metric:"Gen Req",Value:totals.gcgr},{Metric:"OH&P",Value:totals.ohp}]),"Summary");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(allDates),"Combined Curve");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(monthlyData),"Monthly Data");
    if(variances.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(variances),"Variance Alerts");
    sovItems.forEach(s=>{if(byScope[s.id])XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(byScope[s.id].full),s.description.slice(0,28).replace(/[^a-zA-Z0-9 ]/g,""));});
    XLSX.writeFile(wb,"WB_Scope_Curves.xlsx");
  }

  const hasData=sovKeys.length>0;
  const TABS=[
    {id:"dashboard",label:"Dashboard",icon:"📊"},
    {id:"curves",label:"Bell Curves",icon:"📈"},
    {id:"monthly",label:"Monthly View",icon:"📅"},
    {id:"forecast",label:"Forecast",icon:"🔮",badge:variances.length||null,badgeColor:WB.red},
    {id:"progress",label:"Progress",icon:"↻"},
    {id:"links",label:"Link Tasks",icon:"🔗",badge:totals.tasks-totals.linked>0?totals.tasks-totals.linked:null,badgeColor:WB.yellow},
    {id:"subs",label:"Subcontractors",icon:"👷"},
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
              <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:4}}>
                {onBackToProjects&&<button onClick={onBackToProjects} style={{background:"none",border:"none",color:WB.primaryLight,cursor:"pointer",fontSize:11,padding:0,display:"flex",alignItems:"center",gap:4}}>← All Projects</button>}
              </div>
              <h1 style={{fontFamily:"'Syne',sans-serif",fontSize:22,fontWeight:800,margin:0,lineHeight:1.1,color:WB.text}}>
                {project?.name||"Project Spend Baseline Tracker"}
              </h1>
              {user&&<div style={{fontSize:11,color:WB.textDim,marginTop:2}}>{user.email}</div>}
            </div>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",alignSelf:"flex-end"}}>
            {/* Save status */}
            {saveStatus==="saved"&&<span style={{fontSize:11,color:WB.green,display:"flex",alignItems:"center",gap:4}}><Icon.Check color={WB.green}/> Saved</span>}
            {saveStatus==="restored"&&<span style={{fontSize:11,color:WB.primaryLight,display:"flex",alignItems:"center",gap:4}}><Icon.Check color={WB.primaryLight}/> Session restored</span>}
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
          <DropZone label="Upload MS Project XML" accept=".xml" onFile={handleScheduleFile} fileName={schedFile?`${schedFile} · ${workTasks.length} tasks`:""} icon={<Icon.File color={WB.primary}/>} hint="File → Save As → XML Format in MS Project"/>
          <DropZone label="Upload SOV Excel" accept=".xlsx,.xls,.csv" onFile={handleSOVFile} fileName={sovFile?`${sovFile} · ${sovItems.length} scopes`:""} icon={<Icon.Layers color={WB.green}/>} hint="GC/GR and OH&P auto-detected and distributed"/>
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
            <Metric title="Scopes Plotted" value={`${totals.plotted} / ${totals.scopes}`} sub={totals.scopes-totals.plotted>0?`${totals.scopes-totals.plotted} need links`:"all plotting"} color={totals.scopes-totals.plotted>0?WB.yellow:WB.green}/>
            <Metric title="Tasks Linked" value={`${totals.linked} / ${totals.tasks}`} color={WB.primary}/>
            <Metric title="Gen Req" value={currency(totals.gcgr)} sub="distributed" color={WB.orange}/>
            <Metric title="OH&P" value={currency(totals.ohp)} sub="distributed" color="#a855f7"/>
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

        {/* ── Dashboard ── */}
        {activeTab==="dashboard"&&(
          hasData?<Dashboard sovItems={sovItems} allDates={allDates} byScope={byScope} sovKeys={sovKeys} monthlyProgress={monthlyProgress} contractTotal={contractTotal} combinedZoom={combinedZoom} variances={variances}/>
          :<Empty msg="Upload both files and link tasks to scopes to see the dashboard."/>
        )}

        {/* ── Bell Curves ── */}
        {activeTab==="curves"&&(
          <div style={{display:"grid",gap:14}}>
            {!hasData&&<Empty msg="Upload both files and link tasks to generate bell curves."/>}
            {hasData&&(
              <div style={{display:"flex",gap:5,flexWrap:"wrap",alignItems:"center"}}>
                <button onClick={()=>setSelectedScope("all")} style={{padding:"4px 10px",borderRadius:99,border:`1px solid ${selectedScope==="all"?WB.primary:WB.border}`,background:selectedScope==="all"?"#1e3448":"transparent",color:selectedScope==="all"?WB.primaryLight:WB.textDim,cursor:"pointer",fontSize:12}}>All scopes</button>
                {sovKeys.map((sid,i)=>{const s=byScope[sid];const col=PAL[i%PAL.length];return(<button key={sid} onClick={()=>setSelectedScope(sid)} style={{padding:"4px 10px",borderRadius:99,border:`1px solid ${selectedScope===sid?col:WB.border}`,background:selectedScope===sid?`${col}20`:"transparent",color:selectedScope===sid?col:WB.textDim,cursor:"pointer",fontSize:12,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.sov.description.slice(0,22)}</button>);})}
              </div>
            )}
            {hasData&&selectedScope==="all"&&(
              <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:20}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
                  <div><h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 2px",color:WB.text}}>All Scope Bell Curves — Stacked</h2><p style={{color:WB.textDim,fontSize:12,margin:0}}>Stacked scope spend · white line = project total · drag to zoom</p></div>
                  <ZoomBar isZoomed={combinedZoom.isZoomed} onReset={combinedZoom.reset}/>
                </div>
                <div style={{height:440,cursor:"crosshair"}}>
                  <ResponsiveContainer width="100%" height="100%">
                    <AreaChart data={combinedZoom.sliced} margin={{top:8,right:8,left:8,bottom:8}} onMouseDown={combinedZoom.onDown} onMouseMove={combinedZoom.onMove} onMouseUp={combinedZoom.onUp}>
                      <defs>{sovKeys.map((sid,i)=>(<linearGradient key={sid} id={`cg${i}`} x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={PAL[i%PAL.length]} stopOpacity={0.55}/><stop offset="95%" stopColor={PAL[i%PAL.length]} stopOpacity={0.04}/></linearGradient>))}</defs>
                      <CartesianGrid strokeDasharray="3 3" stroke={WB.border}/>
                      <XAxis dataKey="date" tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} minTickGap={55} tickFormatter={fmtShort}/>
                      <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} width={52}/>
                      <Tooltip content={<CTip/>}/>
                      <Legend wrapperStyle={{fontSize:10,color:WB.textDim}}/>
                      {sovKeys.map((sid,i)=>{const s=byScope[sid];const col=PAL[i%PAL.length];const key=s.sov.description.slice(0,18);return(<Area key={sid} type="monotone" dataKey={key} stroke={col} strokeWidth={1.5} fill={`url(#cg${i})`} stackId="s" name={key}/>);})}
                      <Line type="monotone" dataKey="total" stroke="#ffffff" strokeWidth={2.5} dot={false} name="Project Total" strokeOpacity={0.9}/>
                      {combinedZoom.sel?.active&&combinedZoom.sel.start&&combinedZoom.sel.end&&<ReferenceArea x1={combinedZoom.sel.start} x2={combinedZoom.sel.end} fill={WB.primary} fillOpacity={0.08}/>}
                    </AreaChart>
                  </ResponsiveContainer>
                </div>
              </div>
            )}
            {hasData&&selectedScope!=="all"&&byScope[selectedScope]&&(()=>{
              const{sov,tasks:lt,startDate,finishDate}=byScope[selectedScope];
              const col=sovColorMap[selectedScope];const data=getScopeSlice(selectedScope);
              const zoomed=!!scopeZooms[selectedScope];const sel=scopeSels[selectedScope]||{};
              return(
                <div style={{display:"grid",gap:12}}>
                  <div style={{background:WB.card,borderRadius:13,border:`1px solid ${col}40`,padding:16,display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:10}}>
                    <div><div style={{fontSize:9,fontFamily:"monospace",color:WB.textDim,textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Scope</div><div style={{fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:13,color:WB.text}}>{sov.description}</div></div>
                    <div><div style={{fontSize:9,fontFamily:"monospace",color:WB.textDim,textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Contract Value</div><div style={{fontWeight:700,fontSize:15,color:col}}>{currency(sov.scheduledValue)}</div></div>
                    <div><div style={{fontSize:9,fontFamily:"monospace",color:WB.textDim,textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Duration</div><div style={{color:WB.textMuted,fontSize:11}}>{fmtDate(startDate)} → {fmtDate(finishDate)}</div></div>
                    <div><div style={{fontSize:9,fontFamily:"monospace",color:WB.textDim,textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Tasks</div><div style={{fontWeight:600,color:WB.textMuted}}>{lt.length}</div></div>
                  </div>
                  <div style={{background:WB.card,borderRadius:13,border:`1px solid ${WB.border}`,padding:20}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
                      <div><h2 style={{fontFamily:"'Syne',sans-serif",fontSize:15,fontWeight:700,margin:"0 0 2px",color:WB.text}}>{sov.description} — Bell Curve</h2><p style={{color:WB.textDim,fontSize:12,margin:0}}>Drag to zoom</p></div>
                      <ZoomBar isZoomed={zoomed} onReset={()=>setScopeZooms(p=>({...p,[selectedScope]:null}))}/>
                    </div>
                    <div style={{height:280,cursor:"crosshair"}}>
                      <ResponsiveContainer width="100%" height="100%">
                        <AreaChart data={data} margin={{top:8,right:8,left:8,bottom:8}}
                          onMouseDown={e=>{if(e?.activeLabel)setScopeSels(p=>({...p,[selectedScope]:{start:e.activeLabel,end:null,active:true}}));}}
                          onMouseMove={e=>{if(sel?.active&&e?.activeLabel)setScopeSels(p=>({...p,[selectedScope]:{...p[selectedScope],end:e.activeLabel}}));}}
                          onMouseUp={()=>{if(sel?.active&&sel.start&&sel.end&&sel.start!==sel.end){const d=byScope[selectedScope]?.weekly||[];const a=d.findIndex(x=>x.date===sel.start),b=d.findIndex(x=>x.date===sel.end);if(a>=0&&b>=0)setScopeZooms(p=>({...p,[selectedScope]:[Math.min(a,b),Math.max(a,b)]}));}setScopeSels(p=>({...p,[selectedScope]:{active:false}}));}}>
                          <defs><linearGradient id="sg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={col} stopOpacity={0.5}/><stop offset="95%" stopColor={col} stopOpacity={0.02}/></linearGradient></defs>
                          <CartesianGrid strokeDasharray="3 3" stroke={WB.border}/>
                          <XAxis dataKey="date" tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} minTickGap={40} tickFormatter={fmtShort}/>
                          <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} width={52}/>
                          <Tooltip content={<CTip/>}/>
                          <Area type="monotone" dataKey="daily" stroke={col} strokeWidth={2.5} fill="url(#sg)" name="Daily Spend"/>
                          {sel?.active&&sel.start&&sel.end&&<ReferenceArea x1={sel.start} x2={sel.end} fill={col} fillOpacity={0.08}/>}
                        </AreaChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                </div>
              );
            })()}
          </div>
        )}

        {/* ── Monthly View ── */}
        {activeTab==="monthly"&&(
          <div style={{display:"grid",gap:14}}>
            {!hasData&&<Empty msg="Upload both files and link tasks to generate monthly projections."/>}
            {hasData&&(
              <>
                {/* Chart 1: Combined + zoomable */}
                {(()=>{
                  const c1Sliced=c1Right!==null?monthlyData.slice(c1Left,c1Right+1):monthlyData;
                  const c1Zoomed=c1Left>0||c1Right!==null;
                  return(
                    <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:20}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap",gap:8}}>
                        <div><h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 2px",color:WB.text}}>① Combined — Scope Curves + Project Total</h2><p style={{color:WB.textDim,fontSize:12,margin:0}}>Each scope's bell curve across its schedule dates · white line = project monthly total</p></div>
                        {c1Zoomed?<button onClick={()=>{setC1Left(0);setC1Right(null);}} style={{display:"flex",alignItems:"center",gap:5,padding:"4px 9px",background:WB.card,border:`1px solid ${WB.borderLight}`,borderRadius:6,color:WB.textMuted,cursor:"pointer",fontSize:11}}><Icon.ZoomOut/> Reset zoom</button>:<span style={{fontSize:11,color:WB.textDim}}>Drag to zoom</span>}
                      </div>
                      <div style={{height:460,cursor:"crosshair"}}>
                        <ResponsiveContainer width="100%" height="100%">
                          <ComposedChart data={c1Sliced} margin={{top:16,right:20,left:8,bottom:8}}
                            onMouseDown={e=>{if(e?.activeLabel)setC1Sel({start:e.activeLabel,end:null,active:true});}}
                            onMouseMove={e=>{if(c1Sel?.active&&e?.activeLabel)setC1Sel(s=>({...s,end:e.activeLabel}));}}
                            onMouseUp={()=>{if(c1Sel?.active&&c1Sel.start&&c1Sel.end&&c1Sel.start!==c1Sel.end){const a=monthlyData.findIndex(d=>d.month===c1Sel.start),b=monthlyData.findIndex(d=>d.month===c1Sel.end);if(a>=0&&b>=0){setC1Left(Math.min(a,b));setC1Right(Math.max(a,b));}}setC1Sel(null);}}>
                            <CartesianGrid strokeDasharray="3 3" stroke={WB.border}/>
                            <XAxis dataKey="month" tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} minTickGap={10}/>
                            <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} width={58}/>
                            <Tooltip content={({active,payload,label})=>{
                              if(!active||!payload?.length)return null;
                              const items=payload.filter(p=>p.value>0&&p.name!=="▸ Project Total").sort((a,b)=>b.value-a.value);
                              const totalVal=monthlyData.find(d=>d.month===label)?.total||0;
                              return(<div style={{background:WB.card,border:`1px solid ${WB.border}`,borderRadius:10,padding:"10px 13px",fontSize:12,maxWidth:300,boxShadow:"0 8px 32px rgba(0,0,0,.6)"}}>
                                <div style={{color:WB.textDim,marginBottom:5,fontFamily:"monospace",fontSize:10}}>{label}</div>
                                {items.slice(0,10).map(p=>(<div key={p.name} style={{display:"flex",gap:7,alignItems:"center",marginBottom:2}}><span style={{width:7,height:7,borderRadius:"50%",background:p.color,flexShrink:0,display:"inline-block"}}/><span style={{color:WB.textMuted,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontSize:11}}>{p.name}:</span><span style={{fontWeight:600,flexShrink:0}}>{currency(p.value)}</span></div>))}
                                <div style={{borderTop:`1px solid ${WB.borderLight}`,marginTop:7,paddingTop:7,display:"flex",justifyContent:"space-between",gap:8}}><span style={{color:WB.text,fontWeight:700}}>Monthly Total:</span><span style={{color:WB.text,fontWeight:800,fontSize:14}}>{currency(totalVal)}</span></div>
                              </div>);
                            }}/>
                            <Legend wrapperStyle={{fontSize:10,color:WB.textDim}}/>
                            {sovKeys.map((sid,i)=>{
                              const s=byScope[sid];const col=PAL[i%PAL.length];const key=s.sov.description.slice(0,18);
                              const p1=new Date(s.startDate.getFullYear(),s.startDate.getMonth()-1,1);
                              const p2=new Date(s.finishDate.getFullYear(),s.finishDate.getMonth()+2,0);
                              const ps=`${p1.getFullYear()}-${String(p1.getMonth()+1).padStart(2,"0")}`;
                              const pe=`${p2.getFullYear()}-${String(p2.getMonth()+1).padStart(2,"0")}`;
                              return(<Line key={sid} type="monotone" dataKey={d=>d.month>=ps&&d.month<=pe?(d[key]||0):null} stroke={col} strokeWidth={2} dot={false} name={key} connectNulls={false} activeDot={{r:4,fill:col,strokeWidth:0}}/>);
                            })}
                            <Line type="monotone" dataKey="total" stroke="#ffffff" strokeWidth={3} dot={{fill:"#ffffff",r:3,strokeWidth:0}} activeDot={{r:6,fill:"#ffffff",stroke:WB.primary,strokeWidth:2}} name="▸ Project Total" strokeOpacity={1} legendType="line"/>
                            {c1Sel?.active&&c1Sel.start&&c1Sel.end&&<ReferenceArea x1={c1Sel.start} x2={c1Sel.end} fill={WB.primary} fillOpacity={0.08}/>}
                          </ComposedChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                  );
                })()}

                {/* Chart 2: Project total only */}
                <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:20}}>
                  <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 4px",color:WB.text}}>② Project Total Bell Curve — Monthly</h2>
                  <p style={{color:WB.textDim,fontSize:12,marginBottom:14}}>Overarching project spend across all scopes combined, by month</p>
                  <div style={{height:300}}>
                    <ResponsiveContainer width="100%" height="100%">
                      <AreaChart data={monthlyData} margin={{top:16,right:20,left:8,bottom:8}}>
                        <defs><linearGradient id="totMonthGrad" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#ffffff" stopOpacity={0.2}/><stop offset="95%" stopColor="#ffffff" stopOpacity={0.02}/></linearGradient></defs>
                        <CartesianGrid strokeDasharray="3 3" stroke={WB.border}/>
                        <XAxis dataKey="month" tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} minTickGap={10}/>
                        <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} width={58}/>
                        <Tooltip content={({active,payload,label})=>{
                          if(!active||!payload?.length)return null;
                          const totalVal=monthlyData.find(d=>d.month===label)?.total||0;
                          return(<div style={{background:WB.card,border:`1px solid ${WB.border}`,borderRadius:10,padding:"10px 13px",boxShadow:"0 8px 32px rgba(0,0,0,.6)"}}><div style={{color:WB.textDim,marginBottom:5,fontFamily:"monospace",fontSize:10}}>{label}</div><div style={{display:"flex",justifyContent:"space-between",gap:16}}><span style={{color:WB.text,fontWeight:700}}>Monthly Total:</span><span style={{color:WB.text,fontWeight:800,fontSize:14}}>{currency(totalVal)}</span></div></div>);
                        }}/>
                        <Area type="monotone" dataKey="total" stroke="#ffffff" strokeWidth={3} fill="url(#totMonthGrad)" name="Project Total" dot={{fill:"#ffffff",r:4,strokeWidth:0}} activeDot={{r:6,fill:"#ffffff",stroke:WB.primary,strokeWidth:2}}/>
                      </AreaChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                {/* Chart 3: Individual scopes zoomed */}
                {(()=>{
                  const sliced=mRight!==null?monthlyData.slice(mLeft,mRight+1):monthlyData;
                  const isZoomed=mLeft>0||mRight!==null;
                  return(
                    <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:20}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap",gap:8}}>
                        <div><h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 2px",color:WB.text}}>③ Individual Scope Curves — Zoomed</h2><p style={{color:WB.textDim,fontSize:12,margin:0}}>Drag to zoom in on any time window</p></div>
                        {isZoomed?<button onClick={()=>{setMLeft(0);setMRight(null);}} style={{display:"flex",alignItems:"center",gap:5,padding:"4px 9px",background:WB.card,border:`1px solid ${WB.borderLight}`,borderRadius:6,color:WB.textMuted,cursor:"pointer",fontSize:11}}><Icon.ZoomOut/> Reset zoom</button>:<span style={{fontSize:11,color:WB.textDim}}>Drag to zoom</span>}
                      </div>
                      <div style={{height:480,cursor:"crosshair"}}>
                        <ResponsiveContainer width="100%" height="100%">
                          <ComposedChart data={sliced} margin={{top:16,right:20,left:8,bottom:8}}
                            onMouseDown={e=>{if(e?.activeLabel)setMSel({start:e.activeLabel,end:null,active:true});}}
                            onMouseMove={e=>{if(mSel?.active&&e?.activeLabel)setMSel(s=>({...s,end:e.activeLabel}));}}
                            onMouseUp={()=>{if(mSel?.active&&mSel.start&&mSel.end&&mSel.start!==mSel.end){const a=monthlyData.findIndex(d=>d.month===mSel.start),b=monthlyData.findIndex(d=>d.month===mSel.end);if(a>=0&&b>=0){setMLeft(Math.min(a,b));setMRight(Math.max(a,b));}}setMSel(null);}}>
                            <CartesianGrid strokeDasharray="3 3" stroke={WB.border}/>
                            <XAxis dataKey="month" tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} minTickGap={8}/>
                            <YAxis tickFormatter={v=>`$${Math.round(v/1000)}k`} tick={{fontSize:9,fill:WB.textDim,fontFamily:"monospace"}} width={58}/>
                            <Tooltip content={({active,payload,label})=>{
                              if(!active||!payload?.length)return null;
                              const items=payload.filter(p=>p.value>0).sort((a,b)=>b.value-a.value);
                              const totalVal=monthlyData.find(d=>d.month===label)?.total||0;
                              return(<div style={{background:WB.card,border:`1px solid ${WB.border}`,borderRadius:10,padding:"10px 13px",fontSize:12,maxWidth:300,boxShadow:"0 8px 32px rgba(0,0,0,.6)"}}>
                                <div style={{color:WB.textDim,marginBottom:5,fontFamily:"monospace",fontSize:10}}>{label}</div>
                                {items.slice(0,10).map(p=>(<div key={p.name} style={{display:"flex",gap:7,alignItems:"center",marginBottom:2}}><span style={{width:7,height:7,borderRadius:"50%",background:p.color,flexShrink:0,display:"inline-block"}}/><span style={{color:WB.textMuted,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontSize:11}}>{p.name}:</span><span style={{fontWeight:600,flexShrink:0}}>{currency(p.value)}</span></div>))}
                                <div style={{borderTop:`1px solid ${WB.borderLight}`,marginTop:7,paddingTop:7,display:"flex",justifyContent:"space-between",gap:8}}><span style={{color:WB.text,fontWeight:700}}>Monthly Total:</span><span style={{color:WB.text,fontWeight:800,fontSize:14}}>{currency(totalVal)}</span></div>
                              </div>);
                            }}/>
                            <Legend wrapperStyle={{fontSize:10,color:WB.textDim}}/>
                            {sovKeys.map((sid,i)=>{
                              const s=byScope[sid];const col=PAL[i%PAL.length];const key=s.sov.description.slice(0,18);
                              const p1=new Date(s.startDate.getFullYear(),s.startDate.getMonth()-1,1);
                              const p2=new Date(s.finishDate.getFullYear(),s.finishDate.getMonth()+2,0);
                              const ps=`${p1.getFullYear()}-${String(p1.getMonth()+1).padStart(2,"0")}`;
                              const pe=`${p2.getFullYear()}-${String(p2.getMonth()+1).padStart(2,"0")}`;
                              return(<Line key={sid} type="monotone" dataKey={d=>d.month>=ps&&d.month<=pe?(d[key]||0):null} stroke={col} strokeWidth={2.5} dot={{fill:col,r:3,strokeWidth:0}} name={key} connectNulls={false} activeDot={{r:5,fill:col,stroke:"#fff",strokeWidth:1}}/>);
                            })}
                            {mSel?.active&&mSel.start&&mSel.end&&<ReferenceArea x1={mSel.start} x2={mSel.end} fill={WB.primary} fillOpacity={0.08}/>}
                          </ComposedChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                  );
                })()}

                {/* Monthly totals table */}
                <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:20,overflowX:"auto"}}>
                  <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 4px",color:WB.text}}>Monthly Subcontractor Totals</h2>
                  <p style={{color:WB.textDim,fontSize:12,marginBottom:14}}>Sum of all scope spend per month</p>
                  <table style={{width:"100%",minWidth:500,borderCollapse:"collapse",fontSize:12}}>
                    <thead>
                      <tr style={{background:WB.bg}}>
                        <th style={{padding:"7px 12px",textAlign:"left",color:WB.textDim,fontFamily:"monospace",fontSize:9,borderBottom:`1px solid ${WB.border}`,textTransform:"uppercase",letterSpacing:.5}}>Month</th>
                        {sovKeys.map((sid,i)=>{const s=byScope[sid];const col=PAL[i%PAL.length];return(<th key={sid} style={{padding:"7px 10px",textAlign:"right",color:col,fontFamily:"monospace",fontSize:9,borderBottom:`1px solid ${WB.border}`,textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{s.sov.description.slice(0,14)}</th>);})}
                        <th style={{padding:"7px 12px",textAlign:"right",color:"#ffffff",fontFamily:"monospace",fontSize:9,borderBottom:`1px solid ${WB.border}`,textTransform:"uppercase",letterSpacing:.5}}>Monthly Total</th>
                        <th style={{padding:"7px 12px",textAlign:"right",color:WB.primary,fontFamily:"monospace",fontSize:9,borderBottom:`1px solid ${WB.border}`,textTransform:"uppercase",letterSpacing:.5}}>Cumulative</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(()=>{let running=0;return monthlyData.map((row,i)=>{running+=row.total;return(<tr key={row.month} style={{borderBottom:`1px solid ${WB.border}`,background:i%2===0?"transparent":WB.bg}}><td style={{padding:"7px 12px",fontFamily:"monospace",color:WB.textDim,fontSize:11,whiteSpace:"nowrap"}}>{fmtMonth(row.month)}</td>{sovKeys.map(sid=>{const key=byScope[sid].sov.description.slice(0,18);const col=sovColorMap[sid];return(<td key={sid} style={{padding:"7px 10px",textAlign:"right",fontFamily:"monospace",color:row[key]>0?col:WB.textDim,fontSize:11}}>{row[key]>0?currency(row[key]):"—"}</td>);})}<td style={{padding:"7px 12px",textAlign:"right",fontFamily:"monospace",color:WB.text,fontWeight:700,fontSize:11}}>{currency(row.total)}</td><td style={{padding:"7px 12px",textAlign:"right",fontFamily:"monospace",color:WB.primary,fontSize:11}}>{currency(running)}</td></tr>);});})()}
                    </tbody>
                    <tfoot>
                      <tr style={{borderTop:`2px solid ${WB.border}`}}>
                        <td style={{padding:"8px 12px",fontWeight:700,color:WB.text}}>TOTAL</td>
                        {sovKeys.map(sid=>{const key=byScope[sid].sov.description.slice(0,18);const col=sovColorMap[sid];const total=monthlyData.reduce((s,r)=>s+(r[key]||0),0);return(<td key={sid} style={{padding:"8px 10px",textAlign:"right",fontFamily:"monospace",color:col,fontWeight:700,fontSize:11}}>{currency(total)}</td>);})}
                        <td style={{padding:"8px 12px",textAlign:"right",fontFamily:"monospace",color:WB.green,fontWeight:700,fontSize:13}}>{currency(monthlyData.reduce((s,r)=>s+r.total,0))}</td>
                        <td style={{padding:"8px 12px",textAlign:"right",fontFamily:"monospace",color:WB.primary,fontWeight:700,fontSize:13}}>{currency(monthlyData.reduce((s,r)=>s+r.total,0))}</td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              </>
            )}
          </div>
        )}

        {/* ── Forecast ── */}
        {activeTab==="forecast"&&(
          hasData?<ForecastPanel forecasts={forecasts} variances={variances}/>
          :<Empty msg="Upload both files, link scopes, and enter monthly progress to see forecasts."/>
        )}

        {/* ── Progress ── */}
        {activeTab==="progress"&&(
          hasData?<MonthlyProgressPanel sovItems={sovItems} monthlyProgress={monthlyProgress} onUpdate={setMonthlyProgress} byScope={byScope} sovKeys={sovKeys}/>
          :<Empty msg="Upload both files and link scopes to enter progress updates."/>
        )}

        {/* ── Link Tasks ── */}
        {activeTab==="links"&&(
          workTasks.length&&sovItems.length
            ?<LinkManager tasks={workTasks} sovItems={sovItems} linksMap={linksMap} onLinksChange={setLinksMap}/>
            :<div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:40,textAlign:"center"}}><div style={{fontSize:13,color:WB.textDim}}>Upload both files to use the link manager.</div></div>
        )}

        {/* ── SOV Detail ── */}
        {activeTab==="sov"&&(
          <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:18,overflowX:"auto"}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 4px",color:WB.text}}>SOV Line Items — Distributed Values</h2>
            <p style={{color:WB.textDim,fontSize:12,marginBottom:14}}>Base + Gen Req + OH&P = contract value per scope</p>
            {sovItems.length?(
              <table style={{width:"100%",minWidth:660,borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{background:WB.bg}}>{["Ph","Description","Base","+ Gen Req","+ OH&P","Contract Value","Tasks"].map(h=><th key={h} style={{padding:"7px 9px",textAlign:"left",color:WB.textDim,fontFamily:"monospace",fontSize:9,borderBottom:`1px solid ${WB.border}`,textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                <tbody>
                  {sovItems.map((s,i)=>{const col=PAL[i%PAL.length];const count=(linksMap[s.id]||[]).length;return(<tr key={s.id} style={{borderBottom:`1px solid ${WB.border}`,background:i%2===0?"transparent":WB.bg}}><td style={{padding:"7px 9px",fontFamily:"monospace",color:WB.textDim,fontSize:10}}>{s.phase||"—"}</td><td style={{padding:"7px 9px",color:WB.text,maxWidth:170,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.description}</td><td style={{padding:"7px 9px",color:WB.textMuted,fontFamily:"monospace"}}>{currency(s.baseValue)}</td><td style={{padding:"7px 9px",color:WB.orange,fontFamily:"monospace",fontSize:10}}>{currency(s.gcgrShare)}</td><td style={{padding:"7px 9px",color:"#a855f7",fontFamily:"monospace",fontSize:10}}>{currency(s.ohpShare)}</td><td style={{padding:"7px 9px",fontWeight:700,color:col,fontFamily:"monospace"}}>{currency(s.scheduledValue)}</td><td style={{padding:"7px 9px"}}><span style={{fontSize:10,color:count>0?col:WB.textDim,background:count>0?`${col}20`:WB.bg,padding:"1px 7px",borderRadius:99,border:`1px solid ${count>0?`${col}40`:WB.border}`}}>{count}</span></td></tr>);})}
                </tbody>
                <tfoot><tr style={{borderTop:`2px solid ${WB.border}`}}><td colSpan={2} style={{padding:"8px 9px",fontWeight:700,color:WB.text}}>TOTAL</td><td style={{padding:"8px 9px",fontWeight:700,color:WB.textMuted,fontFamily:"monospace"}}>{currency(totals.base)}</td><td style={{padding:"8px 9px",fontWeight:700,color:WB.orange,fontFamily:"monospace"}}>{currency(totals.gcgr)}</td><td style={{padding:"8px 9px",fontWeight:700,color:"#a855f7",fontFamily:"monospace"}}>{currency(totals.ohp)}</td><td style={{padding:"8px 9px",fontWeight:700,color:WB.green,fontFamily:"monospace",fontSize:14}}>{currency(totals.contract)}</td><td/></tr></tfoot>
              </table>
            ):<Empty msg="Upload your SOV to see the breakdown."/>}
          </div>
        )}


        {/* ── Subcontractors ── */}
        {activeTab==="subs"&&(
          <SubcontractorModule
            subcontractors={subcontractors}
            onUpdate={setSubcontractors}
            sovItems={sovItems}
            byScope={byScope}
            sovKeys={sovKeys}
            monthlyData={monthlyData}
            sovColorMap={sovColorMap}
            selectedSub={selectedSub}
            onSelectSub={setSelectedSub}
          />
        )}

        {/* ── Validation ── */}
        {activeTab==="validation"&&(
          <div style={{background:WB.card,borderRadius:14,border:`1px solid ${WB.border}`,padding:18}}>
            <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:700,margin:"0 0 4px",color:WB.text}}>Validation Log</h2>
            <p style={{color:WB.textDim,fontSize:12,marginBottom:14}}>Parse issues from uploaded files.</p>
            {parseErrors.length?(<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(270px,1fr))",gap:7}}>{parseErrors.map((e,i)=>(<div key={i} style={{background:WB.redBg,border:`1px solid ${WB.redBorder}`,borderRadius:8,padding:"9px 12px"}}><div style={{fontFamily:"monospace",fontSize:9,color:WB.red,marginBottom:2}}>ROW {e.row||"—"}</div><div style={{color:WB.text,fontSize:12}}>{e.issue}</div></div>))}</div>)
            :(<div style={{display:"flex",alignItems:"center",gap:9,padding:"11px 14px",background:WB.greenBg,border:`1px solid ${WB.greenBorder}`,borderRadius:8}}><Icon.Check color={WB.green}/><span style={{color:"#86efac",fontSize:12}}>{rawTasks.length?"No issues found.":"Upload a schedule to validate."}</span></div>)}
          </div>
        )}

      </div>
    </div>
  );
}