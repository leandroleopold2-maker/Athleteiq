import { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, BarChart, Bar, RadarChart, Radar, PolarGrid,
  PolarAngleAxis, PolarRadiusAxis, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer,
} from "recharts";

// ── F1 Design Tokens ──────────────────────────────────────────────────────────
const F = {
  carbon:    "#0A0A0A",
  asphalt:   "#141414",
  panel:     "#1C1C1C",
  panelBorder:"#2A2A2A",
  red:       "#E8002D",
  redDim:    "#9B001E",
  white:     "#F0F0F0",
  silver:    "#A8A8A8",
  dim:       "#5A5A5A",
  ghost:     "#2E2E2E",
  yellow:    "#FFD700",
  teal:      "#00D2FF",
  green:     "#39B54A",
  fontF1:    "'Titillium Web', 'Arial Narrow', sans-serif",
  fontMono:  "'Share Tech Mono', 'Courier New', monospace",
};

const TEAM_COLORS = [F.red, F.teal, F.yellow, "#FF8700", "#00D2BE", "#DC143C", "#0600EF", "#006F62", "#B6BABD", "#FFFFFF"];

// ── References ────────────────────────────────────────────────────────────────
const JUMP_REFS = {
  cmj:      { excellent:42, good:35, average:28, unit:"cm", label:"CMJ",      lowerIsBetter:false },
  sj:       { excellent:38, good:31, average:24, unit:"cm", label:"SJ",       lowerIsBetter:false },
  dj:       { excellent:38, good:30, average:22, unit:"cm", label:"DJ",       lowerIsBetter:false },
  rsi:      { excellent:2.0,good:1.5,average:1.0,unit:"",  label:"RSI",      lowerIsBetter:false },
  fuerza:   { excellent:3500,good:2800,average:2200,unit:"N",label:"Fuerza", lowerIsBetter:false },
  potencia: { excellent:5000,good:3500,average:2500,unit:"W",label:"Potencia",lowerIsBetter:false },
};
const VEL_REFS = {
  "10m":{ excellent:1.65,good:1.75,average:1.85,unit:"s",label:"10m",lowerIsBetter:true },
  "20m":{ excellent:2.80,good:2.95,average:3.10,unit:"s",label:"20m",lowerIsBetter:true },
  "30m":{ excellent:3.90,good:4.10,average:4.30,unit:"s",label:"30m",lowerIsBetter:true },
  "40m":{ excellent:4.90,good:5.15,average:5.40,unit:"s",label:"40m",lowerIsBetter:true },
  "60m":{ excellent:6.80,good:7.10,average:7.40,unit:"s",label:"60m",lowerIsBetter:true },
};

function scoreVal(ref,value) {
  if(!ref||value==null) return null;
  const{excellent,good,average,lowerIsBetter}=ref;
  if(lowerIsBetter){
    if(value<=excellent) return 100; if(value<=good) return 75; if(value<=average) return 50; return 25;
  } else {
    if(value>=excellent) return 100; if(value>=good) return 75; if(value>=average) return 50; return 25;
  }
}
function getLevel(s) {
  if(s>=90) return{label:"ELITE",    color:F.yellow};
  if(s>=70) return{label:"BUENO",    color:F.green};
  if(s>=50) return{label:"PROMEDIO", color:F.teal};
  return         {label:"BAJO",      color:F.red};
}

// ── Parsers ───────────────────────────────────────────────────────────────────
function parseNum(str) {
  if(!str||str==="---"||str==="") return null;
  return parseFloat(String(str).replace(",",".").replace(/[^\d.-]/g,""))||null;
}
function parseMyJumpDate(str) {
  const m=String(str).match(/(\d+)\.(\d+)\.(\d+)/);
  if(!m) return str;
  return`${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
}
function parsePhotoDate(str) {
  const m=String(str).match(/(\d+)\/(\d+)\/(\d+)/);
  if(!m) return str;
  return`${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
}
function jumpTypeKey(tipo) {
  if(!tipo) return"cmj"; const t=tipo.toLowerCase();
  if(t.includes("cmj")) return"cmj"; if(t.includes("sj")&&!t.includes("rsi")) return"sj";
  if(t.includes("dj")||t.includes("drop")) return"dj"; return"cmj";
}
function detectDistance(info) {
  if(!info) return null; const m=String(info).match(/(\d+)\s*m/i); return m?m[1]+"m":null;
}
function parseMyJumpCSV(text) {
  const lines=text.trim().split("\n"); if(lines.length<2) return[];
  const headers=lines[0].split(";").map(h=>h.trim());
  return lines.slice(1).map(line=>{const vals=line.split(";");const obj={};headers.forEach((h,i)=>{obj[h]=(vals[i]||"").trim();});return obj;});
}
function parsePhotoFinishCSV(text) {
  const lines=text.trim().split("\n"); if(lines.length<2) return[];
  let cDate="",cInfo=""; const rows=[];
  lines.slice(1).forEach(line=>{
    const cols=line.split(",").map(c=>c.trim());
    if(cols[0]) cDate=cols[0]; if(cols[1]) cInfo=cols[1];
    const athlete=cols[3]||""; const time=cols[5]||"";
    if(!athlete) return;
    rows.push({date:cDate,info:cInfo,athlete,time});
  });
  return rows;
}

// ── F1 Components ─────────────────────────────────────────────────────────────

// Telemetry bar — the signature element
const TelemetryBar = () => {
  const [pos, setPos] = useState(0);
  useEffect(()=>{
    const id=setInterval(()=>setPos(p=>(p+0.4)%100),30);
    return()=>clearInterval(id);
  },[]);
  const segments=[F.red,F.teal,F.yellow,F.red,F.dim,F.teal,F.dim,F.red];
  return (
    <div style={{height:3,background:F.ghost,overflow:"hidden",position:"relative"}}>
      <div style={{position:"absolute",left:`${pos}%`,top:0,height:"100%",width:"30%",
        background:`linear-gradient(90deg, transparent, ${F.red}, ${F.teal}, ${F.yellow}, transparent)`,
        transition:"none",opacity:0.8}}/>
      <div style={{display:"flex",height:"100%"}}>
        {segments.map((c,i)=>(
          <div key={i} style={{flex:1,background:c,opacity:0.3}}/>
        ))}
      </div>
    </div>
  );
};

const CustomTooltip=({active,payload,label})=>{
  if(!active||!payload?.length) return null;
  return(
    <div style={{background:F.panel,border:`1px solid ${F.red}`,borderRadius:4,padding:"8px 14px"}}>
      <div style={{fontFamily:F.fontMono,fontSize:10,color:F.red,marginBottom:4,letterSpacing:2}}>{label}</div>
      {payload.map((p,i)=>(
        <div key={i} style={{color:p.color,fontSize:11,fontFamily:F.fontMono,margin:"1px 0"}}>
          {p.name}: <span style={{color:F.white,fontWeight:700}}>{typeof p.value==="number"?p.value.toFixed(2):p.value}</span>
        </div>
      ))}
    </div>
  );
};

// F1-style sector badge
const SectorBadge=({score})=>{
  const lv=getLevel(score);
  return(
    <div style={{display:"inline-flex",alignItems:"center",gap:6,padding:"2px 8px",border:`1px solid ${lv.color}`,borderRadius:2,background:lv.color+"18"}}>
      <div style={{width:6,height:6,borderRadius:"50%",background:lv.color,boxShadow:`0 0 6px ${lv.color}`}}/>
      <span style={{fontFamily:F.fontMono,fontSize:10,color:lv.color,letterSpacing:2}}>{lv.label}</span>
    </div>
  );
};

// Score display with F1 timing panel aesthetic
const TimingScore=({score,size=64})=>{
  const lv=getLevel(score);
  const r=size/2-5; const circ=2*Math.PI*r; const dash=(score/100)*circ;
  return(
    <div style={{position:"relative",width:size,height:size}}>
      <svg width={size} height={size} style={{transform:"rotate(-90deg)"}}>
        <circle cx={size/2} cy={size/2} r={r} fill="none" stroke={F.ghost} strokeWidth={3}/>
        <circle cx={size/2} cy={size/2} r={r} fill="none" stroke={lv.color} strokeWidth={3}
          strokeDasharray={`${dash} ${circ}`} strokeLinecap="butt"
          style={{transition:"stroke-dasharray 0.8s ease",filter:`drop-shadow(0 0 4px ${lv.color})`}}/>
      </svg>
      <div style={{position:"absolute",inset:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"}}>
        <div style={{fontFamily:F.fontMono,fontSize:size*0.26,fontWeight:700,color:lv.color,lineHeight:1}}>{score}</div>
      </div>
    </div>
  );
};

// Data panel header
const PanelHeader=({children,accent})=>(
  <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14,paddingBottom:8,borderBottom:`1px solid ${F.panelBorder}`}}>
    <div style={{width:3,height:14,background:accent||F.red,borderRadius:1}}/>
    <div style={{fontFamily:F.fontMono,fontSize:10,color:F.silver,letterSpacing:2,textTransform:"uppercase"}}>{children}</div>
  </div>
);

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [players,  setPlayers]  = useState([]);
  const [jumpRecs, setJumpRecs] = useState([]);
  const [velRecs,  setVelRecs]  = useState([]);
  const [customRecs,setCustomRecs]=useState([]); // from xlsx/docx/generic
  const [tab,      setTab]      = useState("import");
  const [selPlayer,  setSelPlayer]  = useState(null);
  const [compareA,   setCompareA]   = useState(null);
  const [compareB,   setCompareB]   = useState(null);
  const [filterTeam, setFilterTeam] = useState("Todos");
  const [msgs,       setMsgs]       = useState({});

  // Generic file mapper state
  const [genericData,  setGenericData]  = useState(null); // raw rows from xlsx/docx
  const [genericCols,  setGenericCols]  = useState([]);   // column names
  const [colMap,       setColMap]       = useState({});   // user mapping
  const [showMapper,   setShowMapper]   = useState(false);

  const jumpRef=useRef(); const velRef=useRef(); const genericRef=useRef();

  function setMsg(key,val){setMsgs(m=>({...m,[key]:val}));}

  function mergePlayer(name,team,cur){
    const key=name.toLowerCase().trim();
    const ex=cur.find(p=>p.name.toLowerCase().trim()===key);
    if(ex) return{players:cur,id:ex.id};
    const np={id:Date.now()+Math.random(),name:name.trim(),team:team||""};
    return{players:[...cur,np],id:np.id};
  }

  // ── MyJump import ─────────────────────────────────────────────────────────
  function handleJumpImport(e){
    const file=e.target.files[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=(ev)=>{
      const rows=parseMyJumpCSV(ev.target.result);
      if(!rows.length){setMsg("jump","❌ No se pudo leer.");return;}
      let cur=[...players]; const newRecs=[];
      rows.forEach(row=>{
        const name=(row["Nombre"]||"").trim(); const team=(row["Equipo"]||"").trim(); if(!name) return;
        const{players:up,id:pid}=mergePlayer(name,team,cur); cur=up;
        const tipo=(row["Tipo de salto"]||"").trim(); const fecha=parseMyJumpDate(row["Fecha"]||"");
        newRecs.push({id:Date.now()+Math.random(),playerId:pid,date:fecha.slice(0,7),fullDate:fecha,
          jumpType:jumpTypeKey(tipo),jumpTypeRaw:tipo,
          altura:parseNum(row["Altura de salto (cm)"]),rsi:parseNum(row["RSI mod (m/s)"]),
          fuerza:parseNum(row["Fuerza (N)"]),potencia:parseNum(row["Potencia (W)"]),
          statusColor:row["Color estado de forma"]||"",team});
      });
      setPlayers(cur); setJumpRecs(r=>[...r,...newRecs]);
      if(!selPlayer&&cur.length){setSelPlayer(cur[0].id);setCompareA(cur[0].id);setCompareB(cur[1]?.id||cur[0].id);}
      setMsg("jump",`✓ ${newRecs.length} registros · ${[...new Set(newRecs.map(r=>r.playerId))].length} atletas`);
      setTab("dashboard");
    };
    reader.readAsText(file,"UTF-8");
  }

  // ── PhotoFinish import ────────────────────────────────────────────────────
  function handleVelImport(e){
    const file=e.target.files[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=(ev)=>{
      const rows=parsePhotoFinishCSV(ev.target.result);
      if(!rows.length){setMsg("vel","❌ No se pudo leer.");return;}
      let cur=[...players]; const newRecs=[];
      rows.forEach(row=>{
        const name=row.athlete.trim(); if(!name) return;
        const{players:up,id:pid}=mergePlayer(name,"",cur); cur=up;
        const fecha=parsePhotoDate(row.date); const dist=detectDistance(row.info);
        const t=parseNum(row.time); if(!t) return;
        newRecs.push({id:Date.now()+Math.random(),playerId:pid,date:fecha.slice(0,7),fullDate:fecha,distance:dist,testName:row.info,time:t});
      });
      setPlayers(cur); setVelRecs(r=>[...r,...newRecs]);
      if(!selPlayer&&cur.length){setSelPlayer(cur[0].id);setCompareA(cur[0].id);setCompareB(cur[1]?.id||cur[0].id);}
      setMsg("vel",`✓ ${newRecs.length} registros · ${[...new Set(newRecs.map(r=>r.playerId))].length} atletas`);
      setTab("dashboard");
    };
    reader.readAsText(file,"UTF-8");
  }

  // ── Generic import (xlsx / csv) ───────────────────────────────────────────
  function handleGenericImport(e){
    const file=e.target.files[0]; if(!file) return;
    const ext=file.name.split(".").pop().toLowerCase();
    const reader=new FileReader();
    reader.onload=(ev)=>{
      try{
        let rows=[];
        if(ext==="xlsx"||ext==="xls"){
          const wb=XLSX.read(ev.target.result,{type:"array"});
          const ws=wb.Sheets[wb.SheetNames[0]];
          rows=XLSX.utils.sheet_to_json(ws,{defval:""});
        } else {
          // generic CSV (comma or semicolon)
          const text=new TextDecoder().decode(ev.target.result);
          const lines=text.trim().split("\n"); if(lines.length<2) return;
          const sep=lines[0].includes(";")?";":","
          const headers=lines[0].split(sep).map(h=>h.trim());
          rows=lines.slice(1).map(line=>{const vals=line.split(sep);const obj={};headers.forEach((h,i)=>{obj[h]=(vals[i]||"").trim();});return obj;});
        }
        if(!rows.length){setMsg("gen","❌ Archivo vacío.");return;}
        const cols=Object.keys(rows[0]);
        setGenericData(rows); setGenericCols(cols);
        setColMap({name:"",date:"",metric1:"",metric1label:"",metric2:"",metric2label:"",team:""});
        setShowMapper(true);
        setMsg("gen",`✓ ${rows.length} filas · ${cols.length} columnas detectadas. Mapeá las columnas abajo.`);
      }catch(err){setMsg("gen","❌ Error al leer el archivo.");}
    };
    if(ext==="xlsx"||ext==="xls") reader.readAsArrayBuffer(file);
    else reader.readAsArrayBuffer(file);
  }

  function applyMapping(){
    if(!genericData||!colMap.name) return;
    let cur=[...players]; const newRecs=[];
    genericData.forEach(row=>{
      const name=String(row[colMap.name]||"").trim(); if(!name) return;
      const team=colMap.team?String(row[colMap.team]||"").trim():"";
      const{players:up,id:pid}=mergePlayer(name,team,cur); cur=up;
      const dateRaw=colMap.date?String(row[colMap.date]||""):"";
      const fecha=dateRaw||new Date().toISOString().slice(0,10);
      const rec={id:Date.now()+Math.random(),playerId:pid,date:fecha.slice(0,7),fullDate:fecha,source:"custom",team};
      if(colMap.metric1&&colMap.metric1label) rec[colMap.metric1label]=parseNum(row[colMap.metric1]);
      if(colMap.metric2&&colMap.metric2label) rec[colMap.metric2label]=parseNum(row[colMap.metric2]);
      if(colMap.metric3&&colMap.metric3label) rec[colMap.metric3label]=parseNum(row[colMap.metric3]);
      newRecs.push(rec);
    });
    setPlayers(cur); setCustomRecs(r=>[...r,...newRecs]);
    if(!selPlayer&&cur.length){setSelPlayer(cur[0].id);setCompareA(cur[0].id);setCompareB(cur[1]?.id||cur[0].id);}
    setShowMapper(false);
    setMsg("gen",`✓ ${newRecs.length} registros importados con mapeo personalizado`);
    setTab("dashboard");
  }

  const hasData=players.length>0;
  const teams=useMemo(()=>["Todos",...Array.from(new Set(players.map(p=>p.team).filter(Boolean)))],[players]);
  const visible=useMemo(()=>players.filter(p=>filterTeam==="Todos"||p.team===filterTeam),[players,filterTeam]);

  const pJR=(pid)=>jumpRecs.filter(r=>r.playerId===pid).sort((a,b)=>a.fullDate.localeCompare(b.fullDate));
  const pVR=(pid)=>velRecs.filter(r=>r.playerId===pid).sort((a,b)=>a.fullDate.localeCompare(b.fullDate));
  const bestJump=(pid)=>{const rs=pJR(pid).filter(r=>r.altura!=null);return rs.reduce((b,r)=>r.altura>(b?.altura??0)?r:b,null);};
  const bestVel=(pid,dist)=>{const rs=pVR(pid).filter(r=>r.distance===dist&&r.time!=null);return rs.reduce((b,r)=>r.time<(b?.time??999)?r:b,null);};
  const lastJump=(pid)=>{const rs=pJR(pid).sort((a,b)=>b.fullDate.localeCompare(a.fullDate));return rs[0]||null;};
  const allDist=useMemo(()=>Array.from(new Set(velRecs.map(r=>r.distance).filter(Boolean))).sort(),[velRecs]);

  const overallScore=(pid)=>{
    const sc=[];
    const bj=bestJump(pid);
    if(bj){sc.push(scoreVal(JUMP_REFS[bj.jumpType]||JUMP_REFS.cmj,bj.altura)??0);if(bj.rsi!=null)sc.push(scoreVal(JUMP_REFS.rsi,bj.rsi)??0);}
    allDist.forEach(d=>{const bv=bestVel(pid,d);if(bv)sc.push(scoreVal(VEL_REFS[d],bv.time)??0);});
    return sc.length?Math.round(sc.reduce((a,b)=>a+b,0)/sc.length):0;
  };

  const rankingData=useMemo(()=>
    visible.filter(p=>(pJR(p.id).length+pVR(p.id).length)>0)
      .map(p=>({name:p.name.split(" ")[0],score:overallScore(p.id)}))
      .sort((a,b)=>b.score-a.score),
    [players,jumpRecs,velRecs,filterTeam]);

  const jumpEvo=(pid)=>{
    const byDate={};
    pJR(pid).forEach(r=>{
      if(!byDate[r.fullDate]) byDate[r.fullDate]={fecha:r.fullDate.slice(8)+"/"+r.fullDate.slice(5,7),_r:[]};
      byDate[r.fullDate]._r.push(r);
    });
    return Object.values(byDate).map(d=>{
      const cmjs=d._r.filter(r=>r.jumpType==="cmj"&&r.altura!=null);
      const djs=d._r.filter(r=>r.jumpType==="dj"&&r.altura!=null);
      const rsif=d._r.filter(r=>r.rsi!=null);
      return{fecha:d.fecha,CMJ:cmjs.length?Math.max(...cmjs.map(r=>r.altura)):null,DJ:djs.length?Math.max(...djs.map(r=>r.altura)):null,RSI:rsif.length?Math.max(...rsif.map(r=>r.rsi)):null};
    });
  };

  const velEvo=(pid)=>{
    const byDate={};
    pVR(pid).forEach(r=>{
      if(!byDate[r.fullDate]) byDate[r.fullDate]={fecha:r.fullDate.slice(8)+"/"+r.fullDate.slice(5,7)};
      if(!byDate[r.fullDate][r.distance]||r.time<byDate[r.fullDate][r.distance]) byDate[r.fullDate][r.distance]=r.time;
    });
    return Object.values(byDate);
  };

  // ── Shared styles ─────────────────────────────────────────────────────────
  const panel={background:F.panel,border:`1px solid ${F.panelBorder}`,borderRadius:4,padding:18,marginBottom:14};
  const sel={background:F.ghost,border:`1px solid ${F.panelBorder}`,borderRadius:3,color:F.white,padding:"6px 10px",fontSize:12,outline:"none",fontFamily:F.fontMono,cursor:"pointer"};
  const lbl={fontSize:9,color:F.dim,letterSpacing:2,marginBottom:4,display:"block",textTransform:"uppercase",fontFamily:F.fontMono};
  const inp={...sel,width:"100%",boxSizing:"border-box"};

  // ── Import View ───────────────────────────────────────────────────────────
  const ImportView=()=>(
    <div style={{maxWidth:800,margin:"0 auto"}}>
      <div style={{marginBottom:20,paddingBottom:16,borderBottom:`1px solid ${F.panelBorder}`}}>
        <div style={{fontFamily:F.fontMono,fontSize:11,color:F.red,letterSpacing:3,marginBottom:6}}>SISTEMA DE IMPORTACIÓN</div>
        <div style={{fontFamily:F.fontMono,fontSize:13,color:F.silver}}>Subí archivos desde cualquier app de evaluación</div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:16}}>
        {/* MyJump */}
        {[
          {key:"jump",label:"MYJUMP LAB",sub:"Salto · CMJ / SJ / DJ / RSI",ref:jumpRef,onChange:handleJumpImport,accent:F.red,ext:".csv,.txt",details:["CMJ, SJ, DJ","RSI mod","Fuerza / Potencia","Estado de forma"]},
          {key:"vel",label:"PHOTOFINISH",sub:"Velocidad · 10m / 30m / 60m",ref:velRef,onChange:handleVelImport,accent:F.teal,ext:".csv,.txt",details:["Tiempos por distancia","Detección automática","Múltiples atletas"]},
          {key:"gen",label:"EXCEL / CSV",sub:"Cualquier planilla de datos",ref:genericRef,onChange:handleGenericImport,accent:F.yellow,ext:".xlsx,.xls,.csv",details:["Cualquier app","Mapeás columnas vos","Múltiples métricas"]},
        ].map(({key,label,sub,ref:fref,onChange,accent,ext,details})=>(
          <div key={key} style={{...panel,cursor:"pointer",borderTop:`2px solid ${accent}`,transition:"border-color .15s"}}
            onClick={()=>fref.current?.click()}>
            <div style={{fontFamily:F.fontMono,fontSize:10,color:accent,letterSpacing:2,marginBottom:4}}>{label}</div>
            <div style={{fontFamily:F.fontMono,fontSize:11,color:F.silver,marginBottom:12}}>{sub}</div>
            <div style={{border:`1px dashed ${F.ghost}`,borderRadius:3,padding:"16px 10px",textAlign:"center",marginBottom:10,background:F.asphalt}}>
              <div style={{fontSize:20,marginBottom:4,color:accent}}>↑</div>
              <div style={{fontFamily:F.fontMono,fontSize:10,color:F.dim,letterSpacing:1}}>SUBIR ARCHIVO</div>
              <input ref={fref} type="file" accept={ext} style={{display:"none"}} onChange={onChange}/>
            </div>
            {details.map(d=>(
              <div key={d} style={{fontSize:10,color:F.dim,fontFamily:F.fontMono,padding:"2px 0",borderBottom:`1px solid ${F.ghost}`}}>
                <span style={{color:accent,marginRight:5}}>—</span>{d}
              </div>
            ))}
            {msgs[key]&&(
              <div style={{marginTop:8,fontSize:10,color:msgs[key].startsWith("✓")?F.green:F.red,fontFamily:F.fontMono,padding:"4px 8px",background:msgs[key].startsWith("✓")?F.green+"14":F.red+"14",borderRadius:3}}>
                {msgs[key]}
              </div>
            )}
          </div>
        ))}
      </div>

      {/* Column mapper */}
      {showMapper&&genericData&&(
        <div style={{...panel,borderTop:`2px solid ${F.yellow}`}}>
          <PanelHeader accent={F.yellow}>MAPEO DE COLUMNAS — {genericData.length} FILAS DETECTADAS</PanelHeader>
          <div style={{marginBottom:10,fontFamily:F.fontMono,fontSize:10,color:F.dim}}>
            Columnas disponibles: {genericCols.join(" · ")}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:14}}>
            {[
              {k:"name",   label:"Columna NOMBRE (obligatorio)"},
              {k:"date",   label:"Columna FECHA"},
              {k:"team",   label:"Columna EQUIPO"},
              {k:"metric1",label:"Métrica 1 — columna"},
              {k:"metric1label",label:"Métrica 1 — nombre"},
              {k:"metric2",label:"Métrica 2 — columna"},
              {k:"metric2label",label:"Métrica 2 — nombre"},
              {k:"metric3",label:"Métrica 3 — columna"},
              {k:"metric3label",label:"Métrica 3 — nombre"},
            ].map(({k,label})=>(
              <div key={k}>
                <label style={lbl}>{label}</label>
                {k.includes("label")?
                  <input style={inp} placeholder="ej: CMJ, RPE, 30m..." value={colMap[k]||""} onChange={e=>setColMap(m=>({...m,[k]:e.target.value}))}/>
                  :
                  <select style={sel} value={colMap[k]||""} onChange={e=>setColMap(m=>({...m,[k]:e.target.value}))}>
                    <option value="">— ninguna —</option>
                    {genericCols.map(c=><option key={c} value={c}>{c}</option>)}
                  </select>
                }
              </div>
            ))}
          </div>
          {/* Preview */}
          <div style={{marginBottom:14,overflowX:"auto"}}>
            <div style={{fontFamily:F.fontMono,fontSize:9,color:F.dim,letterSpacing:2,marginBottom:6}}>PREVIEW — 3 FILAS</div>
            <table style={{borderCollapse:"collapse",fontSize:10,fontFamily:F.fontMono}}>
              <thead><tr>{genericCols.map(c=><th key={c} style={{padding:"4px 10px",color:F.red,borderBottom:`1px solid ${F.panelBorder}`,whiteSpace:"nowrap",letterSpacing:1}}>{c}</th>)}</tr></thead>
              <tbody>{genericData.slice(0,3).map((r,i)=>(
                <tr key={i} style={{borderBottom:`1px solid ${F.ghost}`}}>{genericCols.map(c=><td key={c} style={{padding:"4px 10px",color:F.silver,whiteSpace:"nowrap"}}>{String(r[c]).slice(0,20)}</td>)}</tr>
              ))}</tbody>
            </table>
          </div>
          <button onClick={applyMapping} style={{padding:"8px 24px",background:F.yellow,color:F.carbon,border:"none",borderRadius:3,fontFamily:F.fontMono,fontSize:11,fontWeight:700,letterSpacing:2,cursor:"pointer"}}>
            IMPORTAR CON ESTE MAPEO
          </button>
        </div>
      )}

      {hasData&&(
        <div style={{...panel,borderTop:`2px solid ${F.green}`,display:"flex",alignItems:"center",gap:20,flexWrap:"wrap"}}>
          <div style={{fontFamily:F.fontMono,fontSize:9,color:F.green,letterSpacing:3}}>DATOS EN SISTEMA</div>
          {[{v:players.length,l:"ATLETAS"},{v:jumpRecs.length,l:"SALTOS"},{v:velRecs.length,l:"VEL."},{v:customRecs.length,l:"PERSONALIZADOS"}].map(({v,l})=>(
            <div key={l} style={{textAlign:"center"}}>
              <div style={{fontFamily:F.fontMono,fontSize:18,color:F.white,fontWeight:700}}>{v}</div>
              <div style={{fontFamily:F.fontMono,fontSize:9,color:F.dim,letterSpacing:2}}>{l}</div>
            </div>
          ))}
          <button onClick={()=>setTab("dashboard")} style={{marginLeft:"auto",padding:"8px 20px",background:F.red,color:F.white,border:"none",borderRadius:3,fontFamily:F.fontMono,fontSize:11,letterSpacing:2,cursor:"pointer",fontWeight:700}}>
            VER DASHBOARD →
          </button>
        </div>
      )}
    </div>
  );

  // ── Dashboard ─────────────────────────────────────────────────────────────
  const Dashboard=()=>{
    const vis=visible.filter(p=>(pJR(p.id).length+pVR(p.id).length)>0);
    return(
      <div>
        <div style={{display:"flex",gap:12,marginBottom:16,alignItems:"flex-end",flexWrap:"wrap"}}>
          <div><label style={lbl}>EQUIPO</label>
            <select style={sel} value={filterTeam} onChange={e=>setFilterTeam(e.target.value)}>
              {teams.map(t=><option key={t}>{t}</option>)}
            </select>
          </div>
          <div style={{marginLeft:"auto",display:"flex",gap:8}}>
            {[{v:vis.length,l:"ATLETAS",c:F.red},{v:jumpRecs.length,l:"SALTOS",c:F.teal},{v:velRecs.length,l:"VEL",c:F.yellow}].map(({v,l,c})=>(
              <div key={l} style={{...panel,marginBottom:0,padding:"8px 14px",textAlign:"center",borderTop:`2px solid ${c}`}}>
                <div style={{fontFamily:F.fontMono,fontSize:20,color:F.white,fontWeight:700}}>{v}</div>
                <div style={{fontFamily:F.fontMono,fontSize:9,color:F.dim,letterSpacing:2}}>{l}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Timing-board style ranking */}
        <div style={panel}>
          <PanelHeader>CLASIFICACIÓN GENERAL</PanelHeader>
          <div>
            {rankingData.map((d,i)=>{
              const lv=getLevel(d.score); const col=TEAM_COLORS[i%TEAM_COLORS.length];
              return(
                <div key={d.name} style={{display:"flex",alignItems:"center",gap:12,padding:"8px 0",borderBottom:`1px solid ${F.ghost}`}}>
                  <div style={{fontFamily:F.fontMono,fontSize:11,color:F.dim,width:24,textAlign:"right"}}>{String(i+1).padStart(2,"0")}</div>
                  <div style={{width:4,height:32,background:col,borderRadius:1}}/>
                  <div style={{fontFamily:F.fontMono,fontSize:13,color:F.white,flex:1,letterSpacing:1}}>{d.fullName||d.name}</div>
                  <SectorBadge score={d.score}/>
                  <div style={{width:120,height:4,background:F.ghost,borderRadius:2,overflow:"hidden"}}>
                    <div style={{width:`${d.score}%`,height:"100%",background:lv.color,borderRadius:2,transition:"width .8s ease",boxShadow:`0 0 6px ${lv.color}`}}/>
                  </div>
                  <div style={{fontFamily:F.fontMono,fontSize:14,color:lv.color,width:36,textAlign:"right",fontWeight:700}}>{d.score}</div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Player cards */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:10,marginBottom:14}}>
          {vis.map((p,i)=>{
            const sc=overallScore(p.id); const bj=bestJump(p.id); const lj=lastJump(p.id);
            const bv=allDist.length?bestVel(p.id,allDist[0]):null; const col=TEAM_COLORS[i%TEAM_COLORS.length];
            return(
              <div key={p.id} style={{...panel,marginBottom:0,cursor:"pointer",borderLeft:`3px solid ${col}`,transition:"background .15s"}}
                onClick={()=>{setSelPlayer(p.id);setTab("evolution");}}
                onMouseEnter={e=>e.currentTarget.style.background=F.ghost}
                onMouseLeave={e=>e.currentTarget.style.background=F.panel}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
                  <div>
                    <div style={{fontFamily:F.fontMono,fontSize:13,color:F.white,letterSpacing:1}}>{p.name.split(" ")[0].toUpperCase()}</div>
                    <div style={{fontFamily:F.fontMono,fontSize:9,color:F.dim,letterSpacing:1,marginTop:2}}>{p.name.split(" ").slice(1).join(" ").toUpperCase()}</div>
                    <div style={{fontFamily:F.fontMono,fontSize:9,color:col,letterSpacing:1,marginTop:3}}>{p.team||"—"}</div>
                  </div>
                  <TimingScore score={sc} size={52}/>
                </div>
                <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                  {bj&&<div style={{background:F.ghost,borderRadius:2,padding:"3px 7px",fontFamily:F.fontMono,fontSize:10}}>
                    <span style={{color:F.dim}}>{bj.jumpTypeRaw} </span>
                    <span style={{color:F.teal}}>{bj.altura?.toFixed(1)}cm</span>
                  </div>}
                  {bv&&<div style={{background:F.ghost,borderRadius:2,padding:"3px 7px",fontFamily:F.fontMono,fontSize:10}}>
                    <span style={{color:F.dim}}>{allDist[0]} </span>
                    <span style={{color:F.yellow}}>{bv.time?.toFixed(2)}s</span>
                  </div>}
                </div>
              </div>
            );
          })}
        </div>

        {allDist.length>0&&(
          <div style={panel}>
            <PanelHeader accent={F.yellow}>MEJORES TIEMPOS POR DISTANCIA</PanelHeader>
            <ResponsiveContainer width="100%" height={Math.max(140,vis.length*24)}>
              <BarChart data={vis.filter(p=>pVR(p.id).length>0).map(p=>({name:p.name.split(" ")[0],...Object.fromEntries(allDist.map(d=>[d,bestVel(p.id,d)?.time??null]))}))} layout="vertical">
                <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} horizontal={false}/>
                <XAxis type="number" tick={{fill:F.dim,fontSize:10,fontFamily:F.fontMono}} unit="s"/>
                <YAxis dataKey="name" type="category" tick={{fill:F.silver,fontSize:10,fontFamily:F.fontMono}} width={70}/>
                <Tooltip content={<CustomTooltip/>}/>
                <Legend wrapperStyle={{fontSize:10,fontFamily:F.fontMono}}/>
                {allDist.map((d,i)=><Bar key={d} dataKey={d} fill={TEAM_COLORS[i]} radius={[0,3,3,0]}/>)}
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>
    );
  };

  // ── Evolution ─────────────────────────────────────────────────────────────
  const Evolution=()=>{
    const pid=selPlayer||players[0]?.id;
    const jevo=jumpEvo(pid); const vevo=velEvo(pid);
    const bj=bestJump(pid); const jrecs=pJR(pid); const vrecs=pVR(pid);
    const p=players.find(x=>x.id===pid);
    return(
      <div>
        <div style={{display:"flex",gap:12,marginBottom:16,alignItems:"flex-end",flexWrap:"wrap"}}>
          <div><label style={lbl}>ATLETA</label>
            <select style={sel} value={pid} onChange={e=>setSelPlayer(parseInt(e.target.value))}>
              {players.filter(p=>(pJR(p.id).length+pVR(p.id).length)>0).map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          </div>
          {p&&<div style={{fontFamily:F.fontMono,fontSize:9,color:F.dim,letterSpacing:2,marginLeft:8}}>{p.team||""}</div>}
        </div>

        <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
          {[
            {v:jrecs.length+vrecs.length,l:"REGISTROS",c:F.silver},
            {v:bj?.altura?.toFixed(1)+"cm"??"—",l:"MEJOR SALTO",c:F.teal},
            {v:bj?.rsi?.toFixed(2)??"—",l:"MEJOR RSI",c:F.yellow},
            ...allDist.map(d=>({v:bestVel(pid,d)?.time?.toFixed(2)+"s"??"—",l:d,c:F.red})),
          ].map(({v,l,c})=>(
            <div key={l} style={{...panel,marginBottom:0,padding:"8px 14px",flex:1,minWidth:80,borderTop:`2px solid ${c}`}}>
              <div style={{fontFamily:F.fontMono,fontSize:16,color:F.white,fontWeight:700}}>{v}</div>
              <div style={{fontFamily:F.fontMono,fontSize:9,color:F.dim,letterSpacing:2}}>{l}</div>
            </div>
          ))}
        </div>

        {jevo.length>0&&(
          <div style={panel}>
            <PanelHeader>ALTURA DE SALTO (cm)</PanelHeader>
            <ResponsiveContainer width="100%" height={200}>
              <LineChart data={jevo}>
                <CartesianGrid strokeDasharray="3 3" stroke={F.ghost}/>
                <XAxis dataKey="fecha" tick={{fill:F.dim,fontSize:10,fontFamily:F.fontMono}}/>
                <YAxis tick={{fill:F.dim,fontSize:10,fontFamily:F.fontMono}} unit="cm"/>
                <Tooltip content={<CustomTooltip/>}/>
                <Legend wrapperStyle={{fontSize:10,fontFamily:F.fontMono}}/>
                <Line dataKey="CMJ" stroke={F.teal} strokeWidth={2} dot={{r:4,fill:F.teal}} connectNulls style={{filter:`drop-shadow(0 0 4px ${F.teal})`}}/>
                <Line dataKey="DJ" stroke={F.red} strokeWidth={2} dot={{r:4,fill:F.red}} connectNulls/>
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}

        {vevo.length>0&&allDist.length>0&&(
          <div style={panel}>
            <PanelHeader accent={F.yellow}>VELOCIDAD (s)</PanelHeader>
            <ResponsiveContainer width="100%" height={200}>
              <LineChart data={vevo}>
                <CartesianGrid strokeDasharray="3 3" stroke={F.ghost}/>
                <XAxis dataKey="fecha" tick={{fill:F.dim,fontSize:10,fontFamily:F.fontMono}}/>
                <YAxis tick={{fill:F.dim,fontSize:10,fontFamily:F.fontMono}} unit="s"/>
                <Tooltip content={<CustomTooltip/>}/>
                <Legend wrapperStyle={{fontSize:10,fontFamily:F.fontMono}}/>
                {allDist.map((d,i)=><Line key={d} dataKey={d} stroke={TEAM_COLORS[i]} strokeWidth={2} dot={{r:4}} connectNulls/>)}
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}

        {jevo.some(d=>d.RSI)&&(
          <div style={panel}>
            <PanelHeader accent={F.yellow}>RSI MOD</PanelHeader>
            <ResponsiveContainer width="100%" height={160}>
              <LineChart data={jevo}>
                <CartesianGrid strokeDasharray="3 3" stroke={F.ghost}/>
                <XAxis dataKey="fecha" tick={{fill:F.dim,fontSize:10,fontFamily:F.fontMono}}/>
                <YAxis tick={{fill:F.dim,fontSize:10,fontFamily:F.fontMono}}/>
                <Tooltip content={<CustomTooltip/>}/>
                <Line dataKey="RSI" stroke={F.yellow} strokeWidth={2} dot={{r:4,fill:F.yellow}} connectNulls style={{filter:`drop-shadow(0 0 4px ${F.yellow})`}}/>
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}

        {vrecs.length>0&&(
          <div style={panel}>
            <PanelHeader accent={F.yellow}>HISTORIAL VELOCIDAD</PanelHeader>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                <thead><tr style={{borderBottom:`1px solid ${F.panelBorder}`}}>
                  {["FECHA","TEST","DIST","TIEMPO","NIVEL"].map(h=>(
                    <th key={h} style={{padding:"5px 10px",color:F.dim,textAlign:"left",fontFamily:F.fontMono,fontSize:9,letterSpacing:2}}>{h}</th>
                  ))}
                </tr></thead>
                <tbody>{vrecs.map((r,i)=>{
                  const ref=VEL_REFS[r.distance]; const sc=scoreVal(ref,r.time);
                  return<tr key={i} style={{borderBottom:`1px solid ${F.ghost}`}}>
                    <td style={{padding:"6px 10px",color:F.teal,fontFamily:F.fontMono,fontSize:11}}>{r.fullDate}</td>
                    <td style={{padding:"6px 10px",color:F.silver,fontFamily:F.fontMono,fontSize:11}}>{r.testName}</td>
                    <td style={{padding:"6px 10px",color:F.silver,fontFamily:F.fontMono,fontSize:11}}>{r.distance||"—"}</td>
                    <td style={{padding:"6px 10px",fontFamily:F.fontMono,fontSize:13,color:F.white,fontWeight:700}}>{r.time?.toFixed(2)}s</td>
                    <td style={{padding:"6px 10px"}}><SectorBadge score={sc??0}/></td>
                  </tr>;
                })}</tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    );
  };

  // ── Compare ───────────────────────────────────────────────────────────────
  const Compare=()=>{
    const pidA=compareA||players[0]?.id; const pidB=compareB||players[1]?.id;
    const nA=players.find(p=>p.id===pidA)?.name.split(" ")[0]||"A";
    const nB=players.find(p=>p.id===pidB)?.name.split(" ")[0]||"B";
    const radarM=[
      {key:"cmj",label:"CMJ",vA:bestJump(pidA)?.altura,vB:bestJump(pidB)?.altura,ref:JUMP_REFS.cmj},
      {key:"rsi",label:"RSI",vA:bestJump(pidA)?.rsi,vB:bestJump(pidB)?.rsi,ref:JUMP_REFS.rsi},
      ...allDist.map(d=>({key:d,label:d,vA:bestVel(pidA,d)?.time,vB:bestVel(pidB,d)?.time,ref:VEL_REFS[d]})),
    ];
    const radarD=radarM.map(m=>({metric:m.label,[nA]:scoreVal(m.ref,m.vA)??0,[nB]:scoreVal(m.ref,m.vB)??0}));
    const evoA=jumpEvo(pidA); const evoB=jumpEvo(pidB);
    const allF=Array.from(new Set([...evoA.map(d=>d.fecha),...evoB.map(d=>d.fecha)])).sort();
    const merged=allF.map(f=>({fecha:f,[nA]:evoA.find(d=>d.fecha===f)?.CMJ??null,[nB]:evoB.find(d=>d.fecha===f)?.CMJ??null}));
    const scA=overallScore(pidA); const scB=overallScore(pidB);
    return(
      <div>
        <div style={{display:"flex",gap:12,marginBottom:16,flexWrap:"wrap"}}>
          {[{val:pidA,set:setCompareA,label:"ATLETA A",col:F.teal},{val:pidB,set:setCompareB,label:"ATLETA B",col:F.red}].map(({val,set,label,col})=>(
            <div key={label} style={{flex:1,minWidth:160}}>
              <label style={{...lbl,color:col}}>{label}</label>
              <select style={sel} value={val||""} onChange={e=>set(parseInt(e.target.value))}>
                {players.filter(p=>(pJR(p.id).length+pVR(p.id).length)>0).map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
            </div>
          ))}
        </div>

        {/* Head to head scores */}
        <div style={{...panel,display:"flex",alignItems:"center",gap:0,marginBottom:14,padding:0,overflow:"hidden"}}>
          {[{pid:pidA,sc:scA,col:F.teal,name:nA},{pid:pidB,sc:scB,col:F.red,name:nB}].map(({pid,sc,col,name},i)=>(
            <div key={pid} style={{flex:1,padding:"16px 20px",borderRight:i===0?`1px solid ${F.panelBorder}`:"none",textAlign:"center"}}>
              <div style={{fontFamily:F.fontMono,fontSize:10,color:col,letterSpacing:3,marginBottom:10}}>{name.toUpperCase()}</div>
              <TimingScore score={sc} size={72}/>
              <div style={{marginTop:8}}><SectorBadge score={sc}/></div>
            </div>
          ))}
          <div style={{padding:"16px 24px",textAlign:"center"}}>
            <div style={{fontFamily:F.fontMono,fontSize:9,color:F.dim,letterSpacing:2,marginBottom:6}}>DIFERENCIA</div>
            <div style={{fontFamily:F.fontMono,fontSize:28,color:Math.abs(scA-scB)<5?F.silver:scA>scB?F.teal:F.red,fontWeight:700}}>
              {scA===scB?"=":`${Math.abs(scA-scB)} pts`}
            </div>
          </div>
        </div>

        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          <div style={panel}>
            <PanelHeader>RADAR COMPARATIVO</PanelHeader>
            <ResponsiveContainer width="100%" height={260}>
              <RadarChart data={radarD}>
                <PolarGrid stroke={F.ghost}/>
                <PolarAngleAxis dataKey="metric" tick={{fill:F.silver,fontSize:10,fontFamily:F.fontMono}}/>
                <PolarRadiusAxis domain={[0,100]} tick={{fill:F.dim,fontSize:9,fontFamily:F.fontMono}}/>
                <Radar name={nA} dataKey={nA} stroke={F.teal} fill={F.teal} fillOpacity={0.15} strokeWidth={2} style={{filter:`drop-shadow(0 0 3px ${F.teal})`}}/>
                <Radar name={nB} dataKey={nB} stroke={F.red} fill={F.red} fillOpacity={0.15} strokeWidth={2} style={{filter:`drop-shadow(0 0 3px ${F.red})`}}/>
                <Legend wrapperStyle={{fontSize:10,fontFamily:F.fontMono}}/>
                <Tooltip content={<CustomTooltip/>}/>
              </RadarChart>
            </ResponsiveContainer>
          </div>

          <div style={panel}>
            <PanelHeader>MÉTRICAS VS ÉLITE</PanelHeader>
            <div style={{display:"flex",flexDirection:"column",gap:6}}>
              {radarM.map(({key,label,vA,vB,ref})=>(
                <div key={key} style={{background:F.asphalt,borderRadius:3,padding:"8px 12px",border:`1px solid ${F.ghost}`}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                    <span style={{fontFamily:F.fontMono,fontSize:10,color:F.silver,letterSpacing:1}}>{label}</span>
                    {ref&&<span style={{fontSize:9,color:F.dim,fontFamily:F.fontMono}}>ÉLITE: {ref.excellent}{ref.unit}</span>}
                  </div>
                  <div style={{display:"flex",gap:6}}>
                    {[{name:nA,val:vA,col:F.teal},{name:nB,val:vB,col:F.red}].map(({name,val,col})=>{
                      const sc=scoreVal(ref,val); const lv=getLevel(sc??0);
                      return(
                        <div key={name} style={{flex:1,textAlign:"center",background:F.panel,borderRadius:2,padding:"5px 4px",borderTop:`1px solid ${col}`}}>
                          <div style={{fontSize:9,color:col,marginBottom:2,fontFamily:F.fontMono,letterSpacing:1}}>{name}</div>
                          <div style={{fontSize:14,fontWeight:700,fontFamily:F.fontMono,color:val!=null?lv.color:F.ghost}}>
                            {val!=null?`${val>10?val.toFixed(1):val.toFixed(2)}${ref?.unit||""}`:  "—"}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>

        {merged.some(d=>d[nA]||d[nB])&&(
          <div style={{...panel,marginTop:12}}>
            <PanelHeader>EVOLUCIÓN CMJ COMPARADA</PanelHeader>
            <ResponsiveContainer width="100%" height={200}>
              <LineChart data={merged}>
                <CartesianGrid strokeDasharray="3 3" stroke={F.ghost}/>
                <XAxis dataKey="fecha" tick={{fill:F.dim,fontSize:10,fontFamily:F.fontMono}}/>
                <YAxis tick={{fill:F.dim,fontSize:10,fontFamily:F.fontMono}} unit="cm"/>
                <Tooltip content={<CustomTooltip/>}/>
                <Legend wrapperStyle={{fontSize:10,fontFamily:F.fontMono}}/>
                <Line dataKey={nA} stroke={F.teal} strokeWidth={2} dot={{r:4}} connectNulls style={{filter:`drop-shadow(0 0 4px ${F.teal})`}}/>
                <Line dataKey={nB} stroke={F.red} strokeWidth={2} dot={{r:4}} connectNulls style={{filter:`drop-shadow(0 0 4px ${F.red})`}}/>
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>
    );
  };

  // ── Shell ─────────────────────────────────────────────────────────────────
  return(
    <div style={{minHeight:"100vh",background:F.carbon,fontFamily:F.fontF1,color:F.white}}>
      <link href="https://fonts.googleapis.com/css2?family=Titillium+Web:wght@300;400;600;700&family=Share+Tech+Mono&display=swap" rel="stylesheet"/>

      {/* Header */}
      <div style={{background:F.asphalt,padding:"0 24px",display:"flex",alignItems:"center",height:52,borderBottom:`1px solid ${F.panelBorder}`}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <div style={{width:4,height:28,background:F.red}}/>
          <div>
            <div style={{fontFamily:F.fontMono,fontSize:16,color:F.white,letterSpacing:3,lineHeight:1}}>ATHLETE<span style={{color:F.red}}>IQ</span></div>
            <div style={{fontFamily:F.fontMono,fontSize:8,color:F.dim,letterSpacing:3}}>PERFORMANCE SYSTEM</div>
          </div>
        </div>
        <nav style={{marginLeft:32,display:"flex",gap:2}}>
          {[
            {key:"import",label:"IMPORTAR"},
            {key:"dashboard",label:"DASHBOARD",d:!hasData},
            {key:"evolution",label:"EVOLUCIÓN",d:!hasData},
            {key:"compare",label:"COMPARAR",d:!hasData},
          ].map(({key,label,d})=>(
            <button key={key} onClick={()=>!d&&setTab(key)} style={{
              padding:"0 16px",height:52,border:"none",borderBottom:tab===key?`2px solid ${F.red}`:"2px solid transparent",
              cursor:d?"default":"pointer",fontFamily:F.fontMono,fontSize:10,letterSpacing:2,
              background:"transparent",color:tab===key?F.white:d?F.ghost:F.dim,
              transition:"all .15s",
            }}>{label}</button>
          ))}
        </nav>
        {hasData&&<div style={{marginLeft:"auto",fontFamily:F.fontMono,fontSize:9,color:F.dim,letterSpacing:2}}>{players.length} ATL · {jumpRecs.length+velRecs.length+customRecs.length} REG</div>}
      </div>

      {/* Telemetry bar — signature element */}
      <TelemetryBar/>

      {/* Content */}
      <div style={{padding:"20px 24px",maxWidth:1160,margin:"0 auto"}}>
        {tab==="import"    && <ImportView/>}
        {tab==="dashboard" && hasData && <Dashboard/>}
        {tab==="evolution" && hasData && <Evolution/>}
        {tab==="compare"   && hasData && <Compare/>}
      </div>
    </div>
  );
}
