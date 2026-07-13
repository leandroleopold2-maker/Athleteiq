import { useState, useMemo, useRef } from "react";
import {
  LineChart, Line, BarChart, Bar, RadarChart, Radar, PolarGrid,
  PolarAngleAxis, PolarRadiusAxis, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, ScatterChart, Scatter, ZAxis
} from "recharts";

// ── References ────────────────────────────────────────────────────────────────
const JUMP_REFS = {
  cmj:      { excellent:42, good:35, average:28, unit:"cm", label:"CMJ",      lowerIsBetter:false },
  sj:       { excellent:38, good:31, average:24, unit:"cm", label:"SJ",       lowerIsBetter:false },
  dj:       { excellent:38, good:30, average:22, unit:"cm", label:"DJ",       lowerIsBetter:false },
  rsi:      { excellent:2.0,good:1.5,average:1.0,unit:"",  label:"RSI mod",  lowerIsBetter:false },
  fuerza:   { excellent:3500,good:2800,average:2200,unit:"N",label:"Fuerza", lowerIsBetter:false },
  potencia: { excellent:5000,good:3500,average:2500,unit:"W",label:"Potencia",lowerIsBetter:false },
};
const VEL_REFS = {
  "10m": { excellent:1.65, good:1.75, average:1.85, unit:"s", label:"10m", lowerIsBetter:true },
  "20m": { excellent:2.80, good:2.95, average:3.10, unit:"s", label:"20m", lowerIsBetter:true },
  "30m": { excellent:3.90, good:4.10, average:4.30, unit:"s", label:"30m", lowerIsBetter:true },
  "40m": { excellent:4.90, good:5.15, average:5.40, unit:"s", label:"40m", lowerIsBetter:true },
  "60m": { excellent:6.80, good:7.10, average:7.40, unit:"s", label:"60m", lowerIsBetter:true },
};

function score(ref, value) {
  if (!ref || value == null) return null;
  const { excellent, good, average, lowerIsBetter } = ref;
  if (lowerIsBetter) {
    if (value <= excellent) return 100;
    if (value <= good)      return 75;
    if (value <= average)   return 50;
    return 25;
  } else {
    if (value >= excellent) return 100;
    if (value >= good)      return 75;
    if (value >= average)   return 50;
    return 25;
  }
}

function getLevel(s) {
  if (s >= 90) return { label:"Elite",     color:"#00e5a0" };
  if (s >= 70) return { label:"Bueno",     color:"#7dd4f0" };
  if (s >= 50) return { label:"Promedio",  color:"#f0c060" };
  return           { label:"A mejorar", color:"#f07060" };
}

function getStatusColor(c) {
  if (c==="green")  return "#00e5a0";
  if (c==="red")    return "#f07060";
  if (c==="yellow") return "#f0c060";
  return "#2a4a5a";
}

function parseNum(str) {
  if (!str || str==="---"||str==="") return null;
  return parseFloat(str.replace(",",".").replace(/[^\d.-]/g,"")) || null;
}

function parseMyJumpDate(str) {
  const m = str.match(/(\d+)\.(\d+)\.(\d+)/);
  if (!m) return str;
  return `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
}

function parsePhotoDate(str) {
  // "23/4/2026" → "2026-04-23"
  const m = str.match(/(\d+)\/(\d+)\/(\d+)/);
  if (!m) return str;
  return `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
}

function jumpTypeKey(tipo) {
  if (!tipo) return "cmj";
  const t = tipo.toLowerCase();
  if (t.includes("cmj")) return "cmj";
  if (t.includes("sj")&&!t.includes("rsi")) return "sj";
  if (t.includes("dj")||t.includes("drop")) return "dj";
  if (t.includes("aba")) return "abalakov";
  return "cmj";
}

// Detect distance from test name like "30 m Test", "10m", "30M", etc.
function detectDistance(info) {
  if (!info) return null;
  const m = info.match(/(\d+)\s*m/i);
  return m ? m[1]+"m" : null;
}

// ── CSV parsers ───────────────────────────────────────────────────────────────
function parseMyJumpCSV(text) {
  const lines = text.trim().split("\n");
  if (lines.length < 2) return [];
  const headers = lines[0].split(";").map(h=>h.trim());
  return lines.slice(1).map(line => {
    const vals = line.split(";");
    const obj = {};
    headers.forEach((h,i)=>{ obj[h]=(vals[i]||"").trim(); });
    return obj;
  });
}

function parsePhotoFinishCSV(text) {
  const lines = text.trim().split("\n");
  if (lines.length < 2) return [];
  // skip header row
  const dataLines = lines.slice(1);
  let currentDate = "";
  let currentInfo = "";
  const rows = [];
  dataLines.forEach(line => {
    const cols = line.split(",").map(c=>c.trim());
    if (cols[0]) currentDate = cols[0];
    if (cols[1]) currentInfo = cols[1];
    const num     = cols[2]||"";
    const athlete = cols[3]||"";
    const time    = cols[5]||"";
    if (!athlete) return;
    rows.push({ date: currentDate, info: currentInfo, num, athlete, time });
  });
  return rows;
}

const COLORS = ["#00e5a0","#7dd4f0","#f0c060","#f07060","#c084fc","#fb923c","#34d399","#60a5fa","#fbbf24","#f87171"];

const CustomTooltip = ({ active, payload, label }) => {
  if (!active||!payload?.length) return null;
  return (
    <div style={{ background:"#0f1923", border:"1px solid #1e3a4a", borderRadius:8, padding:"10px 14px" }}>
      <p style={{ color:"#7dd4f0", fontFamily:"Barlow Condensed,sans-serif", fontSize:13, marginBottom:4 }}>{label}</p>
      {payload.map((p,i)=>(
        <p key={i} style={{ color:p.color, fontSize:12, margin:"2px 0" }}>
          {p.name}: <strong>{typeof p.value==="number"?p.value.toFixed(2):p.value}</strong>
        </p>
      ))}
    </div>
  );
};

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [players,  setPlayers]  = useState([]);
  const [jumpRecs, setJumpRecs] = useState([]);
  const [velRecs,  setVelRecs]  = useState([]);
  const [tab,      setTab]      = useState("import");
  const [selPlayer,  setSelPlayer]  = useState(null);
  const [compareA,   setCompareA]   = useState(null);
  const [compareB,   setCompareB]   = useState(null);
  const [jumpMsg,    setJumpMsg]    = useState("");
  const [velMsg,     setVelMsg]     = useState("");
  const [filterTeam, setFilterTeam] = useState("Todos");
  const jumpRef = useRef();
  const velRef  = useRef();

  let _idC = 1;
  const uid = () => _idC++;

  // ── Merge player by name ──────────────────────────────────────────────────
  function mergePlayer(name, team, currentPlayers) {
    const key = name.toLowerCase().trim();
    const existing = currentPlayers.find(p => p.name.toLowerCase().trim() === key);
    if (existing) return { players: currentPlayers, id: existing.id };
    const newP = { id: Date.now() + Math.random(), name: name.trim(), team: team||"" };
    return { players: [...currentPlayers, newP], id: newP.id };
  }

  // ── Import MyJump ─────────────────────────────────────────────────────────
  function handleJumpImport(e) {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const rows = parseMyJumpCSV(ev.target.result);
      if (!rows.length) { setJumpMsg("❌ No se pudo leer el archivo."); return; }
      let curPlayers = [...players];
      const newRecs = [];
      rows.forEach(row => {
        const name = (row["Nombre"]||"").trim();
        const team = (row["Equipo"]||"").trim();
        if (!name) return;
        const { players: updated, id: pid } = mergePlayer(name, team, curPlayers);
        curPlayers = updated;
        const tipo  = (row["Tipo de salto"]||"").trim();
        const fecha = parseMyJumpDate(row["Fecha"]||"");
        newRecs.push({
          id: Date.now()+Math.random(), playerId: pid,
          date: fecha.slice(0,7), fullDate: fecha,
          jumpType: jumpTypeKey(tipo), jumpTypeRaw: tipo,
          altura:    parseNum(row["Altura de salto (cm)"]),
          rsi:       parseNum(row["RSI mod (m/s)"]),
          fuerza:    parseNum(row["Fuerza (N)"]),
          potencia:  parseNum(row["Potencia (W)"]),
          vuelo:     parseNum(row["Tiempo de vuelo (ms)"]),
          contacto:  parseNum(row["Tiempo de contacto (ms)"]),
          velocidad: parseNum(row["Velocidad (m/s)"]),
          impulse:   parseNum(row["Impulse (N*kg)"]),
          statusColor: row["Color estado de forma"]||"",
          team,
        });
      });
      setPlayers(curPlayers);
      setJumpRecs(r => [...r, ...newRecs]);
      if (!selPlayer && curPlayers.length) { setSelPlayer(curPlayers[0].id); setCompareA(curPlayers[0].id); setCompareB(curPlayers[1]?.id||curPlayers[0].id); }
      setJumpMsg(`✓ ${newRecs.length} registros de salto importados.`);
      setTab("dashboard");
    };
    reader.readAsText(file,"UTF-8");
  }

  // ── Import PhotoFinish ────────────────────────────────────────────────────
  function handleVelImport(e) {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const rows = parsePhotoFinishCSV(ev.target.result);
      if (!rows.length) { setVelMsg("❌ No se pudo leer el archivo."); return; }
      let curPlayers = [...players];
      const newRecs = [];
      rows.forEach(row => {
        const name = row.athlete.trim(); if (!name) return;
        const { players: updated, id: pid } = mergePlayer(name, "", curPlayers);
        curPlayers = updated;
        const fecha = parsePhotoDate(row.date);
        const dist  = detectDistance(row.info);
        const t     = parseNum(row.time);
        if (!t) return;
        newRecs.push({
          id: Date.now()+Math.random(), playerId: pid,
          date: fecha.slice(0,7), fullDate: fecha,
          distance: dist, testName: row.info,
          time: t,
        });
      });
      setPlayers(curPlayers);
      setVelRecs(r => [...r, ...newRecs]);
      if (!selPlayer && curPlayers.length) { setSelPlayer(curPlayers[0].id); setCompareA(curPlayers[0].id); setCompareB(curPlayers[1]?.id||curPlayers[0].id); }
      setVelMsg(`✓ ${newRecs.length} registros de velocidad importados.`);
      setTab("dashboard");
    };
    reader.readAsText(file,"UTF-8");
  }

  const hasData = players.length > 0;
  const teams = useMemo(() => ["Todos", ...Array.from(new Set(players.map(p=>p.team).filter(Boolean)))], [players]);

  const visiblePlayers = useMemo(() =>
    players.filter(p => filterTeam==="Todos" || p.team===filterTeam),
    [players, filterTeam]);

  // ── Helpers ───────────────────────────────────────────────────────────────
  const pJumpRecs = (pid) => jumpRecs.filter(r=>r.playerId===pid).sort((a,b)=>a.fullDate.localeCompare(b.fullDate));
  const pVelRecs  = (pid) => velRecs.filter(r=>r.playerId===pid).sort((a,b)=>a.fullDate.localeCompare(b.fullDate));

  const bestJump  = (pid) => { const rs=pJumpRecs(pid).filter(r=>r.altura!=null); return rs.reduce((b,r)=>r.altura>(b?.altura??0)?r:b,null); };
  const bestVel   = (pid, dist) => { const rs=pVelRecs(pid).filter(r=>r.distance===dist&&r.time!=null); return rs.reduce((b,r)=>r.time<(b?.time??999)?r:b,null); };
  const lastJump  = (pid) => { const rs=pJumpRecs(pid).sort((a,b)=>b.fullDate.localeCompare(a.fullDate)); return rs[0]||null; };

  const allDistances = useMemo(() => Array.from(new Set(velRecs.map(r=>r.distance).filter(Boolean))).sort(), [velRecs]);

  const overallScore = (pid) => {
    const scores = [];
    const bj = bestJump(pid);
    if (bj) {
      scores.push(score(JUMP_REFS[bj.jumpType]||JUMP_REFS.cmj, bj.altura)??0);
      if (bj.rsi!=null) scores.push(score(JUMP_REFS.rsi, bj.rsi)??0);
    }
    allDistances.forEach(d => {
      const bv = bestVel(pid, d);
      if (bv) scores.push(score(VEL_REFS[d], bv.time)??0);
    });
    return scores.length ? Math.round(scores.reduce((a,b)=>a+b,0)/scores.length) : 0;
  };

  const rankingData = useMemo(() =>
    visiblePlayers
      .filter(p=>(pJumpRecs(p.id).length+pVelRecs(p.id).length)>0)
      .map(p=>({ name:p.name.split(" ")[0], fullName:p.name, score:overallScore(p.id) }))
      .sort((a,b)=>b.score-a.score),
    [players, jumpRecs, velRecs, filterTeam]);

  const jumpEvoData = (pid) => {
    const byDate={};
    pJumpRecs(pid).forEach(r=>{
      if(!byDate[r.fullDate]) byDate[r.fullDate]={ fecha:r.fullDate.slice(8)+"/"+r.fullDate.slice(5,7), _r:[] };
      byDate[r.fullDate]._r.push(r);
    });
    return Object.values(byDate).map(d=>{
      const cmjs=d._r.filter(r=>r.jumpType==="cmj"&&r.altura!=null);
      const djs =d._r.filter(r=>r.jumpType==="dj"&&r.altura!=null);
      const rsif=d._r.filter(r=>r.rsi!=null);
      return { fecha:d.fecha, CMJ:cmjs.length?Math.max(...cmjs.map(r=>r.altura)):null, DJ:djs.length?Math.max(...djs.map(r=>r.altura)):null, RSI:rsif.length?Math.max(...rsif.map(r=>r.rsi)):null };
    });
  };

  const velEvoData = (pid) => {
    const byDate={};
    pVelRecs(pid).forEach(r=>{
      if(!byDate[r.fullDate]) byDate[r.fullDate]={ fecha:r.fullDate.slice(8)+"/"+r.fullDate.slice(5,7) };
      // keep best time per distance per date
      if(!byDate[r.fullDate][r.distance]||r.time<byDate[r.fullDate][r.distance])
        byDate[r.fullDate][r.distance]=r.time;
    });
    return Object.values(byDate);
  };

  // ── Styles ────────────────────────────────────────────────────────────────
  const s = {
    app:  { minHeight:"100vh", background:"#080f17", fontFamily:"'Barlow',sans-serif", color:"#d8eaf4" },
    header: { background:"linear-gradient(135deg,#0a1520,#0d2035)", borderBottom:"1px solid #1e3a4a", padding:"16px 24px", display:"flex", alignItems:"center", gap:14 },
    logo: { fontFamily:"'Barlow Condensed',sans-serif", fontWeight:800, fontSize:21, letterSpacing:2, color:"#00e5a0", textTransform:"uppercase" },
    tagline: { fontSize:11, color:"#4a7a94", letterSpacing:1, marginTop:2 },
    nav: { display:"flex", gap:4, padding:"10px 24px", background:"#0a1520", borderBottom:"1px solid #132030", flexWrap:"wrap" },
    navBtn: (a,d) => ({ padding:"7px 16px", borderRadius:6, border:"none", cursor:d?"default":"pointer", fontFamily:"'Barlow Condensed',sans-serif", fontWeight:600, fontSize:13, letterSpacing:1, textTransform:"uppercase", background:a?"#00e5a0":"transparent", color:a?"#080f17":"#4a7a94", opacity:d?0.35:1 }),
    main: { padding:"20px 24px", maxWidth:1140, margin:"0 auto" },
    card: { background:"#0d1e2d", border:"1px solid #1a3045", borderRadius:12, padding:18, marginBottom:18 },
    cardTitle: { fontFamily:"'Barlow Condensed',sans-serif", fontWeight:700, fontSize:15, letterSpacing:1.5, textTransform:"uppercase", color:"#7dd4f0", marginBottom:14 },
    grid2: { display:"grid", gridTemplateColumns:"1fr 1fr", gap:16 },
    badge: (sc) => ({ display:"inline-block", padding:"2px 9px", borderRadius:20, fontSize:11, fontWeight:700, letterSpacing:1, background:getLevel(sc).color+"22", color:getLevel(sc).color, border:`1px solid ${getLevel(sc).color}44` }),
    bigScore: (sc) => ({ fontSize:44, fontWeight:800, fontFamily:"'Barlow Condensed',sans-serif", color:getLevel(sc).color, lineHeight:1 }),
    select: { background:"#0a1824", border:"1px solid #1e3a4a", borderRadius:6, color:"#d8eaf4", padding:"7px 11px", fontSize:13, outline:"none", fontFamily:"'Barlow',sans-serif" },
    label: { fontSize:11, color:"#4a7a94", letterSpacing:1, marginBottom:4, display:"block", textTransform:"uppercase" },
    stat: { background:"#0a1824", borderRadius:8, padding:"10px 14px", textAlign:"center", flex:1, minWidth:90 },
    statVal: { fontSize:20, fontWeight:700, fontFamily:"'Barlow Condensed',sans-serif", color:"#00e5a0" },
    statLbl: { fontSize:10, color:"#4a7a94", marginTop:2, letterSpacing:1 },
    uploadBox: (hover) => ({ border:`2px dashed ${hover?"#00e5a0":"#1e3a4a"}`, borderRadius:10, padding:"28px 20px", textAlign:"center", cursor:"pointer", transition:"border-color .2s" }),
    msgBox: (ok) => ({ padding:"10px 14px", borderRadius:8, background:ok?"#00e5a01a":"#f0706022", border:`1px solid ${ok?"#00e5a044":"#f0706044"}`, fontSize:13, color:ok?"#00e5a0":"#f07060", marginTop:10 }),
  };

  // ── Import View ───────────────────────────────────────────────────────────
  const [hJ, setHJ] = useState(false);
  const [hV, setHV] = useState(false);

  const ImportView = () => (
    <div style={s.grid2}>
      {/* MyJump */}
      <div style={s.card}>
        <div style={s.cardTitle}>🦘 MyJump Lab</div>
        <p style={{ fontSize:13, color:"#7dd4f0", marginBottom:16, lineHeight:1.7 }}>Exportá el CSV desde MyJump Lab y subilo acá. Importa CMJ, SJ, DJ, RSI, Fuerza y Potencia.</p>
        <div style={s.uploadBox(hJ)} onClick={()=>jumpRef.current?.click()} onMouseEnter={()=>setHJ(true)} onMouseLeave={()=>setHJ(false)}>
          <div style={{ fontSize:32, marginBottom:8 }}>⬆️</div>
          <div style={{ fontFamily:"Barlow Condensed", fontSize:15, color:"#7dd4f0", letterSpacing:1 }}>Subir CSV de MyJump</div>
          <div style={{ fontSize:11, color:"#4a7a94", marginTop:4 }}>MyJumpLab_MyJump01.csv</div>
          <input ref={jumpRef} type="file" accept=".csv,.txt" style={{ display:"none" }} onChange={handleJumpImport} />
        </div>
        {jumpMsg && <div style={s.msgBox(jumpMsg.startsWith("✓"))}>{jumpMsg}</div>}
        <div style={{ marginTop:16, background:"#0a1824", borderRadius:8, padding:"12px 14px" }}>
          <div style={{ fontSize:11, color:"#4a7a94", letterSpacing:1, marginBottom:6 }}>MÉTRICAS</div>
          {["Altura de salto (cm)","RSI mod","Fuerza (N)","Potencia (W)","Tipo de salto","Color estado de forma"].map(c=>(
            <div key={c} style={{ fontSize:12, color:"#7dd4f0", padding:"2px 0", borderBottom:"1px solid #132030" }}>✓ {c}</div>
          ))}
        </div>
      </div>

      {/* PhotoFinish */}
      <div style={s.card}>
        <div style={s.cardTitle}>⚡ PhotoFinish</div>
        <p style={{ fontSize:13, color:"#7dd4f0", marginBottom:16, lineHeight:1.7 }}>Exportá el CSV desde PhotoFinish y subilo acá. Importa tiempos de velocidad por distancia (10m, 20m, 30m…).</p>
        <div style={s.uploadBox(hV)} onClick={()=>velRef.current?.click()} onMouseEnter={()=>setHV(true)} onMouseLeave={()=>setHV(false)}>
          <div style={{ fontSize:32, marginBottom:8 }}>⬆️</div>
          <div style={{ fontFamily:"Barlow Condensed", fontSize:15, color:"#7dd4f0", letterSpacing:1 }}>Subir CSV de PhotoFinish</div>
          <div style={{ fontSize:11, color:"#4a7a94", marginTop:4 }}>photo-finish-export.csv</div>
          <input ref={velRef} type="file" accept=".csv,.txt" style={{ display:"none" }} onChange={handleVelImport} />
        </div>
        {velMsg && <div style={s.msgBox(velMsg.startsWith("✓"))}>{velMsg}</div>}
        <div style={{ marginTop:16, background:"#0a1824", borderRadius:8, padding:"12px 14px" }}>
          <div style={{ fontSize:11, color:"#4a7a94", letterSpacing:1, marginBottom:6 }}>MÉTRICAS</div>
          {["Tiempo final (s)","Distancia (10m, 20m, 30m…)","Fecha del test","Nombre del atleta"].map(c=>(
            <div key={c} style={{ fontSize:12, color:"#7dd4f0", padding:"2px 0", borderBottom:"1px solid #132030" }}>✓ {c}</div>
          ))}
        </div>
      </div>

      {hasData && (
        <div style={{ ...s.card, gridColumn:"1/-1", background:"#0a1824", border:"1px solid #00e5a033" }}>
          <div style={{ display:"flex", gap:16, alignItems:"center", flexWrap:"wrap" }}>
            <span style={{ color:"#00e5a0", fontSize:13 }}>✓ Datos cargados:</span>
            <span style={{ fontSize:13 }}>{players.length} jugadores</span>
            <span style={{ fontSize:13 }}>·  {jumpRecs.length} registros de salto</span>
            <span style={{ fontSize:13 }}>·  {velRecs.length} registros de velocidad</span>
            <button style={{ marginLeft:"auto", padding:"7px 18px", borderRadius:7, border:"none", cursor:"pointer", fontFamily:"Barlow Condensed", fontWeight:700, fontSize:13, letterSpacing:1, background:"#00e5a0", color:"#080f17" }} onClick={()=>setTab("dashboard")}>
              Ver análisis →
            </button>
          </div>
        </div>
      )}
    </div>
  );

  // ── Dashboard ─────────────────────────────────────────────────────────────
  const Dashboard = () => {
    const visible = visiblePlayers.filter(p=>(pJumpRecs(p.id).length+pVelRecs(p.id).length)>0);
    return (
      <div>
        <div style={{ display:"flex", gap:12, marginBottom:16, flexWrap:"wrap", alignItems:"flex-end" }}>
          <div><label style={s.label}>Equipo</label>
            <select style={s.select} value={filterTeam} onChange={e=>setFilterTeam(e.target.value)}>
              {teams.map(t=><option key={t}>{t}</option>)}
            </select>
          </div>
          <div style={{ marginLeft:"auto", display:"flex", gap:10 }}>
            <div style={s.stat}><div style={s.statVal}>{visible.length}</div><div style={s.statLbl}>Jugadores</div></div>
            <div style={s.stat}><div style={s.statVal}>{jumpRecs.length}</div><div style={s.statLbl}>Saltos</div></div>
            <div style={s.stat}><div style={s.statVal}>{velRecs.length}</div><div style={s.statLbl}>Velocidades</div></div>
          </div>
        </div>

        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(185px,1fr))", gap:12, marginBottom:20 }}>
          {visible.map((p,i)=>{
            const sc=overallScore(p.id); const lv=getLevel(sc);
            const bj=bestJump(p.id); const lj=lastJump(p.id);
            const bv30=bestVel(p.id,"30m");
            return (
              <div key={p.id}
                style={{ ...s.card, borderTop:`3px solid ${COLORS[i%COLORS.length]}`, cursor:"pointer", marginBottom:0, padding:14 }}
                onClick={()=>{ setSelPlayer(p.id); setTab("evolution"); }}>
                <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
                  <div style={{ fontFamily:"Barlow Condensed", fontWeight:700, fontSize:15, lineHeight:1.2 }}>{p.name}</div>
                  <span style={s.badge(sc)}>{lv.label}</span>
                </div>
                <div style={{ fontSize:11, color:"#4a7a94", marginBottom:8 }}>{p.team||"—"}</div>
                <div style={s.bigScore(sc)}>{sc}</div>
                <div style={{ fontSize:10, color:"#4a7a94", marginTop:2 }}>score general</div>
                <div style={{ marginTop:10, display:"flex", gap:5, flexWrap:"wrap" }}>
                  {bj && <div style={{ background:"#0a1824", borderRadius:5, padding:"3px 8px", fontSize:12 }}>
                    <span style={{ color:"#4a7a94" }}>{bj.jumpTypeRaw} </span>
                    <span style={{ fontWeight:600, color:getStatusColor(lj?.statusColor) }}>{bj.altura?.toFixed(1)}cm</span>
                  </div>}
                  {bv30 && <div style={{ background:"#0a1824", borderRadius:5, padding:"3px 8px", fontSize:12 }}>
                    <span style={{ color:"#4a7a94" }}>30m </span>
                    <span style={{ fontWeight:600, color:"#7dd4f0" }}>{bv30.time?.toFixed(2)}s</span>
                  </div>}
                </div>
              </div>
            );
          })}
        </div>

        <div style={s.card}>
          <div style={s.cardTitle}>🏆 Ranking general</div>
          <ResponsiveContainer width="100%" height={Math.max(160,rankingData.length*26)}>
            <BarChart data={rankingData} layout="vertical">
              <CartesianGrid strokeDasharray="3 3" stroke="#1a3045" horizontal={false} />
              <XAxis type="number" domain={[0,100]} tick={{ fill:"#4a7a94", fontSize:11 }} />
              <YAxis dataKey="name" type="category" tick={{ fill:"#7dd4f0", fontSize:11, fontFamily:"Barlow Condensed" }} width={80} />
              <Tooltip content={<CustomTooltip />} />
              <Bar dataKey="score" fill="#00e5a0" radius={[0,6,6,0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        {allDistances.length > 0 && (
          <div style={s.card}>
            <div style={s.cardTitle}>⚡ Mejores tiempos por distancia</div>
            <ResponsiveContainer width="100%" height={Math.max(160, visiblePlayers.length*24)}>
              <BarChart
                data={visiblePlayers.filter(p=>pVelRecs(p.id).length>0).map(p=>({
                  name: p.name.split(" ")[0],
                  ...Object.fromEntries(allDistances.map(d=>[d, bestVel(p.id,d)?.time??null]))
                }))}
                layout="vertical">
                <CartesianGrid strokeDasharray="3 3" stroke="#1a3045" horizontal={false} />
                <XAxis type="number" tick={{ fill:"#4a7a94", fontSize:11 }} unit="s" />
                <YAxis dataKey="name" type="category" tick={{ fill:"#7dd4f0", fontSize:11, fontFamily:"Barlow Condensed" }} width={80} />
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize:12, fontFamily:"Barlow Condensed" }} />
                {allDistances.map((d,i)=><Bar key={d} dataKey={d} fill={COLORS[i]} radius={[0,4,4,0]} />)}
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>
    );
  };

  // ── Evolution ─────────────────────────────────────────────────────────────
  const Evolution = () => {
    const pid  = selPlayer||players[0]?.id;
    const jevo = jumpEvoData(pid);
    const vevo = velEvoData(pid);
    const bj   = bestJump(pid);
    const jrecs= pJumpRecs(pid);
    const vrecs= pVelRecs(pid);

    return (
      <div>
        <div style={{ display:"flex", gap:12, marginBottom:16, alignItems:"flex-end", flexWrap:"wrap" }}>
          <div><label style={s.label}>Jugador</label>
            <select style={s.select} value={pid} onChange={e=>setSelPlayer(parseInt(e.target.value))}>
              {players.filter(p=>(pJumpRecs(p.id).length+pVelRecs(p.id).length)>0).map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          </div>
        </div>

        <div style={{ display:"flex", gap:10, marginBottom:16, flexWrap:"wrap" }}>
          {[
            { val:jrecs.length+vrecs.length, lbl:"Total registros" },
            { val:bj?.altura?.toFixed(1)+"cm"??"—", lbl:"Mejor CMJ" },
            { val:bj?.rsi?.toFixed(2)??"—",          lbl:"Mejor RSI" },
            ...allDistances.map(d=>({ val:bestVel(pid,d)?.time?.toFixed(2)+"s"??"—", lbl:`Mejor ${d}` })),
          ].map(({val,lbl})=>(
            <div key={lbl} style={s.stat}><div style={s.statVal}>{val}</div><div style={s.statLbl}>{lbl}</div></div>
          ))}
        </div>

        {jevo.length>0 && (
          <div style={s.card}>
            <div style={s.cardTitle}>Evolución salto (cm)</div>
            <ResponsiveContainer width="100%" height={210}>
              <LineChart data={jevo}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1a3045" />
                <XAxis dataKey="fecha" tick={{ fill:"#4a7a94", fontSize:11 }} />
                <YAxis tick={{ fill:"#4a7a94", fontSize:11 }} unit="cm" />
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize:12, fontFamily:"Barlow Condensed" }} />
                {["CMJ","DJ"].map((k,i)=><Line key={k} dataKey={k} stroke={COLORS[i]} strokeWidth={2.5} dot={{ r:5 }} connectNulls />)}
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}

        {vevo.length>0 && allDistances.length>0 && (
          <div style={s.card}>
            <div style={s.cardTitle}>Evolución velocidad (s)</div>
            <ResponsiveContainer width="100%" height={210}>
              <LineChart data={vevo}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1a3045" />
                <XAxis dataKey="fecha" tick={{ fill:"#4a7a94", fontSize:11 }} />
                <YAxis tick={{ fill:"#4a7a94", fontSize:11 }} unit="s" />
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize:12, fontFamily:"Barlow Condensed" }} />
                {allDistances.map((d,i)=><Line key={d} dataKey={d} stroke={COLORS[i+2]} strokeWidth={2.5} dot={{ r:5 }} connectNulls />)}
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}

        {jevo.some(d=>d.RSI) && (
          <div style={s.card}>
            <div style={s.cardTitle}>Evolución RSI mod</div>
            <ResponsiveContainer width="100%" height={180}>
              <LineChart data={jevo}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1a3045" />
                <XAxis dataKey="fecha" tick={{ fill:"#4a7a94", fontSize:11 }} />
                <YAxis tick={{ fill:"#4a7a94", fontSize:11 }} />
                <Tooltip content={<CustomTooltip />} />
                <Line dataKey="RSI" stroke="#f0c060" strokeWidth={2.5} dot={{ r:5 }} connectNulls />
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}

        {/* Historial velocidad */}
        {vrecs.length>0 && (
          <div style={s.card}>
            <div style={s.cardTitle}>Historial velocidad</div>
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                <thead><tr style={{ borderBottom:"1px solid #1a3045" }}>
                  {["Fecha","Test","Distancia","Tiempo (s)","Nivel"].map(h=><th key={h} style={{ padding:"6px 10px", color:"#4a7a94", textAlign:"left", fontFamily:"Barlow Condensed", letterSpacing:1, fontWeight:600 }}>{h}</th>)}
                </tr></thead>
                <tbody>
                  {vrecs.map((r,i)=>{
                    const ref=VEL_REFS[r.distance]; const sc=score(ref,r.time); const lv=getLevel(sc??0);
                    return <tr key={i} style={{ borderBottom:"1px solid #132030" }}>
                      <td style={{ padding:"6px 10px", color:"#7dd4f0" }}>{r.fullDate}</td>
                      <td style={{ padding:"6px 10px" }}>{r.testName}</td>
                      <td style={{ padding:"6px 10px" }}>{r.distance||"—"}</td>
                      <td style={{ padding:"6px 10px", fontWeight:600 }}>{r.time?.toFixed(2)}</td>
                      <td style={{ padding:"6px 10px" }}><span style={s.badge(sc??0)}>{lv.label}</span></td>
                    </tr>;
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    );
  };

  // ── Compare ───────────────────────────────────────────────────────────────
  const Compare = () => {
    const pidA=compareA||players[0]?.id; const pidB=compareB||players[1]?.id;
    const nA=players.find(p=>p.id===pidA)?.name.split(" ")[0]||"A";
    const nB=players.find(p=>p.id===pidB)?.name.split(" ")[0]||"B";

    const radarMetrics = [
      { key:"cmj",  label:"CMJ",  vA:bestJump(pidA)?.altura,  vB:bestJump(pidB)?.altura,  ref:JUMP_REFS.cmj },
      { key:"rsi",  label:"RSI",  vA:bestJump(pidA)?.rsi,     vB:bestJump(pidB)?.rsi,     ref:JUMP_REFS.rsi },
      ...allDistances.map(d=>({ key:d, label:d, vA:bestVel(pidA,d)?.time, vB:bestVel(pidB,d)?.time, ref:VEL_REFS[d] })),
    ];
    const radarD = radarMetrics.map(m=>({ metric:m.label, [nA]:score(m.ref,m.vA)??0, [nB]:score(m.ref,m.vB)??0 }));

    const evoA=jumpEvoData(pidA); const evoB=jumpEvoData(pidB);
    const allF=Array.from(new Set([...evoA.map(d=>d.fecha),...evoB.map(d=>d.fecha)])).sort();
    const merged=allF.map(f=>({ fecha:f, [nA]:evoA.find(d=>d.fecha===f)?.CMJ??null, [nB]:evoB.find(d=>d.fecha===f)?.CMJ??null }));

    return (
      <div>
        <div style={{ display:"flex", gap:12, marginBottom:16, flexWrap:"wrap" }}>
          {[{val:pidA,set:setCompareA,label:"Jugador A",col:"#00e5a0"},{val:pidB,set:setCompareB,label:"Jugador B",col:"#7dd4f0"}].map(({val,set,label,col})=>(
            <div key={label} style={{ flex:1, minWidth:160 }}>
              <label style={{ ...s.label, color:col }}>{label}</label>
              <select style={s.select} value={val||""} onChange={e=>set(parseInt(e.target.value))}>
                {players.filter(p=>(pJumpRecs(p.id).length+pVelRecs(p.id).length)>0).map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
            </div>
          ))}
        </div>

        <div style={s.grid2}>
          <div style={s.card}>
            <div style={s.cardTitle}>Radar comparativo</div>
            <ResponsiveContainer width="100%" height={270}>
              <RadarChart data={radarD}>
                <PolarGrid stroke="#1a3045" />
                <PolarAngleAxis dataKey="metric" tick={{ fill:"#7dd4f0", fontSize:10, fontFamily:"Barlow Condensed" }} />
                <PolarRadiusAxis domain={[0,100]} tick={{ fill:"#4a7a94", fontSize:9 }} />
                <Radar name={nA} dataKey={nA} stroke="#00e5a0" fill="#00e5a0" fillOpacity={0.2} strokeWidth={2} />
                <Radar name={nB} dataKey={nB} stroke="#7dd4f0" fill="#7dd4f0" fillOpacity={0.2} strokeWidth={2} />
                <Legend wrapperStyle={{ fontSize:12, fontFamily:"Barlow Condensed" }} />
                <Tooltip content={<CustomTooltip />} />
              </RadarChart>
            </ResponsiveContainer>
          </div>

          <div style={s.card}>
            <div style={s.cardTitle}>Métricas vs referencia</div>
            <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
              {radarMetrics.map(({key,label,vA,vB,ref})=>(
                <div key={key} style={{ background:"#0a1824", borderRadius:8, padding:"8px 12px" }}>
                  <div style={{ display:"flex", justifyContent:"space-between", marginBottom:5 }}>
                    <span style={{ fontFamily:"Barlow Condensed", fontSize:13, color:"#7dd4f0" }}>{label}</span>
                    {ref && <span style={{ fontSize:10, color:"#4a7a94" }}>ref: {ref.excellent}{ref.unit}</span>}
                  </div>
                  <div style={{ display:"flex", gap:8 }}>
                    {[{name:nA,val:vA,col:"#00e5a0"},{name:nB,val:vB,col:"#7dd4f0"}].map(({name,val,col})=>{
                      const sc=score(ref,val); const lv=getLevel(sc??0);
                      return (
                        <div key={name} style={{ flex:1, textAlign:"center", background:"#0d1e2d", borderRadius:6, padding:"5px 4px" }}>
                          <div style={{ fontSize:10, color:col, marginBottom:2 }}>{name}</div>
                          <div style={{ fontSize:14, fontWeight:700, color:val!=null?lv.color:"#2a4a5a" }}>
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

        {merged.some(d=>d[nA]||d[nB]) && (
          <div style={s.card}>
            <div style={s.cardTitle}>Evolución CMJ comparada</div>
            <ResponsiveContainer width="100%" height={210}>
              <LineChart data={merged}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1a3045" />
                <XAxis dataKey="fecha" tick={{ fill:"#4a7a94", fontSize:11 }} />
                <YAxis tick={{ fill:"#4a7a94", fontSize:11 }} unit="cm" />
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize:12, fontFamily:"Barlow Condensed" }} />
                <Line dataKey={nA} stroke="#00e5a0" strokeWidth={2.5} dot={{ r:5 }} connectNulls />
                <Line dataKey={nB} stroke="#7dd4f0" strokeWidth={2.5} dot={{ r:5 }} connectNulls />
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>
    );
  };

  return (
    <div style={s.app}>
      <link href="https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600&family=Barlow+Condensed:wght@600;700;800&display=swap" rel="stylesheet" />
      <div style={s.header}>
        <div>
          <div style={s.logo}>⚡ AthleteIQ</div>
          <div style={s.tagline}>Performance Tracker · Fútbol</div>
        </div>
        {hasData && <div style={{ marginLeft:"auto", fontSize:12, color:"#4a7a94" }}>{players.length} jugadores · {jumpRecs.length} saltos · {velRecs.length} velocidades</div>}
      </div>
      <div style={s.nav}>
        {[
          { key:"import",    label:"📂 Importar" },
          { key:"dashboard", label:"Dashboard",  d:!hasData },
          { key:"evolution", label:"Evolución",  d:!hasData },
          { key:"compare",   label:"Comparar",   d:!hasData },
        ].map(({key,label,d})=>(
          <button key={key} style={s.navBtn(tab===key,d)} onClick={()=>!d&&setTab(key)}>{label}</button>
        ))}
      </div>
      <div style={s.main}>
        {tab==="import"    && <ImportView />}
        {tab==="dashboard" && hasData && <Dashboard />}
        {tab==="evolution" && hasData && <Evolution />}
        {tab==="compare"   && hasData && <Compare />}
      </div>
    </div>
  );
}
