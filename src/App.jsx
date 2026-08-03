import { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, BarChart, Bar, RadarChart, Radar, PolarGrid,
  PolarAngleAxis, PolarRadiusAxis, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, ScatterChart, Scatter, ZAxis,
} from "recharts";

// ── F1 Design Tokens ──────────────────────────────────────────────────────────
const F = {
  carbon:     "#0A0A0A",
  asphalt:    "#141414",
  panel:      "#1C1C1C",
  panelBorder:"#2A2A2A",
  red:        "#E8002D",
  white:      "#F0F0F0",
  silver:     "#A8A8A8",
  dim:        "#5A5A5A",
  ghost:      "#2E2E2E",
  yellow:     "#FFD700",
  teal:       "#00D2FF",
  green:      "#39B54A",
  purple:     "#B084F5",
  orange:     "#FF8700",
  fontF1:     "'Titillium Web', 'Arial Narrow', sans-serif",
  fontMono:   "'Share Tech Mono', 'Courier New', monospace",
};

const TEAM_COLORS = [F.teal, F.red, F.yellow, "#FF8700", "#00D2BE", "#DC143C", "#0600EF", "#006F62", "#B6BABD", "#FFFFFF"];

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
// Asimetría de movilidad: cuanto menor, mejor. No entra al score de rendimiento —
// es un indicador de riesgo / calidad de movimiento aparte.
const MOBILITY_REF = { excellent:5, good:10, average:15, unit:"%", label:"Asimetría", lowerIsBetter:true };
// Fuerza relativa (1RM / peso corporal) para los ejercicios principales.
const STRENGTH_REFS = {
  squat:        { excellent:2.0, good:1.6, average:1.2, unit:"x PC", label:"Squat rel.",  lowerIsBetter:false },
  "bench-press":{ excellent:1.2, good:1.0, average:0.8, unit:"x PC", label:"Bench rel.",   lowerIsBetter:false },
};

// ── Banco de ejercicios correctivos ──────────────────────────────────────────
// Curado a partir de los propios programas de Leandro (Fuerza Estructural I-V,
// Readaptación LCA, planillas Fuerza-Potencia) — no genérico.
// ── Referencias de campo (fútbol) provistas por Leandro — usadas SOLO en el Informe ──
// El Dashboard y Comparar siguen siendo percentiles relativos al plantel (sin tocar).
const REFERENCE_RANGES = {
  cmj:           { min: 42.0, max: 54.5, unit: "cm",   lowerIsBetter: false, label: "CMJ" },
  dj:            { min: 35.0, max: 45.0, unit: "cm",   lowerIsBetter: false, label: "DJ" },
  rsi:           { min: 1.8,  max: 2.5,  unit: "",     lowerIsBetter: false, label: "RSI (DJ)" },
  "10m":         { min: 1.62, max: 1.75, unit: "s",    lowerIsBetter: true,  label: "Sprint 10m" },
  "20m":         { min: 2.75, max: 2.95, unit: "s",    lowerIsBetter: true,  label: "Sprint 20m" },
  "30m":         { min: 3.85, max: 4.10, unit: "s",    lowerIsBetter: true,  label: "Sprint 30m" },
  squat:         { min: 1.7,  max: 2.2,  unit: "x PC", lowerIsBetter: false, label: "Squat 1RM rel." },
  "bench-press": { min: 1.2,  max: 1.5,  unit: "x PC", lowerIsBetter: false, label: "Bench 1RM rel." },
  dorsiflexion:  { min: 38,   max: 45,   unit: "°",    lowerIsBetter: false, label: "Dorsiflexión tobillo" },
  hipIR:         { min: 30,   max: 45,   unit: "°",    lowerIsBetter: false, label: "Rotación interna cadera" },
};
// El piso/techo ÉLITE de cada rango es tu dato real. BUENO/PROMEDIO se extrapolan
// proporcionalmente (±12% / ±25%) hasta que me pases esos cortes específicos —
// ajustables acá mismo si conseguís valores más precisos.
function bandFromRange(range) {
  const { min, max, lowerIsBetter } = range;
  if (lowerIsBetter) {
    const excellent = max;
    return { ...range, excellent, good: excellent * 1.12, average: excellent * 1.25 };
  }
  const excellent = min;
  return { ...range, excellent, good: excellent * 0.88, average: excellent * 0.75 };
}
const REFERENCE_BANDS = Object.fromEntries(Object.entries(REFERENCE_RANGES).map(([k, v]) => [k, bandFromRange(v)]));

function bandScore(ref, value) {
  if (!ref || value == null) return null;
  const { excellent, good, average, lowerIsBetter } = ref;
  if (lowerIsBetter) {
    if (value <= excellent) return 100; if (value <= good) return 75;
    if (value <= average) return 50; return 25;
  } else {
    if (value >= excellent) return 100; if (value >= good) return 75;
    if (value >= average) return 50; return 25;
  }
}
function bandLevel(score) {
  if (score >= 90) return { label: "ELITE", color: F.yellow };
  if (score >= 70) return { label: "BUENO", color: F.green };
  if (score >= 50) return { label: "PROMEDIO", color: F.teal };
  return { label: "BAJO", color: F.red };
}
const ReportBadge = ({ score }) => {
  if (score == null) return <span style={{ fontFamily: F.fontMono, fontSize: 9, color: F.ghost }}>—</span>;
  const lv = bandLevel(score);
  return (
    <div style={{ display: "inline-flex", alignItems: "center", gap: 5, padding: "2px 8px",
      border: `1px solid ${lv.color}`, borderRadius: 2, background: lv.color + "18" }}>
      <div style={{ width: 5, height: 5, borderRadius: "50%", background: lv.color, boxShadow: `0 0 5px ${lv.color}` }} />
      <span style={{ fontFamily: F.fontMono, fontSize: 9, color: lv.color, letterSpacing: 2 }}>{lv.label}</span>
    </div>
  );
};

const EXERCISE_BANK = {
  salto: {
    label: "Salto / Potencia reactiva",
    fuente: "Programa de Readaptación LCA · Planillas Fuerza-Potencia",
    items: [
      { name: "Pogos reactivos", sets: "3×10", note: "Contacto corto con el suelo, tobillo rígido" },
      { name: "Drop jump progresivo (30-45-60cm)", sets: "3×4", note: "Progresión de fuerza reactiva" },
      { name: "Sentadilla con salto", sets: "3×6-8" },
      { name: "Salto al cajón a 1 pierna", sets: "3×4+4" },
      { name: "Búlgaras con salto (carga progresiva)", sets: "3×4+4" },
    ],
  },
  rsi: {
    label: "RSI / Ciclo estiramiento-acortamiento",
    fuente: "Índice de Fuerza Reactiva (Comunidad Lift) · Readaptación LCA",
    items: [
      { name: "Pogos reactivos (tiempo de contacto <200ms)", sets: "3×10" },
      { name: "Drop jump + salto horizontal", sets: "3×4" },
      { name: "Hurdle hops (vallas bajas)", sets: "3×6" },
      { name: "Saltos unilaterales alternados", sets: "3×4+4" },
    ],
  },
  velocidad: {
    label: "Velocidad / Aceleración",
    fuente: "Programa de Readaptación LCA",
    items: [
      { name: "Aceleraciones 5-10m", sets: "4×4-5" },
      { name: "Técnica de carrera + aceleración", sets: "3×20m" },
      { name: "Aceleración 10m + freno controlado en 5m", sets: "4×3" },
      { name: "Cambios de dirección 45°-90°", sets: "3-4×3+3" },
    ],
  },
  squat: {
    label: "Fuerza tren inferior (Squat)",
    fuente: "Fuerza Estructural I-V · Día Fuerza (planillas propias)",
    items: [
      { name: "Sentadilla profunda con barra", sets: "3×6-8" },
      { name: "Sentadilla búlgara", sets: "3×6+6" },
      { name: "Peso muerto convencional / rumano", sets: "3×6-8" },
      { name: "Hip thrust con barra", sets: "3×8-12" },
      { name: "Sentadilla copa / Prensa", sets: "3×8" },
    ],
  },
  bench: {
    label: "Fuerza tren superior (Bench)",
    fuente: "Fuerza Estructural I-V · Día Fuerza (planillas propias)",
    items: [
      { name: "Press plano con barra", sets: "3×6-8" },
      { name: "Press inclinado con barra", sets: "3×6-8" },
      { name: "Flexiones de brazos (lastradas si es posible)", sets: "3×8-12" },
      { name: "Press de hombros con mancuernas", sets: "3×8" },
    ],
  },
  movilidad: {
    label: "Movilidad / Control neuromuscular",
    fuente: "Fuerza Estructural I-V · Síndrome femoropatelar (Comunidad Lift)",
    items: [
      { name: "Movilidad de cadera 90-90", sets: "2-3×6-10+6-10" },
      { name: "Movilidad de tobillo", sets: "2-3×6-10+6-10" },
      { name: "Movilidad de aductores", sets: "2×10" },
      { name: "Activación de glúteo medio (monster walk / step down)", sets: "3×6-10+6-10" },
      { name: "Control neuromuscular unilateral (pistol squat asistido)", sets: "3×6+6" },
    ],
  },
};

// percentileColor: gradiente visual según posición relativa dentro del plantel cargado.
// No es una clasificación (ELITE/BUENO) — esa vendrá del informe una vez cargadas
// referencias/investigación reales. Acá solo indica mejor/peor que el resto del grupo.
function percentileColor(pct) {
  if (pct == null) return F.ghost;
  if (pct >= 75) return F.green;
  if (pct >= 50) return F.teal;
  if (pct >= 25) return F.yellow;
  return F.red;
}

// ── ID generator (string, no floats) ─────────────────────────────────────────
let _seq = 1;
const uid = () => `id_${Date.now()}_${_seq++}`;

// ── Parsers ───────────────────────────────────────────────────────────────────
function parseNum(str) {
  if (!str || str === "---" || str === "") return null;
  return parseFloat(String(str).replace(",", ".").replace(/[^\d.-]/g, "")) || null;
}
function parseMyJumpDate(str) {
  const m = String(str).match(/(\d+)\.(\d+)\.(\d+)/);
  if (!m) return str;
  return `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
}
function parsePhotoDate(str) {
  const m = String(str).match(/(\d+)\/(\d+)\/(\d+)/);
  if (!m) return str;
  return `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
}
function jumpTypeKey(tipo) {
  if (!tipo) return "cmj"; const t = tipo.toLowerCase();
  if (t.includes("cmj")) return "cmj";
  if (t.includes("sj") && !t.includes("rsi")) return "sj";
  if (t.includes("dj") || t.includes("drop")) return "dj";
  return "cmj";
}
function detectDistance(info) {
  if (!info) return null;
  const m = String(info).match(/(\d+)\s*m/i);
  return m ? m[1] + "m" : null;
}
function exerciseKey(name) {
  if (!name) return null; const t = name.toLowerCase();
  if (t.includes("squat") || t.includes("sentadilla")) return "squat";
  if (t.includes("bench") || t.includes("banca")) return "bench-press";
  return null;
}
// Generic MyJump-family semicolon CSV parser — also used for MyROM and MyLift exports,
// which share the same delimiter/header structure.
function parseMyJumpCSV(text) {
  const lines = text.trim().split("\n"); if (lines.length < 2) return [];
  const headers = lines[0].split(";").map(h => h.trim());
  return lines.slice(1).map(line => {
    const vals = line.split(";"); const obj = {};
    headers.forEach((h, i) => { obj[h] = (vals[i] || "").trim(); });
    return obj;
  });
}
function parsePhotoFinishCSV(text) {
  const lines = text.trim().split("\n"); if (lines.length < 2) return [];
  let cDate = "", cInfo = ""; const rows = [];
  lines.slice(1).forEach(line => {
    const cols = line.split(",").map(c => c.trim());
    if (cols[0]) cDate = cols[0]; if (cols[1]) cInfo = cols[1];
    const athlete = cols[3] || ""; const time = cols[5] || "";
    if (!athlete) return;
    rows.push({ date: cDate, info: cInfo, athlete, time });
  });
  return rows;
}

// ── F1 UI Components ──────────────────────────────────────────────────────────
const TelemetryBar = () => {
  const [pos, setPos] = useState(0);
  useEffect(() => {
    const id = setInterval(() => setPos(p => (p + 0.4) % 100), 30);
    return () => clearInterval(id);
  }, []);
  return (
    <div style={{ height: 3, background: F.ghost, overflow: "hidden", position: "relative" }}>
      <div style={{ position: "absolute", left: `${pos}%`, top: 0, height: "100%", width: "30%",
        background: `linear-gradient(90deg, transparent, ${F.red}, ${F.teal}, ${F.yellow}, transparent)`, opacity: 0.8 }} />
    </div>
  );
};

const CustomTooltip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: F.panel, border: `1px solid ${F.red}`, borderRadius: 4, padding: "8px 14px" }}>
      <div style={{ fontFamily: F.fontMono, fontSize: 10, color: F.red, marginBottom: 4, letterSpacing: 2 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color, fontSize: 11, fontFamily: F.fontMono, margin: "1px 0" }}>
          {p.name}: <span style={{ color: F.white, fontWeight: 700 }}>{typeof p.value === "number" ? p.value.toFixed(2) : p.value}</span>
        </div>
      ))}
    </div>
  );
};

// Chip de percentil — muestra "P{n}" (posición relativa dentro del plantel), sin etiquetas de calidad.
const PercentileChip = ({ pct }) => {
  if (pct == null) return <span style={{ fontFamily: F.fontMono, fontSize: 9, color: F.ghost }}>—</span>;
  const col = percentileColor(pct);
  return (
    <div style={{ display: "inline-flex", alignItems: "center", gap: 5, padding: "2px 8px",
      border: `1px solid ${col}`, borderRadius: 2, background: col + "18" }}>
      <div style={{ width: 5, height: 5, borderRadius: "50%", background: col, boxShadow: `0 0 5px ${col}` }} />
      <span style={{ fontFamily: F.fontMono, fontSize: 9, color: col, letterSpacing: 1 }}>P{pct}</span>
    </div>
  );
};

const TimingScore = ({ score, size = 60 }) => {
  const col = percentileColor(score);
  const val = score ?? 0;
  const r = size / 2 - 5; const circ = 2 * Math.PI * r; const dash = (val / 100) * circ;
  return (
    <div style={{ position: "relative", width: size, height: size }}>
      <svg width={size} height={size} style={{ transform: "rotate(-90deg)" }}>
        <circle cx={size/2} cy={size/2} r={r} fill="none" stroke={F.ghost} strokeWidth={3} />
        <circle cx={size/2} cy={size/2} r={r} fill="none" stroke={col} strokeWidth={3}
          strokeDasharray={`${dash} ${circ}`} strokeLinecap="butt"
          style={{ filter: `drop-shadow(0 0 4px ${col})` }} />
      </svg>
      <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <div style={{ fontFamily: F.fontMono, fontSize: size * 0.26, fontWeight: 700, color: col, lineHeight: 1 }}>{score ?? "—"}</div>
      </div>
    </div>
  );
};

const PanelHeader = ({ children, accent }) => (
  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14,
    paddingBottom: 8, borderBottom: `1px solid ${F.panelBorder}` }}>
    <div style={{ width: 3, height: 14, background: accent || F.red, borderRadius: 1 }} />
    <div style={{ fontFamily: F.fontMono, fontSize: 10, color: F.silver, letterSpacing: 2, textTransform: "uppercase" }}>{children}</div>
  </div>
);

const BackBtn = ({ onClick }) => (
  <button onClick={onClick} style={{ display: "inline-flex", alignItems: "center", gap: 6,
    padding: "6px 14px", background: "transparent", border: `1px solid ${F.panelBorder}`,
    borderRadius: 3, color: F.silver, fontFamily: F.fontMono, fontSize: 10, letterSpacing: 2,
    cursor: "pointer", marginBottom: 16, transition: "border-color .15s" }}
    onMouseEnter={e => e.currentTarget.style.borderColor = F.red}
    onMouseLeave={e => e.currentTarget.style.borderColor = F.panelBorder}>
    ← VOLVER
  </button>
);

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [players,    setPlayers]    = useState([]);
  const [jumpRecs,   setJumpRecs]   = useState([]);
  const [velRecs,    setVelRecs]    = useState([]);
  const [romRecs,    setRomRecs]    = useState([]);
  const [liftRecs,   setLiftRecs]   = useState([]);
  const [customRecs, setCustomRecs] = useState([]);
  const [tab,        setTab]        = useState("import");
  const [prevTab,    setPrevTab]    = useState(null);

  // Use string IDs throughout
  const [selPlayer, setSelPlayer] = useState(null); // string id
  const [compareA,  setCompareA]  = useState(null); // string id
  const [compareB,  setCompareB]  = useState(null); // string id

  const [filterTeam,  setFilterTeam]  = useState("Todos");
  const [msgs,        setMsgs]        = useState({});
  const [showDelete,  setShowDelete]  = useState(false);
  const [genericData, setGenericData] = useState(null);
  const [genericCols, setGenericCols] = useState([]);
  const [colMap,      setColMap]      = useState({});
  const [showMapper,  setShowMapper]  = useState(false);
  const [liftExercise,setLiftExercise]= useState(null);

  const jumpRef    = useRef();
  const velRef     = useRef();
  const romRef     = useRef();
  const liftRef    = useRef();
  const genericRef = useRef();

  function setMsg(key, val) { setMsgs(m => ({ ...m, [key]: val })); }

  function goTo(t) { setPrevTab(tab); setTab(t); }
  function goBack() { if (prevTab) { setTab(prevTab); setPrevTab(null); } else setTab("dashboard"); }

  // ── Player merge (string IDs) ─────────────────────────────────────────────
  function mergePlayer(name, team, cur) {
    const key = name.toLowerCase().trim();
    const ex = cur.find(p => p.name.toLowerCase().trim() === key);
    if (ex) return { players: cur, id: ex.id };
    const np = { id: uid(), name: name.trim(), team: team || "" };
    return { players: [...cur, np], id: np.id };
  }

  // ── Delete functions ──────────────────────────────────────────────────────
  function deleteAllData() {
    setPlayers([]); setJumpRecs([]); setVelRecs([]); setRomRecs([]); setLiftRecs([]); setCustomRecs([]);
    setSelPlayer(null); setCompareA(null); setCompareB(null);
    setMsgs({}); setShowDelete(false); setTab("import");
  }
  function deleteJumpRecs()  { setJumpRecs([]);   setMsg("del", "✓ Registros de salto eliminados"); }
  function deleteVelRecs()   { setVelRecs([]);    setMsg("del", "✓ Registros de velocidad eliminados"); }
  function deleteRomRecs()   { setRomRecs([]);    setMsg("del", "✓ Registros de movilidad eliminados"); }
  function deleteLiftRecs()  { setLiftRecs([]);   setMsg("del", "✓ Registros de VBT/RM eliminados"); }
  function deleteCustomRecs(){ setCustomRecs([]); setMsg("del", "✓ Registros personalizados eliminados"); }
  function deletePlayer(pid) {
    setPlayers(p => p.filter(x => x.id !== pid));
    setJumpRecs(r => r.filter(x => x.playerId !== pid));
    setVelRecs(r  => r.filter(x => x.playerId !== pid));
    setRomRecs(r  => r.filter(x => x.playerId !== pid));
    setLiftRecs(r => r.filter(x => x.playerId !== pid));
    setCustomRecs(r=> r.filter(x => x.playerId !== pid));
    if (selPlayer === pid) setSelPlayer(null);
    if (compareA  === pid) setCompareA(null);
    if (compareB  === pid) setCompareB(null);
  }

  // ── MyJump import ─────────────────────────────────────────────────────────
  function handleJumpImport(e) {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const rows = parseMyJumpCSV(ev.target.result);
      if (!rows.length) { setMsg("jump", "❌ No se pudo leer."); return; }
      let cur = [...players]; const newRecs = [];
      rows.forEach(row => {
        const name = (row["Nombre"] || "").trim(); const team = (row["Equipo"] || "").trim();
        if (!name) return;
        const { players: up, id: pid } = mergePlayer(name, team, cur); cur = up;
        const tipo = (row["Tipo de salto"] || "").trim();
        const fecha = parseMyJumpDate(row["Fecha"] || "");
        newRecs.push({
          id: uid(), playerId: pid, date: fecha.slice(0, 7), fullDate: fecha,
          jumpType: jumpTypeKey(tipo), jumpTypeRaw: tipo,
          altura:   parseNum(row["Altura de salto (cm)"]),
          rsi:      parseNum(row["RSI mod (m/s)"]),
          fuerza:   parseNum(row["Fuerza (N)"]),
          potencia: parseNum(row["Potencia (W)"]),
          statusColor: row["Color estado de forma"] || "", team,
        });
      });
      setPlayers(cur); setJumpRecs(r => [...r, ...newRecs]);
      if (!selPlayer && cur.length) {
        setSelPlayer(cur[0].id);
        setCompareA(cur[0].id);
        setCompareB(cur[1]?.id || cur[0].id);
      }
      setMsg("jump", `✓ ${newRecs.length} registros · ${[...new Set(newRecs.map(r => r.playerId))].length} atletas`);
      setTab("dashboard");
    };
    reader.readAsText(file, "UTF-8");
    e.target.value = "";
  }

  // ── PhotoFinish import ────────────────────────────────────────────────────
  function handleVelImport(e) {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const rows = parsePhotoFinishCSV(ev.target.result);
      if (!rows.length) { setMsg("vel", "❌ No se pudo leer."); return; }
      let cur = [...players]; const newRecs = [];
      rows.forEach(row => {
        const name = row.athlete.trim(); if (!name) return;
        const { players: up, id: pid } = mergePlayer(name, "", cur); cur = up;
        const fecha = parsePhotoDate(row.date);
        const dist = detectDistance(row.info);
        const t = parseNum(row.time); if (!t) return;
        newRecs.push({ id: uid(), playerId: pid, date: fecha.slice(0, 7), fullDate: fecha, distance: dist, testName: row.info, time: t });
      });
      setPlayers(cur); setVelRecs(r => [...r, ...newRecs]);
      if (!selPlayer && cur.length) { setSelPlayer(cur[0].id); setCompareA(cur[0].id); setCompareB(cur[1]?.id || cur[0].id); }
      setMsg("vel", `✓ ${newRecs.length} registros · ${[...new Set(newRecs.map(r => r.playerId))].length} atletas`);
      setTab("dashboard");
    };
    reader.readAsText(file, "UTF-8");
    e.target.value = "";
  }

  // ── MyROM (movilidad) import ──────────────────────────────────────────────
  function handleRomImport(e) {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const rows = parseMyJumpCSV(ev.target.result);
      if (!rows.length) { setMsg("rom", "❌ No se pudo leer."); return; }
      let cur = [...players]; const newRecs = [];
      rows.forEach(row => {
        const name = (row["Nombre"] || "").trim(); const team = (row["Equipo"] || "").trim();
        if (!name) return;
        const { players: up, id: pid } = mergePlayer(name, team, cur); cur = up;
        const fecha = parseMyJumpDate(row["Fecha"] || "");
        newRecs.push({
          id: uid(), playerId: pid, date: fecha.slice(0, 7), fullDate: fecha,
          test: (row["Test"] || "").trim(),
          angI: parseNum(row["Ángulo izq(º)"]),
          angD: parseNum(row["Ángulo drch(º)"]),
          asim: parseNum(row["Asimetría (%)"]),
          statusColor: row["Color estado de forma"] || "", team,
        });
      });
      setPlayers(cur); setRomRecs(r => [...r, ...newRecs]);
      if (!selPlayer && cur.length) { setSelPlayer(cur[0].id); setCompareA(cur[0].id); setCompareB(cur[1]?.id || cur[0].id); }
      setMsg("rom", `✓ ${newRecs.length} registros · ${[...new Set(newRecs.map(r => r.playerId))].length} atletas`);
      setTab("dashboard");
    };
    reader.readAsText(file, "UTF-8");
    e.target.value = "";
  }

  // ── MyLift (VBT / RM) import ──────────────────────────────────────────────
  function handleLiftImport(e) {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const rows = parseMyJumpCSV(ev.target.result);
      if (!rows.length) { setMsg("lift", "❌ No se pudo leer."); return; }
      let cur = [...players]; const newRecs = [];
      rows.forEach(row => {
        const name = (row["Nombre"] || "").trim(); const team = (row["Equipo"] || "").trim();
        if (!name) return;
        const { players: up, id: pid } = mergePlayer(name, team, cur); cur = up;
        const fecha = parseMyJumpDate(row["Fecha"] || "");
        const exercise = (row["Nombre del ejercicio"] || "").trim();
        const loads = [1, 2, 3, 4].map(i => parseNum(row[`Carga ${i} (kg)`]));
        const vels  = [1, 2, 3, 4].map(i => parseNum(row[`Velocidad media ${i} (m/s)`]));
        const points = loads.map((l, i) => ({ load: l, vel: vels[i] })).filter(p => p.load != null && p.vel != null);
        newRecs.push({
          id: uid(), playerId: pid, date: fecha.slice(0, 7), fullDate: fecha,
          exercise, exKey: exerciseKey(exercise),
          bodyweight: parseNum(row["Peso corporal(kg)"]),
          rm: parseNum(row["1-RM (kg)"]),
          points, team,
        });
      });
      setPlayers(cur); setLiftRecs(r => [...r, ...newRecs]);
      if (!selPlayer && cur.length) { setSelPlayer(cur[0].id); setCompareA(cur[0].id); setCompareB(cur[1]?.id || cur[0].id); }
      setMsg("lift", `✓ ${newRecs.length} registros · ${[...new Set(newRecs.map(r => r.playerId))].length} atletas`);
      setTab("dashboard");
    };
    reader.readAsText(file, "UTF-8");
    e.target.value = "";
  }

  // ── Generic import ────────────────────────────────────────────────────────
  function handleGenericImport(e) {
    const file = e.target.files[0]; if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        let rows = [];
        if (ext === "xlsx" || ext === "xls") {
          const wb = XLSX.read(ev.target.result, { type: "array" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        } else {
          const text = new TextDecoder().decode(ev.target.result);
          const lines = text.trim().split("\n"); if (lines.length < 2) return;
          const sep = lines[0].includes(";") ? ";" : ",";
          const headers = lines[0].split(sep).map(h => h.trim());
          rows = lines.slice(1).map(line => {
            const vals = line.split(sep); const obj = {};
            headers.forEach((h, i) => { obj[h] = (vals[i] || "").trim(); });
            return obj;
          });
        }
        if (!rows.length) { setMsg("gen", "❌ Archivo vacío."); return; }
        const cols = Object.keys(rows[0]);
        setGenericData(rows); setGenericCols(cols);
        setColMap({ name: "", date: "", metric1: "", metric1label: "", metric2: "", metric2label: "", metric3: "", metric3label: "", team: "" });
        setShowMapper(true);
        setMsg("gen", `✓ ${rows.length} filas · ${cols.length} columnas. Mapeá las columnas abajo.`);
      } catch { setMsg("gen", "❌ Error al leer el archivo."); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  }

  function applyMapping() {
    if (!genericData || !colMap.name) return;
    let cur = [...players]; const newRecs = [];
    genericData.forEach(row => {
      const name = String(row[colMap.name] || "").trim(); if (!name) return;
      const team = colMap.team ? String(row[colMap.team] || "").trim() : "";
      const { players: up, id: pid } = mergePlayer(name, team, cur); cur = up;
      const dateRaw = colMap.date ? String(row[colMap.date] || "") : "";
      const fecha = dateRaw || new Date().toISOString().slice(0, 10);
      const rec = { id: uid(), playerId: pid, date: fecha.slice(0, 7), fullDate: fecha, source: "custom", team };
      if (colMap.metric1 && colMap.metric1label) rec[colMap.metric1label] = parseNum(row[colMap.metric1]);
      if (colMap.metric2 && colMap.metric2label) rec[colMap.metric2label] = parseNum(row[colMap.metric2]);
      if (colMap.metric3 && colMap.metric3label) rec[colMap.metric3label] = parseNum(row[colMap.metric3]);
      newRecs.push(rec);
    });
    setPlayers(cur); setCustomRecs(r => [...r, ...newRecs]);
    if (!selPlayer && cur.length) { setSelPlayer(cur[0].id); setCompareA(cur[0].id); setCompareB(cur[1]?.id || cur[0].id); }
    setShowMapper(false);
    setMsg("gen", `✓ ${newRecs.length} registros importados`);
    setTab("dashboard");
  }

  // ── Derived data ──────────────────────────────────────────────────────────
  const hasData = players.length > 0;
  const teams   = useMemo(() => ["Todos", ...Array.from(new Set(players.map(p => p.team).filter(Boolean)))], [players]);
  const visible = useMemo(() => players.filter(p => filterTeam === "Todos" || p.team === filterTeam), [players, filterTeam]);

  const pJR = (pid) => jumpRecs.filter(r => r.playerId === pid).sort((a, b) => a.fullDate.localeCompare(b.fullDate));
  const pVR = (pid) => velRecs.filter(r => r.playerId === pid).sort((a, b) => a.fullDate.localeCompare(b.fullDate));
  const pRR = (pid) => romRecs.filter(r => r.playerId === pid).sort((a, b) => a.fullDate.localeCompare(b.fullDate));
  const pLR = (pid) => liftRecs.filter(r => r.playerId === pid).sort((a, b) => a.fullDate.localeCompare(b.fullDate));

  const bestJump = (pid) => { const rs = pJR(pid).filter(r => r.altura != null); return rs.reduce((b, r) => r.altura > (b?.altura ?? 0) ? r : b, null); };
  const bestJumpByType = (pid, type) => { const rs = pJR(pid).filter(r => r.jumpType === type && r.altura != null); return rs.reduce((b, r) => r.altura > (b?.altura ?? -Infinity) ? r : b, null); };
  const bestRsi = (pid) => { const rs = pJR(pid).filter(r => r.rsi != null); return rs.reduce((b, r) => r.rsi > (b?.rsi ?? -Infinity) ? r : b, null); };
  // Detecta a qué referencia de movilidad corresponde el nombre del test (Dorsiflexión / Cadera).
  function matchRomRef(testName) {
    if (!testName) return null; const t = testName.toLowerCase();
    if (t.includes("dorsiflex")) return "dorsiflexion";
    if (t.includes("cadera") || t.includes("rotaci")) return "hipIR";
    return null;
  }
  const bestVel  = (pid, dist) => { const rs = pVR(pid).filter(r => r.distance === dist && r.time != null); return rs.reduce((b, r) => r.time < (b?.time ?? 999) ? r : b, null); };
  const lastJump = (pid) => { const rs = pJR(pid).sort((a, b) => b.fullDate.localeCompare(a.fullDate)); return rs[0] || null; };
  const allDist  = useMemo(() => Array.from(new Set(velRecs.map(r => r.distance).filter(Boolean))).sort(), [velRecs]);
  const romTests = useMemo(() => Array.from(new Set(romRecs.map(r => r.test).filter(Boolean))).sort(), [romRecs]);
  const liftExercises = useMemo(() => Array.from(new Set(liftRecs.map(r => r.exercise).filter(Boolean))).sort(), [liftRecs]);

  // Most recent mobility record per player (any test) — used as risk indicator.
  const lastAsym = (pid) => { const rs = pRR(pid).filter(r => r.asim != null); return rs[rs.length - 1] || null; };
  // Best (max) 1RM for a given exercise-family, with its bodyweight for relative strength.
  const bestRM = (pid, exKey) => {
    const rs = pLR(pid).filter(r => r.exKey === exKey && r.rm != null);
    return rs.reduce((b, r) => r.rm > (b?.rm ?? 0) ? r : b, null);
  };
  const relStrength = (pid, exKey) => {
    const r = bestRM(pid, exKey);
    if (!r || !r.bodyweight) return null;
    return r.rm / r.bodyweight;
  };

  // ── Comparación directa dentro del plantel (percentiles, sin referencias fijas) ──
  // percentileRank: % del grupo al que el atleta supera en esa métrica (100 = mejor del grupo).
  function percentileRank(list, targetPid, lowerIsBetter) {
    if (!list || list.length < 2) return null;
    const target = list.find(x => x.pid === targetPid);
    if (!target || target.value == null) return null;
    const better = list.filter(x => x.pid !== targetPid && (lowerIsBetter ? x.value > target.value : x.value < target.value)).length;
    return Math.round((better / (list.length - 1)) * 100);
  }
  const jumpGroup = useMemo(() => visible.map(p => ({ pid: p.id, value: bestJump(p.id)?.altura ?? null })).filter(x => x.value != null),
    [players, jumpRecs, filterTeam]);
  const rsiGroup = useMemo(() => visible.map(p => ({ pid: p.id, value: bestJump(p.id)?.rsi ?? null })).filter(x => x.value != null),
    [players, jumpRecs, filterTeam]);
  const distGroups = useMemo(() => Object.fromEntries(allDist.map(d => [d, visible.map(p => ({ pid: p.id, value: bestVel(p.id, d)?.time ?? null })).filter(x => x.value != null)])),
    [players, velRecs, filterTeam, allDist]);
  const strengthGroups = useMemo(() => Object.fromEntries(Object.keys(STRENGTH_REFS).map(k => [k, visible.map(p => ({ pid: p.id, value: relStrength(p.id, k) })).filter(x => x.value != null)])),
    [players, liftRecs, filterTeam]);
  const asymGroup = useMemo(() => visible.map(p => ({ pid: p.id, value: lastAsym(p.id)?.asim ?? null })).filter(x => x.value != null),
    [players, romRecs, filterTeam]);

  // Promedio de percentiles disponibles para un atleta — usado para ordenar el plantel.
  const avgPercentile = (pid) => {
    const pcts = [];
    const jp = percentileRank(jumpGroup, pid, false); if (jp != null) pcts.push(jp);
    const rp = percentileRank(rsiGroup, pid, false); if (rp != null) pcts.push(rp);
    allDist.forEach(d => { const vp = percentileRank(distGroups[d], pid, true); if (vp != null) pcts.push(vp); });
    Object.keys(STRENGTH_REFS).forEach(k => { const sp = percentileRank(strengthGroups[k], pid, false); if (sp != null) pcts.push(sp); });
    return pcts.length ? Math.round(pcts.reduce((a, b) => a + b, 0) / pcts.length) : null;
  };

  const rankingData = useMemo(() =>
    visible.filter(p => (pJR(p.id).length + pVR(p.id).length + pLR(p.id).length) > 0)
      .map(p => ({ id: p.id, name: p.name.split(" ")[0], fullName: p.name, pct: avgPercentile(p.id) ?? 0 }))
      .sort((a, b) => b.pct - a.pct),
    [players, jumpRecs, velRecs, liftRecs, filterTeam]);

  const jumpEvo = (pid) => {
    const byDate = {};
    pJR(pid).forEach(r => {
      if (!byDate[r.fullDate]) byDate[r.fullDate] = { fecha: r.fullDate.slice(8) + "/" + r.fullDate.slice(5, 7), _r: [] };
      byDate[r.fullDate]._r.push(r);
    });
    return Object.values(byDate).map(d => {
      const cmjs = d._r.filter(r => r.jumpType === "cmj" && r.altura != null);
      const djs  = d._r.filter(r => r.jumpType === "dj"  && r.altura != null);
      const rsif = d._r.filter(r => r.rsi != null);
      return { fecha: d.fecha, CMJ: cmjs.length ? Math.max(...cmjs.map(r => r.altura)) : null, DJ: djs.length ? Math.max(...djs.map(r => r.altura)) : null, RSI: rsif.length ? Math.max(...rsif.map(r => r.rsi)) : null };
    });
  };

  const velEvo = (pid) => {
    const byDate = {};
    pVR(pid).forEach(r => {
      if (!byDate[r.fullDate]) byDate[r.fullDate] = { fecha: r.fullDate.slice(8) + "/" + r.fullDate.slice(5, 7) };
      if (!byDate[r.fullDate][r.distance] || r.time < byDate[r.fullDate][r.distance]) byDate[r.fullDate][r.distance] = r.time;
    });
    return Object.values(byDate);
  };

  // Asimetría (%) por fecha, una serie por tipo de test presente.
  const romEvo = (pid) => {
    const byDate = {};
    pRR(pid).forEach(r => {
      if (!byDate[r.fullDate]) byDate[r.fullDate] = { fecha: r.fullDate.slice(8) + "/" + r.fullDate.slice(5, 7) };
      if (r.asim != null) byDate[r.fullDate][r.test || "Test"] = r.asim;
    });
    return Object.values(byDate);
  };

  // 1RM por fecha para un ejercicio dado (toma el máximo si hay más de un test el mismo día).
  const liftRmEvo = (pid, exercise) => {
    const byDate = {};
    pLR(pid).filter(r => r.exercise === exercise && r.rm != null).forEach(r => {
      const key = r.fullDate;
      if (!byDate[key] || r.rm > byDate[key].rm) byDate[key] = { fecha: r.fullDate.slice(8) + "/" + r.fullDate.slice(5, 7), rm: r.rm };
    });
    return Object.values(byDate).sort((a, b) => a.fecha.localeCompare(b.fecha));
  };

  // Perfil carga-velocidad del test más reciente de un ejercicio.
  const latestLiftProfile = (pid, exercise) => {
    const rs = pLR(pid).filter(r => r.exercise === exercise);
    if (!rs.length) return { points: [], date: null };
    const latest = rs[rs.length - 1];
    return { points: latest.points, date: latest.fullDate };
  };

  // ── Styles ────────────────────────────────────────────────────────────────
  const panel = { background: F.panel, border: `1px solid ${F.panelBorder}`, borderRadius: 4, padding: 18, marginBottom: 14 };
  const sel   = { background: F.ghost, border: `1px solid ${F.panelBorder}`, borderRadius: 3, color: F.white, padding: "6px 10px", fontSize: 12, outline: "none", fontFamily: F.fontMono, cursor: "pointer" };
  const lbl   = { fontSize: 9, color: F.dim, letterSpacing: 2, marginBottom: 4, display: "block", textTransform: "uppercase", fontFamily: F.fontMono };
  const inp   = { ...sel, width: "100%", boxSizing: "border-box" };
  const dangerBtn = { padding: "6px 14px", background: "transparent", border: `1px solid ${F.red}`, borderRadius: 3, color: F.red, fontFamily: F.fontMono, fontSize: 10, letterSpacing: 2, cursor: "pointer" };

  // ── Import View ───────────────────────────────────────────────────────────
  const ImportView = () => (
    <div style={{ maxWidth: 980, margin: "0 auto" }}>
      <div style={{ marginBottom: 20, paddingBottom: 16, borderBottom: `1px solid ${F.panelBorder}` }}>
        <div style={{ fontFamily: F.fontMono, fontSize: 11, color: F.red, letterSpacing: 3, marginBottom: 6 }}>SISTEMA DE IMPORTACIÓN</div>
        <div style={{ fontFamily: F.fontMono, fontSize: 13, color: F.silver }}>Subí archivos desde cualquier app de evaluación</div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12, marginBottom: 16 }}>
        {[
          { key:"jump", label:"MYJUMP LAB", sub:"CMJ / SJ / DJ / RSI", ref:jumpRef, onChange:handleJumpImport, accent:F.red,    ext:".csv,.txt", details:["CMJ, SJ, DJ","RSI mod","Fuerza / Potencia","Estado de forma"] },
          { key:"vel",  label:"PHOTOFINISH",sub:"10m / 30m / 60m",      ref:velRef,  onChange:handleVelImport,  accent:F.teal,   ext:".csv,.txt", details:["Tiempos por distancia","Detección automática","Múltiples atletas"] },
          { key:"rom",  label:"MOVILIDAD",  sub:"MyROM · Asimetrías",   ref:romRef,  onChange:handleRomImport,  accent:F.purple, ext:".csv,.txt", details:["Ángulo izq / drch","Asimetría %","Múltiples tests"] },
          { key:"lift", label:"VBT / RM",   sub:"MyLift · Carga-velocidad", ref:liftRef, onChange:handleLiftImport, accent:F.orange, ext:".csv,.txt", details:["1-RM estimado","Perfil carga-velocidad","Fuerza relativa"] },
          { key:"gen",  label:"EXCEL / CSV", sub:"Cualquier planilla",   ref:genericRef,onChange:handleGenericImport,accent:F.yellow,ext:".xlsx,.xls,.csv",details:["Cualquier app","Mapeás columnas vos","Múltiples métricas"] },
        ].map(({ key, label, sub, ref: fref, onChange, accent, ext, details }) => (
          <div key={key} style={{ ...panel, cursor: "pointer", borderTop: `2px solid ${accent}` }}
            onClick={() => fref.current?.click()}>
            <div style={{ fontFamily: F.fontMono, fontSize: 10, color: accent, letterSpacing: 2, marginBottom: 4 }}>{label}</div>
            <div style={{ fontFamily: F.fontMono, fontSize: 11, color: F.silver, marginBottom: 12 }}>{sub}</div>
            <div style={{ border: `1px dashed ${F.ghost}`, borderRadius: 3, padding: "16px 10px", textAlign: "center", marginBottom: 10, background: F.asphalt }}>
              <div style={{ fontSize: 20, color: accent, marginBottom: 4 }}>↑</div>
              <div style={{ fontFamily: F.fontMono, fontSize: 10, color: F.dim, letterSpacing: 1 }}>SUBIR ARCHIVO</div>
              <input ref={fref} type="file" accept={ext} style={{ display: "none" }} onChange={onChange} />
            </div>
            {details.map(d => (
              <div key={d} style={{ fontSize: 10, color: F.dim, fontFamily: F.fontMono, padding: "2px 0", borderBottom: `1px solid ${F.ghost}` }}>
                <span style={{ color: accent, marginRight: 5 }}>—</span>{d}
              </div>
            ))}
            {msgs[key] && (
              <div style={{ marginTop: 8, fontSize: 10, color: msgs[key].startsWith("✓") ? F.green : F.red, fontFamily: F.fontMono, padding: "4px 8px", background: msgs[key].startsWith("✓") ? F.green + "14" : F.red + "14", borderRadius: 3 }}>
                {msgs[key]}
              </div>
            )}
          </div>
        ))}
      </div>

      {/* Column mapper */}
      {showMapper && genericData && (
        <div style={{ ...panel, borderTop: `2px solid ${F.yellow}` }}>
          <PanelHeader accent={F.yellow}>MAPEO DE COLUMNAS — {genericData.length} FILAS</PanelHeader>
          <div style={{ marginBottom: 10, fontFamily: F.fontMono, fontSize: 10, color: F.dim }}>
            Columnas: {genericCols.join(" · ")}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10, marginBottom: 14 }}>
            {[
              { k:"name",        label:"NOMBRE (obligatorio)" },
              { k:"date",        label:"FECHA" },
              { k:"team",        label:"EQUIPO" },
              { k:"metric1",     label:"MÉTRICA 1 — columna" },
              { k:"metric1label",label:"MÉTRICA 1 — nombre" },
              { k:"metric2",     label:"MÉTRICA 2 — columna" },
              { k:"metric2label",label:"MÉTRICA 2 — nombre" },
              { k:"metric3",     label:"MÉTRICA 3 — columna" },
              { k:"metric3label",label:"MÉTRICA 3 — nombre" },
            ].map(({ k, label }) => (
              <div key={k}>
                <label style={lbl}>{label}</label>
                {k.includes("label")
                  ? <input style={inp} placeholder="ej: CMJ, RPE..." value={colMap[k] || ""} onChange={e => setColMap(m => ({ ...m, [k]: e.target.value }))} />
                  : <select style={{ ...sel, width: "100%" }} value={colMap[k] || ""} onChange={e => setColMap(m => ({ ...m, [k]: e.target.value }))}>
                      <option value="">— ninguna —</option>
                      {genericCols.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                }
              </div>
            ))}
          </div>
          <div style={{ marginBottom: 14, overflowX: "auto" }}>
            <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 2, marginBottom: 6 }}>PREVIEW — 3 FILAS</div>
            <table style={{ borderCollapse: "collapse", fontSize: 10, fontFamily: F.fontMono }}>
              <thead><tr>{genericCols.map(c => <th key={c} style={{ padding: "4px 10px", color: F.red, borderBottom: `1px solid ${F.panelBorder}`, whiteSpace: "nowrap" }}>{c}</th>)}</tr></thead>
              <tbody>{genericData.slice(0, 3).map((r, i) => (
                <tr key={i} style={{ borderBottom: `1px solid ${F.ghost}` }}>
                  {genericCols.map(c => <td key={c} style={{ padding: "4px 10px", color: F.silver, whiteSpace: "nowrap" }}>{String(r[c]).slice(0, 20)}</td>)}
                </tr>
              ))}</tbody>
            </table>
          </div>
          <button onClick={applyMapping} style={{ padding: "8px 24px", background: F.yellow, color: F.carbon, border: "none", borderRadius: 3, fontFamily: F.fontMono, fontSize: 11, fontWeight: 700, letterSpacing: 2, cursor: "pointer" }}>
            IMPORTAR CON ESTE MAPEO
          </button>
        </div>
      )}

      {hasData && (
        <div style={{ ...panel, borderTop: `2px solid ${F.green}`, display: "flex", alignItems: "center", gap: 20, flexWrap: "wrap" }}>
          <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.green, letterSpacing: 3 }}>DATOS EN SISTEMA</div>
          {[{ v: players.length, l: "ATLETAS" }, { v: jumpRecs.length, l: "SALTOS" }, { v: velRecs.length, l: "VEL." }, { v: romRecs.length, l: "MOVIL." }, { v: liftRecs.length, l: "VBT/RM" }, { v: customRecs.length, l: "CUSTOM" }].map(({ v, l }) => (
            <div key={l} style={{ textAlign: "center" }}>
              <div style={{ fontFamily: F.fontMono, fontSize: 18, color: F.white, fontWeight: 700 }}>{v}</div>
              <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 2 }}>{l}</div>
            </div>
          ))}
          <button onClick={() => setTab("dashboard")} style={{ marginLeft: "auto", padding: "8px 20px", background: F.red, color: F.white, border: "none", borderRadius: 3, fontFamily: F.fontMono, fontSize: 11, letterSpacing: 2, cursor: "pointer", fontWeight: 700 }}>
            VER DASHBOARD →
          </button>
        </div>
      )}
    </div>
  );

  // ── Delete Panel ──────────────────────────────────────────────────────────
  const DeletePanel = () => (
    <div style={{ ...panel, borderTop: `2px solid ${F.red}`, maxWidth: 700, margin: "0 auto" }}>
      <PanelHeader accent={F.red}>GESTIÓN DE DATOS</PanelHeader>
      {msgs.del && <div style={{ fontFamily: F.fontMono, fontSize: 11, color: F.green, marginBottom: 14, padding: "6px 10px", background: F.green + "14", borderRadius: 3 }}>{msgs.del}</div>}

      <div style={{ marginBottom: 20 }}>
        <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 2, marginBottom: 10 }}>BORRAR POR TIPO</div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button onClick={deleteJumpRecs}   style={dangerBtn}>BORRAR SALTOS ({jumpRecs.length})</button>
          <button onClick={deleteVelRecs}    style={dangerBtn}>BORRAR VELOCIDAD ({velRecs.length})</button>
          <button onClick={deleteRomRecs}    style={dangerBtn}>BORRAR MOVILIDAD ({romRecs.length})</button>
          <button onClick={deleteLiftRecs}   style={dangerBtn}>BORRAR VBT/RM ({liftRecs.length})</button>
          <button onClick={deleteCustomRecs} style={dangerBtn}>BORRAR CUSTOM ({customRecs.length})</button>
        </div>
      </div>

      <div style={{ marginBottom: 20 }}>
        <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 2, marginBottom: 10 }}>BORRAR POR ATLETA</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {players.map(p => (
            <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "6px 10px", background: F.asphalt, borderRadius: 3 }}>
              <div style={{ fontFamily: F.fontMono, fontSize: 12, color: F.white, flex: 1 }}>{p.name}</div>
              <div style={{ fontFamily: F.fontMono, fontSize: 10, color: F.dim }}>
                {pJR(p.id).length} saltos · {pVR(p.id).length} vel. · {pRR(p.id).length} movil. · {pLR(p.id).length} vbt
              </div>
              <button onClick={() => deletePlayer(p.id)} style={{ ...dangerBtn, padding: "4px 10px" }}>BORRAR</button>
            </div>
          ))}
        </div>
      </div>

      <div style={{ borderTop: `1px solid ${F.panelBorder}`, paddingTop: 16 }}>
        <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 2, marginBottom: 10 }}>ZONA DE PELIGRO</div>
        <button onClick={deleteAllData} style={{ ...dangerBtn, background: F.red + "22", fontWeight: 700 }}>
          ⚠ BORRAR TODO Y EMPEZAR DE CERO
        </button>
      </div>
    </div>
  );

  // ── Dashboard ─────────────────────────────────────────────────────────────
  const Dashboard = () => {
    const vis = visible.filter(p => (pJR(p.id).length + pVR(p.id).length + pLR(p.id).length) > 0);
    return (
      <div>
        <div style={{ display: "flex", gap: 12, marginBottom: 16, alignItems: "flex-end", flexWrap: "wrap" }}>
          <div>
            <label style={lbl}>EQUIPO</label>
            <select style={sel} value={filterTeam} onChange={e => setFilterTeam(e.target.value)}>
              {teams.map(t => <option key={t}>{t}</option>)}
            </select>
          </div>
          <div style={{ marginLeft: "auto", display: "flex", gap: 8, flexWrap: "wrap" }}>
            {[{ v: vis.length, l: "ATLETAS", c: F.red }, { v: jumpRecs.length, l: "SALTOS", c: F.teal }, { v: velRecs.length, l: "VEL", c: F.yellow }, { v: romRecs.length, l: "MOVIL", c: F.purple }, { v: liftRecs.length, l: "VBT", c: F.orange }].map(({ v, l, c }) => (
              <div key={l} style={{ ...panel, marginBottom: 0, padding: "8px 14px", textAlign: "center", borderTop: `2px solid ${c}` }}>
                <div style={{ fontFamily: F.fontMono, fontSize: 20, color: F.white, fontWeight: 700 }}>{v}</div>
                <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 2 }}>{l}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Timing board ranking */}
        <div style={panel}>
          <PanelHeader>COMPARATIVA DEL PLANTEL <span style={{ color: F.dim, fontSize: 9 }}>(percentil relativo al grupo cargado — salto + velocidad + fuerza relativa)</span></PanelHeader>
          {rankingData.map((d, i) => {
            const col = TEAM_COLORS[i % TEAM_COLORS.length]; const barCol = percentileColor(d.pct);
            return (
              <div key={d.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "8px 0", borderBottom: `1px solid ${F.ghost}`, cursor: "pointer" }}
                onClick={() => { setSelPlayer(d.id); goTo("evolution"); }}>
                <div style={{ fontFamily: F.fontMono, fontSize: 11, color: F.dim, width: 24, textAlign: "right" }}>{String(i + 1).padStart(2, "0")}</div>
                <div style={{ width: 4, height: 32, background: col, borderRadius: 1 }} />
                <div style={{ fontFamily: F.fontMono, fontSize: 13, color: F.white, flex: 1, letterSpacing: 1 }}>{d.fullName}</div>
                <PercentileChip pct={d.pct} />
                <div style={{ width: 120, height: 4, background: F.ghost, borderRadius: 2, overflow: "hidden" }}>
                  <div style={{ width: `${d.pct}%`, height: "100%", background: barCol, borderRadius: 2, boxShadow: `0 0 6px ${barCol}` }} />
                </div>
                <div style={{ fontFamily: F.fontMono, fontSize: 14, color: barCol, width: 36, textAlign: "right", fontWeight: 700 }}>{d.pct}</div>
              </div>
            );
          })}
        </div>

        {/* Player cards */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(200px,1fr))", gap: 10, marginBottom: 14 }}>
          {vis.map((p, i) => {
            const pct = avgPercentile(p.id); const bj = bestJump(p.id);
            const bv = allDist.length ? bestVel(p.id, allDist[0]) : null; const col = TEAM_COLORS[i % TEAM_COLORS.length];
            const asym = lastAsym(p.id); const asymPct = asym ? percentileRank(asymGroup, p.id, true) : null;
            const risk = asymPct != null && asymPct < 25;
            return (
              <div key={p.id} style={{ ...panel, marginBottom: 0, cursor: "pointer", borderLeft: `3px solid ${col}` }}
                onClick={() => { setSelPlayer(p.id); goTo("evolution"); }}
                onMouseEnter={e => e.currentTarget.style.background = F.ghost}
                onMouseLeave={e => e.currentTarget.style.background = F.panel}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
                  <div>
                    <div style={{ fontFamily: F.fontMono, fontSize: 13, color: F.white, letterSpacing: 1 }}>{p.name.split(" ")[0].toUpperCase()}</div>
                    <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 1, marginTop: 2 }}>{p.name.split(" ").slice(1).join(" ").toUpperCase()}</div>
                    <div style={{ fontFamily: F.fontMono, fontSize: 9, color: col, letterSpacing: 1, marginTop: 3 }}>{p.team || "—"}</div>
                  </div>
                  <TimingScore score={pct} size={52} />
                </div>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {bj && <div style={{ background: F.ghost, borderRadius: 2, padding: "3px 7px", fontFamily: F.fontMono, fontSize: 10 }}>
                    <span style={{ color: F.dim }}>{bj.jumpTypeRaw} </span>
                    <span style={{ color: F.teal }}>{bj.altura?.toFixed(1)}cm</span>
                  </div>}
                  {bv && <div style={{ background: F.ghost, borderRadius: 2, padding: "3px 7px", fontFamily: F.fontMono, fontSize: 10 }}>
                    <span style={{ color: F.dim }}>{allDist[0]} </span>
                    <span style={{ color: F.yellow }}>{bv.time?.toFixed(2)}s</span>
                  </div>}
                  {risk && <div style={{ background: F.red + "18", border: `1px solid ${F.red}`, borderRadius: 2, padding: "3px 7px", fontFamily: F.fontMono, fontSize: 10, color: F.red }}>
                    ⚠ ASIMETRÍA {asym.asim?.toFixed(1)}% <span style={{ color: F.dim }}>(peor cuarto del plantel)</span>
                  </div>}
                </div>
              </div>
            );
          })}
        </div>

        {allDist.length > 0 && (
          <div style={panel}>
            <PanelHeader accent={F.yellow}>MEJORES TIEMPOS</PanelHeader>
            <ResponsiveContainer width="100%" height={Math.max(140, vis.length * 24)}>
              <BarChart data={vis.filter(p => pVR(p.id).length > 0).map(p => ({ name: p.name.split(" ")[0], ...Object.fromEntries(allDist.map(d => [d, bestVel(p.id, d)?.time ?? null])) }))} layout="vertical">
                <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} horizontal={false} />
                <XAxis type="number" tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} unit="s" />
                <YAxis dataKey="name" type="category" tick={{ fill: F.silver, fontSize: 10, fontFamily: F.fontMono }} width={70} />
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize: 10, fontFamily: F.fontMono }} />
                {allDist.map((d, i) => <Bar key={d} dataKey={d} fill={TEAM_COLORS[i]} radius={[0, 3, 3, 0]} />)}
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>
    );
  };

  // ── Evolution ─────────────────────────────────────────────────────────────
  const Evolution = () => {
    const activePlayers = players.filter(p => (pJR(p.id).length + pVR(p.id).length + pRR(p.id).length + pLR(p.id).length) > 0);
    const pid = selPlayer && activePlayers.find(p => p.id === selPlayer) ? selPlayer : activePlayers[0]?.id;
    const jevo = pid ? jumpEvo(pid) : []; const vevo = pid ? velEvo(pid) : []; const revo = pid ? romEvo(pid) : [];
    const bj = pid ? bestJump(pid) : null; const jrecs = pid ? pJR(pid) : []; const vrecs = pid ? pVR(pid) : [];
    const rrecs = pid ? pRR(pid) : []; const lrecs = pid ? pLR(pid) : [];
    const p = players.find(x => x.id === pid);
    const playerExercises = useMemo(() => Array.from(new Set(lrecs.map(r => r.exercise))), [pid, liftRecs]);
    const curExercise = liftExercise && playerExercises.includes(liftExercise) ? liftExercise : playerExercises[0];
    const rmEvo = curExercise ? liftRmEvo(pid, curExercise) : [];
    const profile = curExercise ? latestLiftProfile(pid, curExercise) : { points: [], date: null };
    const curRelStrength = curExercise ? relStrength(pid, exerciseKey(curExercise)) : null;

    return (
      <div>
        <BackBtn onClick={goBack} />
        <div style={{ display: "flex", gap: 12, marginBottom: 16, alignItems: "flex-end", flexWrap: "wrap" }}>
          <div>
            <label style={lbl}>ATLETA</label>
            <select style={sel} value={pid || ""} onChange={e => setSelPlayer(e.target.value)}>
              {activePlayers.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          </div>
          {p && <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 2 }}>{p.team || ""}</div>}
          <button onClick={() => goTo("informe")} style={{ marginLeft: "auto", padding: "8px 16px", background: "transparent", border: `1px solid ${F.yellow}`, borderRadius: 3, color: F.yellow, fontFamily: F.fontMono, fontSize: 10, letterSpacing: 2, cursor: "pointer", fontWeight: 700 }}>
            📄 DESCARGAR INFORME
          </button>
        </div>

        <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
          {[
            { v: jrecs.length + vrecs.length + rrecs.length + lrecs.length, l: "REGISTROS",    c: F.silver },
            { v: bj?.altura?.toFixed(1) + "cm" ?? "—",           l: "MEJOR SALTO", c: F.teal },
            { v: bj?.rsi?.toFixed(2) ?? "—",                     l: "MEJOR RSI",   c: F.yellow },
            ...allDist.map(d => ({ v: bestVel(pid, d)?.time?.toFixed(2) + "s" ?? "—", l: d, c: F.red })),
          ].map(({ v, l, c }) => (
            <div key={l} style={{ ...panel, marginBottom: 0, padding: "8px 14px", flex: 1, minWidth: 80, borderTop: `2px solid ${c}` }}>
              <div style={{ fontFamily: F.fontMono, fontSize: 16, color: F.white, fontWeight: 700 }}>{v}</div>
              <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 2 }}>{l}</div>
            </div>
          ))}
        </div>

        {jevo.length > 0 && (
          <div style={panel}>
            <PanelHeader>ALTURA DE SALTO (cm)</PanelHeader>
            <ResponsiveContainer width="100%" height={200}>
              <LineChart data={jevo}>
                <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} />
                <YAxis tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} unit="cm" />
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize: 10, fontFamily: F.fontMono }} />
                <Line dataKey="CMJ" stroke={F.teal} strokeWidth={2} dot={{ r: 4, fill: F.teal }} connectNulls />
                <Line dataKey="DJ"  stroke={F.red}  strokeWidth={2} dot={{ r: 4, fill: F.red  }} connectNulls />
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}

        {vevo.length > 0 && allDist.length > 0 && (
          <div style={panel}>
            <PanelHeader accent={F.yellow}>VELOCIDAD (s)</PanelHeader>
            <ResponsiveContainer width="100%" height={200}>
              <LineChart data={vevo}>
                <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} />
                <YAxis tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} unit="s" />
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize: 10, fontFamily: F.fontMono }} />
                {allDist.map((d, i) => <Line key={d} dataKey={d} stroke={TEAM_COLORS[i]} strokeWidth={2} dot={{ r: 4 }} connectNulls />)}
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}

        {jevo.some(d => d.RSI) && (
          <div style={panel}>
            <PanelHeader accent={F.yellow}>RSI MOD</PanelHeader>
            <ResponsiveContainer width="100%" height={160}>
              <LineChart data={jevo}>
                <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} />
                <YAxis tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} />
                <Tooltip content={<CustomTooltip />} />
                <Line dataKey="RSI" stroke={F.yellow} strokeWidth={2} dot={{ r: 4, fill: F.yellow }} connectNulls />
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}

        {/* Movilidad */}
        {revo.length > 0 && (
          <div style={panel}>
            <PanelHeader accent={F.purple}>MOVILIDAD — ASIMETRÍA (%)</PanelHeader>
            <ResponsiveContainer width="100%" height={180}>
              <LineChart data={revo}>
                <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} />
                <YAxis tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} unit="%" />
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize: 10, fontFamily: F.fontMono }} />
                {romTests.map((t, i) => <Line key={t} dataKey={t} stroke={[F.purple, F.orange, F.teal, F.yellow][i % 4]} strokeWidth={2} dot={{ r: 4 }} connectNulls />)}
              </LineChart>
            </ResponsiveContainer>
            <div style={{ marginTop: 10, fontFamily: F.fontMono, fontSize: 9, color: F.dim }}>
              Sin referencia fija cargada todavía — comparado contra el resto del plantel en la columna PERCENTIL.
            </div>
          </div>
        )}
        {rrecs.length > 0 && (
          <div style={panel}>
            <PanelHeader accent={F.purple}>HISTORIAL MOVILIDAD</PanelHeader>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                <thead><tr style={{ borderBottom: `1px solid ${F.panelBorder}` }}>
                  {["FECHA", "TEST", "ÁNG. IZQ", "ÁNG. DRCH", "ASIMETRÍA", "PERCENTIL"].map(h => (
                    <th key={h} style={{ padding: "5px 10px", color: F.dim, textAlign: "left", fontFamily: F.fontMono, fontSize: 9, letterSpacing: 2 }}>{h}</th>
                  ))}
                </tr></thead>
                <tbody>{rrecs.map((r, i) => {
                  const isLatest = i === rrecs.length - 1;
                  const pct = isLatest ? percentileRank(asymGroup, pid, true) : null;
                  return <tr key={i} style={{ borderBottom: `1px solid ${F.ghost}` }}>
                    <td style={{ padding: "6px 10px", color: F.teal, fontFamily: F.fontMono, fontSize: 11 }}>{r.fullDate}</td>
                    <td style={{ padding: "6px 10px", color: F.silver, fontFamily: F.fontMono, fontSize: 11 }}>{r.test}</td>
                    <td style={{ padding: "6px 10px", color: F.silver, fontFamily: F.fontMono, fontSize: 11 }}>{r.angI?.toFixed(1) ?? "—"}º</td>
                    <td style={{ padding: "6px 10px", color: F.silver, fontFamily: F.fontMono, fontSize: 11 }}>{r.angD?.toFixed(1) ?? "—"}º</td>
                    <td style={{ padding: "6px 10px", fontFamily: F.fontMono, fontSize: 13, color: F.white, fontWeight: 700 }}>{r.asim?.toFixed(1)}%</td>
                    <td style={{ padding: "6px 10px" }}>{isLatest ? <PercentileChip pct={pct} /> : <span style={{ color: F.ghost, fontFamily: F.fontMono, fontSize: 9 }}>—</span>}</td>
                  </tr>;
                })}</tbody>
              </table>
            </div>
          </div>
        )}

        {/* VBT / RM */}
        {lrecs.length > 0 && (
          <div style={panel}>
            <PanelHeader accent={F.orange}>VBT / RM {playerExercises.length > 1 && (
              <select style={{ ...sel, marginLeft: 10 }} value={curExercise} onChange={e => setLiftExercise(e.target.value)}>
                {playerExercises.map(ex => <option key={ex} value={ex}>{ex}</option>)}
              </select>
            )}</PanelHeader>
            <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
              <div style={{ ...panel, marginBottom: 0, padding: "8px 14px", borderTop: `2px solid ${F.orange}` }}>
                <div style={{ fontFamily: F.fontMono, fontSize: 16, color: F.white, fontWeight: 700 }}>{rmEvo[rmEvo.length - 1]?.rm?.toFixed(1) ?? "—"} kg</div>
                <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 2 }}>1RM ACTUAL</div>
              </div>
              {curRelStrength != null && (
                <div style={{ ...panel, marginBottom: 0, padding: "8px 14px", borderTop: `2px solid ${F.orange}` }}>
                  <div style={{ fontFamily: F.fontMono, fontSize: 16, color: F.white, fontWeight: 700 }}>{curRelStrength.toFixed(2)}x PC</div>
                  <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 2 }}>FUERZA RELATIVA</div>
                </div>
              )}
            </div>
            {rmEvo.length > 1 && (
              <ResponsiveContainer width="100%" height={180}>
                <LineChart data={rmEvo}>
                  <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                  <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} />
                  <YAxis tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} unit="kg" />
                  <Tooltip content={<CustomTooltip />} />
                  <Line dataKey="rm" name="1RM" stroke={F.orange} strokeWidth={2} dot={{ r: 4, fill: F.orange }} connectNulls />
                </LineChart>
              </ResponsiveContainer>
            )}
            {profile.points.length > 0 && (
              <div style={{ marginTop: 14 }}>
                <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 2, marginBottom: 8 }}>
                  PERFIL CARGA-VELOCIDAD · {profile.date}
                </div>
                <ResponsiveContainer width="100%" height={200}>
                  <ScatterChart>
                    <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                    <XAxis dataKey="load" name="Carga" unit="kg" tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} type="number" />
                    <YAxis dataKey="vel" name="Vel. media" unit="m/s" tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} type="number" />
                    <Tooltip content={<CustomTooltip />} cursor={{ strokeDasharray: "3 3" }} />
                    <Scatter data={profile.points} fill={F.orange} line={{ stroke: F.orange, strokeWidth: 1 }} shape="circle" />
                  </ScatterChart>
                </ResponsiveContainer>
              </div>
            )}
          </div>
        )}

        {vrecs.length > 0 && (
          <div style={panel}>
            <PanelHeader accent={F.yellow}>HISTORIAL VELOCIDAD</PanelHeader>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                <thead><tr style={{ borderBottom: `1px solid ${F.panelBorder}` }}>
                  {["FECHA", "TEST", "DIST", "TIEMPO", "PERCENTIL"].map(h => (
                    <th key={h} style={{ padding: "5px 10px", color: F.dim, textAlign: "left", fontFamily: F.fontMono, fontSize: 9, letterSpacing: 2 }}>{h}</th>
                  ))}
                </tr></thead>
                <tbody>{vrecs.map((r, i) => {
                  const isBest = r.time === bestVel(pid, r.distance)?.time;
                  const pct = isBest ? percentileRank(distGroups[r.distance], pid, true) : null;
                  return <tr key={i} style={{ borderBottom: `1px solid ${F.ghost}` }}>
                    <td style={{ padding: "6px 10px", color: F.teal, fontFamily: F.fontMono, fontSize: 11 }}>{r.fullDate}</td>
                    <td style={{ padding: "6px 10px", color: F.silver, fontFamily: F.fontMono, fontSize: 11 }}>{r.testName}</td>
                    <td style={{ padding: "6px 10px", color: F.silver, fontFamily: F.fontMono, fontSize: 11 }}>{r.distance || "—"}</td>
                    <td style={{ padding: "6px 10px", fontFamily: F.fontMono, fontSize: 13, color: F.white, fontWeight: 700 }}>{r.time?.toFixed(2)}s</td>
                    <td style={{ padding: "6px 10px" }}>{isBest ? <PercentileChip pct={pct} /> : <span style={{ color: F.ghost, fontFamily: F.fontMono, fontSize: 9 }}>—</span>}</td>
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
  const Compare = () => {
    const activePlayers = players.filter(p => (pJR(p.id).length + pVR(p.id).length + pRR(p.id).length + pLR(p.id).length) > 0);
    const pidA = compareA && activePlayers.find(p => p.id === compareA) ? compareA : activePlayers[0]?.id;
    const pidB = compareB && activePlayers.find(p => p.id === compareB) ? compareB : activePlayers[1]?.id || activePlayers[0]?.id;
    const nA = players.find(p => p.id === pidA)?.name.split(" ")[0] || "A";
    const nB = players.find(p => p.id === pidB)?.name.split(" ")[0] || "B";

    // Comparación directa dato-contra-dato: cada métrica trae su grupo (roster) para calcular
    // el percentil de A y B, sin ninguna referencia fija externa.
    const radarM = [
      { key:"cmj", label:"CMJ", vA:bestJump(pidA)?.altura, vB:bestJump(pidB)?.altura, unit:"cm", lowerIsBetter:false, group:jumpGroup },
      { key:"rsi", label:"RSI", vA:bestJump(pidA)?.rsi,    vB:bestJump(pidB)?.rsi,    unit:"", lowerIsBetter:false, group:rsiGroup },
      ...allDist.map(d => ({ key:d, label:d, vA:bestVel(pidA,d)?.time, vB:bestVel(pidB,d)?.time, unit:"s", lowerIsBetter:true, group:distGroups[d] })),
      ...Object.keys(STRENGTH_REFS).map(k => ({ key:k, label:STRENGTH_REFS[k].label, vA:relStrength(pidA,k), vB:relStrength(pidB,k), unit:"x PC", lowerIsBetter:false, group:strengthGroups[k] })),
    ];
    const radarD = radarM.map(m => ({ metric: m.label, [nA]: percentileRank(m.group, pidA, m.lowerIsBetter) ?? 0, [nB]: percentileRank(m.group, pidB, m.lowerIsBetter) ?? 0 }));

    const evoA = jumpEvo(pidA); const evoB = jumpEvo(pidB);
    const allF = Array.from(new Set([...evoA.map(d => d.fecha), ...evoB.map(d => d.fecha)])).sort();
    const merged = allF.map(f => ({ fecha: f, [nA]: evoA.find(d => d.fecha === f)?.CMJ ?? null, [nB]: evoB.find(d => d.fecha === f)?.CMJ ?? null }));

    const asymA = lastAsym(pidA); const asymB = lastAsym(pidB);
    const pctA = avgPercentile(pidA); const pctB = avgPercentile(pidB);

    return (
      <div>
        <BackBtn onClick={goBack} />
        <div style={{ display: "flex", gap: 12, marginBottom: 16, flexWrap: "wrap" }}>
          {[
            { val: pidA, set: setCompareA, label: "ATLETA A", col: F.teal },
            { val: pidB, set: setCompareB, label: "ATLETA B", col: F.red  },
          ].map(({ val, set, label, col }) => (
            <div key={label} style={{ flex: 1, minWidth: 160 }}>
              <label style={{ ...lbl, color: col }}>{label}</label>
              <select style={{ ...sel, width: "100%" }} value={val || ""}
                onChange={e => set(e.target.value)}>
                {activePlayers.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
            </div>
          ))}
        </div>

        {/* Head to head */}
        <div style={{ ...panel, display: "flex", alignItems: "center", gap: 0, marginBottom: 14, padding: 0, overflow: "hidden" }}>
          {[{ pid: pidA, pct: pctA, col: F.teal, name: nA }, { pid: pidB, pct: pctB, col: F.red, name: nB }].map(({ pid, pct, col, name }, i) => (
            <div key={pid} style={{ flex: 1, padding: "16px 20px", borderRight: i === 0 ? `1px solid ${F.panelBorder}` : "none", textAlign: "center" }}>
              <div style={{ fontFamily: F.fontMono, fontSize: 10, color: col, letterSpacing: 3, marginBottom: 10 }}>{name.toUpperCase()}</div>
              <TimingScore score={pct} size={72} />
              <div style={{ marginTop: 8 }}><PercentileChip pct={pct} /></div>
            </div>
          ))}
          <div style={{ padding: "16px 24px", textAlign: "center" }}>
            <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 2, marginBottom: 6 }}>DIFERENCIA</div>
            <div style={{ fontFamily: F.fontMono, fontSize: 28, color: pctA == null || pctB == null || Math.abs(pctA - pctB) < 5 ? F.silver : pctA > pctB ? F.teal : F.red, fontWeight: 700 }}>
              {pctA == null || pctB == null ? "—" : pctA === pctB ? "=" : `${Math.abs(pctA - pctB)} pts`}
            </div>
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <div style={panel}>
            <PanelHeader>RADAR COMPARATIVO <span style={{ color: F.dim, fontSize: 9 }}>(percentil vs. plantel)</span></PanelHeader>
            <ResponsiveContainer width="100%" height={260}>
              <RadarChart data={radarD}>
                <PolarGrid stroke={F.ghost} />
                <PolarAngleAxis dataKey="metric" tick={{ fill: F.silver, fontSize: 10, fontFamily: F.fontMono }} />
                <PolarRadiusAxis domain={[0, 100]} tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} />
                <Radar name={nA} dataKey={nA} stroke={F.teal} fill={F.teal} fillOpacity={0.15} strokeWidth={2} />
                <Radar name={nB} dataKey={nB} stroke={F.red}  fill={F.red}  fillOpacity={0.15} strokeWidth={2} />
                <Legend wrapperStyle={{ fontSize: 10, fontFamily: F.fontMono }} />
                <Tooltip content={<CustomTooltip />} />
              </RadarChart>
            </ResponsiveContainer>
          </div>

          <div style={panel}>
            <PanelHeader>MÉTRICAS — COMPARACIÓN DIRECTA</PanelHeader>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {radarM.map(({ key, label, vA, vB, unit, lowerIsBetter }) => {
                const winner = vA == null || vB == null ? null : (lowerIsBetter ? vA < vB : vA > vB) ? "A" : (vA === vB ? null : "B");
                return (
                  <div key={key} style={{ background: F.asphalt, borderRadius: 3, padding: "8px 12px", border: `1px solid ${F.ghost}` }}>
                    <div style={{ fontFamily: F.fontMono, fontSize: 10, color: F.silver, letterSpacing: 1, marginBottom: 5 }}>{label}</div>
                    <div style={{ display: "flex", gap: 6 }}>
                      {[{ name: nA, val: vA, col: F.teal, isWinner: winner === "A" }, { name: nB, val: vB, col: F.red, isWinner: winner === "B" }].map(({ name, val, col, isWinner }) => (
                        <div key={name} style={{ flex: 1, textAlign: "center", background: F.panel, borderRadius: 2, padding: "5px 4px", borderTop: `2px solid ${isWinner ? F.green : col}` }}>
                          <div style={{ fontSize: 9, color: col, marginBottom: 2, fontFamily: F.fontMono, letterSpacing: 1 }}>{name}{isWinner ? " ▲" : ""}</div>
                          <div style={{ fontSize: 14, fontWeight: 700, fontFamily: F.fontMono, color: val != null ? (isWinner ? F.green : F.white) : F.ghost }}>
                            {val != null ? `${val > 10 ? val.toFixed(1) : val.toFixed(2)}${unit}` : "—"}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>

        {(asymA || asymB) && (
          <div style={{ ...panel, marginTop: 12, borderTop: `2px solid ${F.purple}` }}>
            <PanelHeader accent={F.purple}>MOVILIDAD — ÚLTIMA ASIMETRÍA</PanelHeader>
            <div style={{ display: "flex", gap: 10 }}>
              {[{ name: nA, r: asymA, col: F.teal }, { name: nB, r: asymB, col: F.red }].map(({ name, r, col }) => (
                <div key={name} style={{ flex: 1, textAlign: "center", background: F.asphalt, borderRadius: 3, padding: "10px 8px", borderTop: `1px solid ${col}` }}>
                  <div style={{ fontSize: 9, color: col, fontFamily: F.fontMono, letterSpacing: 1, marginBottom: 4 }}>{name}</div>
                  {r ? (
                    <>
                      <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, marginBottom: 2 }}>{r.test}</div>
                      <div style={{ fontSize: 18, fontWeight: 700, fontFamily: F.fontMono, color: F.white }}>{r.asim?.toFixed(1)}%</div>
                    </>
                  ) : <div style={{ fontFamily: F.fontMono, fontSize: 12, color: F.ghost }}>Sin datos</div>}
                </div>
              ))}
            </div>
          </div>
        )}

        {merged.some(d => d[nA] || d[nB]) && (
          <div style={{ ...panel, marginTop: 12 }}>
            <PanelHeader>EVOLUCIÓN CMJ COMPARADA</PanelHeader>
            <ResponsiveContainer width="100%" height={200}>
              <LineChart data={merged}>
                <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} />
                <YAxis tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} unit="cm" />
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize: 10, fontFamily: F.fontMono }} />
                <Line dataKey={nA} stroke={F.teal} strokeWidth={2} dot={{ r: 4 }} connectNulls />
                <Line dataKey={nB} stroke={F.red}  strokeWidth={2} dot={{ r: 4 }} connectNulls />
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>
    );
  };

  // ── Informe descargable ──────────────────────────────────────────────────
  const ReportView = () => {
    const activePlayers = players.filter(p => (pJR(p.id).length + pVR(p.id).length + pRR(p.id).length + pLR(p.id).length) > 0);
    const pid = selPlayer && activePlayers.find(p => p.id === selPlayer) ? selPlayer : activePlayers[0]?.id;
    const p = players.find(x => x.id === pid);
    if (!p) return null;

    const cmjRec = bestJumpByType(pid, "cmj"); const djRec = bestJumpByType(pid, "dj"); const rsiRec = bestRsi(pid);
    const jevo = jumpEvo(pid); const vevo = velEvo(pid); const revo = romEvo(pid);
    const asym = lastAsym(pid);

    // Métricas clasificadas contra tus referencias reales de fútbol (REFERENCE_BANDS).
    const metrics = [];
    if (cmjRec?.altura != null) metrics.push({ label: "CMJ", value: cmjRec.altura, unit: "cm", ref: REFERENCE_BANDS.cmj, category: "salto" });
    if (djRec?.altura != null) metrics.push({ label: "DJ", value: djRec.altura, unit: "cm", ref: REFERENCE_BANDS.dj, category: "salto" });
    if (rsiRec?.rsi != null) metrics.push({ label: "RSI (DJ)", value: rsiRec.rsi, unit: "", ref: REFERENCE_BANDS.rsi, category: "rsi" });
    allDist.forEach(d => {
      const bv = bestVel(pid, d); if (!bv) return;
      metrics.push({ label: `Sprint ${d}`, value: bv.time, unit: "s", ref: REFERENCE_BANDS[d] || null, category: "velocidad" });
    });
    Object.keys(STRENGTH_REFS).forEach(k => {
      const rs = relStrength(pid, k); if (rs == null) return;
      metrics.push({ label: STRENGTH_REFS[k].label, value: rs, unit: "x PC", ref: REFERENCE_BANDS[k], category: k === "squat" ? "squat" : "bench" });
    });
    // Movilidad: ángulo absoluto (el lado más limitado) contra tu referencia de ROM.
    const romTestsForPlayer = Array.from(new Set(pRR(pid).map(r => r.test)));
    romTestsForPlayer.forEach(testName => {
      const refKey = matchRomRef(testName); if (!refKey) return;
      const recs = pRR(pid).filter(r => r.test === testName); const latest = recs[recs.length - 1]; if (!latest) return;
      const sides = [latest.angI, latest.angD].filter(v => v != null);
      if (!sides.length) return;
      const worse = Math.min(...sides);
      metrics.push({ label: `${testName} (lado más limitado)`, value: worse, unit: "°", ref: REFERENCE_BANDS[refKey], category: "movilidad" });
    });

    const classified = metrics.map(m => ({ ...m, score: m.ref ? bandScore(m.ref, m.value) : null }));
    const fortalezas = classified.filter(m => m.score >= 70);
    const promedio    = classified.filter(m => m.score >= 50 && m.score < 70);
    const debilidades = classified.filter(m => m.score != null && m.score < 50);
    const weakCategories = Array.from(new Set(debilidades.map(m => m.category)));

    return (
      <div className="report-page" style={{ background: F.carbon, maxWidth: 860, margin: "0 auto" }}>
        <style>{`
          @media print {
            .no-print { display: none !important; }
            .report-page { background: #fff !important; color: #111 !important; }
            .report-page * { color: #111 !important; background: transparent !important; border-color: #ccc !important; box-shadow: none !important; }
            .report-page .report-panel { border: 1px solid #ccc !important; page-break-inside: avoid; }
          }
        `}</style>

        <div className="no-print" style={{ display: "flex", gap: 10, marginBottom: 16 }}>
          <BackBtn onClick={goBack} />
          <button onClick={() => window.print()} style={{ padding: "6px 18px", background: F.yellow, color: F.carbon, border: "none", borderRadius: 3, fontFamily: F.fontMono, fontSize: 10, fontWeight: 700, letterSpacing: 2, cursor: "pointer" }}>
            🖨 IMPRIMIR / GUARDAR PDF
          </button>
        </div>

        <div className="report-panel" style={{ ...panel, borderTop: `3px solid ${F.red}` }}>
          <div>
            <div style={{ fontFamily: F.fontMono, fontSize: 20, color: F.white, letterSpacing: 2, fontWeight: 700 }}>{p.name}</div>
            <div style={{ fontFamily: F.fontMono, fontSize: 11, color: F.dim, marginTop: 4 }}>{p.team || "—"} · Informe generado {new Date().toLocaleDateString("es-AR")}</div>
          </div>
        </div>

        <div className="report-panel" style={panel}>
          <PanelHeader>MÉTRICAS ACTUALES <span style={{ color: F.dim, fontSize: 9 }}>(vs. referencia de fútbol)</span></PanelHeader>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead><tr style={{ borderBottom: `1px solid ${F.panelBorder}` }}>
              {["MÉTRICA", "VALOR", "RANGO ÉLITE", "NIVEL"].map(h => (
                <th key={h} style={{ padding: "6px 10px", color: F.dim, textAlign: "left", fontFamily: F.fontMono, fontSize: 9, letterSpacing: 2 }}>{h}</th>
              ))}
            </tr></thead>
            <tbody>{classified.map((m, i) => (
              <tr key={i} style={{ borderBottom: `1px solid ${F.ghost}` }}>
                <td style={{ padding: "7px 10px", color: F.silver, fontFamily: F.fontMono }}>{m.label}</td>
                <td style={{ padding: "7px 10px", color: F.white, fontFamily: F.fontMono, fontWeight: 700 }}>{m.value?.toFixed(2)}{m.unit}</td>
                <td style={{ padding: "7px 10px", color: F.dim, fontFamily: F.fontMono, fontSize: 11 }}>{m.ref ? `${m.ref.min}–${m.ref.max}${m.ref.unit}` : "—"}</td>
                <td style={{ padding: "7px 10px" }}><ReportBadge score={m.score} /></td>
              </tr>
            ))}</tbody>
          </table>
          {classified.length === 0 && <div style={{ fontFamily: F.fontMono, fontSize: 11, color: F.dim, marginTop: 8 }}>Sin datos suficientes todavía.</div>}
          {asym && (
            <div style={{ marginTop: 10, fontFamily: F.fontMono, fontSize: 10, color: F.dim }}>
              Asimetría {asym.test}: <span style={{ color: F.white }}>{asym.asim?.toFixed(1)}%</span> (informativo — sin umbral de referencia cargado)
            </div>
          )}
        </div>

        {jevo.length > 0 && (
          <div className="report-panel" style={panel}>
            <PanelHeader>EVOLUCIÓN — ALTURA DE SALTO (cm)</PanelHeader>
            <ResponsiveContainer width="100%" height={180}>
              <LineChart data={jevo}>
                <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} />
                <YAxis tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} unit="cm" />
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize: 10, fontFamily: F.fontMono }} />
                <Line dataKey="CMJ" stroke={F.teal} strokeWidth={2} dot={{ r: 3 }} connectNulls />
                <Line dataKey="DJ"  stroke={F.red}  strokeWidth={2} dot={{ r: 3 }} connectNulls />
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}

        {vevo.length > 0 && allDist.length > 0 && (
          <div className="report-panel" style={panel}>
            <PanelHeader accent={F.yellow}>EVOLUCIÓN — VELOCIDAD (s)</PanelHeader>
            <ResponsiveContainer width="100%" height={180}>
              <LineChart data={vevo}>
                <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} />
                <YAxis tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} unit="s" />
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize: 10, fontFamily: F.fontMono }} />
                {allDist.map((d, i) => <Line key={d} dataKey={d} stroke={TEAM_COLORS[i]} strokeWidth={2} dot={{ r: 3 }} connectNulls />)}
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}

        {revo.length > 0 && (
          <div className="report-panel" style={panel}>
            <PanelHeader accent={F.purple}>EVOLUCIÓN — MOVILIDAD (ASIMETRÍA %)</PanelHeader>
            <ResponsiveContainer width="100%" height={160}>
              <LineChart data={revo}>
                <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} />
                <YAxis tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} unit="%" />
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize: 10, fontFamily: F.fontMono }} />
                {romTests.map((t, i) => <Line key={t} dataKey={t} stroke={[F.purple, F.orange, F.teal, F.yellow][i % 4]} strokeWidth={2} dot={{ r: 3 }} connectNulls />)}
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}

        <div className="report-panel" style={{ ...panel, borderTop: `2px solid ${F.green}` }}>
          <PanelHeader accent={F.green}>CONCLUSIÓN</PanelHeader>
          {fortalezas.length > 0 && (
            <div style={{ marginBottom: 10 }}>
              <span style={{ fontFamily: F.fontMono, fontSize: 10, color: F.green, letterSpacing: 1 }}>FORTALEZAS: </span>
              <span style={{ fontFamily: F.fontF1, fontSize: 13, color: F.silver }}>
                {fortalezas.map(m => `${m.label} (${m.value?.toFixed(1)}${m.unit})`).join(", ")}. Dentro o por encima del rango élite de referencia.
              </span>
            </div>
          )}
          {promedio.length > 0 && (
            <div style={{ marginBottom: 10 }}>
              <span style={{ fontFamily: F.fontMono, fontSize: 10, color: F.teal, letterSpacing: 1 }}>A CONSOLIDAR: </span>
              <span style={{ fontFamily: F.fontF1, fontSize: 13, color: F.silver }}>{promedio.map(m => m.label).join(", ")}.</span>
            </div>
          )}
          {debilidades.length > 0 ? (
            <div>
              <span style={{ fontFamily: F.fontMono, fontSize: 10, color: F.red, letterSpacing: 1 }}>A MEJORAR: </span>
              <span style={{ fontFamily: F.fontF1, fontSize: 13, color: F.silver }}>
                {debilidades.map(m => `${m.label} (${m.value?.toFixed(1)}${m.unit} vs. ${m.ref.min}–${m.ref.max}${m.ref.unit} élite)`).join(", ")}. Se recomienda trabajo complementario dirigido en las próximas semanas.
              </span>
            </div>
          ) : classified.length > 0 ? (
            <div style={{ fontFamily: F.fontF1, fontSize: 13, color: F.silver }}>No se detectan puntos débiles marcados frente a la referencia — perfil parejo o por encima en todas las métricas evaluadas.</div>
          ) : (
            <div style={{ fontFamily: F.fontF1, fontSize: 13, color: F.dim }}>Sin datos suficientes para evaluar todavía.</div>
          )}
        </div>

        {weakCategories.length > 0 && (
          <div className="report-panel" style={{ ...panel, borderTop: `2px solid ${F.orange}` }}>
            <PanelHeader accent={F.orange}>PLAN DE EJERCICIOS COMPLEMENTARIO</PanelHeader>
            {weakCategories.map(cat => {
              const bank = EXERCISE_BANK[cat]; if (!bank) return null;
              return (
                <div key={cat} style={{ marginBottom: 16 }}>
                  <div style={{ fontFamily: F.fontMono, fontSize: 11, color: F.orange, letterSpacing: 1, marginBottom: 2 }}>{bank.label.toUpperCase()}</div>
                  <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, marginBottom: 8 }}>Fuente: {bank.fuente}</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                    {bank.items.map((it, i) => (
                      <div key={i} style={{ display: "flex", justifyContent: "space-between", background: F.asphalt, borderRadius: 3, padding: "6px 10px" }}>
                        <span style={{ fontFamily: F.fontF1, fontSize: 12, color: F.white }}>{it.name}{it.note ? <span style={{ color: F.dim, fontSize: 11 }}> — {it.note}</span> : null}</span>
                        <span style={{ fontFamily: F.fontMono, fontSize: 11, color: F.teal, whiteSpace: "nowrap", marginLeft: 10 }}>{it.sets}</span>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}
            <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, marginTop: 4 }}>
              Plan sugerido a partir de tus propios protocolos y tu referencia de campo. Ajustar cargas, series y progresión según criterio profesional.
            </div>
          </div>
        )}
      </div>
    );
  };

  // ── Shell ─────────────────────────────────────────────────────────────────
  return (
    <div style={{ minHeight: "100vh", background: F.carbon, fontFamily: F.fontF1, color: F.white }}>
      <link href="https://fonts.googleapis.com/css2?family=Titillium+Web:wght@300;400;600;700&family=Share+Tech+Mono&display=swap" rel="stylesheet" />

      <div className="no-print" style={{ background: F.asphalt, padding: "0 24px", display: "flex", alignItems: "center", height: 52, borderBottom: `1px solid ${F.panelBorder}` }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 4, height: 28, background: F.red }} />
          <div>
            <div style={{ fontFamily: F.fontMono, fontSize: 16, color: F.white, letterSpacing: 3, lineHeight: 1 }}>
              ATHLETE<span style={{ color: F.red }}>IQ</span>
            </div>
            <div style={{ fontFamily: F.fontMono, fontSize: 8, color: F.dim, letterSpacing: 3 }}>PERFORMANCE SYSTEM</div>
          </div>
        </div>

        <nav style={{ marginLeft: 32, display: "flex", gap: 2 }}>
          {[
            { key:"import",    label:"IMPORTAR" },
            { key:"dashboard", label:"DASHBOARD", d:!hasData },
            { key:"evolution", label:"EVOLUCIÓN",  d:!hasData },
            { key:"compare",   label:"COMPARAR",   d:!hasData },
            { key:"delete",    label:"GESTIÓN",    d:!hasData },
          ].map(({ key, label, d }) => (
            <button key={key} onClick={() => !d && goTo(key)} style={{
              padding: "0 14px", height: 52, border: "none",
              borderBottom: tab === key ? `2px solid ${key === "delete" ? F.red : F.red}` : "2px solid transparent",
              cursor: d ? "default" : "pointer", fontFamily: F.fontMono, fontSize: 10, letterSpacing: 2,
              background: "transparent", color: tab === key ? F.white : d ? F.ghost : F.dim,
              transition: "all .15s",
            }}>{label}</button>
          ))}
        </nav>

        {hasData && <div style={{ marginLeft: "auto", fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 2 }}>
          {players.length} ATL · {jumpRecs.length + velRecs.length + romRecs.length + liftRecs.length + customRecs.length} REG
        </div>}
      </div>

      <div className="no-print"><TelemetryBar /></div>

      <div style={{ padding: "20px 24px", maxWidth: 1160, margin: "0 auto" }}>
        {tab === "import"    && <ImportView />}
        {tab === "dashboard" && hasData && <Dashboard />}
        {tab === "evolution" && hasData && <Evolution />}
        {tab === "compare"   && hasData && <Compare />}
        {tab === "delete"    && hasData && <DeletePanel />}
        {tab === "informe"   && hasData && <ReportView />}
      </div>
    </div>
  );
}
