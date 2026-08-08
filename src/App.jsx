import { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, BarChart, Bar, RadarChart, Radar, PolarGrid,
  PolarAngleAxis, PolarRadiusAxis, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, ScatterChart, Scatter,
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

// ── Referencias de campo (fútbol) — usadas para clasificar en el Informe ──────
// El Dashboard/Comparar usan percentiles relativos al plantel (más abajo), sin tocar esto.
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
// El piso/techo ÉLITE de cada rango es el dato real que pasaste. BUENO/PROMEDIO se
// extrapolan proporcionalmente (±12% / ±25%) — ajustable acá si conseguís cortes más precisos.
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
  if (score == null) return { label: "—", color: F.ghost };
  if (score >= 90) return { label: "ELITE", color: F.yellow };
  if (score >= 70) return { label: "BUENO", color: F.green };
  if (score >= 50) return { label: "PROMEDIO", color: F.teal };
  return { label: "BAJO", color: F.red };
}
const ReportBadge = ({ score }) => {
  const lv = bandLevel(score);
  if (score == null) return <span style={{ fontFamily: F.fontMono, fontSize: 9, color: F.ghost }}>—</span>;
  return (
    <div style={{ display: "inline-flex", alignItems: "center", gap: 5, padding: "2px 8px",
      border: `1px solid ${lv.color}`, borderRadius: 2, background: lv.color + "18" }}>
      <div style={{ width: 5, height: 5, borderRadius: "50%", background: lv.color, boxShadow: `0 0 5px ${lv.color}` }} />
      <span style={{ fontFamily: F.fontMono, fontSize: 9, color: lv.color, letterSpacing: 2 }}>{lv.label}</span>
    </div>
  );
};

// ── Banco de ejercicios correctivos (curado de los programas propios de Leandro) ──
const EXERCISE_BANK = {
  salto: {
    label: "Salto / Potencia reactiva",
    fuente: "Programa de Readaptación LCA · Planillas Fuerza-Potencia",
    items: [
      { name: "Pogos reactivos", sets: "3×10" },
      { name: "Drop jump progresivo (30-45-60cm)", sets: "3×4" },
      { name: "Sentadilla con salto", sets: "3×6-8" },
      { name: "Salto al cajón a 1 pierna", sets: "3×4+4" },
      { name: "Búlgaras con salto (carga progresiva)", sets: "3×4+4" },
    ],
  },
  rsi: {
    label: "RSI / Ciclo estiramiento-acortamiento",
    fuente: "Índice de Fuerza Reactiva (Comunidad Lift) · Readaptación LCA",
    items: [
      { name: "Pogos reactivos (contacto <200ms)", sets: "3×10" },
      { name: "Drop jump + salto horizontal", sets: "3×4" },
      { name: "Hurdle hops (vallas bajas)", sets: "3×6" },
    ],
  },
  velocidad: {
    label: "Velocidad / Aceleración",
    fuente: "Programa de Readaptación LCA",
    items: [
      { name: "Aceleraciones 5-10m", sets: "4×4-5" },
      { name: "Técnica de carrera + aceleración", sets: "3×20m" },
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
    ],
  },
  bench: {
    label: "Fuerza tren superior (Bench)",
    fuente: "Fuerza Estructural I-V · Día Fuerza (planillas propias)",
    items: [
      { name: "Press plano con barra", sets: "3×6-8" },
      { name: "Press inclinado con barra", sets: "3×6-8" },
      { name: "Flexiones lastradas", sets: "3×8-12" },
    ],
  },
  movilidad: {
    label: "Movilidad / Control neuromuscular",
    fuente: "Fuerza Estructural I-V · Síndrome femoropatelar (Comunidad Lift)",
    items: [
      { name: "Movilidad de cadera 90-90", sets: "2-3×6-10+6-10" },
      { name: "Movilidad de tobillo", sets: "2-3×6-10+6-10" },
      { name: "Activación de glúteo medio (monster walk)", sets: "3×6-10+6-10" },
    ],
  },
};

function percentileColor(pct) {
  if (pct == null) return F.ghost;
  if (pct >= 75) return F.green;
  if (pct >= 50) return F.teal;
  if (pct >= 25) return F.yellow;
  return F.red;
}

// ── ID generator ──────────────────────────────────────────────────────────────
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
// Detecta automáticamente el tipo de archivo de texto (CSV/TXT) por su contenido.
function classifyCSV(text) {
  const firstLine = (text.split("\n")[0] || "");
  if (firstLine.includes(";")) {
    const headers = firstLine.split(";").map(h => h.trim());
    if (headers.some(h => h.includes("Altura de salto"))) return "jump";
    if (headers.some(h => h.includes("Ángulo"))) return "rom";
    if (headers.some(h => h.includes("1-RM") || h.includes("Nombre del ejercicio"))) return "lift";
    return "generic";
  }
  const testRows = parsePhotoFinishCSV(text);
  const validTimes = testRows.filter(r => parseNum(r.time) != null).length;
  if (testRows.length > 0 && validTimes >= Math.ceil(testRows.length * 0.5)) return "vel";
  return "generic";
}
const IMG_EXT = ["png", "jpg", "jpeg", "webp", "gif"];
function readFileAsText(file) { return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsText(file, "UTF-8"); }); }
function readFileAsDataURL(file) { return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(file); }); }
function readFileAsArrayBuffer(file) { return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsArrayBuffer(file); }); }

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
    <div style={{ background: F.panel, border: `1px solid ${F.red}`, borderRadius: 4, padding: "6px 12px" }}>
      <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.red, marginBottom: 3, letterSpacing: 2 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color, fontSize: 10, fontFamily: F.fontMono, margin: "1px 0" }}>
          {p.name}: <span style={{ color: F.white, fontWeight: 700 }}>{typeof p.value === "number" ? p.value.toFixed(2) : p.value}</span>
        </div>
      ))}
    </div>
  );
};

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
    <div style={{ position: "relative", width: size, height: size, flexShrink: 0 }}>
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

// Avatar del deportista — foto si existe, si no iniciales.
const Avatar = ({ player, size = 40, ring }) => {
  const initials = (player?.name || "?").split(" ").filter(Boolean).slice(0, 2).map(w => w[0]).join("").toUpperCase();
  const border = ring ? `2px solid ${ring}` : `1px solid ${F.panelBorder}`;
  if (player?.photo) {
    return <img src={player.photo} alt={player.name} style={{ width: size, height: size, borderRadius: "50%", objectFit: "cover", border, flexShrink: 0 }} />;
  }
  return (
    <div style={{ width: size, height: size, borderRadius: "50%", background: F.ghost, display: "flex",
      alignItems: "center", justifyContent: "center", fontFamily: F.fontMono, fontWeight: 700,
      fontSize: size * 0.34, color: F.silver, border, flexShrink: 0 }}>
      {initials}
    </div>
  );
};

const PanelHeader = ({ children, accent }) => (
  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8,
    paddingBottom: 6, borderBottom: `1px solid ${F.panelBorder}` }}>
    <div style={{ width: 3, height: 12, background: accent || F.red, borderRadius: 1 }} />
    <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.silver, letterSpacing: 2, textTransform: "uppercase" }}>{children}</div>
  </div>
);

const BackBtn = ({ onClick }) => (
  <button onClick={onClick} style={{ display: "inline-flex", alignItems: "center", gap: 6,
    padding: "6px 14px", background: "transparent", border: `1px solid ${F.panelBorder}`,
    borderRadius: 3, color: F.silver, fontFamily: F.fontMono, fontSize: 10, letterSpacing: 2,
    cursor: "pointer", transition: "border-color .15s" }}
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

  const [selPlayer, setSelPlayer] = useState(null);
  const [compareA,  setCompareA]  = useState(null);
  const [compareB,  setCompareB]  = useState(null);

  const [filterTeam,  setFilterTeam]  = useState("Todos");
  const [msgs,        setMsgs]        = useState({});
  const [genericData, setGenericData] = useState(null);
  const [genericCols, setGenericCols] = useState([]);
  const [genericFileName, setGenericFileName] = useState("");
  const [genericQueue, setGenericQueue] = useState([]); // archivos pendientes de mapeo manual
  const [colMap,      setColMap]      = useState({});
  const [showMapper,  setShowMapper]  = useState(false);
  const [liftExercise,setLiftExercise]= useState(null);
  const [importing,   setImporting]   = useState(false);

  const unifiedRef = useRef();

  function setMsg(key, val) { setMsgs(m => ({ ...m, [key]: val })); }
  function goTo(t) { setPrevTab(tab); setTab(t); }
  function goBack() { if (prevTab) { setTab(prevTab); setPrevTab(null); } else setTab("dashboard"); }

  // ── Player merge ──────────────────────────────────────────────────────────
  function mergePlayer(name, team, cur) {
    const key = name.toLowerCase().trim();
    const ex = cur.find(p => p.name.toLowerCase().trim() === key);
    if (ex) return { players: team && !ex.team ? cur.map(p => p.id === ex.id ? { ...p, team } : p) : cur, id: ex.id };
    const np = { id: uid(), name: name.trim(), team: team || "", photo: null };
    return { players: [...cur, np], id: np.id };
  }
  function attachPhoto(cur, name, dataUrl) {
    const { players: up, id: pid } = mergePlayer(name, "", cur);
    return up.map(p => p.id === pid ? { ...p, photo: dataUrl } : p);
  }

  // Constructores puros de registros a partir de filas ya parseadas — usados tanto
  // por el importador unificado como por el mapeo manual.
  function buildJumpRecords(rows, cur) {
    const recs = [];
    rows.forEach(row => {
      const name = (row["Nombre"] || "").trim(); const team = (row["Equipo"] || "").trim();
      if (!name) return;
      const { players: up, id: pid } = mergePlayer(name, team, cur); cur = up;
      const tipo = (row["Tipo de salto"] || "").trim();
      const fecha = parseMyJumpDate(row["Fecha"] || "");
      recs.push({
        id: uid(), playerId: pid, date: fecha.slice(0, 7), fullDate: fecha,
        jumpType: jumpTypeKey(tipo), jumpTypeRaw: tipo,
        altura: parseNum(row["Altura de salto (cm)"]), rsi: parseNum(row["RSI mod (m/s)"]),
        fuerza: parseNum(row["Fuerza (N)"]), potencia: parseNum(row["Potencia (W)"]),
        statusColor: row["Color estado de forma"] || "", team,
      });
    });
    return { cur, recs };
  }
  function buildVelRecords(rows, cur) {
    const recs = [];
    rows.forEach(row => {
      const name = row.athlete.trim(); if (!name) return;
      const { players: up, id: pid } = mergePlayer(name, "", cur); cur = up;
      const fecha = parsePhotoDate(row.date); const dist = detectDistance(row.info);
      const t = parseNum(row.time); if (!t) return;
      recs.push({ id: uid(), playerId: pid, date: fecha.slice(0, 7), fullDate: fecha, distance: dist, testName: row.info, time: t });
    });
    return { cur, recs };
  }
  function buildRomRecords(rows, cur) {
    const recs = [];
    rows.forEach(row => {
      const name = (row["Nombre"] || "").trim(); const team = (row["Equipo"] || "").trim();
      if (!name) return;
      const { players: up, id: pid } = mergePlayer(name, team, cur); cur = up;
      const fecha = parseMyJumpDate(row["Fecha"] || "");
      recs.push({
        id: uid(), playerId: pid, date: fecha.slice(0, 7), fullDate: fecha,
        test: (row["Test"] || "").trim(), angI: parseNum(row["Ángulo izq(º)"]), angD: parseNum(row["Ángulo drch(º)"]),
        asim: parseNum(row["Asimetría (%)"]), statusColor: row["Color estado de forma"] || "", team,
      });
    });
    return { cur, recs };
  }
  function buildLiftRecords(rows, cur) {
    const recs = [];
    rows.forEach(row => {
      const name = (row["Nombre"] || "").trim(); const team = (row["Equipo"] || "").trim();
      if (!name) return;
      const { players: up, id: pid } = mergePlayer(name, team, cur); cur = up;
      const fecha = parseMyJumpDate(row["Fecha"] || "");
      const exercise = (row["Nombre del ejercicio"] || "").trim();
      const loads = [1, 2, 3, 4].map(i => parseNum(row[`Carga ${i} (kg)`]));
      const vels  = [1, 2, 3, 4].map(i => parseNum(row[`Velocidad media ${i} (m/s)`]));
      const points = loads.map((l, i) => ({ load: l, vel: vels[i] })).filter(x => x.load != null && x.vel != null);
      recs.push({
        id: uid(), playerId: pid, date: fecha.slice(0, 7), fullDate: fecha, exercise, exKey: exerciseKey(exercise),
        bodyweight: parseNum(row["Peso corporal(kg)"]), rm: parseNum(row["1-RM (kg)"]), points, team,
      });
    });
    return { cur, recs };
  }

  // ── Delete functions ──────────────────────────────────────────────────────
  function deleteAllData() {
    setPlayers([]); setJumpRecs([]); setVelRecs([]); setRomRecs([]); setLiftRecs([]); setCustomRecs([]);
    setSelPlayer(null); setCompareA(null); setCompareB(null);
    setMsgs({}); setTab("import");
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

  // ── Importador unificado: cualquier mezcla de archivos, todos a la vez ──────
  async function handleUnifiedImport(e) {
    const files = Array.from(e.target.files || []);
    e.target.value = "";
    if (!files.length) return;
    setImporting(true);
    let cur = [...players];
    let newJump = [], newVel = [], newRom = [], newLift = [];
    const genericFiles = [];
    let photosAdded = 0, filesFailed = 0;

    for (const file of files) {
      const ext = file.name.split(".").pop().toLowerCase();
      try {
        if (IMG_EXT.includes(ext)) {
          const dataUrl = await readFileAsDataURL(file);
          const rawName = file.name.replace(/\.[^.]+$/, "").replace(/[_\-]+/g, " ").trim();
          cur = attachPhoto(cur, rawName, dataUrl);
          photosAdded++;
          continue;
        }
        if (ext === "xlsx" || ext === "xls") {
          const buf = await readFileAsArrayBuffer(file);
          const wb = XLSX.read(buf, { type: "array" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
          if (!rows.length) { filesFailed++; continue; }
          genericFiles.push({ rows, cols: Object.keys(rows[0]), fileName: file.name });
          continue;
        }
        const text = await readFileAsText(file);
        const type = classifyCSV(text);
        if (type === "jump") {
          const { cur: c2, recs } = buildJumpRecords(parseMyJumpCSV(text), cur); cur = c2; newJump = newJump.concat(recs);
        } else if (type === "rom") {
          const { cur: c2, recs } = buildRomRecords(parseMyJumpCSV(text), cur); cur = c2; newRom = newRom.concat(recs);
        } else if (type === "lift") {
          const { cur: c2, recs } = buildLiftRecords(parseMyJumpCSV(text), cur); cur = c2; newLift = newLift.concat(recs);
        } else if (type === "vel") {
          const { cur: c2, recs } = buildVelRecords(parsePhotoFinishCSV(text), cur); cur = c2; newVel = newVel.concat(recs);
        } else {
          const lines = text.trim().split("\n");
          if (lines.length < 2) { filesFailed++; continue; }
          const sep = lines[0].includes(";") ? ";" : ",";
          const headers = lines[0].split(sep).map(h => h.trim());
          const rows = lines.slice(1).map(line => {
            const vals = line.split(sep); const obj = {};
            headers.forEach((h, i) => { obj[h] = (vals[i] || "").trim(); });
            return obj;
          });
          genericFiles.push({ rows, cols: headers, fileName: file.name });
        }
      } catch { filesFailed++; }
    }

    setPlayers(cur);
    if (newJump.length) setJumpRecs(r => [...r, ...newJump]);
    if (newVel.length) setVelRecs(r => [...r, ...newVel]);
    if (newRom.length) setRomRecs(r => [...r, ...newRom]);
    if (newLift.length) setLiftRecs(r => [...r, ...newLift]);
    if (!selPlayer && cur.length) { setSelPlayer(cur[0].id); setCompareA(cur[0].id); setCompareB(cur[1]?.id || cur[0].id); }

    const parts = [];
    if (newJump.length) parts.push(`${newJump.length} saltos`);
    if (newVel.length) parts.push(`${newVel.length} vel.`);
    if (newRom.length) parts.push(`${newRom.length} movil.`);
    if (newLift.length) parts.push(`${newLift.length} vbt/rm`);
    if (photosAdded) parts.push(`${photosAdded} fotos`);
    let m = parts.length ? `✓ ${parts.join(" · ")}` : "";
    if (genericFiles.length) m += `${m ? " · " : ""}${genericFiles.length} archivo(s) necesitan mapeo manual`;
    if (filesFailed) m += `${m ? " · " : ""}⚠ ${filesFailed} no se pudieron leer`;
    setMsg("unified", m || "❌ No se pudo leer ningún archivo.");
    setImporting(false);

    if (genericFiles.length) {
      setGenericQueue(genericFiles.slice(1));
      loadMapperFile(genericFiles[0]);
    } else if (newJump.length || newVel.length || newRom.length || newLift.length) {
      setTab("dashboard");
    }
  }

  function loadMapperFile(gf) {
    setGenericData(gf.rows); setGenericCols(gf.cols); setGenericFileName(gf.fileName);
    setColMap({ name: "", date: "", metric1: "", metric1label: "", metric2: "", metric2label: "", metric3: "", metric3label: "", team: "" });
    setShowMapper(true);
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
    setMsg("gen", `✓ ${newRecs.length} registros importados de ${genericFileName}`);

    if (genericQueue.length) {
      const next = genericQueue[0];
      setGenericQueue(q => q.slice(1));
      loadMapperFile(next);
    } else {
      setShowMapper(false);
      setTab("dashboard");
    }
  }
  function skipMapperFile() {
    if (genericQueue.length) {
      const next = genericQueue[0];
      setGenericQueue(q => q.slice(1));
      loadMapperFile(next);
    } else {
      setShowMapper(false);
    }
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
  const bestVel  = (pid, dist) => { const rs = pVR(pid).filter(r => r.distance === dist && r.time != null); return rs.reduce((b, r) => r.time < (b?.time ?? 999) ? r : b, null); };
  const allDist  = useMemo(() => Array.from(new Set(velRecs.map(r => r.distance).filter(Boolean))).sort(), [velRecs]);
  const romTests = useMemo(() => Array.from(new Set(romRecs.map(r => r.test).filter(Boolean))).sort(), [romRecs]);
  const lastAsym = (pid) => { const rs = pRR(pid).filter(r => r.asim != null); return rs[rs.length - 1] || null; };
  const bestRM = (pid, exKey) => { const rs = pLR(pid).filter(r => r.exKey === exKey && r.rm != null); return rs.reduce((b, r) => r.rm > (b?.rm ?? 0) ? r : b, null); };
  const relStrength = (pid, exKey) => { const r = bestRM(pid, exKey); if (!r || !r.bodyweight) return null; return r.rm / r.bodyweight; };
  function matchRomRef(testName) {
    if (!testName) return null; const t = testName.toLowerCase();
    if (t.includes("dorsiflex")) return "dorsiflexion";
    if (t.includes("cadera") || t.includes("rotaci")) return "hipIR";
    return null;
  }

  // ── Percentiles (comparación directa dentro del plantel cargado) ───────────
  function percentileRank(list, targetPid, lowerIsBetter) {
    if (!list || list.length < 2) return null;
    const target = list.find(x => x.pid === targetPid);
    if (!target || target.value == null) return null;
    const better = list.filter(x => x.pid !== targetPid && (lowerIsBetter ? x.value > target.value : x.value < target.value)).length;
    return Math.round((better / (list.length - 1)) * 100);
  }
  const jumpGroup = useMemo(() => visible.map(p => ({ pid: p.id, value: bestJump(p.id)?.altura ?? null })).filter(x => x.value != null), [players, jumpRecs, filterTeam]);
  const rsiGroup  = useMemo(() => visible.map(p => ({ pid: p.id, value: bestJump(p.id)?.rsi ?? null })).filter(x => x.value != null), [players, jumpRecs, filterTeam]);
  const distGroups = useMemo(() => Object.fromEntries(allDist.map(d => [d, visible.map(p => ({ pid: p.id, value: bestVel(p.id, d)?.time ?? null })).filter(x => x.value != null)])), [players, velRecs, filterTeam, allDist]);
  const strengthGroups = useMemo(() => Object.fromEntries(Object.keys(REFERENCE_RANGES).filter(k => k === "squat" || k === "bench-press").map(k => [k, visible.map(p => ({ pid: p.id, value: relStrength(p.id, k) })).filter(x => x.value != null)])), [players, liftRecs, filterTeam]);
  const asymGroup = useMemo(() => visible.map(p => ({ pid: p.id, value: lastAsym(p.id)?.asim ?? null })).filter(x => x.value != null), [players, romRecs, filterTeam]);

  const avgPercentile = (pid) => {
    const pcts = [];
    const jp = percentileRank(jumpGroup, pid, false); if (jp != null) pcts.push(jp);
    const rp = percentileRank(rsiGroup, pid, false); if (rp != null) pcts.push(rp);
    allDist.forEach(d => { const vp = percentileRank(distGroups[d], pid, true); if (vp != null) pcts.push(vp); });
    ["squat", "bench-press"].forEach(k => { const sp = percentileRank(strengthGroups[k], pid, false); if (sp != null) pcts.push(sp); });
    return pcts.length ? Math.round(pcts.reduce((a, b) => a + b, 0) / pcts.length) : null;
  };

  const rankingData = useMemo(() =>
    visible.filter(p => (pJR(p.id).length + pVR(p.id).length + pLR(p.id).length) > 0)
      .map(p => ({ id: p.id, name: p.name.split(" ")[0], fullName: p.name, photo: p.photo, pct: avgPercentile(p.id) ?? 0 }))
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
  const romEvo = (pid) => {
    const byDate = {};
    pRR(pid).forEach(r => {
      if (!byDate[r.fullDate]) byDate[r.fullDate] = { fecha: r.fullDate.slice(8) + "/" + r.fullDate.slice(5, 7) };
      if (r.asim != null) byDate[r.fullDate][r.test || "Test"] = r.asim;
    });
    return Object.values(byDate);
  };
  const liftRmEvo = (pid, exercise) => {
    const byDate = {};
    pLR(pid).filter(r => r.exercise === exercise && r.rm != null).forEach(r => {
      const key = r.fullDate;
      if (!byDate[key] || r.rm > byDate[key].rm) byDate[key] = { fecha: r.fullDate.slice(8) + "/" + r.fullDate.slice(5, 7), rm: r.rm };
    });
    return Object.values(byDate).sort((a, b) => a.fecha.localeCompare(b.fecha));
  };
  const latestLiftProfile = (pid, exercise) => {
    const rs = pLR(pid).filter(r => r.exercise === exercise);
    if (!rs.length) return { points: [], date: null };
    const latest = rs[rs.length - 1];
    return { points: latest.points, date: latest.fullDate };
  };

  // Devolución + plan corto para una métrica evaluada (usado en Informe).
  function feedbackFor(category, score) {
    const bank = EXERCISE_BANK[category]; if (!bank) return null;
    if (score == null || score >= 70) return null;
    return bank;
  }

  // ── Styles ────────────────────────────────────────────────────────────────
  const panel = { background: F.panel, border: `1px solid ${F.panelBorder}`, borderRadius: 4, padding: 12, marginBottom: 10, display: "flex", flexDirection: "column", minHeight: 0 };
  const sel   = { background: F.ghost, border: `1px solid ${F.panelBorder}`, borderRadius: 3, color: F.white, padding: "6px 10px", fontSize: 12, outline: "none", fontFamily: F.fontMono, cursor: "pointer" };
  const lbl   = { fontSize: 9, color: F.dim, letterSpacing: 2, marginBottom: 4, display: "block", textTransform: "uppercase", fontFamily: F.fontMono };
  const inp   = { ...sel, width: "100%", boxSizing: "border-box" };
  const dangerBtn = { padding: "6px 14px", background: "transparent", border: `1px solid ${F.red}`, borderRadius: 3, color: F.red, fontFamily: F.fontMono, fontSize: 10, letterSpacing: 2, cursor: "pointer" };

  // ── Import View (unificado, multi-archivo, auto-detección) ─────────────────
  const ImportView = () => (
    <div style={{ maxWidth: 760, margin: "0 auto" }}>
      <div style={{ marginBottom: 16, paddingBottom: 12, borderBottom: `1px solid ${F.panelBorder}` }}>
        <div style={{ fontFamily: F.fontMono, fontSize: 11, color: F.red, letterSpacing: 3, marginBottom: 6 }}>SISTEMA DE IMPORTACIÓN</div>
        <div style={{ fontFamily: F.fontMono, fontSize: 13, color: F.silver }}>Subí todos tus archivos juntos — la app detecta solo si es MyJump, PhotoFinish, MyROM, MyLift, Excel o una foto</div>
      </div>

      <div style={{ ...panel, cursor: importing ? "default" : "pointer", borderTop: `2px solid ${F.red}`, alignItems: "center", padding: 30 }}
        onClick={() => !importing && unifiedRef.current?.click()}>
        <div style={{ fontSize: 32, color: F.red, marginBottom: 8 }}>{importing ? "⏳" : "↑"}</div>
        <div style={{ fontFamily: F.fontMono, fontSize: 13, color: F.white, letterSpacing: 2, marginBottom: 4 }}>
          {importing ? "PROCESANDO..." : "SUBIR TODOS LOS ARCHIVOS"}
        </div>
        <div style={{ fontFamily: F.fontMono, fontSize: 10, color: F.dim, textAlign: "center" }}>
          MyJump Lab (saltos) · MyROM · MyLift · PhotoFinish · Excel/CSV · Fotos de los deportistas — todo junto, sin límite de archivos
        </div>
        <input ref={unifiedRef} type="file" multiple accept=".csv,.txt,.xlsx,.xls,.png,.jpg,.jpeg,.webp,.gif" style={{ display: "none" }} onChange={handleUnifiedImport} />
        {msgs.unified && (
          <div style={{ marginTop: 14, fontSize: 11, color: msgs.unified.startsWith("✓") || msgs.unified.startsWith("❌") === false ? F.green : F.red, fontFamily: F.fontMono, padding: "6px 12px", background: F.asphalt, borderRadius: 3, textAlign: "center" }}>
            {msgs.unified}
          </div>
        )}
      </div>

      <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, textAlign: "center", marginBottom: 16 }}>
        Tip: para que la foto se asocie sola, nombrá el archivo igual que el jugador (ej. "Juan Perez.jpg")
      </div>

      {showMapper && genericData && (
        <div style={{ ...panel, borderTop: `2px solid ${F.yellow}` }}>
          <PanelHeader accent={F.yellow}>MAPEO — {genericFileName} ({genericData.length} filas){genericQueue.length > 0 ? ` · quedan ${genericQueue.length} archivo(s) más` : ""}</PanelHeader>
          <div style={{ marginBottom: 10, fontFamily: F.fontMono, fontSize: 10, color: F.dim }}>Columnas: {genericCols.join(" · ")}</div>
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
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={applyMapping} style={{ padding: "8px 24px", background: F.yellow, color: F.carbon, border: "none", borderRadius: 3, fontFamily: F.fontMono, fontSize: 11, fontWeight: 700, letterSpacing: 2, cursor: "pointer" }}>
              IMPORTAR
            </button>
            <button onClick={skipMapperFile} style={{ padding: "8px 24px", background: "transparent", border: `1px solid ${F.panelBorder}`, color: F.dim, borderRadius: 3, fontFamily: F.fontMono, fontSize: 11, letterSpacing: 2, cursor: "pointer" }}>
              OMITIR ESTE ARCHIVO
            </button>
          </div>
        </div>
      )}

      {hasData && (
        <div style={{ ...panel, borderTop: `2px solid ${F.green}`, flexDirection: "row", flexWrap: "wrap", alignItems: "center", gap: 20 }}>
          <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.green, letterSpacing: 3 }}>DATOS EN SISTEMA</div>
          {[{ v: players.length, l: "ATLETAS" }, { v: jumpRecs.length, l: "SALTOS" }, { v: velRecs.length, l: "VEL." }, { v: romRecs.length, l: "MOVIL." }, { v: liftRecs.length, l: "VBT/RM" }, { v: players.filter(p=>p.photo).length, l: "FOTOS" }].map(({ v, l }) => (
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
              <Avatar player={p} size={28} />
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
        <button onClick={deleteAllData} style={{ ...dangerBtn, background: F.red + "22", fontWeight: 700 }}>⚠ BORRAR TODO Y EMPEZAR DE CERO</button>
      </div>
    </div>
  );

  // ── Dashboard ─────────────────────────────────────────────────────────────
  const Dashboard = () => {
    const vis = visible.filter(p => (pJR(p.id).length + pVR(p.id).length + pLR(p.id).length) > 0);
    return (
      <div>
        <div style={{ display: "flex", gap: 12, marginBottom: 12, alignItems: "flex-end", flexWrap: "wrap" }}>
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

        <div style={panel}>
          <PanelHeader>COMPARATIVA DEL PLANTEL <span style={{ color: F.dim, fontSize: 9 }}>(percentil relativo al grupo cargado)</span></PanelHeader>
          {rankingData.map((d, i) => {
            const barCol = percentileColor(d.pct);
            return (
              <div key={d.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "6px 0", borderBottom: `1px solid ${F.ghost}`, cursor: "pointer" }}
                onClick={() => { setSelPlayer(d.id); goTo("evolution"); }}>
                <div style={{ fontFamily: F.fontMono, fontSize: 11, color: F.dim, width: 20, textAlign: "right" }}>{String(i + 1).padStart(2, "0")}</div>
                <Avatar player={d} size={30} />
                <div style={{ fontFamily: F.fontMono, fontSize: 12, color: F.white, flex: 1, letterSpacing: 1 }}>{d.fullName}</div>
                <PercentileChip pct={d.pct} />
                <div style={{ width: 100, height: 4, background: F.ghost, borderRadius: 2, overflow: "hidden" }}>
                  <div style={{ width: `${d.pct}%`, height: "100%", background: barCol, borderRadius: 2, boxShadow: `0 0 6px ${barCol}` }} />
                </div>
              </div>
            );
          })}
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(200px,1fr))", gap: 10, marginBottom: 14 }}>
          {vis.map((p, i) => {
            const pct = avgPercentile(p.id); const bj = bestJump(p.id);
            const bv = allDist.length ? bestVel(p.id, allDist[0]) : null; const col = TEAM_COLORS[i % TEAM_COLORS.length];
            const asym = lastAsym(p.id); const asymPct = asym ? percentileRank(asymGroup, p.id, true) : null;
            const risk = asymPct != null && asymPct < 25;
            return (
              <div key={p.id} style={{ ...panel, marginBottom: 0, cursor: "pointer", borderLeft: `3px solid ${col}` }}
                onClick={() => { setSelPlayer(p.id); goTo("evolution"); }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8, gap: 8 }}>
                  <div style={{ display: "flex", gap: 8, alignItems: "center", minWidth: 0 }}>
                    <Avatar player={p} size={38} />
                    <div style={{ minWidth: 0 }}>
                      <div style={{ fontFamily: F.fontMono, fontSize: 12, color: F.white, letterSpacing: 1, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{p.name.split(" ")[0].toUpperCase()}</div>
                      <div style={{ fontFamily: F.fontMono, fontSize: 9, color: col, letterSpacing: 1 }}>{p.team || "—"}</div>
                    </div>
                  </div>
                  <TimingScore score={pct} size={44} />
                </div>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {bj && <div style={{ background: F.ghost, borderRadius: 2, padding: "3px 7px", fontFamily: F.fontMono, fontSize: 10 }}>
                    <span style={{ color: F.dim }}>{bj.jumpTypeRaw} </span><span style={{ color: F.teal }}>{bj.altura?.toFixed(1)}cm</span>
                  </div>}
                  {bv && <div style={{ background: F.ghost, borderRadius: 2, padding: "3px 7px", fontFamily: F.fontMono, fontSize: 10 }}>
                    <span style={{ color: F.dim }}>{allDist[0]} </span><span style={{ color: F.yellow }}>{bv.time?.toFixed(2)}s</span>
                  </div>}
                  {risk && <div style={{ background: F.red + "18", border: `1px solid ${F.red}`, borderRadius: 2, padding: "3px 7px", fontFamily: F.fontMono, fontSize: 10, color: F.red }}>
                    ⚠ ASIMETRÍA {asym.asim?.toFixed(1)}%
                  </div>}
                </div>
              </div>
            );
          })}
        </div>

        {allDist.length > 0 && (
          <div style={panel}>
            <PanelHeader accent={F.yellow}>MEJORES TIEMPOS</PanelHeader>
            <ResponsiveContainer width="100%" height={Math.max(120, vis.length * 22)}>
              <BarChart data={vis.filter(p => pVR(p.id).length > 0).map(p => ({ name: p.name.split(" ")[0], ...Object.fromEntries(allDist.map(d => [d, bestVel(p.id, d)?.time ?? null])) }))} layout="vertical">
                <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} horizontal={false} />
                <XAxis type="number" tick={{ fill: F.dim, fontSize: 10, fontFamily: F.fontMono }} unit="s" />
                <YAxis dataKey="name" type="category" tick={{ fill: F.silver, fontSize: 10, fontFamily: F.fontMono }} width={65} />
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

  // ── Evolution (layout condensado en grilla, todo en una pantalla) ──────────
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

    const cardH = 155;
    return (
      <div style={{ display: "flex", flexDirection: "column", height: "calc(100vh - 110px)" }}>
        <div style={{ display: "flex", gap: 10, marginBottom: 8, alignItems: "center", flexWrap: "wrap" }}>
          <Avatar player={p} size={36} />
          <div>
            <label style={lbl}>ATLETA</label>
            <select style={sel} value={pid || ""} onChange={e => setSelPlayer(e.target.value)}>
              {activePlayers.map(pl => <option key={pl.id} value={pl.id}>{pl.name}</option>)}
            </select>
          </div>
          {p && <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 2 }}>{p.team || ""}</div>}
          <button onClick={() => goTo("informe")} style={{ marginLeft: "auto", padding: "6px 14px", background: "transparent", border: `1px solid ${F.yellow}`, borderRadius: 3, color: F.yellow, fontFamily: F.fontMono, fontSize: 9, letterSpacing: 2, cursor: "pointer", fontWeight: 700 }}>
            📄 DESCARGAR INFORME
          </button>
        </div>

        <div style={{ display: "flex", gap: 6, marginBottom: 8, flexWrap: "wrap" }}>
          {[
            { v: jrecs.length + vrecs.length + rrecs.length + lrecs.length, l: "REG.", c: F.silver },
            { v: bj?.altura != null ? bj.altura.toFixed(1) + "cm" : "—", l: "SALTO", c: F.teal },
            { v: bj?.rsi?.toFixed(2) ?? "—", l: "RSI", c: F.yellow },
            ...allDist.map(d => ({ v: bestVel(pid, d)?.time != null ? bestVel(pid, d).time.toFixed(2) + "s" : "—", l: d, c: F.red })),
          ].map(({ v, l, c }) => (
            <div key={l} style={{ ...panel, marginBottom: 0, padding: "6px 12px", flex: 1, minWidth: 70, borderTop: `2px solid ${c}` }}>
              <div style={{ fontFamily: F.fontMono, fontSize: 14, color: F.white, fontWeight: 700 }}>{v}</div>
              <div style={{ fontFamily: F.fontMono, fontSize: 8, color: F.dim, letterSpacing: 1 }}>{l}</div>
            </div>
          ))}
        </div>

        <div style={{ flex: 1, minHeight: 0, display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gridAutoRows: `${cardH}px`, gap: 8, overflowY: "auto" }}>
          {jevo.length > 0 && (
            <div style={{ ...panel, marginBottom: 0 }}>
              <PanelHeader>SALTO (cm)</PanelHeader>
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={jevo}>
                  <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                  <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} />
                  <YAxis tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} unit="cm" width={30} />
                  <Tooltip content={<CustomTooltip />} />
                  <Legend wrapperStyle={{ fontSize: 9, fontFamily: F.fontMono }} />
                  <Line dataKey="CMJ" stroke={F.teal} strokeWidth={2} dot={{ r: 3 }} connectNulls />
                  <Line dataKey="DJ" stroke={F.red} strokeWidth={2} dot={{ r: 3 }} connectNulls />
                </LineChart>
              </ResponsiveContainer>
            </div>
          )}
          {jevo.some(d => d.RSI) && (
            <div style={{ ...panel, marginBottom: 0 }}>
              <PanelHeader accent={F.yellow}>RSI MOD</PanelHeader>
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={jevo}>
                  <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                  <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} />
                  <YAxis tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} width={30} />
                  <Tooltip content={<CustomTooltip />} />
                  <Line dataKey="RSI" stroke={F.yellow} strokeWidth={2} dot={{ r: 3 }} connectNulls />
                </LineChart>
              </ResponsiveContainer>
            </div>
          )}
          {vevo.length > 0 && allDist.length > 0 && (
            <div style={{ ...panel, marginBottom: 0 }}>
              <PanelHeader accent={F.yellow}>VELOCIDAD (s)</PanelHeader>
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={vevo}>
                  <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                  <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} />
                  <YAxis tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} unit="s" width={30} />
                  <Tooltip content={<CustomTooltip />} />
                  <Legend wrapperStyle={{ fontSize: 9, fontFamily: F.fontMono }} />
                  {allDist.map((d, i) => <Line key={d} dataKey={d} stroke={TEAM_COLORS[i]} strokeWidth={2} dot={{ r: 3 }} connectNulls />)}
                </LineChart>
              </ResponsiveContainer>
            </div>
          )}
          {revo.length > 0 && (
            <div style={{ ...panel, marginBottom: 0 }}>
              <PanelHeader accent={F.purple}>MOVILIDAD — ASIMETRÍA (%)</PanelHeader>
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={revo}>
                  <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                  <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} />
                  <YAxis tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} unit="%" width={30} />
                  <Tooltip content={<CustomTooltip />} />
                  <Legend wrapperStyle={{ fontSize: 9, fontFamily: F.fontMono }} />
                  {romTests.map((t, i) => <Line key={t} dataKey={t} stroke={[F.purple, F.orange, F.teal, F.yellow][i % 4]} strokeWidth={2} dot={{ r: 3 }} connectNulls />)}
                </LineChart>
              </ResponsiveContainer>
            </div>
          )}
          {lrecs.length > 0 && rmEvo.length > 1 && (
            <div style={{ ...panel, marginBottom: 0 }}>
              <PanelHeader accent={F.orange}>1RM {playerExercises.length > 1 && (
                <select style={{ ...sel, marginLeft: 8, padding: "2px 6px", fontSize: 9 }} value={curExercise} onChange={e => setLiftExercise(e.target.value)}>
                  {playerExercises.map(ex => <option key={ex} value={ex}>{ex}</option>)}
                </select>
              )}</PanelHeader>
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={rmEvo}>
                  <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                  <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} />
                  <YAxis tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} unit="kg" width={30} />
                  <Tooltip content={<CustomTooltip />} />
                  <Line dataKey="rm" name="1RM" stroke={F.orange} strokeWidth={2} dot={{ r: 3, fill: F.orange }} connectNulls />
                </LineChart>
              </ResponsiveContainer>
            </div>
          )}
          {profile.points.length > 0 && (
            <div style={{ ...panel, marginBottom: 0 }}>
              <PanelHeader accent={F.orange}>PERFIL CARGA-VEL. · {profile.date}</PanelHeader>
              <ResponsiveContainer width="100%" height="100%">
                <ScatterChart>
                  <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                  <XAxis dataKey="load" name="Carga" unit="kg" tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} type="number" />
                  <YAxis dataKey="vel" name="Vel." unit="m/s" tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} type="number" width={30} />
                  <Tooltip content={<CustomTooltip />} cursor={{ strokeDasharray: "3 3" }} />
                  <Scatter data={profile.points} fill={F.orange} line={{ stroke: F.orange, strokeWidth: 1 }} shape="circle" />
                </ScatterChart>
              </ResponsiveContainer>
            </div>
          )}
          {vrecs.length > 0 && (
            <div style={{ ...panel, marginBottom: 0, overflowY: "auto" }}>
              <PanelHeader accent={F.yellow}>HISTORIAL VELOCIDAD</PanelHeader>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 10 }}>
                <tbody>{vrecs.slice().reverse().map((r, i) => (
                  <tr key={i} style={{ borderBottom: `1px solid ${F.ghost}` }}>
                    <td style={{ padding: "3px 6px", color: F.teal, fontFamily: F.fontMono }}>{r.fullDate}</td>
                    <td style={{ padding: "3px 6px", color: F.silver, fontFamily: F.fontMono }}>{r.distance || "—"}</td>
                    <td style={{ padding: "3px 6px", color: F.white, fontFamily: F.fontMono, fontWeight: 700 }}>{r.time?.toFixed(2)}s</td>
                  </tr>
                ))}</tbody>
              </table>
            </div>
          )}
          {rrecs.length > 0 && (
            <div style={{ ...panel, marginBottom: 0, overflowY: "auto" }}>
              <PanelHeader accent={F.purple}>HISTORIAL MOVILIDAD</PanelHeader>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 10 }}>
                <tbody>{rrecs.slice().reverse().map((r, i) => (
                  <tr key={i} style={{ borderBottom: `1px solid ${F.ghost}` }}>
                    <td style={{ padding: "3px 6px", color: F.teal, fontFamily: F.fontMono }}>{r.fullDate}</td>
                    <td style={{ padding: "3px 6px", color: F.silver, fontFamily: F.fontMono }}>{r.test}</td>
                    <td style={{ padding: "3px 6px", color: F.white, fontFamily: F.fontMono, fontWeight: 700 }}>{r.asim?.toFixed(1)}%</td>
                  </tr>
                ))}</tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    );
  };

  // ── Compare (condensado) ────────────────────────────────────────────────────
  const Compare = () => {
    const activePlayers = players.filter(p => (pJR(p.id).length + pVR(p.id).length + pRR(p.id).length + pLR(p.id).length) > 0);
    const pidA = compareA && activePlayers.find(p => p.id === compareA) ? compareA : activePlayers[0]?.id;
    const pidB = compareB && activePlayers.find(p => p.id === compareB) ? compareB : activePlayers[1]?.id || activePlayers[0]?.id;
    const pA = players.find(p => p.id === pidA); const pB = players.find(p => p.id === pidB);
    const nA = pA?.name.split(" ")[0] || "A"; const nB = pB?.name.split(" ")[0] || "B";

    const radarM = [
      { key:"cmj", label:"CMJ", vA:bestJump(pidA)?.altura, vB:bestJump(pidB)?.altura, unit:"cm", lowerIsBetter:false, group:jumpGroup },
      { key:"rsi", label:"RSI", vA:bestJump(pidA)?.rsi,    vB:bestJump(pidB)?.rsi,    unit:"", lowerIsBetter:false, group:rsiGroup },
      ...allDist.map(d => ({ key:d, label:d, vA:bestVel(pidA,d)?.time, vB:bestVel(pidB,d)?.time, unit:"s", lowerIsBetter:true, group:distGroups[d] })),
      ...["squat","bench-press"].map(k => ({ key:k, label:REFERENCE_RANGES[k].label, vA:relStrength(pidA,k), vB:relStrength(pidB,k), unit:"x PC", lowerIsBetter:false, group:strengthGroups[k] })),
    ];
    const radarD = radarM.map(m => ({ metric: m.label, [nA]: percentileRank(m.group, pidA, m.lowerIsBetter) ?? 0, [nB]: percentileRank(m.group, pidB, m.lowerIsBetter) ?? 0 }));

    const asymA = lastAsym(pidA); const asymB = lastAsym(pidB);
    const pctA = avgPercentile(pidA); const pctB = avgPercentile(pidB);

    return (
      <div style={{ display: "flex", flexDirection: "column", height: "calc(100vh - 110px)" }}>
        <div style={{ display: "flex", gap: 10, marginBottom: 8, flexWrap: "wrap" }}>
          {[
            { val: pidA, set: setCompareA, label: "ATLETA A", col: F.teal, pl: pA },
            { val: pidB, set: setCompareB, label: "ATLETA B", col: F.red, pl: pB },
          ].map(({ val, set, label, col, pl }) => (
            <div key={label} style={{ flex: 1, minWidth: 160, display: "flex", alignItems: "center", gap: 8 }}>
              <Avatar player={pl} size={32} ring={col} />
              <div style={{ flex: 1 }}>
                <label style={{ ...lbl, color: col }}>{label}</label>
                <select style={{ ...sel, width: "100%" }} value={val || ""} onChange={e => set(e.target.value)}>
                  {activePlayers.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
                </select>
              </div>
            </div>
          ))}
        </div>

        <div style={{ flex: 1, minHeight: 0, display: "grid", gridTemplateColumns: "1fr 1fr", gridTemplateRows: "auto 1fr", gap: 8, overflowY: "auto" }}>
          <div style={{ ...panel, gridColumn: "1 / -1", flexDirection: "row", alignItems: "center", padding: "10px 16px", marginBottom: 0 }}>
            {[{ pct: pctA, col: F.teal, name: nA }, { pct: pctB, col: F.red, name: nB }].map(({ pct, col, name }) => (
              <div key={name} style={{ flex: 1, display: "flex", alignItems: "center", gap: 10, justifyContent: "center" }}>
                <TimingScore score={pct} size={48} />
                <div>
                  <div style={{ fontFamily: F.fontMono, fontSize: 10, color: col, letterSpacing: 2 }}>{name.toUpperCase()}</div>
                  <PercentileChip pct={pct} />
                </div>
              </div>
            ))}
            <div style={{ textAlign: "center", paddingLeft: 16, borderLeft: `1px solid ${F.panelBorder}` }}>
              <div style={{ fontFamily: F.fontMono, fontSize: 8, color: F.dim, letterSpacing: 1 }}>DIF.</div>
              <div style={{ fontFamily: F.fontMono, fontSize: 20, color: pctA == null || pctB == null || Math.abs(pctA - pctB) < 5 ? F.silver : pctA > pctB ? F.teal : F.red, fontWeight: 700 }}>
                {pctA == null || pctB == null ? "—" : pctA === pctB ? "=" : `${Math.abs(pctA - pctB)}`}
              </div>
            </div>
          </div>

          <div style={{ ...panel, marginBottom: 0 }}>
            <PanelHeader>RADAR <span style={{ color: F.dim, fontSize: 8 }}>(percentil vs. plantel)</span></PanelHeader>
            <ResponsiveContainer width="100%" height="100%">
              <RadarChart data={radarD}>
                <PolarGrid stroke={F.ghost} />
                <PolarAngleAxis dataKey="metric" tick={{ fill: F.silver, fontSize: 9, fontFamily: F.fontMono }} />
                <PolarRadiusAxis domain={[0, 100]} tick={{ fill: F.dim, fontSize: 8, fontFamily: F.fontMono }} />
                <Radar name={nA} dataKey={nA} stroke={F.teal} fill={F.teal} fillOpacity={0.15} strokeWidth={2} />
                <Radar name={nB} dataKey={nB} stroke={F.red}  fill={F.red}  fillOpacity={0.15} strokeWidth={2} />
                <Legend wrapperStyle={{ fontSize: 9, fontFamily: F.fontMono }} />
                <Tooltip content={<CustomTooltip />} />
              </RadarChart>
            </ResponsiveContainer>
          </div>

          <div style={{ ...panel, marginBottom: 0, overflowY: "auto" }}>
            <PanelHeader>COMPARACIÓN DIRECTA</PanelHeader>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              {radarM.map(({ key, label, vA, vB, unit, lowerIsBetter }) => {
                const winner = vA == null || vB == null ? null : (lowerIsBetter ? vA < vB : vA > vB) ? "A" : (vA === vB ? null : "B");
                return (
                  <div key={key} style={{ display: "flex", alignItems: "center", gap: 6, background: F.asphalt, borderRadius: 3, padding: "4px 8px" }}>
                    <span style={{ fontFamily: F.fontMono, fontSize: 9, color: F.silver, width: 60 }}>{label}</span>
                    <span style={{ flex: 1, textAlign: "center", fontFamily: F.fontMono, fontSize: 11, fontWeight: 700, color: winner === "A" ? F.green : F.white }}>{vA != null ? `${vA > 10 ? vA.toFixed(1) : vA.toFixed(2)}${unit}` : "—"}</span>
                    <span style={{ fontSize: 9, color: F.dim }}>vs</span>
                    <span style={{ flex: 1, textAlign: "center", fontFamily: F.fontMono, fontSize: 11, fontWeight: 700, color: winner === "B" ? F.green : F.white }}>{vB != null ? `${vB > 10 ? vB.toFixed(1) : vB.toFixed(2)}${unit}` : "—"}</span>
                  </div>
                );
              })}
              {(asymA || asymB) && (
                <div style={{ display: "flex", alignItems: "center", gap: 6, background: F.asphalt, borderRadius: 3, padding: "4px 8px", marginTop: 4, borderTop: `1px solid ${F.purple}` }}>
                  <span style={{ fontFamily: F.fontMono, fontSize: 9, color: F.purple, width: 60 }}>Asimetría</span>
                  <span style={{ flex: 1, textAlign: "center", fontFamily: F.fontMono, fontSize: 11, color: F.white }}>{asymA ? `${asymA.asim?.toFixed(1)}%` : "—"}</span>
                  <span style={{ fontSize: 9, color: F.dim }}>vs</span>
                  <span style={{ flex: 1, textAlign: "center", fontFamily: F.fontMono, fontSize: 11, color: F.white }}>{asymB ? `${asymB.asim?.toFixed(1)}%` : "—"}</span>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    );
  };

  // ── Informe (per-metric: devolución + plan en cada dato que esté bajo) ─────
  const ReportView = () => {
    const activePlayers = players.filter(p => (pJR(p.id).length + pVR(p.id).length + pRR(p.id).length + pLR(p.id).length) > 0);
    const pid = selPlayer && activePlayers.find(p => p.id === selPlayer) ? selPlayer : activePlayers[0]?.id;
    const p = players.find(x => x.id === pid);
    if (!p) return null;

    const cmjRec = bestJumpByType(pid, "cmj"); const djRec = bestJumpByType(pid, "dj"); const rsiRec = bestRsi(pid);
    const jevo = jumpEvo(pid); const vevo = velEvo(pid); const revo = romEvo(pid);
    const asym = lastAsym(pid);

    const metrics = [];
    if (cmjRec?.altura != null) metrics.push({ label: "CMJ", value: cmjRec.altura, unit: "cm", ref: REFERENCE_BANDS.cmj, category: "salto" });
    if (djRec?.altura != null) metrics.push({ label: "DJ", value: djRec.altura, unit: "cm", ref: REFERENCE_BANDS.dj, category: "salto" });
    if (rsiRec?.rsi != null) metrics.push({ label: "RSI (DJ)", value: rsiRec.rsi, unit: "", ref: REFERENCE_BANDS.rsi, category: "rsi" });
    allDist.forEach(d => { const bv = bestVel(pid, d); if (bv) metrics.push({ label: `Sprint ${d}`, value: bv.time, unit: "s", ref: REFERENCE_BANDS[d] || null, category: "velocidad" }); });
    ["squat", "bench-press"].forEach(k => {
      const rs = relStrength(pid, k); if (rs == null) return;
      metrics.push({ label: REFERENCE_RANGES[k].label, value: rs, unit: "x PC", ref: REFERENCE_BANDS[k], category: k === "squat" ? "squat" : "bench" });
    });
    const romTestsForPlayer = Array.from(new Set(pRR(pid).map(r => r.test)));
    romTestsForPlayer.forEach(testName => {
      const refKey = matchRomRef(testName); if (!refKey) return;
      const recs = pRR(pid).filter(r => r.test === testName); const latest = recs[recs.length - 1]; if (!latest) return;
      const sides = [latest.angI, latest.angD].filter(v => v != null); if (!sides.length) return;
      metrics.push({ label: `${testName}`, value: Math.min(...sides), unit: "°", ref: REFERENCE_BANDS[refKey], category: "movilidad" });
    });

    const classified = metrics.map(m => ({ ...m, score: m.ref ? bandScore(m.ref, m.value) : null }));

    return (
      <div className="report-page" style={{ background: F.carbon, maxWidth: 1000, margin: "0 auto" }}>
        <style>{`
          @media print {
            .no-print { display: none !important; }
            .report-page, .report-page * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; }
            @page { size: A4 landscape; margin: 10mm; }
          }
        `}</style>

        <div className="no-print" style={{ display: "flex", gap: 10, marginBottom: 12 }}>
          <BackBtn onClick={goBack} />
          <button onClick={() => window.print()} style={{ padding: "6px 18px", background: F.yellow, color: F.carbon, border: "none", borderRadius: 3, fontFamily: F.fontMono, fontSize: 10, fontWeight: 700, letterSpacing: 2, cursor: "pointer" }}>
            🖨 IMPRIMIR / GUARDAR PDF
          </button>
          <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, alignSelf: "center" }}>Tip: en el diálogo de impresión, activá "Gráficos de fondo" para que salga a color</div>
        </div>

        <div className="report-panel" style={{ ...panel, borderTop: `3px solid ${F.red}`, flexDirection: "row", alignItems: "center", gap: 14 }}>
          <Avatar player={p} size={56} />
          <div>
            <div style={{ fontFamily: F.fontMono, fontSize: 18, color: F.white, letterSpacing: 2, fontWeight: 700 }}>{p.name}</div>
            <div style={{ fontFamily: F.fontMono, fontSize: 10, color: F.dim, marginTop: 2 }}>{p.team || "—"} · Informe generado {new Date().toLocaleDateString("es-AR")}</div>
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(230px, 1fr))", gap: 8, marginBottom: 10 }}>
          {classified.map((m, i) => {
            const fb = feedbackFor(m.category, m.score);
            return (
              <div key={i} className="report-panel" style={{ ...panel, marginBottom: 0, borderTop: `2px solid ${m.score != null ? bandLevel(m.score).color : F.ghost}` }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
                  <span style={{ fontFamily: F.fontMono, fontSize: 10, color: F.silver, letterSpacing: 1 }}>{m.label}</span>
                  <ReportBadge score={m.score} />
                </div>
                <div style={{ fontFamily: F.fontMono, fontSize: 18, color: F.white, fontWeight: 700, marginBottom: 2 }}>{m.value?.toFixed(2)}{m.unit}</div>
                <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim, marginBottom: fb ? 8 : 0 }}>{m.ref ? `Élite: ${m.ref.min}–${m.ref.max}${m.ref.unit}` : "Sin referencia"}</div>
                {fb && (
                  <div style={{ borderTop: `1px solid ${F.ghost}`, paddingTop: 6 }}>
                    <div style={{ fontFamily: F.fontF1, fontSize: 11, color: F.orange, marginBottom: 4 }}>→ A mejorar. Plan sugerido:</div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                      {fb.items.slice(0, 3).map((it, j) => (
                        <div key={j} style={{ display: "flex", justifyContent: "space-between", fontFamily: F.fontF1, fontSize: 10 }}>
                          <span style={{ color: F.white }}>{it.name}</span>
                          <span style={{ color: F.teal, fontFamily: F.fontMono, marginLeft: 6, whiteSpace: "nowrap" }}>{it.sets}</span>
                        </div>
                      ))}
                    </div>
                    <div style={{ fontFamily: F.fontMono, fontSize: 8, color: F.dim, marginTop: 4 }}>Fuente: {fb.fuente}</div>
                  </div>
                )}
              </div>
            );
          })}
          {classified.length === 0 && <div style={{ fontFamily: F.fontMono, fontSize: 11, color: F.dim }}>Sin datos suficientes todavía.</div>}
          {asym && (
            <div className="report-panel" style={{ ...panel, marginBottom: 0 }}>
              <div style={{ fontFamily: F.fontMono, fontSize: 10, color: F.purple, marginBottom: 4 }}>ASIMETRÍA {asym.test}</div>
              <div style={{ fontFamily: F.fontMono, fontSize: 18, color: F.white, fontWeight: 700 }}>{asym.asim?.toFixed(1)}%</div>
              <div style={{ fontFamily: F.fontMono, fontSize: 9, color: F.dim }}>Informativo — sin umbral de referencia cargado</div>
            </div>
          )}
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 8 }}>
          {jevo.length > 0 && (
            <div className="report-panel" style={{ ...panel, height: 170 }}>
              <PanelHeader>EVOLUCIÓN — SALTO (cm)</PanelHeader>
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={jevo}>
                  <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                  <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} />
                  <YAxis tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} unit="cm" width={30} />
                  <Tooltip content={<CustomTooltip />} />
                  <Legend wrapperStyle={{ fontSize: 9, fontFamily: F.fontMono }} />
                  <Line dataKey="CMJ" stroke={F.teal} strokeWidth={2} dot={{ r: 3 }} connectNulls />
                  <Line dataKey="DJ"  stroke={F.red}  strokeWidth={2} dot={{ r: 3 }} connectNulls />
                </LineChart>
              </ResponsiveContainer>
            </div>
          )}
          {vevo.length > 0 && allDist.length > 0 && (
            <div className="report-panel" style={{ ...panel, height: 170 }}>
              <PanelHeader accent={F.yellow}>EVOLUCIÓN — VELOCIDAD (s)</PanelHeader>
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={vevo}>
                  <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                  <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} />
                  <YAxis tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} unit="s" width={30} />
                  <Tooltip content={<CustomTooltip />} />
                  <Legend wrapperStyle={{ fontSize: 9, fontFamily: F.fontMono }} />
                  {allDist.map((d, i) => <Line key={d} dataKey={d} stroke={TEAM_COLORS[i]} strokeWidth={2} dot={{ r: 3 }} connectNulls />)}
                </LineChart>
              </ResponsiveContainer>
            </div>
          )}
          {revo.length > 0 && (
            <div className="report-panel" style={{ ...panel, height: 170 }}>
              <PanelHeader accent={F.purple}>EVOLUCIÓN — MOVILIDAD (%)</PanelHeader>
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={revo}>
                  <CartesianGrid strokeDasharray="3 3" stroke={F.ghost} />
                  <XAxis dataKey="fecha" tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} />
                  <YAxis tick={{ fill: F.dim, fontSize: 9, fontFamily: F.fontMono }} unit="%" width={30} />
                  <Tooltip content={<CustomTooltip />} />
                  <Legend wrapperStyle={{ fontSize: 9, fontFamily: F.fontMono }} />
                  {romTests.map((t, i) => <Line key={t} dataKey={t} stroke={[F.purple, F.orange, F.teal, F.yellow][i % 4]} strokeWidth={2} dot={{ r: 3 }} connectNulls />)}
                </LineChart>
              </ResponsiveContainer>
            </div>
          )}
        </div>
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
              borderBottom: tab === key ? `2px solid ${F.red}` : "2px solid transparent",
              cursor: d ? "default" : "pointer", fontFamily: F.fontMono, fontSize: 10, letterSpacing: 2,
              background: "transparent", color: tab === key ? F.white : d ? F.ghost : F.dim,
            }}>{label}</button>
          ))}
        </nav>
        {hasData && <div style={{ marginLeft: "auto", fontFamily: F.fontMono, fontSize: 9, color: F.dim, letterSpacing: 2 }}>
          {players.length} ATL · {jumpRecs.length + velRecs.length + romRecs.length + liftRecs.length + customRecs.length} REG
        </div>}
      </div>

      <div className="no-print"><TelemetryBar /></div>

      <div style={{ padding: "14px 20px", maxWidth: 1280, margin: "0 auto" }}>
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
