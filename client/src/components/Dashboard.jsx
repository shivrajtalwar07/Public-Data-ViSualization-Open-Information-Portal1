import { useState, useRef, useEffect } from "react";
import axios from "axios";
import * as XLSX from 'xlsx';
import {
  LineChart, Line, BarChart, Bar, PieChart, Pie, Cell,
  AreaChart, Area, RadarChart, Radar, PolarGrid, PolarAngleAxis,
  ResponsiveContainer, XAxis, YAxis, CartesianGrid, Tooltip, Legend
} from "recharts";

/* ─── GLOBAL STYLES ─────────────────────────────────────────────────────── */
const CSS = `
  @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800;900&family=DM+Sans:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

  :root {
    --teal:    #0ea5e9;
    --tealDk:  #0284c7;
    --coral:   #f43f5e;
    --amber:   #f59e0b;
    --violet:  #8b5cf6;
    --emerald: #10b981;
    --pink:    #ec4899;
    --bg:      #f0f4ff;
    --card:    #ffffff;
    --border:  #e2e8f0;
    --text:    #0f172a;
    --muted:   #64748b;
    --dim:     #cbd5e1;
  }

  * { margin:0; padding:0; box-sizing:border-box; }
  body { background:var(--bg); color:var(--text); font-family:'DM Sans',sans-serif; overflow-x:hidden; }
  ::-webkit-scrollbar { width:6px; }
  ::-webkit-scrollbar-track { background:var(--bg); }
  ::-webkit-scrollbar-thumb { background:var(--dim); border-radius:3px; }

  .shell { display:flex; min-height:100vh; }
  .sidebar {
    width:260px; min-width:260px; background:var(--card);
    border-right:1px solid var(--border); display:flex;
    flex-direction:column; position:sticky; top:0; height:100vh; overflow-y:auto;
    transition:width .3s;
  }
  .sidebar.collapsed { width:68px; min-width:68px; }
  .main { flex:1; display:flex; flex-direction:column; overflow-x:hidden; }

  .topbar {
    background:var(--card); border-bottom:1px solid var(--border);
    padding:0 28px; height:64px; display:flex; align-items:center;
    gap:16px; position:sticky; top:0; z-index:50; box-shadow:0 1px 12px #0001;
  }
  .topbar-title { font-family:'Outfit',sans-serif; font-weight:800; font-size:18px; }
  .topbar-title span { background:linear-gradient(135deg,var(--teal),var(--violet)); -webkit-background-clip:text; -webkit-text-fill-color:transparent; }

  .hero {
    background:linear-gradient(135deg,#0ea5e9 0%,#8b5cf6 50%,#f43f5e 100%);
    padding:40px 32px; position:relative; overflow:hidden;
  }
  .hero::before {
    content:''; position:absolute; inset:0;
    background:url("data:image/svg+xml,%3Csvg width='60' height='60' viewBox='0 0 60 60' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='none' fill-rule='evenodd'%3E%3Cg fill='%23ffffff' fill-opacity='0.06'%3E%3Cpath d='M36 34v-4h-2v4h-4v2h4v4h2v-4h4v-2h-4zm0-30V0h-2v4h-4v2h4v4h2V6h4V4h-4zM6 34v-4H4v4H0v2h4v4h2v-4h4v-2H6zM6 4V0H4v4H0v2h4v4h2V6h4V4H6z'/%3E%3C/g%3E%3C/g%3E%3C/svg%3E");
  }
  .hero-tag {
    display:inline-flex; align-items:center; gap:6px;
    background:#ffffff22; border:1px solid #ffffff44;
    padding:4px 12px; border-radius:20px; font-size:12px; font-weight:600;
    color:#fff; margin-bottom:12px; backdrop-filter:blur(8px);
  }
  .hero h1 { font-family:'Outfit',sans-serif; font-size:36px; font-weight:900; color:#fff; line-height:1.15; margin-bottom:8px; }
  .hero p { color:#ffffffcc; font-size:15px; max-width:520px; line-height:1.6; }
  .hero-stats { display:flex; gap:24px; margin-top:24px; flex-wrap:wrap; }
  .hero-stat { background:#ffffff18; backdrop-filter:blur(8px); border:1px solid #ffffff33; border-radius:12px; padding:12px 20px; }
  .hero-stat .val { font-family:'Outfit',sans-serif; font-size:24px; font-weight:800; color:#fff; }
  .hero-stat .lbl { font-size:11px; color:#ffffffaa; margin-top:2px; }

  .sb-logo { padding:18px 16px; border-bottom:1px solid var(--border); display:flex; align-items:center; gap:10px; }
  .sb-logo-icon { width:36px; height:36px; border-radius:10px; background:linear-gradient(135deg,var(--teal),var(--violet)); display:flex; align-items:center; justify-content:center; }
  .sb-logo-text { font-family:'Outfit',sans-serif; font-weight:800; font-size:15px; }
  .sb-logo-text span { color:var(--teal); }
  .sb-section { padding:14px 16px 4px; font-size:10px; font-weight:700; color:var(--dim); letter-spacing:.1em; text-transform:uppercase; }
  .sb-item {
    display:flex; align-items:center; gap:10px; padding:10px 16px;
    margin:2px 8px; border-radius:10px; cursor:pointer;
    font-size:14px; font-weight:500; color:var(--muted);
    transition:all .18s; white-space:nowrap; overflow:hidden;
  }
  .sb-item:hover { background:#f1f5f9; color:var(--text); }
  .sb-item.active { background:linear-gradient(135deg,#e0f2fe,#ede9fe); color:var(--teal); font-weight:600; }
  .sb-item svg { min-width:18px; }
  .sb-badge { margin-left:auto; background:var(--coral); color:#fff; font-size:10px; font-weight:700; border-radius:10px; padding:2px 7px; }
  .sb-divider { height:1px; background:var(--border); margin:8px 16px; }
  .sb-bottom { margin-top:auto; padding:12px; border-top:1px solid var(--border); display:flex; align-items:center; gap:10px; }

  .card { background:var(--card); border:1px solid var(--border); border-radius:16px; padding:20px; box-shadow:0 1px 8px #0002; transition:box-shadow .2s,transform .2s; }
  .card:hover { box-shadow:0 4px 20px #0003; }
  .card-title { font-family:'Outfit',sans-serif; font-size:15px; font-weight:700; color:var(--text); margin-bottom:4px; }
  .card-sub { font-size:12px; color:var(--muted); }

  .stat-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:16px; padding:24px 28px 0; }
  .stat-card {
    background:var(--card); border-radius:16px; padding:20px 22px;
    border:1px solid var(--border); position:relative; overflow:hidden;
    box-shadow:0 1px 8px #0001; transition:all .2s;
  }
  .stat-card:hover { transform:translateY(-2px); box-shadow:0 8px 24px #0002; }
  .stat-card::before { content:''; position:absolute; top:-30px; right:-30px; width:90px; height:90px; border-radius:50%; opacity:.12; }
  .stat-card.teal::before { background:var(--teal); }
  .stat-card.coral::before { background:var(--coral); }
  .stat-card.amber::before { background:var(--amber); }
  .stat-card.violet::before { background:var(--violet); }
  .stat-card.emerald::before { background:var(--emerald); }
  .stat-label { font-size:12px; font-weight:600; color:var(--muted); text-transform:uppercase; letter-spacing:.05em; margin-bottom:8px; display:flex; align-items:center; gap:6px; }
  .stat-val { font-family:'Outfit',sans-serif; font-size:32px; font-weight:800; line-height:1; margin-bottom:6px; }
  .stat-val.teal  { color:var(--teal); }
  .stat-val.coral { color:var(--coral); }
  .stat-val.amber { color:var(--amber); }
  .stat-val.violet{ color:var(--violet); }
  .stat-val.emerald{color:var(--emerald);}
  .stat-change { font-size:12px; font-weight:600; display:flex; align-items:center; gap:3px; }
  .stat-change.up { color:var(--emerald); }
  .stat-change.dn { color:var(--coral); }

  .charts-area { padding:20px 28px; display:flex; flex-direction:column; gap:20px; }
  .charts-row { display:grid; gap:16px; }
  .charts-row.two { grid-template-columns:1fr 1fr; }
  .charts-row.three { grid-template-columns:2fr 1fr; }

  .filter-panel {
    width:280px; min-width:280px; background:var(--card);
    border-left:1px solid var(--border); padding:20px; overflow-y:auto;
    height:calc(100vh - 64px); position:sticky; top:64px;
    display:flex; flex-direction:column; gap:20px;
  }
  .filter-title { font-family:'Outfit',sans-serif; font-size:14px; font-weight:700; margin-bottom:12px; display:flex; align-items:center; gap:6px; }
  .filter-group { display:flex; flex-direction:column; gap:8px; }
  .filter-label { font-size:11px; font-weight:700; color:var(--muted); text-transform:uppercase; letter-spacing:.08em; }
  select {
    width:100%; padding:9px 12px; border:1px solid var(--border); border-radius:10px;
    background:var(--bg); color:var(--text); font-size:13px; font-family:'DM Sans',sans-serif;
    cursor:pointer; outline:none; transition:border-color .18s;
  }
  select:focus { border-color:var(--teal); }
  .range-row { display:flex; gap:8px; }
  .range-row select { flex:1; }
  .checkbox-group { display:flex; flex-direction:column; gap:8px; }
  .cb-item { display:flex; align-items:center; gap:8px; cursor:pointer; font-size:13px; font-weight:500; padding:6px 10px; border-radius:8px; transition:background .15s; }
  .cb-item:hover { background:var(--bg); }
  .cb-box { width:16px; height:16px; border-radius:4px; border:2px solid var(--dim); display:flex; align-items:center; justify-content:center; transition:all .15s; flex-shrink:0; cursor:pointer; }
  .cb-box.checked { background:var(--teal); border-color:var(--teal); }
  .color-dot { width:8px; height:8px; border-radius:50%; flex-shrink:0; }

  .btn { display:inline-flex; align-items:center; gap:6px; padding:10px 20px; border-radius:10px; border:none; cursor:pointer; font-family:'DM Sans',sans-serif; font-size:13px; font-weight:600; transition:all .18s; text-decoration:none; }
  .btn-primary { background:linear-gradient(135deg,var(--teal),var(--violet)); color:#fff; box-shadow:0 4px 14px #0ea5e944; }
  .btn-primary:hover { transform:translateY(-1px); box-shadow:0 6px 20px #0ea5e966; }
  .btn-outline { background:var(--card); border:1px solid var(--border); color:var(--text); }
  .btn-outline:hover { border-color:var(--teal); color:var(--teal); }
  .btn-coral { background:linear-gradient(135deg,var(--coral),var(--pink)); color:#fff; }
  .btn-amber { background:linear-gradient(135deg,var(--amber),#f97316); color:#fff; }
  .btn-emerald { background:linear-gradient(135deg,var(--emerald),#059669); color:#fff; }
  .btn-sm { padding:6px 14px; font-size:12px; }
  .btn:disabled { opacity:.6; cursor:not-allowed; transform:none !important; }

  .table-wrap { overflow-x:auto; border-radius:12px; border:1px solid var(--border); }
  table { width:100%; border-collapse:collapse; }
  thead th { padding:11px 16px; text-align:left; font-size:11px; font-weight:700; color:var(--muted); text-transform:uppercase; letter-spacing:.07em; background:var(--bg); border-bottom:1px solid var(--border); }
  tbody tr { border-bottom:1px solid rgba(226,232,240,.4); transition:background .15s; }
  tbody tr:last-child { border-bottom:none; }
  tbody tr:hover { background:#f8faff; }
  tbody td { padding:11px 16px; font-size:13px; color:var(--text); }
  .td-mono { font-family:'JetBrains Mono',monospace; font-size:12px; }

  .badge { display:inline-flex; align-items:center; padding:3px 9px; border-radius:20px; font-size:11px; font-weight:700; }
  .badge-teal    { background:#e0f2fe; color:var(--teal); }
  .badge-coral   { background:#ffe4e6; color:var(--coral); }
  .badge-amber   { background:#fef3c7; color:#d97706; }
  .badge-violet  { background:#ede9fe; color:var(--violet); }
  .badge-emerald { background:#d1fae5; color:#059669; }

  .search-wrap { position:relative; flex:1; max-width:340px; }
  .search-wrap input { width:100%; padding:9px 12px 9px 36px; border:1px solid var(--border); border-radius:10px; background:var(--bg); font-size:13px; font-family:'DM Sans',sans-serif; outline:none; transition:border-color .18s; }
  .search-wrap input:focus { border-color:var(--teal); }
  .search-icon { position:absolute; left:10px; top:50%; transform:translateY(-50%); color:var(--muted); pointer-events:none; }

  .dataset-tabs { display:flex; gap:8px; flex-wrap:wrap; padding:20px 28px 0; }
  .ds-tab { padding:8px 18px; border-radius:20px; border:2px solid var(--border); background:var(--card); font-size:13px; font-weight:600; cursor:pointer; transition:all .18s; display:flex; align-items:center; gap:6px; }
  .ds-tab:hover { border-color:var(--teal); color:var(--teal); }
  .ds-tab.active { background:linear-gradient(135deg,var(--teal),var(--violet)); border-color:transparent; color:#fff; box-shadow:0 4px 14px #0ea5e944; }

  .city-chips { display:flex; gap:8px; flex-wrap:wrap; margin-top:8px; }
  .city-chip { display:flex; align-items:center; gap:6px; padding:6px 12px; border-radius:20px; font-size:12px; font-weight:600; cursor:pointer; transition:all .18s; }
  .city-chip.sel { box-shadow:0 2px 10px #0002; }

  .upload-zone {
    border:2px dashed var(--border); border-radius:16px; padding:48px 24px;
    text-align:center; cursor:pointer; transition:all .2s; background:var(--bg);
  }
  .upload-zone:hover, .upload-zone.drag { border-color:var(--teal); background:#e0f2fe22; }
  .upload-icon { width:64px; height:64px; border-radius:50%; background:linear-gradient(135deg,#e0f2fe,#ede9fe); display:flex; align-items:center; justify-content:center; margin:0 auto 16px; }
  .progress-bar { height:8px; background:var(--border); border-radius:4px; overflow:hidden; margin-top:6px; }
  .progress-fill { height:100%; border-radius:4px; background:linear-gradient(90deg,var(--teal),var(--violet)); transition:width .4s ease; }

  @keyframes fadeUp { from{opacity:0;transform:translateY(16px)} to{opacity:1;transform:translateY(0)} }
  @keyframes pulse  { 0%,100%{opacity:1} 50%{opacity:.5} }
  @keyframes spin   { from{transform:rotate(0deg)} to{transform:rotate(360deg)} }
  .fade-up   { animation:fadeUp .35s ease both; }
  .fade-up-1 { animation:fadeUp .35s .05s ease both; }
  .fade-up-2 { animation:fadeUp .35s .10s ease both; }
  .fade-up-3 { animation:fadeUp .35s .15s ease both; }
  .fade-up-4 { animation:fadeUp .35s .20s ease both; }
  .pulse { animation:pulse 2s infinite; }

  .live-dot { width:8px; height:8px; border-radius:50%; background:var(--emerald); display:inline-block; position:relative; }
  .live-dot::after { content:''; position:absolute; inset:-3px; border-radius:50%; background:var(--emerald); opacity:.3; animation:pulse 1.5s infinite; }

  .dl-card { display:flex; align-items:center; gap:14px; padding:16px; border:1px solid var(--border); border-radius:12px; background:var(--card); transition:all .2s; }
  .dl-card:hover { border-color:var(--teal); box-shadow:0 4px 16px #0ea5e922; transform:translateY(-1px); }
  .dl-icon { width:44px; height:44px; border-radius:12px; display:flex; align-items:center; justify-content:center; font-size:20px; flex-shrink:0; }

  .notif { position:fixed; bottom:24px; right:24px; background:var(--text); color:#fff; padding:12px 20px; border-radius:12px; font-size:13px; font-weight:500; z-index:1000; animation:fadeUp .3s ease; box-shadow:0 8px 24px #0004; display:flex; align-items:center; gap:8px; }

  .chart-tabs { display:flex; background:var(--bg); border-radius:10px; padding:4px; gap:4px; margin-bottom:16px; }
  .ct-tab { flex:1; padding:7px; text-align:center; border-radius:8px; cursor:pointer; font-size:12px; font-weight:600; color:var(--muted); transition:all .18s; }
  .ct-tab.active { background:var(--card); color:var(--teal); box-shadow:0 1px 4px #0001; }

  @media(max-width:900px) {
    .filter-panel { display:none; }
    .charts-row.two   { grid-template-columns:1fr; }
    .charts-row.three { grid-template-columns:1fr; }
  }
`;

/* ─── CONSTANTS ─────────────────────────────────────────────────────────── */
const COLORS = ["#0ea5e9", "#8b5cf6", "#f43f5e", "#f59e0b", "#10b981", "#ec4899", "#06b6d4", "#84cc16"];

const CITIES = [
  { name: "Mumbai", color: "#0ea5e9" },
  { name: "Delhi", color: "#8b5cf6" },
  { name: "Bangalore", color: "#10b981" },
  { name: "Chennai", color: "#f43f5e" },
  { name: "Hyderabad", color: "#f59e0b" },
  { name: "Pune", color: "#ec4899" },
];

const DATASETS = [
  { key: "population", label: "👥 Population", icon: "👥" },
  { key: "gdp", label: "💰 GDP", icon: "💰" },
  { key: "health", label: "🏥 Healthcare", icon: "🏥" },
  { key: "education", label: "🎓 Education", icon: "🎓" },
  { key: "environment", label: "🌿 Environment", icon: "🌿" },
];

const TS = {
  contentStyle: { background: "#fff", border: "1px solid #e2e8f0", borderRadius: 10, fontFamily: "DM Sans", fontSize: 12 },
  labelStyle: { color: "#64748b", fontWeight: 600 },
};

/* ─── DATA GENERATORS ───────────────────────────────────────────────────── */
function makeYearData(ds) {
  const base = { population: 820, gdp: 420, health: 310, education: 550, environment: 200 };
  return ["2020", "2021", "2022", "2023", "2024", "2025", "2026"].map((y, i) => ({
    year: y,
    Mumbai: +(base[ds] + Math.random() * 80 + i * 15).toFixed(1),
    Delhi: +(base[ds] + Math.random() * 80 + i * 12).toFixed(1),
    Bangalore: +(base[ds] + Math.random() * 80 + i * 18).toFixed(1),
    Chennai: +(base[ds] + Math.random() * 80 + i * 10).toFixed(1),
    Hyderabad: +(base[ds] + Math.random() * 80 + i * 14).toFixed(1),
    Pune: +(base[ds] + Math.random() * 80 + i * 9).toFixed(1),
  }));
}

function makePieData(ds) {
  const labels = {
    population: ["Urban", "Rural", "Suburban", "Tribal"],
    gdp: ["Services", "Manufacturing", "Agriculture", "Construction"],
    health: ["Hospitals", "Clinics", "PHC", "Others"],
    education: ["Primary", "Secondary", "Higher", "Vocational"],
    environment: ["Forest", "Wetland", "Agricultural", "Urban"],
  };
  const vals = [40, 28, 20, 12];
  return (labels[ds] || labels.population).map((n, i) => ({
    name: n, value: vals[i] + Math.floor(Math.random() * 10 - 5)
  }));
}

function makeBarData() {
  return CITIES.map(c => ({
    city: c.name,
    value: +(200 + Math.random() * 600).toFixed(0),
    target: +(300 + Math.random() * 400).toFixed(0),
  }));
}

const INITIAL_DATASETS = [
  { id: 1, name: "Census 2024 - Urban Areas", category: "Population", size: "2.4 MB", rows: 14200, status: "Active", date: "2024-03-01" },
  { id: 2, name: "GDP Q4 2023 Report", category: "Economics", size: "1.1 MB", rows: 8400, status: "Active", date: "2024-02-20" },
  { id: 3, name: "Healthcare Infrastructure", category: "Health", size: "3.8 MB", rows: 22100, status: "Processing", date: "2024-03-05" },
  { id: 4, name: "Literacy Rate Survey", category: "Education", size: "0.9 MB", rows: 5600, status: "Active", date: "2024-02-15" },
  { id: 5, name: "Air Quality Index 2023", category: "Environment", size: "4.2 MB", rows: 31500, status: "Active", date: "2024-03-08" },
];

/* ─── SVG ICON HELPER ───────────────────────────────────────────────────── */
const I = ({ d, size = 18, c = "currentColor", sw = 2, className = "" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none"
    className={className}
    stroke={c} strokeWidth={sw} strokeLinecap="round" strokeLinejoin="round">
    <path d={d} />
  </svg>
);

const ic = {
  home: "M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z",
  chart: "M18 20V10M12 20V4M6 20v-6",
  pie: "M21.21 15.89A10 10 0 118 2.83M22 12A10 10 0 0012 2v10z",
  table: "M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2",
  dl: "M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3",
  admin: "M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z",
  filter: "M22 3H2l8 9.46V19l4 2V12.46L22 3z",
  upload: "M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12",
  check: "M20 6L9 17l-5-5",
  x: "M18 6L6 18M6 6l12 12",
  menu: "M3 12h18M3 6h18M3 18h18",
  refresh: "M23 4v6h-6M1 20v-6h6M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15",
  search: "M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0",
  bell: "M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9M13.73 21a2 2 0 01-3.46 0",
  user: "M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2M12 11a4 4 0 100-8 4 4 0 000 8z",
  settings: "M12 20a8 8 0 100-16 8 8 0 000 16zM12 14a2 2 0 100-4 2 2 0 000 4z",
  csv: "M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z",
  pdf: "M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8zM14 2v6h6",
  eye: "M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8zM12 9a3 3 0 100 6 3 3 0 000-6z",
  trash: "M3 6h18M8 6V4h8v2M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6",
  plus: "M12 5v14M5 12h14",
  globe: "M12 22a10 10 0 100-20 10 10 0 000 20zM2 12h20M12 2a15.3 15.3 0 010 20M12 2a15.3 15.3 0 000 20",
  lightning: "M13 2L3 14h9l-1 8 10-12h-9l1-8z",
};

/* ─── NOTIFICATION HOOK ─────────────────────────────────────────────────── */
function useNotif() {
  const [msg, setMsg] = useState(null);
  const show = (m) => { setMsg(m); setTimeout(() => setMsg(null), 2500); };
  return [msg, show];
}

/* ─── FILTER PANEL ──────────────────────────────────────────────────────── */
function FilterPanel({ filters, setFilters, selectedCities, setSelectedCities }) {
  const years = ["All", "2020", "2021", "2022", "2023", "2024", "2025", "2026"];
  const regions = ["All India", "North", "South", "East", "West", "Central", "Northeast"];
  const categories = DATASETS.map((d, i) => ({ key: d.key, label: d.label, color: COLORS[i] }));

  const toggleCat = (key) => {
    setFilters(f => ({
      ...f,
      categories: f.categories.includes(key)
        ? f.categories.filter(x => x !== key)
        : [...f.categories, key],
    }));
  };

  return (
    <div className="filter-panel">
      <div className="filter-title" style={{ fontSize: 15 }}>
        <I d={ic.filter} size={16} c="#0ea5e9" /> Filters &amp; Controls
      </div>

      {/* YEAR */}
      <div className="filter-group">
        <div className="filter-label">📅 Year Range</div>
        <div className="range-row">
          <select value={filters.yearFrom} onChange={e => setFilters(f => ({ ...f, yearFrom: e.target.value }))}>
            {years.map(y => <option key={y}>{y}</option>)}
          </select>
          <select value={filters.yearTo} onChange={e => setFilters(f => ({ ...f, yearTo: e.target.value }))}>
            {years.map(y => <option key={y}>{y}</option>)}
          </select>
        </div>
      </div>

      <div style={{ height: 1, background: "var(--border)", margin: "4px 0" }} />

      {/* REGION */}
      <div className="filter-group">
        <div className="filter-label">🗺 Region</div>
        <select value={filters.region} onChange={e => setFilters(f => ({ ...f, region: e.target.value }))}>
          {regions.map(r => <option key={r}>{r}</option>)}
        </select>
      </div>

      <div style={{ height: 1, background: "var(--border)", margin: "4px 0" }} />

      {/* CATEGORY */}
      <div className="filter-group">
        <div className="filter-label">📊 Category</div>
        <div className="checkbox-group">
          {categories.map(c => (
            <div key={c.key} className="cb-item" onClick={() => toggleCat(c.key)}>
              <div className={`cb-box${filters.categories.includes(c.key) ? " checked" : ""}`}>
                {filters.categories.includes(c.key) &&
                  <I d={ic.check} size={10} c="#fff" sw={3} />}
              </div>
              <span>{c.label}</span>
            </div>
          ))}
        </div>
      </div>

      <div style={{ height: 1, background: "var(--border)", margin: "4px 0" }} />

      {/* CITIES */}
      <div className="filter-group">
        <div className="filter-label">🏙 Compare Cities</div>
        <div className="city-chips">
          {CITIES.map(c => {
            const sel = selectedCities.includes(c.name);
            return (
              <div key={c.name}
                className={`city-chip${sel ? " sel" : ""}`}
                style={{
                  background: sel ? c.color + "22" : "var(--bg)",
                  border: `2px solid ${sel ? c.color : "var(--border)"}`,
                  color: sel ? c.color : "var(--muted)",
                }}
                onClick={() => setSelectedCities(prev =>
                  sel ? prev.filter(x => x !== c.name) : [...prev, c.name]
                )}>
                <div className="color-dot" style={{ background: c.color }} />
                {c.name}
              </div>
            );
          })}
        </div>
      </div>

      <button className="btn btn-primary" style={{ width: "100%", justifyContent: "center" }}>
        <I d={ic.filter} size={14} /> Apply Filters
      </button>
      <button className="btn btn-outline" style={{ width: "100%", justifyContent: "center", marginTop: 8 }}>
        <I d={ic.refresh} size={14} /> Reset
      </button>
    </div>
  );
}

/* ─── CITIZEN DASHBOARD ─────────────────────────────────────────────────── */
function CitizenDashboard({ selectedCities, showNotif, customData }) {
  const [ds, setDs] = useState("population");
  const [chartTab, setChartTab] = useState("line");
  const [yearData, setYearData] = useState([]);
  const [pieData, setPieData] = useState([]);
  const [barData, setBarData] = useState([]);
  const [isAiLoading, setIsAiLoading] = useState(false);

  const activeCities = selectedCities.length ? selectedCities : ["Mumbai", "Delhi", "Bangalore"];

  // Initialize with mock data
  useEffect(() => {
    if (customData) {
      if (customData.yearData) setYearData(customData.yearData);
      if (customData.pieData) setPieData(customData.pieData);
      if (customData.barData) setBarData(customData.barData);
    } else {
      setYearData(makeYearData(ds));
      setPieData(makePieData(ds));
      setBarData(makeBarData());
    }
  }, [ds, customData]);

  const refreshWithAI = async () => {
    setIsAiLoading(true);
    showNotif("AI is generating real-time data...");

    const prompt = `Generate realistic statistical data for the "${ds}" dataset in India for the years 2020-2026. 
    Cities: Mumbai, Delhi, Bangalore, Chennai, Hyderabad, Pune.
    Include:
    1. "yearData": Array of 7 objects (years 2020-2026). Each object must have 'year' (string) and values for all cities.
    2. "pieData": A breakdown relevant to ${ds} (at least 4 segments with 'name' and 'value').
    3. "barData": Array of objects for each city with 'city', 'value', and 'target'.
    
    Ensure the data reflects realistic trends (e.g., growth in population/GDP). Return ONLY the JSON object.`;

    try {
      // Inline fetch to avoid import issues in this large file
      const API_KEY = "AIzaSyBk7paPnpP523j2LKgRegRZFCCMOKaHmrM";
      const API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent";

      const response = await fetch(`${API_URL}?key=${API_KEY}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [{ parts: [{ text: `${prompt}\n\nReturn ONLY raw JSON.` }] }],
          generationConfig: { response_mime_type: "application/json" }
        })
      });

      if (!response.ok) throw new Error("AI Fetch Failed");
      const data = await response.json();
      const result = JSON.parse(data.candidates[0].content.parts[0].text);

      if (result.yearData) setYearData(result.yearData);
      if (result.pieData) setPieData(result.pieData);
      if (result.barData) setBarData(result.barData);

      showNotif("Dashboard updated with AI insights! ✨");
    } catch (err) {
      console.error(err);
      showNotif("AI Refresh failed. Using default data.");
    } finally {
      setIsAiLoading(false);
    }
  };

  const stats = [
    { label: "Total Records", val: "2.4M", change: "+12%", dir: "up", color: "teal", icon: "📊" },
    { label: "Datasets", val: "48", change: "+5", dir: "up", color: "violet", icon: "🗂" },
    { label: "Cities Covered", val: "128", change: "+8", dir: "up", color: "emerald", icon: "🏙" },
    { label: "Last Updated", val: "Today", change: "Live", dir: "up", color: "amber", icon: "🔄" },
    { label: "Downloads", val: "18.2K", change: "+22%", dir: "up", color: "coral", icon: "⬇" },
  ];

  const downloadCSV = () => {
    if (!yearData || yearData.length === 0) {
      showNotif("No data to export");
      return;
    }

    // 1. Define headers
    const headers = ["Year", ...CITIES.map(c => c.name), "Average"];

    // 2. Map data rows
    const rows = yearData.map(row => {
      const cityVals = CITIES.map(c => row[c.name]);
      const avg = (cityVals.reduce((a, b) => a + b, 0) / cityVals.length).toFixed(1);
      return [row.year, ...cityVals, avg].join(",");
    });

    // 3. Combine headers and rows
    const csvContent = [headers.join(","), ...rows].join("\n");

    // 4. Create blob and download link
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `${ds}_data_export_${new Date().toISOString().split('T')[0]}.csv`);
    link.style.visibility = "hidden";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    showNotif("CSV downloaded successfully! ✓");
  };

  const downloadPDF = () => {
    const reportTitle = `${ds.toUpperCase()} DATA REPORT`;
    const date = new Date().toLocaleString();

    let content = `%PDF-1.4\n`;
    content += `1 0 obj << /Title (${reportTitle}) /Author (DataPortal) >> endobj\n`;
    content += `[REPORT SUMMARY]\n`;
    content += `Dataset: ${ds}\n`;
    content += `Generated on: ${date}\n\n`;
    content += `[DATA TABLE]\n`;
    content += `Year | ` + CITIES.map(c => c.name).join(" | ") + " | Average\n";
    content += `-`.repeat(80) + "\n";

    yearData.forEach(row => {
      const cityVals = CITIES.map(c => row[c.name]);
      const avg = (cityVals.reduce((a, b) => a + b, 0) / cityVals.length).toFixed(1);
      content += `${row.year} | ${cityVals.join(" | ")} | ${avg}\n`;
    });

    const blob = new Blob([content], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `${ds}_report_${new Date().getTime()}.pdf`);
    link.click();

    showNotif("PDF Report generated successfully! ✓");
  };

  return (
    <div style={{ flex: 1, overflowY: "auto" }}>
      {/* HERO */}
      <div className="hero fade-up">
        <div className="hero-tag">
          <span className="live-dot" />Live Portal
        </div>
        <h1>India <span style={{ color: "rgba(255,255,255,0.8)" }}>AI</span> Data Engine</h1>
        <p>Explore, analyze, and download India's most comprehensive public datasets through interactive visualizations and powerful filters.</p>
        <div style={{ marginTop: 24, display: "flex", gap: 12 }}>


        </div>
        <div className="hero-stats">
          {[["2.4M+", "Records"], ["48", "Datasets"], ["128", "Cities"], ["99.9%", "Uptime"]].map(([v, l]) => (
            <div className="hero-stat" key={l}>
              <div className="val">{v}</div>
              <div className="lbl">{l}</div>
            </div>
          ))}
        </div>
      </div>

      {/* KPI */}
      <div className="stat-grid fade-up-1">
        {stats.map(s => (
          <div key={s.label} className={`stat-card ${s.color}`}>
            <div className="stat-label"><span>{s.icon}</span>{s.label}</div>
            <div className={`stat-val ${s.color}`}>{s.val}</div>
            <div className={`stat-change ${s.dir}`}>
              {s.dir === "up" ? "▲" : "▼"} {s.change} vs last month
            </div>
          </div>
        ))}
      </div>

      {/* DATASET SELECTOR */}
      <div style={{ padding: "20px 28px 0" }}>
        <div style={{ fontFamily: "Outfit", fontWeight: 700, fontSize: 14, marginBottom: 10, color: "var(--muted)", textTransform: "uppercase", letterSpacing: ".08em" }}>
          📂 Select Dataset
        </div>
        <div className="dataset-tabs fade-up-2">
          {DATASETS.map(d => (
            <div key={d.key} className={`ds-tab${ds === d.key ? " active" : ""}`} onClick={() => setDs(d.key)}>
              {d.icon} {d.label.replace(/^\S+ /, "")}
            </div>
          ))}
        </div>
      </div>

      {/* CHARTS */}
      <div className="charts-area fade-up-3">
        {/* MAIN CHART */}
        <div className="card">
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
            <div>
              <div className="card-title">📈 Trend Analysis — {ds.charAt(0).toUpperCase() + ds.slice(1)}</div>
              <div className="card-sub">Year-over-year comparison across selected cities</div>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button className={`btn btn-outline btn-sm ${isAiLoading ? 'pulse' : ''}`}
                onClick={refreshWithAI}
                disabled={isAiLoading}>
                <I d={ic.refresh} size={13} className={isAiLoading ? "spin" : ""} />
              </button>
              <button className="btn btn-outline btn-sm" onClick={downloadCSV}>
                <I d={ic.dl} size={13} />
              </button>
            </div>
          </div>

          <div className="chart-tabs">
            {[["line", "📈"], ["bar", "📊"], ["area", "📉"]].map(([k, l]) => (
              <div key={k} className={`ct-tab${chartTab === k ? " active" : ""}`} onClick={() => setChartTab(k)}>{l}</div>
            ))}
          </div>

          <ResponsiveContainer width="100%" height={300}>
            {chartTab === "line" ? (
              <LineChart data={yearData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                <XAxis dataKey="year" tick={{ fontSize: 12, fontFamily: "DM Sans" }} />
                <YAxis tick={{ fontSize: 12, fontFamily: "DM Sans" }} />
                <Tooltip {...TS} />
                <Legend wrapperStyle={{ fontSize: 12, fontFamily: "DM Sans" }} />
                {activeCities.map((c, i) => (
                  <Line key={c} type="monotone" dataKey={c}
                    stroke={CITIES.find(x => x.name === c)?.color || COLORS[i]}
                    strokeWidth={2.5} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                ))}
              </LineChart>
            ) : chartTab === "bar" ? (
              <BarChart data={yearData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                <XAxis dataKey="year" tick={{ fontSize: 12 }} />
                <YAxis tick={{ fontSize: 12 }} />
                <Tooltip {...TS} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                {activeCities.map((c, i) => (
                  <Bar key={c} dataKey={c}
                    fill={CITIES.find(x => x.name === c)?.color || COLORS[i]}
                    radius={[4, 4, 0, 0]} />
                ))}
              </BarChart>
            ) : (
              <AreaChart data={yearData}>
                <defs>
                  {activeCities.map((c, i) => (
                    <linearGradient key={c} id={`ag${i}`} x1="0" y1="0" x2="0" y2="1">
                      <stop offset="0%" stopColor={CITIES.find(x => x.name === c)?.color || COLORS[i]} stopOpacity={0.30} />
                      <stop offset="100%" stopColor={CITIES.find(x => x.name === c)?.color || COLORS[i]} stopOpacity={0.02} />
                    </linearGradient>
                  ))}
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                <XAxis dataKey="year" tick={{ fontSize: 12 }} />
                <YAxis tick={{ fontSize: 12 }} />
                <Tooltip {...TS} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                {activeCities.map((c, i) => (
                  <Area key={c} type="monotone" dataKey={c}
                    stroke={CITIES.find(x => x.name === c)?.color || COLORS[i]}
                    fill={`url(#ag${i})`} strokeWidth={2} />
                ))}
              </AreaChart>
            )}
          </ResponsiveContainer>
        </div>

        {/* PIE + BAR - HIDDEN IF CUSTOM DATA IS ACTIVE */}
        {!customData && (
          <div className="charts-row two">
            <div className="card">
              <div className="card-title">🥧 Distribution — {ds.charAt(0).toUpperCase() + ds.slice(1)}</div>
              <div className="card-sub">Proportional breakdown by segment</div>
              <ResponsiveContainer width="100%" height={260}>
                <PieChart>
                  <Pie data={pieData} cx="50%" cy="50%"
                    innerRadius={60} outerRadius={100} paddingAngle={4} dataKey="value"
                    label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}
                    labelLine={false}>
                    {pieData.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
                  </Pie>
                  <Tooltip {...TS} formatter={v => [`${v}%`, "Share"]} />
                  <Legend iconType="circle" iconSize={9} wrapperStyle={{ fontSize: 12 }} />
                </PieChart>
              </ResponsiveContainer>
            </div>

            <div className="card">
              <div className="card-title">📊 City Comparison</div>
              <div className="card-sub">Actual vs Target across cities</div>
              <ResponsiveContainer width="100%" height={260}>
                <BarChart data={barData} layout="vertical" barCategoryGap="30%">
                  <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                  <XAxis type="number" tick={{ fontSize: 11 }} />
                  <YAxis dataKey="city" type="category" tick={{ fontSize: 11 }} width={75} />
                  <Tooltip {...TS} />
                  <Legend wrapperStyle={{ fontSize: 11 }} />
                  <Bar dataKey="value" name="Actual" radius={[0, 4, 4, 0]}>
                    {barData.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
                  </Bar>
                  <Bar dataKey="target" name="Target" fill="#e2e8f0" radius={[0, 4, 4, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        )}

        {/* DATA TABLE */}
        <div className="card">
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
            <div>
              <div className="card-title">📋 Data Table</div>
              <div className="card-sub">Yearly values for all cities</div>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button className="btn btn-coral btn-sm" onClick={downloadPDF}>
                <I d={ic.pdf} size={13} /> Export PDF
              </button>
              <button className="btn btn-emerald btn-sm" onClick={downloadCSV}>
                <I d={ic.dl} size={13} /> Export CSV
              </button>
            </div>
          </div>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Year</th>
                  {CITIES.map(c => <th key={c.name}>{c.name}</th>)}
                  <th>Avg</th>
                </tr>
              </thead>
              <tbody>
                {yearData.map(row => {
                  const vals = CITIES.map(c => row[c.name]);
                  const avg = (vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(1);
                  return (
                    <tr key={row.year}>
                      <td><span className="badge badge-teal">{row.year}</span></td>
                      {CITIES.map(c => <td key={c.name} className="td-mono">{row[c.name]}</td>)}
                      <td className="td-mono" style={{ fontWeight: 700, color: "var(--violet)" }}>{avg}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ─── CHARTS EXPLORER ───────────────────────────────────────────────────── */
function ChartsExplorer({ showNotif }) {
  const [active, setActive] = useState("line");
  const yearData = makeYearData("gdp");
  const pieData = makePieData("education");
  const barData = makeBarData();
  const [radarData, setRadarData] = useState([]);

  useEffect(() => {
    setRadarData(CITIES.map(c => ({ city: c.name, score: +(40 + Math.random() * 60).toFixed(0) })));
  }, []);

  const downloadCSV = () => {
    let headers = [];
    let rows = [];
    let filename = `export_${active}_${new Date().toISOString().split('T')[0]}.csv`;

    if (active === "line") {
      headers = ["Year", ...CITIES.map(c => c.name)];
      rows = yearData.map(row => [row.year, ...CITIES.map(c => row[c.name])].join(","));
    } else if (active === "bar") {
      headers = ["City", "Value", "Target"];
      rows = barData.map(row => [row.city, row.value, row.target].join(","));
    } else if (active === "pie") {
      headers = ["Name", "Value"];
      rows = pieData.map(row => [row.name, row.value].join(","));
    } else if (active === "radar") {
      headers = ["City", "Score"];
      rows = radarData.map(row => [row.city, row.score].join(","));
    }

    const csvContent = [headers.join(","), ...rows].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", filename);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showNotif(`${active.toUpperCase()} data exported ✓`);
  };

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: "28px" }}>
      <div className="fade-up" style={{ marginBottom: 24 }}>
        <div style={{ fontFamily: "Outfit", fontWeight: 900, fontSize: 28, marginBottom: 4 }}>Interactive Charts</div>
        <div style={{ color: "var(--muted)", fontSize: 14 }}>Explore data through multiple visualization types</div>
      </div>

      <div style={{ display: "flex", gap: 10, marginBottom: 24, flexWrap: "wrap" }} className="fade-up-1">
        {[["line", "📈", "Line Chart"], ["bar", "📊", "Bar Chart"], ["pie", "🥧", "Pie Chart"], ["radar", "🕸", "Radar Chart"]].map(([k, em, l]) => (
          <button key={k} className={`btn ${active === k ? "btn-primary" : "btn-outline"}`} onClick={() => setActive(k)}>
            {em} {l}
          </button>
        ))}
      </div>

      <div className="fade-up-2" style={{ marginBottom: 20 }}>
        <div className="card">
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
            <div className="card-title" style={{ fontSize: 17 }}>
              {active === "line" ? "📈 Line Chart — GDP Trends"
                : active === "bar" ? "📊 Bar Chart — Healthcare Metrics"
                  : active === "pie" ? "🥧 Pie Chart — Education Distribution"
                    : "🕸 Radar Chart — City Performance Score"}
            </div>
            <button className="btn btn-outline btn-sm" onClick={downloadCSV}>
              <I d={ic.dl} size={13} /> Export
            </button>
          </div>
          <ResponsiveContainer width="100%" height={320}>
            {active === "line" ? (
              <LineChart data={yearData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                <XAxis dataKey="year" tick={{ fontSize: 12 }} />
                <YAxis tick={{ fontSize: 12 }} />
                <Tooltip {...TS} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                {CITIES.slice(0, 4).map((c) => (
                  <Line key={c.name} type="monotone" dataKey={c.name}
                    stroke={c.color} strokeWidth={2.5} dot={{ r: 4 }} />
                ))}
              </LineChart>
            ) : active === "bar" ? (
              <BarChart data={barData} barCategoryGap="25%">
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                <XAxis dataKey="city" tick={{ fontSize: 12 }} />
                <YAxis tick={{ fontSize: 12 }} />
                <Tooltip {...TS} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                <Bar dataKey="value" name="Score" radius={[6, 6, 0, 0]}>
                  {barData.map((_, i) => <Cell key={i} fill={COLORS[i]} />)}
                </Bar>
                <Bar dataKey="target" name="Target" fill="#cbd5e1" radius={[6, 6, 0, 0]} />
              </BarChart>
            ) : active === "pie" ? (
              <PieChart>
                <Pie data={pieData} cx="50%" cy="50%" outerRadius={130} paddingAngle={3} dataKey="value"
                  label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}>
                  {pieData.map((_, i) => <Cell key={i} fill={COLORS[i]} />)}
                </Pie>
                <Tooltip {...TS} />
                <Legend iconType="circle" iconSize={10} wrapperStyle={{ fontSize: 12 }} />
              </PieChart>
            ) : (
              <RadarChart cx="50%" cy="50%" outerRadius={130} data={radarData}>
                <PolarGrid stroke="#e2e8f0" />
                <PolarAngleAxis dataKey="city" tick={{ fontSize: 12 }} />
                <Radar dataKey="score" stroke="#0ea5e9" fill="#0ea5e9" fillOpacity={0.25} strokeWidth={2} />
                <Tooltip {...TS} />
              </RadarChart>
            )}
          </ResponsiveContainer>
        </div>
      </div>

      <div className="fade-up-3" style={{ fontFamily: "Outfit", fontWeight: 700, fontSize: 15, marginBottom: 14 }}>📊 Chart Gallery</div>
      <div className="charts-row two fade-up-4">
        {[
          { title: "Population Growth", sub: "2020–2026 trend", data: yearData, type: "area" },
          { title: "City Distribution", sub: "Population share", data: pieData, type: "pie" },
        ].map(({ title, sub, data, type }, idx) => (
          <div key={idx} className="card">
            <div className="card-title">{title}</div>
            <div className="card-sub">{sub}</div>
            <ResponsiveContainer width="100%" height={180}>
              {type === "area" ? (
                <AreaChart data={data.slice(0, 5)}>
                  <defs>
                    <linearGradient id={`mini${idx}`} x1="0" y1="0" x2="0" y2="1">
                      <stop offset="0%" stopColor={COLORS[idx * 2]} stopOpacity={0.35} />
                      <stop offset="100%" stopColor={COLORS[idx * 2]} stopOpacity={0} />
                    </linearGradient>
                  </defs>
                  <XAxis dataKey="year" tick={{ fontSize: 10 }} hide />
                  <YAxis tick={{ fontSize: 10 }} hide />
                  <Tooltip {...TS} />
                  <Area type="monotone" dataKey="Mumbai" stroke={COLORS[0]} fill={`url(#mini${idx})`} strokeWidth={2} />
                </AreaChart>
              ) : (
                <PieChart>
                  <Pie data={data} cx="50%" cy="50%" outerRadius={65} dataKey="value">
                    {data.map((_, i) => <Cell key={i} fill={COLORS[i]} />)}
                  </Pie>
                  <Tooltip {...TS} />
                </PieChart>
              )}
            </ResponsiveContainer>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ─── DOWNLOAD REPORTS ──────────────────────────────────────────────────── */
/* BUG FIX: added missing closing brace for the component                   */
function DownloadReports({ showNotif }) {
  const [selected, setSelected] = useState([]);
  const [loading, setLoading] = useState(null);

  const reports = [
    { id: 1, title: "Annual Population Report 2024", category: "Population", size: "4.2 MB", rows: "14,200", icon: "👥", color: "teal" },
    { id: 2, title: "GDP Growth Analysis Q4 2023", category: "Economics", size: "2.1 MB", rows: "8,400", icon: "💰", color: "violet" },
    { id: 3, title: "Healthcare Infrastructure Survey", category: "Health", size: "6.8 MB", rows: "22,100", icon: "🏥", color: "coral" },
    { id: 4, title: "National Literacy Rate Survey", category: "Education", size: "1.9 MB", rows: "5,600", icon: "🎓", color: "amber" },
    { id: 5, title: "Air Quality Index Report 2023", category: "Environment", size: "3.5 MB", rows: "31,500", icon: "🌿", color: "emerald" },
    { id: 6, title: "Urban Employment Statistics", category: "Economics", size: "2.8 MB", rows: "18,200", icon: "💼", color: "teal" },
    { id: 7, title: "Rural Development Index", category: "Population", size: "1.4 MB", rows: "9,800", icon: "🏡", color: "violet" },
    { id: 8, title: "Renewable Energy Adoption", category: "Environment", size: "2.2 MB", rows: "7,400", icon: "⚡", color: "amber" },
  ];

  const toggleSel = (id) =>
    setSelected(p => p.includes(id) ? p.filter(x => x !== id) : [...p, id]);

  const handleDl = (id, title, fmt) => {
    setLoading(id);

    setTimeout(() => {
      // 1. Generate sample data based on format
      let content = "";
      let type = "";
      let ext = fmt.toLowerCase();

      if (fmt === "CSV") {
        const report = reports.find(r => r.id === id);
        content = `Report Title,Category,Rows,Size,Date\n${title},${report.category},${report.rows},${report.size},${new Date().toLocaleDateString()}\n`;
        // Add some dummy data rows
        for (let i = 1; i <= 5; i++) {
          content += `Record ${i},Data ${Math.random().toFixed(2)},Sample Value,Active,2024-01-01\n`;
        }
        type = "text/csv;charset=utf-8;";
      } else {
        // Structured PDF Simulation
        const date = new Date().toLocaleString();
        content = `%PDF-1.4\n1 0 obj\n<< /Title (${title}) /Author (DataPortal) /Subject (Public Data Report) /CreationDate (D:${new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14)}) >>\nendobj\n\n`;
        content += `==================================================\n`;
        content += `            OFFICIAL DATAPORTAL REPORT            \n`;
        content += `==================================================\n\n`;
        content += `REPORT TITLE: ${title}\n`;
        content += `CATEGORY:     ${reports.find(r => r.id === id).category}\n`;
        content += `EXTRACTED ON: ${date}\n\n`;
        content += `--------------------------------------------------\n`;
        content += `[DATA PREVIEW]\n`;
        content += `Row | Attribute | Value\n`;
        for (let i = 1; i <= 10; i++) {
          content += `${String(i).padStart(3, ' ')} | Sample ID | ${Math.random().toString(36).substring(7).toUpperCase()}\n`;
        }
        content += `--------------------------------------------------\n\n`;
        content += `[END OF REPORT]`;
        type = "application/pdf";
      }

      // 2. Trigger Download
      const blob = new Blob([content], { type });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.setAttribute("href", url);
      link.setAttribute("download", `${title.replace(/\s+/g, '_')}_${new Date().getTime()}.${ext}`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      setLoading(null);
      showNotif(`${title} downloaded as ${fmt} ✓`);
    }, 1200);
  };

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: "28px" }}>
      <div className="fade-up" style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 24 }}>
        <div>
          <div style={{ fontFamily: "Outfit", fontWeight: 900, fontSize: 28, marginBottom: 4 }}>Download Reports</div>
          <div style={{ color: "var(--muted)", fontSize: 14 }}>Download public datasets as CSV or PDF</div>
        </div>
        {selected.length > 0 && (
          <button className="btn btn-primary"
            onClick={() => {
              selected.forEach(id => {
                const report = reports.find(r => r.id === id);
                if (report) handleDl(id, report.title, "CSV");
              });
              showNotif(`${selected.length} reports in queue ✓`);
              setSelected([]);
            }}>
            <I d={ic.dl} size={15} /> Download Selected ({selected.length})
          </button>
        )}
      </div>

      {/* FORMAT CARDS */}
      <div className="charts-row two fade-up-1" style={{ marginBottom: 24 }}>
        {[
          { fmt: "CSV", icon: "📄", color: "#10b981", bg: "#d1fae5", desc: "Machine-readable comma-separated values. Ideal for data analysis with Excel, Python, R." },
          { fmt: "PDF", icon: "📕", color: "#f43f5e", bg: "#ffe4e6", desc: "Formatted report with charts and insights. Perfect for presentations and sharing." },
        ].map(({ fmt, icon, color, bg, desc }) => (
          <div key={fmt} className="card" style={{ borderColor: color + "44", background: `linear-gradient(135deg,${bg},#fff)` }}>
            <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 10 }}>
              <div style={{ width: 48, height: 48, borderRadius: 12, background: color + "22", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 24 }}>{icon}</div>
              <div>
                <div style={{ fontFamily: "Outfit", fontWeight: 800, fontSize: 18, color }}>{fmt} Format</div>
                <div style={{ fontSize: 12, color: "var(--muted)" }}>Available for all datasets</div>
              </div>
            </div>
            <p style={{ fontSize: 13, color: "var(--muted)", lineHeight: 1.6 }}>{desc}</p>
          </div>
        ))}
      </div>

      <div className="fade-up-2" style={{ fontFamily: "Outfit", fontWeight: 700, fontSize: 15, marginBottom: 14 }}>📂 Available Reports</div>

      {/* REPORTS LIST */}
      {/* BUG FIX: removed stray quote in gap:8" → gap:8 */}
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }} className="fade-up-3">
        {reports.map(r => (
          <div key={r.id} className="dl-card"
            style={{
              borderColor: selected.includes(r.id) ? "var(--teal)" : "var(--border)",
              background: selected.includes(r.id) ? "#e0f2fe11" : "var(--card)",
            }}>

            <div className={`cb-box${selected.includes(r.id) ? " checked" : ""}`}
              style={{ cursor: "pointer" }} onClick={() => toggleSel(r.id)}>
              {selected.includes(r.id) && <I d={ic.check} size={10} c="#fff" sw={3} />}
            </div>

            <div className="dl-icon" style={{ background: `var(--${r.color})22` }}>{r.icon}</div>

            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 2, fontFamily: "Outfit" }}>{r.title}</div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                <span className={`badge badge-${r.color}`}>{r.category}</span>
                <span style={{ fontSize: 12, color: "var(--muted)" }}>{r.rows} rows</span>
                <span style={{ fontSize: 12, color: "var(--muted)" }}>•</span>
                <span style={{ fontSize: 12, color: "var(--muted)" }}>{r.size}</span>
              </div>
            </div>

            {/* BUG FIX: removed stray quote in gap:8" */}
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <button className="btn btn-outline btn-sm"
                style={{ borderColor: "#10b981", color: "#10b981" }}
                onClick={() => handleDl(r.id, r.title, "CSV")}
                disabled={loading === r.id}>
                <I d={loading === r.id ? ic.refresh : ic.csv} size={12} c="#10b981" /> CSV
              </button>
              <button className="btn btn-outline btn-sm"
                style={{ borderColor: "#f43f5e", color: "#f43f5e" }}
                onClick={() => handleDl(r.id, r.title, "PDF")}
                disabled={loading === r.id}>
                <I d={loading === r.id ? ic.refresh : ic.pdf} size={12} c="#f43f5e" /> PDF
              </button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}   /* ← BUG FIX: this closing brace was MISSING in the original code */

/* ─── ADMIN PANEL ───────────────────────────────────────────────────────── */
function AdminPanel({ showNotif, onDataUpload, uploads, setUploads }) {
  const [drag, setDrag] = useState(false);
  const [progress, setProgress] = useState(null);
  const [form, setForm] = useState({ name: "", category: "Population", desc: "" });
  const [uploadedLink, setUploadedLink] = useState(null);
  const [selectedFile, setSelectedFile] = useState(null);
  const fileRef = useRef();

  const simulateUpload = () => {
    if (!form.name) { showNotif("Please enter a dataset name"); return; }
    if (!selectedFile) { showNotif("Please select an Excel file first"); return; }

    setProgress(0);
    const iv = setInterval(() => {
      setProgress(p => {
        const next = p + Math.random() * 15;
        if (next >= 100) {
          clearInterval(iv);

          // Parse File
          const reader = new FileReader();
          reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(worksheet);

            // Transform data for charts
            // Expecting Excel with columns like: Year, Mumbai, Delhi, Bangalore, etc. or Name, Value for Pie
            const yearData = json.filter(r => r.Year || r.year).map(r => ({
              year: String(r.Year || r.year),
              Mumbai: r.Mumbai || 0,
              Delhi: r.Delhi || 0,
              Bangalore: r.Bangalore || 0,
              Chennai: r.Chennai || 0,
              Hyderabad: r.Hyderabad || 0,
              Pune: r.Pune || 0,
            }));

            const pieData = json.filter(r => r.Name || r.name).map(r => ({
              name: r.Name || r.name,
              value: r.Value || r.value || 0
            }));

            const barData = json.filter(r => r.City || r.city || r.Location || r.location).map(r => ({
              city: r.City || r.city || r.Location || r.location,
              value: r.Value || r.value || r.Actual || r.actual || 0,
              target: r.Target || r.target || r.Goal || r.goal || 0
            }));

            const chartData = {
              yearData: yearData.length ? yearData : null,
              pieData: pieData.length ? pieData : null,
              barData: barData.length ? barData : null
            };

            onDataUpload(chartData);

            setTimeout(() => {
              setProgress(null);
              setUploads(u => [{
                id: u.length + 1,
                name: form.name,
                category: form.category,
                size: (selectedFile.size / 1024 / 1024).toFixed(2) + " MB",
                rows: json.length,
                status: "Active",
                date: new Date().toISOString().slice(0, 10),
                data: chartData // Store data for later retrieval
              }, ...u]);
              setForm({ name: "", category: "Population", desc: "" });
              setSelectedFile(null);
              showNotif("Dataset uploaded and visualized successfully ✓");
              setUploadedLink("https://colab.research.google.com/drive/1u2TMvvXzQdoVM5wSYogQLx8LnZsFJXrC#scrollTo=RlZhomtKvxZn");
            }, 400);
          };
          reader.readAsArrayBuffer(selectedFile);

          return 100;
        }
        return next;
      });
    }, 120);
  };

  const loadSampleData = () => {
    const sampleYearData = [
      { year: "2020", Mumbai: 100, Delhi: 110, Bangalore: 120, Chennai: 130, Hyderabad: 140, Pune: 150 },
      { year: "2021", Mumbai: 110, Delhi: 120, Bangalore: 130, Chennai: 140, Hyderabad: 150, Pune: 160 },
      { year: "2022", Mumbai: 120, Delhi: 130, Bangalore: 140, Chennai: 150, Hyderabad: 160, Pune: 170 },
      { year: "2023", Mumbai: 130, Delhi: 140, Bangalore: 150, Chennai: 160, Hyderabad: 170, Pune: 180 },
    ];
    const samplePieData = [
      { name: "Urban Core", value: 45 },
      { name: "Suburbia", value: 30 },
      { name: "Rural Area", value: 15 },
      { name: "Industrial", value: 10 },
    ];
    const sampleBarData = [
      { city: "Mumbai", value: 500, target: 450 },
      { city: "Delhi", value: 550, target: 500 },
      { city: "Bangalore", value: 400, target: 420 },
      { city: "Chennai", value: 300, target: 350 },
    ];

    onDataUpload({ yearData: sampleYearData, pieData: samplePieData, barData: sampleBarData });

    // Also add to the table so it's "stored"
    setUploads(u => [{
      id: u.length + 1,
      name: "Sample Test Dataset",
      category: "Tests",
      size: "0.1 MB",
      rows: sampleYearData.length,
      status: "Active",
      date: new Date().toISOString().slice(0, 10),
      data: { yearData: sampleYearData, pieData: samplePieData, barData: sampleBarData }
    }, ...u]);

    showNotif("Sample data loaded and stored! ✨");
    setUploadedLink("https://colab.research.google.com/drive/1u2TMvvXzQdoVM5wSYogQLx8LnZsFJXrC#scrollTo=RlZhomtKvxZn");
  };

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (file) {
      setSelectedFile(file);
      if (!form.name) setForm(f => ({ ...f, name: file.name.split('.')[0] }));
    }
  };

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: "28px" }}>
      <div className="fade-up" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
        <div>
          <div style={{ fontFamily: "Outfit", fontWeight: 900, fontSize: 28 }}>Admin Panel</div>
          <div style={{ color: "var(--muted)", fontSize: 14, marginTop: 2 }}>Upload and manage public datasets</div>
        </div>
        <span className="badge badge-emerald" style={{ fontSize: 12, padding: "5px 12px", display: "flex", alignItems: "center", gap: 6 }}>
          <span className="live-dot" /> System Online
        </span>
      </div>

      {/* ADMIN STATS */}
      <div className="stat-grid fade-up-1" style={{ marginBottom: 24 }}>
        {[
          { label: "Total Datasets", val: uploads.length, color: "teal" },
          { label: "Active", val: uploads.filter(u => u.status === "Active").length, color: "emerald" },
          { label: "Processing", val: uploads.filter(u => u.status === "Processing").length, color: "amber" },
          { label: "Total Records", val: "94.2K", color: "violet" },
        ].map(s => (
          <div key={s.label} className={`stat-card ${s.color}`}>
            <div className="stat-label">{s.label}</div>
            <div className={`stat-val ${s.color}`}>{s.val}</div>
          </div>
        ))}
      </div>

      <div className="charts-row two fade-up-2" style={{ marginBottom: 24, alignItems: "start" }}>
        {/* UPLOAD FORM */}
        <div className="card">
          <div className="card-title" style={{ fontSize: 16, marginBottom: 16 }}>⬆ Upload New Dataset</div>

          <div className={`upload-zone${drag ? " drag" : ""}`}
            onDragOver={e => { e.preventDefault(); setDrag(true); }}
            onDragLeave={() => setDrag(false)}
            onDrop={e => { e.preventDefault(); setDrag(false); const file = e.dataTransfer.files[0]; if (file) setSelectedFile(file); showNotif("File received — fill form to upload"); }}
            onClick={() => fileRef.current.click()}>
            <input ref={fileRef} type="file" style={{ display: "none" }} accept=".csv,.pdf,.xlsx" onChange={handleFileChange} />
            <div className="upload-icon">
              <I d={ic.upload} size={28} c="#0ea5e9" />
            </div>
            <div style={{ fontFamily: "Outfit", fontWeight: 700, fontSize: 15, marginBottom: 6 }}>
              {selectedFile ? `Selected: ${selectedFile.name}` : "Drop files here or click to browse"}
            </div>
            <div style={{ color: "var(--muted)", fontSize: 13 }}>Supports CSV, Excel, PDF — max 50 MB</div>
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 12, marginTop: 16 }}>
            <div>
              <div className="filter-label" style={{ marginBottom: 6 }}>Dataset Name *</div>
              <input value={form.name}
                onChange={e => setForm(f => ({ ...f, name: e.target.value }))}
                placeholder="e.g. Census 2024 Urban Areas"
                style={{ width: "100%", padding: "9px 12px", border: "1px solid var(--border)", borderRadius: 10, fontFamily: "DM Sans", fontSize: 13, outline: "none", background: "var(--bg)" }} />
            </div>
            <div>
              <div className="filter-label" style={{ marginBottom: 6 }}>Category</div>
              <select value={form.category} onChange={e => setForm(f => ({ ...f, category: e.target.value }))}>
                {["Population", "Economics", "Health", "Education", "Environment", "Infrastructure"].map(c => (
                  <option key={c}>{c}</option>
                ))}
              </select>
            </div>
            <div>
              <div className="filter-label" style={{ marginBottom: 6 }}>Description</div>
              <textarea value={form.desc}
                onChange={e => setForm(f => ({ ...f, desc: e.target.value }))}
                placeholder="Brief description of this dataset..."
                rows={3}
                style={{ width: "100%", padding: "9px 12px", border: "1px solid var(--border)", borderRadius: 10, fontFamily: "DM Sans", fontSize: 13, resize: "none", outline: "none", background: "var(--bg)" }} />
            </div>

            {progress !== null && (
              <div>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 4 }}>
                  <span>Uploading…</span><span>{Math.round(progress)}%</span>
                </div>
                <div className="progress-bar">
                  <div className="progress-fill" style={{ width: `${progress}%` }} />
                </div>
              </div>
            )}

            <button className="btn btn-primary" style={{ justifyContent: "center" }}
              onClick={simulateUpload} disabled={progress !== null}>
              {progress !== null
                ? <><I d={ic.refresh} size={14} c="#fff" className="spin" /> Uploading…</>
                : <><I d={ic.upload} size={14} c="#fff" /> Upload Dataset</>}
            </button>

            <button className="btn btn-outline" style={{ marginTop: 12, width: "100%", justifyContent: "center" }}
              onClick={loadSampleData}>
              <I d={ic.lightning} size={14} c="var(--teal)" /> Load Sample Data (Tests)
            </button>

            <button className="btn btn-outline" style={{ marginTop: 8, width: "100%", justifyContent: "center", borderColor: "var(--coral)", color: "var(--coral)" }}
              onClick={() => { onDataUpload(null); showNotif("Dashboard reset to default data"); }}>
              <I d={ic.refresh} size={14} /> Reset Dashboard Data
            </button>

            {uploadedLink && (
              <div style={{ marginTop: 16, padding: 16, borderRadius: 16, background: "linear-gradient(135deg, #f0fdf4, #f8fafc)", border: "1px solid #bbf7d0", animation: "fade-in 0.5s ease-out" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                  <div style={{ width: 24, height: 24, borderRadius: "50%", background: "#22c55e", display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <I d={ic.check} size={12} c="#fff" sw={3} />
                  </div>
                  <div style={{ fontWeight: 800, color: "#166534", fontSize: 14, fontFamily: "Outfit" }}>Upload Complete!</div>
                </div>
                <div style={{ fontSize: 13, color: "#475569", lineHeight: 1.5, marginBottom: 12 }}>
                  Your dataset has been successfully processed. You can now view the analysis in the linked Google Colab notebook.
                </div>
                <a href={uploadedLink} target="_blank" rel="noopener noreferrer"
                  className="btn btn-primary" style={{ width: "100%", justifyContent: "center", background: "#1e3a8a", border: "none" }}>
                  <I d={ic.globe} size={14} /> Open Colab Notebook
                </a>
              </div>
            )}
          </div>
        </div>

        {/* QUICK ACTIONS + SYSTEM HEALTH */}
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <div className="card">
            <div className="card-title" style={{ marginBottom: 12 }}>⚡ Quick Actions</div>
            {[
              { label: "Generate Summary Report", icon: ic.chart, color: "btn-primary" },
              { label: "Export All Datasets", icon: ic.dl, color: "btn-emerald" },
              { label: "Refresh Data Sources", icon: ic.refresh, color: "btn-outline" },
              { label: "View System Logs", icon: ic.table, color: "btn-outline" },
            ].map(({ label, icon, color }) => (
              <button key={label} className={`btn ${color}`}
                style={{ width: "100%", justifyContent: "flex-start", marginBottom: 8 }}
                onClick={() => showNotif(`${label} triggered ✓`)}>
                <I d={icon} size={14} /> {label}
              </button>
            ))}
          </div>

          <div className="card">
            <div className="card-title" style={{ marginBottom: 10 }}>📊 System Health</div>
            {[
              ["API Status", "Online", "badge-emerald"],
              ["DB Connection", "Healthy", "badge-emerald"],
              ["Storage", "62% used", "badge-amber"],
              ["Last Backup", "2h ago", "badge-teal"],
            ].map(([k, v, b]) => (
              <div key={k} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 0", borderBottom: "1px solid var(--border)" }}>
                <span style={{ fontSize: 13, color: "var(--muted)" }}>{k}</span>
                <span className={`badge ${b}`}>{v}</span>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* DATASETS TABLE */}
      <div className="card fade-up-3">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <div className="card-title" style={{ fontSize: 16 }}>📂 Manage Datasets</div>
          <div className="search-wrap">
            <span className="search-icon"><I d={ic.search} size={14} /></span>
            <input placeholder="Search datasets…" />
          </div>
        </div>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Name</th><th>Category</th><th>Records</th><th>Size</th><th>Status</th><th>Date</th><th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {uploads.map(u => (
                <tr key={u.id}>
                  <td style={{ fontWeight: 600, fontFamily: "Outfit" }}>{u.name}</td>
                  <td><span className="badge badge-teal">{u.category}</span></td>
                  <td className="td-mono">{u.rows.toLocaleString()}</td>
                  <td className="td-mono">{u.size}</td>
                  <td>
                    <span className={`badge ${u.status === "Active" ? "badge-emerald" : "badge-amber"}`}
                      style={{ display: "inline-flex", alignItems: "center", gap: 5 }}>
                      {u.status === "Processing" &&
                        <span className="live-dot" style={{ background: "var(--amber)" }} />}
                      {u.status}
                    </span>
                  </td>
                  <td style={{ color: "var(--muted)", fontSize: 12 }}>{u.date}</td>
                  <td>
                    <div style={{ display: "flex", gap: 6 }}>
                      <button className="btn btn-outline btn-sm"
                        onClick={() => {
                          if (u.data) {
                            onDataUpload(u.data);
                            showNotif(`${u.name} loaded into dashboard`);
                          } else {
                            showNotif(`Viewing details for ${u.name}`);
                          }
                        }}>
                        <I d={ic.eye} size={12} />
                      </button>
                      <button className="btn btn-outline btn-sm"
                        style={{ borderColor: "#f43f5e44", color: "#f43f5e" }}
                        onClick={() => { setUploads(p => p.filter(x => x.id !== u.id)); showNotif(`${u.name} deleted`); }}>
                        <I d={ic.trash} size={12} />
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ─── AI CHATBOT ────────────────────────────────────────────────────────── */
function Chartbot() {
  const [messages, setMessages] = useState([]);
  const [allChats, setAllChats] = useState([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const messagesEndRef = useRef(null);

  const API_KEY = "AIzaSyBk7paPnpP523j2LKgRegRZFCCMOKaHmrM";
  const API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent";

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages, loading]);

  const newChat = () => {
    if (messages.length > 0) {
      saveCurrentChat();
    }
    setMessages([]);
    setInput("");
  };

  const saveCurrentChat = () => {
    if (messages.length > 0) {
      const title = messages[0].content.substring(0, 30) + "...";
      if (!allChats.some(c => c.title === title)) {
        setAllChats(prev => [{ title, messages: [...messages] }, ...prev]);
      }
    }
  };

  const loadChat = (index) => {
    setMessages(allChats[index].messages);
  };

  const sendMessage = async () => {
    const text = input.trim();
    if (!text || loading) return;

    const userMsg = { role: "user", content: text };
    const currentMessages = [...messages, userMsg];
    setMessages(currentMessages);
    setInput("");
    setLoading(true);

    try {
      // Role mapping and consolidation for Gemini
      const historyItems = currentMessages
        .filter(m => m.role && m.content && !m.isError)
        .map(m => ({
          role: m.role === "assistant" ? "model" : "user",
          parts: [{ text: m.content }]
        }));

      const finalHistory = [];
      historyItems.forEach(item => {
        if (finalHistory.length > 0 && finalHistory[finalHistory.length - 1].role === item.role) {
          finalHistory[finalHistory.length - 1].parts[0].text += "\n" + item.parts[0].text;
        } else {
          finalHistory.push(item);
        }
      });

      console.log("Sending to AI...", finalHistory);

      const res = await axios.post(`${API_URL}?key=${API_KEY}`,
        { contents: finalHistory },
        { timeout: 30000 }
      );

      const aiText = res.data?.candidates?.[0]?.content?.parts?.[0]?.text;
      if (!aiText) throw new Error("Empty response from AI (possibly blocked)");

      setMessages(prev => [...prev, { role: "assistant", content: aiText }]);
    } catch (err) {
      console.error("Gemini Details:", err.response?.data || err.message);
      const msg = err.response?.data?.error?.message || err.message;
      setMessages(prev => [...prev, { role: "assistant", content: `API Error: ${msg}`, isError: true }]);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (messages.length > 0 && messages.length % 2 === 0) {
      const title = messages[0].content.substring(0, 30);
      setAllChats(prev => {
        const existingIdx = prev.findIndex(c => c.title.startsWith(title));
        if (existingIdx !== -1) {
          const updated = [...prev];
          updated[existingIdx] = { ...updated[existingIdx], messages: [...messages] };
          return updated;
        } else {
          return [{ title: title + "...", messages: [...messages] }, ...prev];
        }
      });
    }
  }, [messages]);

  return (
    <div style={{ flex: 1, display: 'flex', height: '100%', overflow: 'hidden' }}>
      <div style={{ width: '260px', borderRight: '1px solid var(--border)', background: 'var(--card)', display: 'flex', flexDirection: 'column' }}>
        <div style={{ padding: '20px', borderBottom: '1px solid var(--border)' }}>
          <button className="btn btn-primary" style={{ width: '100%', justifyContent: 'center', borderRadius: '10px' }} onClick={newChat}>
            + New Chat
          </button>
        </div>
        <div style={{ flex: 1, overflowY: 'auto', padding: '10px' }}>
          <div style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', padding: '10px', textTransform: 'uppercase' }}>Recent History</div>
          {allChats.length === 0 && <div style={{ padding: '10px', fontSize: '13px', color: 'var(--muted)', textAlign: 'center' }}>No history yet</div>}
          {allChats.map((chat, idx) => (
            <div key={idx} onClick={() => loadChat(idx)} className="sb-item" style={{ margin: '2px 0', borderRadius: '8px' }}>
              💬 {chat.title}
            </div>
          ))}
        </div>
      </div>

      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--bg)' }}>
        <div style={{ padding: '16px 28px', background: 'var(--card)', borderBottom: '1px solid var(--border)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <div style={{ fontFamily: "Outfit", fontWeight: 800, fontSize: '18px' }}>AI Assistant</div>
            <div style={{ fontSize: '12px', color: 'var(--muted)' }}>Powered by Gemini 1.5 Flash</div>
          </div>
        </div>

        <div style={{ flex: 1, overflowY: 'auto', padding: '28px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
          {messages.length === 0 && (
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: 'var(--muted)', textAlign: 'center' }}>
              <div style={{ fontSize: '48px', marginBottom: '20px' }}>🤖</div>
              <h1 style={{ fontFamily: "Outfit", fontWeight: 900, color: 'var(--text)', marginBottom: '8px' }}>Chat AI</h1>
              <p>Start a conversation with our AI assistant</p>
            </div>
          )}

          {messages.map((msg, idx) => (
            <div key={idx} style={{ display: 'flex', gap: '12px', flexDirection: msg.role === 'user' ? 'row-reverse' : 'row', alignItems: 'flex-start' }}>
              <div style={{ width: '32px', height: '32px', borderRadius: '50%', background: msg.role === 'user' ? 'var(--teal)' : 'var(--violet)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '16px', flexShrink: 0, boxShadow: '0 2px 8px #0001' }}>
                {msg.role === 'user' ? '👤' : '🤖'}
              </div>
              <div style={{ maxWidth: '70%', padding: '14px 18px', borderRadius: '16px', fontSize: '14px', lineHeight: 1.6, background: msg.role === 'user' ? 'linear-gradient(135deg, var(--teal), var(--violet))' : 'var(--card)', color: msg.isError ? '#c33' : (msg.role === 'user' ? '#fff' : 'var(--text)'), border: msg.role === 'assistant' ? '1px solid var(--border)' : 'none', boxShadow: '0 2px 12px #00000008', borderTopRightRadius: msg.role === 'user' ? '4px' : '16px', borderTopLeftRadius: msg.role === 'assistant' ? '4px' : '16px', backgroundColor: msg.isError ? '#fee' : undefined }}>
                <div style={{ whiteSpace: 'pre-wrap' }}>{msg.content}</div>
              </div>
            </div>
          ))}

          {loading && (
            <div style={{ display: 'flex', gap: '12px', alignItems: 'flex-start' }}>
              <div style={{ width: '32px', height: '32px', borderRadius: '50%', background: 'var(--violet)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '16px' }}>🤖</div>
              <div style={{ padding: '14px 18px', background: 'var(--card)', borderRadius: '16px', borderTopLeftRadius: '4px', border: '1px solid var(--border)', display: 'flex', gap: '4px' }}>
                <div className="dot-bounce"></div><div className="dot-bounce" style={{ animationDelay: '0.2s' }}></div><div className="dot-bounce" style={{ animationDelay: '0.4s' }}></div>
              </div>
            </div>
          )}
          <div ref={messagesEndRef} />
        </div>

        <div style={{ padding: '20px 28px', background: 'var(--card)', borderTop: '1px solid var(--border)' }}>
          <div style={{ display: 'flex', gap: '12px', background: 'var(--bg)', padding: '8px 12px', borderRadius: '16px', border: '1px solid var(--border)' }}>
            <textarea rows={1} value={input} onChange={(e) => setInput(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendMessage(); } }} placeholder="Ask anything..." style={{ flex: 1, background: 'transparent', border: 'none', outline: 'none', padding: '8px', fontFamily: 'DM Sans', fontSize: '14px', resize: 'none' }} />
            <button className="btn btn-primary" style={{ padding: '0 16px', borderRadius: '12px' }} onClick={sendMessage} disabled={loading || !input.trim()}>Send</button>
          </div>
        </div>
      </div>
      <style>{`
        .dot-bounce { width: 6px; height: 6px; background: var(--muted); borderRadius: 50%; animation: bounce 1.4s infinite ease-in-out; }
        @keyframes bounce { 0%, 80%, 100% { transform: scale(0); } 40% { transform: scale(1.0); } }
      `}</style>
    </div>
  );
}

/* ─── ROOT APP ──────────────────────────────────────────────────────────── */
export default function App() {
  const [page, setPage] = useState("dashboard");
  const [collapsed, setCollapsed] = useState(false);
  const [showFilter, setShowFilter] = useState(true);
  const [notif, showNotif] = useNotif();
  const [user, setUser] = useState({ name: "Admin User", email: "admin@dataportal.gov" });
  const [customData, setCustomData] = useState(null);
  const [uploads, setUploads] = useState(INITIAL_DATASETS);

  useEffect(() => {
    const savedUser = localStorage.getItem("user");
    if (savedUser) {
      const parsedUser = JSON.parse(savedUser);
      setUser(prev => {
        if (JSON.stringify(prev) !== JSON.stringify(parsedUser)) {
          return parsedUser;
        }
        return prev;
      });
    }
  }, []);

  const [filters, setFilters] = useState({
    yearFrom: "2020", yearTo: "2026",
    region: "All India",
    categories: ["population", "gdp", "health", "education", "environment"],
  });
  const [selectedCities, setSelectedCities] = useState(["Mumbai", "Delhi", "Bangalore"]);

  const nav = [
    { key: "dashboard", icon: ic.home, label: "Citizen Dashboard", badge: null },
    { key: "charts", icon: ic.chart, label: "Interactive Charts", badge: null },
    { key: "download", icon: ic.dl, label: "Download Reports", badge: "New" },
    { key: "admin", icon: ic.admin, label: "Admin Panel", badge: null },
    { key: "chatbot", icon: ic.bell, label: "Chatbot", badge: null },
  ];

  return (
    <>
      <style>{`
        ${CSS}
        @keyframes spin { from {transform:rotate(0deg)} to {transform:rotate(360deg)} }
        .spin { animation: spin 1s linear infinite; }
      `}</style>
      <div className="shell">

        {/* ── SIDEBAR ── */}
        <div className={`sidebar${collapsed ? " collapsed" : ""}`}>
          <div className="sb-logo">
            <div className="sb-logo-icon">
              <I d={ic.globe} size={18} c="#fff" />
            </div>
            {!collapsed && <div className="sb-logo-text">Data<span>Portal</span></div>}
          </div>

          {!collapsed && <div className="sb-section">Navigation</div>}
          {nav.map(n => (
            <div key={n.key} className={`sb-item${page === n.key ? " active" : ""}`}
              onClick={() => setPage(n.key)}>
              <I d={n.icon} size={18} />
              {!collapsed && (
                <>
                  <span style={{ flex: 1 }}>{n.label}</span>
                  {n.badge && <span className="sb-badge">{n.badge}</span>}
                </>
              )}
            </div>
          ))}

          {!collapsed && (
            <>
              <div className="sb-divider" />
              <div className="sb-section"></div>
              <div style={{ margin: "4px 12px", padding: "12px", borderRadius: 12, background: "linear-gradient(135deg,#e0f2fe,#ede9fe)", border: "1px solid #bae6fd" }}>
                <div style={{ fontFamily: "Outfit", fontWeight: 800, fontSize: 13, color: "var(--teal)" }}></div>
                <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 4, lineHeight: 1.5 }}>  </div>
                <div style={{ marginTop: 8, display: "flex", gap: 4 }}>
                  <span className="badge badge-teal" style={{ fontSize: 10 }}></span>
                  <span className="badge badge-violet" style={{ fontSize: 10 }}></span>
                </div>
              </div>
            </>
          )}

          {/* BUG FIX: added display:flex to sb-bottom via CSS */}
          <div className="sb-bottom">
            <div style={{ width: 32, height: 32, borderRadius: 10, background: "linear-gradient(135deg,var(--teal),var(--violet))", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
              <I d={ic.user} size={16} c="#fff" />
            </div>
            {!collapsed && (
              <div>
                <div style={{ fontSize: 13, fontWeight: 600 }}>{user.name}</div>
                <div style={{ fontSize: 11, color: "var(--muted)" }}>{user.email}</div>
              </div>
            )}
          </div>
        </div>

        {/* ── MAIN ── */}
        <div className="main">
          {/* TOPBAR */}
          <div className="topbar">
            <button className="btn btn-outline" style={{ padding: "7px 10px" }}
              onClick={() => setCollapsed(c => !c)}>
              <I d={ic.menu} size={18} />
            </button>
            <div className="topbar-title">
              <span>DataPortal</span>
            </div>
            <div className="search-wrap">
              <span className="search-icon"><I d={ic.search} size={14} /></span>
              <input placeholder="Search datasets, cities, reports…" />
            </div>
            <div style={{ marginLeft: "auto", display: "flex", gap: 10, alignItems: "center" }}>
              {page === "dashboard" && (
                <button className={`btn ${showFilter ? "btn-primary" : "btn-outline"} btn-sm`}
                  onClick={() => setShowFilter(f => !f)}>
                  <I d={ic.filter} size={13} /> Filters
                </button>
              )}
              <button className="btn btn-outline" style={{ padding: "7px 10px" }}
                onClick={() => showNotif("No new notifications")}>
                <I d={ic.bell} size={16} />
              </button>
              <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 10px", borderRadius: 10, border: "1px solid var(--border)", cursor: "pointer" }}>
                <div style={{ width: 28, height: 28, borderRadius: 8, background: "linear-gradient(135deg,var(--teal),var(--violet))", display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <I d={ic.user} size={14} c="#fff" />
                </div>
                <div style={{ fontSize: 13, fontWeight: 600 }}>{user.name}</div>
              </div>
            </div>
          </div>

          {/* PAGE CONTENT */}
          <div style={{ display: "flex", flex: 1, minHeight: 0 }}>
            {page === "dashboard" && (
              <>
                <CitizenDashboard
                  key="dash"
                  filters={filters} setFilters={setFilters}
                  selectedCities={selectedCities} setSelectedCities={setSelectedCities}
                  showNotif={showNotif} customData={customData} />
                {showFilter && (
                  <FilterPanel
                    filters={filters} setFilters={setFilters}
                    selectedCities={selectedCities} setSelectedCities={setSelectedCities} />
                )}
              </>
            )}
            {page === "charts" && <ChartsExplorer key="charts" showNotif={showNotif} />}
            {page === "download" && <DownloadReports key="dl" showNotif={showNotif} />}
            {page === "admin" && (
              <AdminPanel
                key="admin"
                showNotif={showNotif}
                onDataUpload={setCustomData}
                uploads={uploads}
                setUploads={setUploads} />
            )}
            {page === "chatbot" && <Chartbot />}
          </div>
        </div>
      </div>

      {notif && (
        <div className="notif">
          <I d={ic.check} size={16} c="#4ade80" /> {notif}
        </div>
      )}
    </>
  );
}