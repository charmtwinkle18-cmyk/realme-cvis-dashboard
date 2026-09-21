"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer,
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  BarChart,
  Bar,
  Legend,
} from "recharts";

const NAV_ITEMS = [
  "Overview",
  "Push Model Performance",
  "Dealer Performance",
  "PS Sales Review",
  "Promoter Score",
  "Promoter Productivity",
  "Daily Zero Sellout",
  "ASM Incentives",
];

const EMPTY_FILTERS = {
  month: "ALL",
  area: "ALL",
  subregion: "ALL",
  customer: "ALL",
  channel: "ALL",
  model: "ALL",
  priceRange: "ALL",
  series: "ALL",
};

function number(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number") {
    const d = XLSX.SSF.parse_date_code(value);
    if (!d) return null;
    return new Date(d.y, d.m - 1, d.d);
  }
  const normalized = String(value).trim().replace(/\//g, "-");
  const direct = new Date(normalized + (/^\d{4}-\d{1,2}-\d{1,2}$/.test(normalized) ? "T00:00:00" : ""));
  return Number.isNaN(direct.getTime()) ? null : direct;
}

function pad(n) { return String(n).padStart(2, "0"); }
function monthKey(d) { return d ? `${d.getFullYear()}-${pad(d.getMonth() + 1)}` : "Unknown"; }
function dateKey(d) { return d ? `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` : "Unknown"; }
function monthLabel(key) {
  if (!/^\d{4}-\d{2}$/.test(key)) return key;
  const [y,m] = key.split("-").map(Number);
  return new Date(y, m - 1, 1).toLocaleDateString("en-PH", { month: "long", year: "numeric" });
}
function shortDate(key) {
  const d = parseDate(key);
  return d ? d.toLocaleDateString("en-PH", { month: "short", day: "numeric" }) : key;
}
function money(v, decimals = 0) {
  return new Intl.NumberFormat("en-PH", { style: "currency", currency: "PHP", minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(number(v));
}
function fmt(v, decimals = 0) {
  return new Intl.NumberFormat("en-PH", { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(number(v));
}
function percent(v) { return Number.isFinite(v) ? `${v >= 0 ? "+" : ""}${v.toFixed(2)}%` : "N/A"; }
function change(current, previous) {
  if (!previous) return null;
  return ((current - previous) / previous) * 100;
}
function unique(values) { return [...new Set(values.filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b))); }

function seriesFromModel(model = "") {
  const m = String(model).toUpperCase();
  if (m.includes("NOTE")) return "Note Series";
  if (/HP\.C\d/.test(m)) return "C Series";
  if (m.includes("GT")) return "GT Series";
  if (/HP\.(8|9|10|11|12|13|14|15|16)/.test(m)) return "Number Series";
  if (m.includes("NARZO")) return "Narzo Series";
  return "Other Series";
}

function priceRange(price) {
  const p = number(price);
  if (!p) return "Unpriced";
  if (p < 7000) return "Below ₱7,000";
  if (p < 10000) return "₱7,000–9,999";
  if (p < 13000) return "₱10,000–12,999";
  if (p < 20000) return "₱13,000–19,999";
  return "₱20,000+";
}

function buildWorkbookData(workbook) {
  const sheet = (name) => workbook.Sheets[name] ? XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: "" }) : [];
  const modelMaster = sheet("Model_Master");
  const modelMap = new Map(modelMaster.map(row => [String(row.Model || "").trim(), { price: number(row.Price), category: row.Category || "" }]));
  const storeMaster = sheet("Store_Master");
  const storeMap = new Map(storeMaster.map(row => [String(row.Store || "").trim(), row]));

  const sales = sheet("Sales_Data").map((row, index) => {
    const d = parseDate(row.Date);
    const model = String(row.Model || "").trim();
    const units = number(row["Units Sold"]);
    const amount = number(row["Sales Amount"]);
    const modelInfo = modelMap.get(model) || {};
    const storeInfo = storeMap.get(String(row.Store || "").trim()) || {};
    const inferredPrice = units ? amount / units : 0;
    const unitPrice = modelInfo.price || inferredPrice;
    return {
      id: index,
      date: d,
      dateKey: dateKey(d),
      month: monthKey(d),
      store: row.Store || "Unknown store",
      area: row.Area || storeInfo.Area || "Unknown",
      subregion: row.Subregion || storeInfo.Subregion || "Unknown",
      customer: row.Customer || "Unknown customer",
      channel: row.Channel || row["Customer Type"] || storeInfo["Store Type"] || "Unknown",
      model,
      units,
      amount,
      unitPrice,
      priceRange: priceRange(unitPrice),
      series: seriesFromModel(model),
    };
  }).filter(r => r.date && r.units >= 0);

  const targets = sheet("Sales_Target");
  const inventory = sheet("Inventory");
  const promoters = sheet("Promoter_Data");
  const areas = sheet("Area_Master");
  return { sales, targets, inventory, promoters, modelMaster, storeMaster, areas };
}

function groupTrend(rows, mode = "Daily") {
  const bucket = new Map();
  rows.forEach(r => {
    let key = r.dateKey;
    if (mode === "Monthly") key = r.month;
    if (mode === "Weekly") {
      const d = new Date(r.date);
      const day = (d.getDay() + 6) % 7;
      d.setDate(d.getDate() - day);
      key = dateKey(d);
    }
    const item = bucket.get(key) || { key, units: 0, amount: 0 };
    item.units += r.units;
    item.amount += r.amount;
    bucket.set(key, item);
  });
  return [...bucket.values()].sort((a,b)=>a.key.localeCompare(b.key)).map(x => ({
    ...x,
    label: mode === "Monthly" ? monthLabel(x.key).replace(/ \d{4}/, "") : shortDate(x.key),
  }));
}

function dealerTrend(rows, mode = "Daily") {
  const totals = new Map();
  rows.forEach(r => totals.set(r.customer, (totals.get(r.customer) || 0) + r.units));
  const dealers = [...totals.entries()].sort((a,b)=>b[1]-a[1]).slice(0,5).map(x=>x[0]);
  const buckets = new Map();
  rows.forEach(r => {
    if (!dealers.includes(r.customer)) return;
    let key = r.dateKey;
    if (mode === "Monthly") key = r.month;
    if (mode === "Weekly") {
      const d = new Date(r.date); const day = (d.getDay()+6)%7; d.setDate(d.getDate()-day); key = dateKey(d);
    }
    if (!buckets.has(key)) buckets.set(key, { key });
    const entry = buckets.get(key);
    entry[r.customer] = (entry[r.customer] || 0) + r.units;
  });
  const data = [...buckets.values()].sort((a,b)=>a.key.localeCompare(b.key)).map(x => ({
    ...x,
    label: mode === "Monthly" ? monthLabel(x.key).replace(/ \d{4}/, "") : shortDate(x.key),
  }));
  return { data, dealers };
}

function useDashboardData() {
  const [data, setData] = useState({ sales: [], targets: [], inventory: [], promoters: [], modelMaster: [], storeMaster: [], areas: [] });
  const [fileName, setFileName] = useState("default-data.xlsx");
  const [loadedAt, setLoadedAt] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  async function loadBuffer(buffer, name) {
    try {
      const wb = XLSX.read(buffer, { type: "array", cellDates: true });
      const parsed = buildWorkbookData(wb);
      if (!parsed.sales.length) throw new Error("The Sales_Data sheet is missing or has no rows.");
      setData(parsed); setFileName(name); setLoadedAt(new Date()); setError("");
    } catch (e) {
      setError(e?.message || "Could not read this workbook.");
    } finally { setLoading(false); }
  }

  async function loadDefault() {
    setLoading(true);
    try {
      const res = await fetch("/default-data.xlsx", { cache: "no-store" });
      if (!res.ok) throw new Error("Default workbook could not be loaded.");
      await loadBuffer(await res.arrayBuffer(), "default-data.xlsx");
    } catch (e) { setError(e?.message || "Default workbook could not be loaded."); setLoading(false); }
  }

  useEffect(() => { loadDefault(); }, []);
  return { data, fileName, loadedAt, loading, error, loadBuffer, loadDefault };
}

export default function Dashboard() {
  const { data, fileName, loadedAt, loading, error, loadBuffer, loadDefault } = useDashboardData();
  const [activePage, setActivePage] = useState("Overview");
  const [filters, setFilters] = useState(EMPTY_FILTERS);
  const [filterHistory, setFilterHistory] = useState([]);
  const [dark, setDark] = useState(false);
  const [showFilters, setShowFilters] = useState(true);
  const [trendMode, setTrendMode] = useState("Daily");
  const [dealerMode, setDealerMode] = useState("Daily");
  const fileRef = useRef(null);

  const sales = data.sales;
  const months = useMemo(() => unique(sales.map(r=>r.month)).reverse(), [sales]);

  useEffect(() => {
    if (months.length && filters.month === "ALL") {
      setFilters(f => ({ ...f, month: months[0] }));
    }
  }, [months.join("|")]);

  const updateFilter = (key, value) => {
    setFilterHistory(h => [...h.slice(-9), filters]);
    setFilters(f => ({ ...f, [key]: value }));
  };
  const clearFilters = () => { setFilterHistory(h=>[...h.slice(-9), filters]); setFilters({ ...EMPTY_FILTERS, month: months[0] || "ALL" }); };
  const undoFilters = () => {
    if (!filterHistory.length) return;
    const previous = filterHistory[filterHistory.length-1];
    setFilterHistory(h=>h.slice(0,-1)); setFilters(previous);
  };

  const filtered = useMemo(() => sales.filter(r =>
    (filters.month === "ALL" || r.month === filters.month) &&
    (filters.area === "ALL" || r.area === filters.area) &&
    (filters.subregion === "ALL" || r.subregion === filters.subregion) &&
    (filters.customer === "ALL" || r.customer === filters.customer) &&
    (filters.channel === "ALL" || r.channel === filters.channel) &&
    (filters.model === "ALL" || r.model === filters.model) &&
    (filters.priceRange === "ALL" || r.priceRange === filters.priceRange) &&
    (filters.series === "ALL" || r.series === filters.series)
  ), [sales, filters]);

  const optionsBase = useMemo(() => sales.filter(r => filters.month === "ALL" || r.month === filters.month), [sales, filters.month]);
  const options = {
    areas: unique(optionsBase.map(r=>r.area)),
    subregions: unique(optionsBase.map(r=>r.subregion)),
    customers: unique(optionsBase.map(r=>r.customer)),
    channels: unique(optionsBase.map(r=>r.channel)),
    models: unique(optionsBase.map(r=>r.model)),
    priceRanges: unique(optionsBase.map(r=>r.priceRange)),
    series: unique(optionsBase.map(r=>r.series)),
  };

  const totalUnits = filtered.reduce((s,r)=>s+r.units,0);
  const totalAmount = filtered.reduce((s,r)=>s+r.amount,0);
  const asp = totalUnits ? totalAmount / totalUnits : 0;
  const above13 = filtered.filter(r=>r.unitPrice >= 13000).reduce((s,r)=>s+r.units,0);
  const activeStores = new Set(filtered.filter(r=>r.units>0).map(r=>r.store)).size;
  const maxDate = filtered.length ? new Date(Math.max(...filtered.map(r=>r.date.getTime()))) : null;
  const elapsedDays = maxDate ? maxDate.getDate() : 0;
  const dailyRunRate = elapsedDays ? totalUnits / elapsedDays : 0;

  const prevMonth = (() => {
    if (!/^\d{4}-\d{2}$/.test(filters.month)) return null;
    const [y,m]=filters.month.split("-").map(Number); const d=new Date(y,m-2,1); return `${d.getFullYear()}-${pad(d.getMonth()+1)}`;
  })();
  const prevRows = prevMonth ? sales.filter(r => r.month === prevMonth) : [];
  const prevUnits = prevRows.reduce((s,r)=>s+r.units,0);
  const prevAmount = prevRows.reduce((s,r)=>s+r.amount,0);
  const prevAsp = prevUnits ? prevAmount/prevUnits : 0;
  const prevAbove = prevRows.filter(r=>r.unitPrice>=13000).reduce((s,r)=>s+r.units,0);
  const prevRunRate = prevRows.length ? prevUnits / Math.max(...prevRows.map(r=>r.date.getDate())) : 0;

  const targetUnits = useMemo(() => {
    const rows = data.targets.filter(t => String(t.Month || "").slice(0,7) === filters.month);
    if (!rows.length) return 0;
    const matchingStores = new Set(filtered.map(r=>r.store));
    const exact = rows.filter(t => matchingStores.has(t.Store));
    if (exact.length) return exact.reduce((s,t)=>s+number(t["Target Units"]),0);
    if (filters.area !== "ALL") {
      const areaRows = rows.filter(t => t.Area === filters.area);
      return areaRows.reduce((s,t)=>s+number(t["Target Units"]),0);
    }
    return 0;
  }, [data.targets, filtered, filters.month, filters.area]);

  const salesTrend = useMemo(()=>groupTrend(filtered, trendMode), [filtered, trendMode]);
  const dealers = useMemo(()=>dealerTrend(filtered, dealerMode), [filtered, dealerMode]);

  const modelRows = useMemo(() => aggregate(filtered, "model").slice(0,20), [filtered]);
  const dealerRows = useMemo(() => aggregate(filtered, "customer").slice(0,20), [filtered]);
  const subregionRows = useMemo(() => aggregate(filtered, "subregion").slice(0,20), [filtered]);
  const storeRows = useMemo(() => aggregate(filtered, "store").slice(0,30), [filtered]);
  const promoterRows = useMemo(() => [...data.promoters].sort((a,b)=>number(b["Units Sold"])-number(a["Units Sold"])), [data.promoters]);
  const areaRows = useMemo(() => aggregate(filtered, "subregion"), [filtered]);

  async function onFile(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    await loadBuffer(await file.arrayBuffer(), file.name);
    event.target.value = "";
  }

  const currentMonthText = filters.month === "ALL" ? "All uploaded months" : monthLabel(filters.month);
  const checkedText = loadedAt ? loadedAt.toLocaleString("en-PH", { dateStyle:"short", timeStyle:"medium" }) : "—";

  return (
    <div className={`dashboard-shell ${dark ? "dark-mode" : ""}`}>
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark">R</div>
          <div><div className="brand-title">realme CVIS</div><div className="brand-subtitle">Sales Intelligence</div></div>
        </div>
        <nav className="nav-list">
          {NAV_ITEMS.map(item => <button key={item} className={`nav-item ${activePage===item?"active":""}`} onClick={()=>setActivePage(item)}>{item}</button>)}
        </nav>
        <div className="sidebar-spacer" />
        <div className="nav-section-label">Data Sources</div>
        <button className="nav-secondary" onClick={()=>setActivePage("Current Smartphone Line-up")}>Current Smartphone Line-up</button>
        <button className="nav-secondary" onClick={()=>setActivePage("Store Performance")}>Store Performance</button>
        <button className="nav-secondary" onClick={()=>setActivePage("Inventory")}>Inventory {data.inventory.length ? "" : "(soon)"}</button>
        <div className="source-status"><span className="status-dot" /> Excel sales</div>
      </aside>

      <main className="main">
        <header className="topbar">
          <div>
            <div className="eyebrow">REALME CVIS</div>
            <h1 className="page-title">{activePage}</h1>
            <div className="page-subtitle">{currentMonthText} performance</div>
          </div>
          <div className="top-actions">
            <button className="btn" onClick={()=>setDark(d=>!d)}>{dark ? "☀ Light mode" : "☾ Night mode"}</button>
            <div className="last-updated">Last updated<br/><strong>{checkedText}</strong></div>
            <label className="btn upload-label">Upload Excel<input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={onFile}/></label>
            <button className="btn accent" onClick={loadDefault}>Refresh data</button>
          </div>
        </header>

        <div className="source-line">Excel sales · {fileName} · {fmt(sales.length)} rows · Last checked {checkedText}</div>
        {error && <div className="notice" style={{marginBottom:16}}><strong>Workbook error:</strong> {error}</div>}
        {loading && !sales.length && <div className="notice">Loading the CVIS workbook…</div>}

        {sales.length > 0 && <>
          {showFilters && <section className="panel filter-panel">
            <div className="filter-actions">
              <button className="btn tiny ghost" onClick={clearFilters}>Clear Filter</button>
              <button className="btn tiny ghost" onClick={undoFilters} disabled={!filterHistory.length}>Undo</button>
            </div>
            <div className="filter-grid">
              <Filter label="Month" value={filters.month} onChange={v=>updateFilter("month",v)} options={months.map(m=>[m,monthLabel(m)])} allLabel="All months" />
              <Filter label="Area" value={filters.area} onChange={v=>updateFilter("area",v)} options={options.areas} allLabel="All areas" />
              <Filter label="Subregions" value={filters.subregion} onChange={v=>updateFilter("subregion",v)} options={options.subregions} allLabel="All subregions" />
              <Filter label="Customer" value={filters.customer} onChange={v=>updateFilter("customer",v)} options={options.customers} allLabel="All customers" />
              <Filter label="Channel" value={filters.channel} onChange={v=>updateFilter("channel",v)} options={options.channels} allLabel="All channels" />
              <Filter label="Model" value={filters.model} onChange={v=>updateFilter("model",v)} options={options.models} allLabel="All models" />
              <Filter label="Price Range" value={filters.priceRange} onChange={v=>updateFilter("priceRange",v)} options={options.priceRanges} allLabel="All price ranges" />
              <Filter label="Smartphone Series" value={filters.series} onChange={v=>updateFilter("series",v)} options={options.series} allLabel="All series" className="span2" />
            </div>
          </section>}
          <div className="section-toolbar"><button className="btn tiny" onClick={()=>setShowFilters(x=>!x)}>{showFilters?"Hide":"Show filters"}</button></div>

          {activePage === "Overview" ? <Overview
            totalUnits={totalUnits} totalAmount={totalAmount} asp={asp} above13={above13} dailyRunRate={dailyRunRate} activeStores={activeStores}
            prevUnits={prevUnits} prevAmount={prevAmount} prevAsp={prevAsp} prevAbove={prevAbove} prevRunRate={prevRunRate}
            targetUnits={targetUnits} salesTrend={salesTrend} trendMode={trendMode} setTrendMode={setTrendMode}
            dealerTrend={dealers.data} dealerNames={dealers.dealers} dealerMode={dealerMode} setDealerMode={setDealerMode}
          /> : <PageContent activePage={activePage} modelRows={modelRows} dealerRows={dealerRows} subregionRows={subregionRows} storeRows={storeRows} promoterRows={promoterRows} areaRows={areaRows} data={data} />}
        </>}
      </main>
    </div>
  );
}

function Filter({ label, value, onChange, options, allLabel, className="" }) {
  return <div className={`filter-field ${className}`}>
    <label>{label}</label>
    <select value={value} onChange={e=>onChange(e.target.value)}>
      <option value="ALL">{allLabel}</option>
      {options.map(o => Array.isArray(o) ? <option key={o[0]} value={o[0]}>{o[1]}</option> : <option key={o} value={o}>{o}</option>)}
    </select>
  </div>;
}

function Kpi({ label, value, caption, previousLabel, previous, current, target, targetLabel }) {
  const delta = previous ? change(current, previous) : null;
  const achievement = target ? (current/target)*100 : null;
  return <div className="kpi-card">
    <div className="kpi-label">{label}</div>
    <div className="kpi-value">{value}</div>
    <div className="kpi-caption">{caption}</div>
    <div className="kpi-subline">{previous ? <>{previousLabel}: {typeof previous === "string" ? previous : fmt(previous, previous % 1 ? 2 : 0)}<div className={`delta ${delta>=0?"good":"bad"}`}>IR {delta>=0?"▲":"▼"} {percent(delta).replace(/[+-]/,"")}</div></> : <>Previous month: <span className="muted">not available</span></>}</div>
    {target ? <div className="kpi-subline">{targetLabel || "Target"}: {fmt(target)}<div className="delta neutral">AR: {achievement.toFixed(2)}%</div></div> : null}
  </div>;
}

function Overview(props) {
  const { totalUnits,totalAmount,asp,above13,dailyRunRate,activeStores,prevUnits,prevAmount,prevAsp,prevAbove,prevRunRate,targetUnits,salesTrend,trendMode,setTrendMode,dealerTrend,dealerNames,dealerMode,setDealerMode } = props;
  return <>
    <section className="kpi-grid">
      <Kpi label="Sales" value={fmt(totalUnits)} caption="units sold" previous={prevUnits} current={totalUnits} previousLabel="Previous month" target={targetUnits} targetLabel="Month target" />
      <Kpi label="Sales Amount (PHP)" value={money(totalAmount,2)} caption="total Sales Amount · PHP" previous={prevAmount} current={totalAmount} previousLabel="Previous month" />
      <Kpi label="ASP (PHP)" value={money(asp,2)} caption="Sales Amount ÷ total units sold" previous={prevAsp} current={asp} previousLabel="Previous month" />
      <Kpi label="Sales above ₱13,000" value={fmt(above13)} caption={`${totalUnits ? ((above13/totalUnits)*100).toFixed(1) : 0}% of units sold`} previous={prevAbove} current={above13} previousLabel="Previous month" />
      <Kpi label="Daily Run Rate" value={fmt(dailyRunRate,1)} caption="units / elapsed calendar day" previous={prevRunRate} current={dailyRunRate} previousLabel="Previous month" />
      <Kpi label="Active Stores" value={fmt(activeStores)} caption="stores with sell-out in selected period" current={activeStores} />
    </section>

    <section className="chart-grid">
      <TrendCard title="Sales Trend" data={salesTrend} mode={trendMode} setMode={setTrendMode} />
      <DealerTrend title="Dealer Sales Trend" data={dealerTrend} dealers={dealerNames} mode={dealerMode} setMode={setDealerMode} />
    </section>
  </>;
}

function ModeButtons({mode,setMode}) {
  return <div className="segmented">{["Daily","Weekly","Monthly"].map(m=><button key={m} className={mode===m?"active":""} onClick={()=>setMode(m)}>{m}</button>)}</div>;
}

function TrendCard({title,data,mode,setMode}) {
  return <div className="chart-card">
    <div className="card-head"><h2 className="card-title">{title}</h2><div className="card-actions"><button className="btn tiny">Hide</button><button className="btn tiny">Expand</button></div></div>
    <ModeButtons mode={mode} setMode={setMode}/>
    <div className="chart-note">Selected Excel rows · units sold</div>
    {data.length ? <div className="chart-wrap"><ResponsiveContainer width="100%" height="100%"><LineChart data={data} margin={{top:12,right:12,left:-12,bottom:8}}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e6e8ec"/><XAxis dataKey="label" tick={{fontSize:11,fill:"#7a8390"}} tickLine={false} axisLine={false}/><YAxis tick={{fontSize:11,fill:"#7a8390"}} tickLine={false} axisLine={false}/><Tooltip formatter={(v)=>[fmt(v),"Units"]}/><Line type="monotone" dataKey="units" stroke="#e6ad00" strokeWidth={3} dot={{r:3}} activeDot={{r:5}}/></LineChart></ResponsiveContainer></div> : <div className="chart-empty">No trend data for the selected filters.</div>}
  </div>;
}

function DealerTrend({title,data,dealers,mode,setMode}) {
  const strokes=["#e6ad00","#2563eb","#16a34a","#7c3aed","#ea580c"];
  return <div className="chart-card">
    <div className="card-head"><h2 className="card-title">{title}</h2><div className="card-actions"><button className="btn tiny">Hide</button><button className="btn tiny">Expand</button></div></div>
    <ModeButtons mode={mode} setMode={setMode}/>
    <div className="chart-note">Top 5 customers by selected-period units</div>
    {data.length ? <div className="chart-wrap"><ResponsiveContainer width="100%" height="100%"><LineChart data={data} margin={{top:12,right:12,left:-12,bottom:8}}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e6e8ec"/><XAxis dataKey="label" tick={{fontSize:11,fill:"#7a8390"}} tickLine={false} axisLine={false}/><YAxis tick={{fontSize:11,fill:"#7a8390"}} tickLine={false} axisLine={false}/><Tooltip/><Legend wrapperStyle={{fontSize:10}}/>{dealers.map((d,i)=><Line key={d} type="monotone" dataKey={d} stroke={strokes[i%strokes.length]} strokeWidth={2.2} dot={false}/>)}</LineChart></ResponsiveContainer></div> : <div className="chart-empty">No dealer trend data for the selected filters.</div>}
  </div>;
}

function aggregate(rows, key) {
  const map = new Map();
  rows.forEach(r => {
    const name = r[key] || "Unknown";
    const x = map.get(name) || { name, units:0, amount:0, stores:new Set() };
    x.units += r.units; x.amount += r.amount; x.stores.add(r.store); map.set(name,x);
  });
  return [...map.values()].map(x=>({name:x.name, units:x.units, amount:x.amount, stores:x.stores.size, asp:x.units?x.amount/x.units:0})).sort((a,b)=>b.units-a.units);
}

function PageContent({ activePage, modelRows, dealerRows, subregionRows, storeRows, promoterRows, areaRows, data }) {
  if (activePage === "Push Model Performance" || activePage === "Current Smartphone Line-up") {
    return <RankingPage title="Model Performance" rows={modelRows} firstLabel="Model" />;
  }
  if (activePage === "Dealer Performance") return <RankingPage title="Dealer Performance" rows={dealerRows} firstLabel="Dealer / Customer" />;
  if (activePage === "PS Sales Review") return <RankingPage title="Subregion Sales Review" rows={subregionRows} firstLabel="Subregion" />;
  if (activePage === "Store Performance") return <RankingPage title="Store Performance" rows={storeRows} firstLabel="Store" />;
  if (activePage === "Promoter Score" || activePage === "Promoter Productivity") {
    return <div className="panel content-panel"><div className="card-head"><h2 className="card-title">{activePage}</h2><span className="pill">{promoterRows.length} promoters</span></div><div className="table-wrap"><table><thead><tr><th>Rank</th><th>Promoter</th><th>Area</th><th>Units Sold</th><th>Status</th></tr></thead><tbody>{promoterRows.slice(0,50).map((r,i)=><tr key={`${r["Promoter ID"]}-${i}`}><td><span className="rank">{i+1}</span></td><td><strong>{r.Name}</strong></td><td>{r.Area}</td><td>{fmt(r["Units Sold"])}</td><td><span className={`pill ${String(r.Status).toUpperCase()==="ACTIVE"?"good":""}`}>{r.Status || "—"}</span></td></tr>)}</tbody></table></div></div>;
  }
  if (activePage === "ASM Incentives") return <RankingPage title="ASM / Area Performance" rows={areaRows} firstLabel="Area / Subregion" />;
  if (activePage === "Daily Zero Sellout") return <div className="panel content-panel"><h2 className="card-title">Daily Zero Sellout</h2><div className="notice" style={{marginTop:14}}>Your Sales_Data contains sell-out transactions only. A true zero-sellout report needs a daily store roster or attendance/store-opening sheet so the dashboard can distinguish “zero sales” from “no record uploaded.”</div></div>;
  if (activePage === "Inventory") return <div className="panel content-panel"><h2 className="card-title">Inventory</h2>{data.inventory.length ? <div className="notice" style={{marginTop:14}}>Inventory data detected: {fmt(data.inventory.length)} rows. This page is ready for stock KPIs.</div> : <div className="notice" style={{marginTop:14}}>The Inventory sheet is currently empty. Add inventory rows to activate stock-on-hand, low-stock, and days-of-inventory analytics.</div>}</div>;
  return <div className="panel content-panel"><div className="notice">This dashboard page is ready to be connected to an additional business rule or Excel sheet.</div></div>;
}

function RankingPage({title,rows,firstLabel}) {
  return <div className="panel content-panel">
    <div className="card-head"><h2 className="card-title">{title}</h2><span className="pill">Top {Math.min(rows.length,30)}</span></div>
    <div className="chart-wrap" style={{height:270, marginTop:14}}><ResponsiveContainer width="100%" height="100%"><BarChart data={rows.slice(0,10)} margin={{top:10,right:16,left:-10,bottom:36}}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e6e8ec"/><XAxis dataKey="name" angle={-20} textAnchor="end" interval={0} tick={{fontSize:10,fill:"#78818f"}}/><YAxis tick={{fontSize:11,fill:"#78818f"}}/><Tooltip formatter={(v)=>fmt(v)}/><Bar dataKey="units" fill="#ffc400" radius={[6,6,0,0]}/></BarChart></ResponsiveContainer></div>
    <div className="table-wrap"><table><thead><tr><th>Rank</th><th>{firstLabel}</th><th>Units</th><th>Sales Amount</th><th>ASP</th><th>Stores</th></tr></thead><tbody>{rows.slice(0,30).map((r,i)=><tr key={r.name}><td><span className="rank">{i+1}</span></td><td><strong>{r.name}</strong></td><td>{fmt(r.units)}</td><td>{money(r.amount)}</td><td>{money(r.asp)}</td><td>{fmt(r.stores)}</td></tr>)}</tbody></table></div>
  </div>;
}
