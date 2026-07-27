import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

// ── ATO TD 2025/4 ─────────────────────────────────────────────────────────────
const RATES = {
  capital:  { B: 34.75, L: 39.10, D: 66.65, I: 24.50 },
  regional: { B: 31.15, L: 35.55, D: 61.30, I: 24.50 },
};
const DEFAULT_SPEND_PCT = 80; // default allowance = 80% of ATO reasonable amounts
const MEALS = [
  { key: "B", label: "Breakfast" },
  { key: "L", label: "Lunch" },
  { key: "D", label: "Dinner" },
];
const ITEMS = [...MEALS, { key: "I", label: "Incidentals" }];

// ── Location database ──────────────────────────────────────────────────────────
const LOCATIONS = [
  { name: "Brisbane",        state: "QLD", type: "capital" },
  { name: "Sydney",          state: "NSW", type: "capital" },
  { name: "Melbourne",       state: "VIC", type: "capital" },
  { name: "Adelaide",        state: "SA",  type: "capital" },
  { name: "Perth",           state: "WA",  type: "capital" },
  { name: "Hobart",          state: "TAS", type: "capital" },
  { name: "Darwin",          state: "NT",  type: "capital" },
  { name: "Canberra",        state: "ACT", type: "capital" },
  { name: "Moranbah",        state: "QLD", type: "capital", note: "High-cost centre" },
  { name: "Weipa",           state: "QLD", type: "capital", note: "High-cost centre" },
  { name: "Karratha",        state: "WA",  type: "capital", note: "High-cost centre" },
  { name: "Port Hedland",    state: "WA",  type: "capital", note: "High-cost centre" },
  { name: "Broome",          state: "WA",  type: "capital", note: "High-cost centre" },
  { name: "Newman",          state: "WA",  type: "capital", note: "High-cost centre" },
  { name: "Kalgoorlie",      state: "WA",  type: "capital", note: "High-cost centre" },
  { name: "Exmouth",         state: "WA",  type: "capital", note: "High-cost centre" },
  { name: "Carnarvon",       state: "WA",  type: "capital", note: "High-cost centre" },
  { name: "Alice Springs",   state: "NT",  type: "capital", note: "High-cost centre" },
  { name: "Katherine",       state: "NT",  type: "capital", note: "High-cost centre" },
  { name: "Nhulunbuy",       state: "NT",  type: "capital", note: "High-cost centre" },
  { name: "Tennant Creek",   state: "NT",  type: "capital", note: "High-cost centre" },
  { name: "Gold Coast",      state: "QLD", type: "regional" },
  { name: "Robina",          state: "QLD", type: "regional", note: "Gold Coast" },
  { name: "Surfers Paradise",state: "QLD", type: "regional", note: "Gold Coast" },
  { name: "Southport",       state: "QLD", type: "regional", note: "Gold Coast" },
  { name: "Broadbeach",      state: "QLD", type: "regional", note: "Gold Coast" },
  { name: "Coomera",         state: "QLD", type: "regional", note: "Gold Coast" },
  { name: "Helensvale",      state: "QLD", type: "regional", note: "Gold Coast" },
  { name: "Varsity Lakes",   state: "QLD", type: "regional", note: "Gold Coast" },
  { name: "Sunshine Coast",  state: "QLD", type: "regional" },
  { name: "Noosa",           state: "QLD", type: "regional" },
  { name: "Mackay",          state: "QLD", type: "regional" },
  { name: "Townsville",      state: "QLD", type: "regional" },
  { name: "Cairns",          state: "QLD", type: "regional" },
  { name: "Rockhampton",     state: "QLD", type: "regional" },
  { name: "Bundaberg",       state: "QLD", type: "regional" },
  { name: "Gladstone",       state: "QLD", type: "regional" },
  { name: "Toowoomba",       state: "QLD", type: "regional" },
  { name: "Mount Isa",       state: "QLD", type: "regional" },
  { name: "Dysart",          state: "QLD", type: "regional" },
  { name: "Emerald",         state: "QLD", type: "regional" },
  { name: "Blackwater",      state: "QLD", type: "regional" },
  { name: "Clermont",        state: "QLD", type: "regional" },
  { name: "Middlemount",     state: "QLD", type: "regional" },
  { name: "Bowen",           state: "QLD", type: "regional" },
  { name: "Collinsville",    state: "QLD", type: "regional" },
  { name: "Charters Towers", state: "QLD", type: "regional" },
  { name: "Longreach",       state: "QLD", type: "regional" },
  { name: "Biloela",         state: "QLD", type: "regional" },
  { name: "Moura",           state: "QLD", type: "regional" },
  { name: "Capella",         state: "QLD", type: "regional" },
  { name: "Springsure",      state: "QLD", type: "regional" },
  { name: "Alpha",           state: "QLD", type: "regional" },
  { name: "Yeppoon",         state: "QLD", type: "regional" },
  { name: "Airlie Beach",    state: "QLD", type: "regional" },
  { name: "Proserpine",      state: "QLD", type: "regional" },
  { name: "Hervey Bay",      state: "QLD", type: "regional" },
  { name: "Maryborough",     state: "QLD", type: "regional" },
  { name: "Newcastle",       state: "NSW", type: "regional" },
  { name: "Wollongong",      state: "NSW", type: "regional" },
  { name: "Singleton",       state: "NSW", type: "regional" },
  { name: "Muswellbrook",    state: "NSW", type: "regional" },
  { name: "Cessnock",        state: "NSW", type: "regional" },
  { name: "Maitland",        state: "NSW", type: "regional" },
  { name: "Lithgow",         state: "NSW", type: "regional" },
  { name: "Geelong",         state: "VIC", type: "regional" },
  { name: "Ballarat",        state: "VIC", type: "regional" },
  { name: "Bendigo",         state: "VIC", type: "regional" },
  { name: "Launceston",      state: "TAS", type: "regional" },
  { name: "Devonport",       state: "TAS", type: "regional" },
  { name: "Whyalla",         state: "SA",  type: "regional" },
  { name: "Port Augusta",    state: "SA",  type: "regional" },
  { name: "Mount Gambier",   state: "SA",  type: "regional" },
  { name: "Geraldton",       state: "WA",  type: "regional" },
  { name: "Albany",          state: "WA",  type: "regional" },
  { name: "Bunbury",         state: "WA",  type: "regional" },
];

// ── Helpers ────────────────────────────────────────────────────────────────────
const STORAGE_KEY = "travel-meal-diary-v5";
const MAX_DAYS    = 31;
const todayISO    = () => new Date().toISOString().slice(0, 10);
const parseDate   = (s) => new Date(s + "T12:00:00");
const fmtLong     = (s) => parseDate(s).toLocaleDateString("en-AU", { day: "numeric", month: "short", year: "numeric" });
const fmtShort    = (s) => parseDate(s).toLocaleDateString("en-AU", { weekday: "short", day: "numeric", month: "short" });
const fmtExcel    = (s) => parseDate(s).toLocaleDateString("en-AU", { day: "2-digit", month: "2-digit", year: "numeric" });
const fmt$        = (n) => `$${Number(n).toFixed(2)}`;

const datesBetween = (start, end) => {
  const dates = []; const fin = parseDate(end); const d = parseDate(start);
  while (d <= fin && dates.length <= MAX_DAYS) {
    dates.push(d.toISOString().slice(0, 10)); d.setDate(d.getDate() + 1);
  }
  return dates;
};

// Pre-trip determination: expected meals → day allowance
const roundDown5 = (n) => Math.floor(n / 5) * 5; // clean payroll figures, always rounds down (stays under ATO caps)
const tripPct = (trip) => trip.spendPct ?? (trip.rateCap ? null : DEFAULT_SPEND_PCT); // legacy trips keep rateCap behaviour
const dayAllowance = (expected, locType, pctOrLegacy, legacyCap) => {
  const r = RATES[locType];
  const mealSum = MEALS.reduce((s, { key }) => s + (expected[key] ? r[key] : 0), 0);
  const atoFull = mealSum + r.I; // expected meals + full incidentals rate
  if (pctOrLegacy == null && legacyCap) {
    // legacy trip saved under the old daily-cap model
    const rawLegacy = mealSum + Math.round(r.I * 0.8 * 100) / 100;
    return { raw: rawLegacy, allowance: roundDown5(Math.min(legacyCap, rawLegacy)), mealSum, atoFull };
  }
  const pct = (pctOrLegacy ?? DEFAULT_SPEND_PCT) / 100;
  return { raw: atoFull * pct, allowance: roundDown5(atoFull * pct), mealSum, atoFull };
};
const tripAllowance = (trip) =>
  trip.days.reduce((s, d) => s + dayAllowance(d.expected, trip.locType, tripPct(trip), trip.rateCap).allowance, 0);

// Post-trip diary: actual spend
const dayActual = (d) => ITEMS.reduce((s, { key }) => s + (parseFloat(d.actual?.[key]) || 0), 0);
const tripActual = (trip) => trip.days.reduce((s, d) => s + dayActual(d), 0);
const dayClaimable = (d, locType) => ITEMS.reduce((s, { key }) => s + Math.min(parseFloat(d.actual?.[key]) || 0, RATES[locType][key]), 0);
const tripClaimable = (trip) => trip.days.reduce((s, d) => s + dayClaimable(d, trip.locType), 0);
const tripHasDiary = (trip) => trip.days.some(d => dayActual(d) > 0);

// ── Component ──────────────────────────────────────────────────────────────────
export default function TravelMealDiary() {
  const [trips, setTrips] = useState([]);
  const [ready, setReady] = useState(false);
  // Determination form
  const [destInput, setDestInput] = useState("");
  const [destLoc,   setDestLoc]   = useState(null);
  const [locType,   setLocType]   = useState(null);
  const [startDate, setStartDate] = useState(todayISO());
  const [endDate,   setEndDate]   = useState(todayISO());
  const [spendPct,  setSpendPct]  = useState(String(DEFAULT_SPEND_PCT));
  const [purpose,   setPurpose]   = useState("");
  const [expDays,   setExpDays]   = useState([]);
  const [sugg,      setSugg]      = useState([]);
  const [showSugg,  setShowSugg]  = useState(false);
  // UI
  const [expanded,  setExpanded]  = useState({});
  const [diaryEdit, setDiaryEdit] = useState({}); // tripId -> editable actuals
  const [flash,     setFlash]     = useState(null);
  const [deleteId,  setDeleteId]  = useState(null);

  useEffect(() => {
    try {
      const r = localStorage.getItem(STORAGE_KEY);
      if (r) setTrips(JSON.parse(r));
    } catch {}
    setReady(true);
  }, []);

  const persist = useCallback(async (data) => {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); } catch {}
  }, []);

  useEffect(() => {
    if (!destInput.trim()) { setSugg([]); return; }
    const q = destInput.toLowerCase();
    setSugg(LOCATIONS.filter(l =>
      l.name.toLowerCase().startsWith(q) ||
      `${l.name} ${l.state}`.toLowerCase().includes(q) ||
      (l.note && l.note.toLowerCase().includes(q))
    ).slice(0, 8));
  }, [destInput]);

  useEffect(() => {
    if (!startDate || !endDate || endDate < startDate || !locType) return;
    const dates = datesBetween(startDate, endDate);
    if (dates.length > MAX_DAYS) return;
    setExpDays(prev => {
      const pm = {}; prev.forEach(d => { pm[d.date] = d; });
      return dates.map(date => pm[date] || { date, B: true, L: true, D: true });
    });
  }, [startDate, endDate, locType]);

  const selectLoc = (loc) => {
    setDestLoc(loc); setDestInput(`${loc.name}, ${loc.state}`);
    setLocType(loc.type); setShowSugg(false); setSugg([]);
  };

  const toggleExp = (date, k) => setExpDays(p => p.map(d => d.date === date ? { ...d, [k]: !d[k] } : d));
  const showFlash = (type, msg) => { setFlash({ type, msg }); setTimeout(() => setFlash(null), 5000); };

  // ── Save determination (pre-trip, locks the allowance) ──────────────────────
  const saveDetermination = async () => {
    if (!destInput.trim()) return showFlash("err", "Enter a destination.");
    if (!locType)          return showFlash("err", "Select a location type.");
    if (!purpose.trim())   return showFlash("err", "Enter a work purpose.");
    if (!expDays.length)   return showFlash("err", "Select a valid date range.");
    const pct = parseFloat(spendPct);
    if (!pct || pct <= 0 || pct > 100) return showFlash("err", "Enter a spend % between 1 and 100.");
    if (!expDays.some(d => d.B || d.L || d.D))
                           return showFlash("err", "Select expected meals for at least one day.");
    const trip = {
      id: Date.now(),
      determinedOn: todayISO(),
      dest: destInput.trim(), locType, startDate, endDate,
      spendPct: pct, purpose: purpose.trim(),
      days: expDays.map(d => ({
        date: d.date,
        expected: { B: d.B, L: d.L, D: d.D },
        actual: { B: "", L: "", D: "", I: "" },
      })),
    };
    const next = [trip, ...trips]; setTrips(next); await persist(next);
    setDestInput(""); setDestLoc(null); setLocType(null);
    setStartDate(todayISO()); setEndDate(todayISO());
    setSpendPct(String(DEFAULT_SPEND_PCT)); setPurpose(""); setExpDays([]);
    showFlash("ok", "Allowance determination saved — this is now the fixed amount to pay in MYOB.");
  };

  // ── Diary editing ────────────────────────────────────────────────────────────
  const startDiary = (trip) => {
    setDiaryEdit(p => ({ ...p, [trip.id]: trip.days.map(d => ({ date: d.date, ...d.actual })) }));
    setExpanded(p => ({ ...p, [trip.id]: true }));
  };
  const setDiaryVal = (tripId, date, key, v) => {
    if (v !== "" && !/^\d{0,4}(\.\d{0,2})?$/.test(v)) return;
    setDiaryEdit(p => ({ ...p, [tripId]: p[tripId].map(d => d.date === date ? { ...d, [key]: v } : d) }));
  };
  const saveDiary = async (tripId) => {
    const edit = diaryEdit[tripId];
    const next = trips.map(t => t.id !== tripId ? t : {
      ...t,
      days: t.days.map(d => {
        const e = edit.find(x => x.date === d.date);
        return e ? { ...d, actual: { B: e.B, L: e.L, D: e.D, I: e.I } } : d;
      }),
    });
    setTrips(next); await persist(next);
    setDiaryEdit(p => { const q = { ...p }; delete q[tripId]; return q; });
    showFlash("ok", "Diary saved.");
  };

  const confirmDelete = async (id) => {
    const next = trips.filter(t => t.id !== id); setTrips(next); await persist(next); setDeleteId(null);
  };

  // ── Excel export ─────────────────────────────────────────────────────────────
  const exportExcel = () => {
    if (!trips.length) return showFlash("err", "No trips to export.");
    const rows = [];
    [...trips].sort((a, b) => a.startDate.localeCompare(b.startDate)).forEach(trip => {
      trip.days.forEach(day => {
        const det = dayAllowance(day.expected, trip.locType, tripPct(trip), trip.rateCap);
        const act = dayActual(day);
        const claim = dayClaimable(day, trip.locType);
        rows.push({
          "Date":                     fmtExcel(day.date),
          "Destination":              trip.dest,
          "Location Type":            trip.locType === "capital" ? "Capital / High-cost" : "Regional",
          "Work Purpose":             trip.purpose,
          "Determination Date":       fmtExcel(trip.determinedOn),
          "Expected Breakfast":       day.expected.B ? "Yes" : "No",
          "Expected Lunch":           day.expected.L ? "Yes" : "No",
          "Expected Dinner":          day.expected.D ? "Yes" : "No",
          "Allowance % of ATO Rate":  trip.spendPct ? `${trip.spendPct}%` : "legacy cap $" + trip.rateCap,
          "ATO Full Rate ($)":        Math.round(det.atoFull * 100) / 100,
          "Day Allowance ($)":        det.allowance,
          "Actual Breakfast ($)":     parseFloat(day.actual.B) || "",
          "Actual Lunch ($)":         parseFloat(day.actual.L) || "",
          "Actual Dinner ($)":        parseFloat(day.actual.D) || "",
          "Actual Incidentals ($)":   parseFloat(day.actual.I) || "",
          "Day Actual Total ($)":     act || "",
          "Day Max Claimable ($)":    act > 0 ? Math.round(claim * 100) / 100 : "",
          "Variance ($)":             act > 0 ? Math.round((claim - det.allowance) * 100) / 100 : "",
          "Claimable Shortfall Flag": act > 0 && claim < det.allowance - 0.005 ? "Yes — allowance exceeds max claimable; gap may be assessable" : "",
          "FY":                       "2025-26",
          "ATO Reference":            "TD 2025/4",
        });
      });
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [
      { wch: 12 }, { wch: 22 }, { wch: 20 }, { wch: 44 }, { wch: 17 },
      { wch: 17 }, { wch: 14 }, { wch: 15 }, { wch: 22 }, { wch: 16 }, { wch: 16 },
      { wch: 18 }, { wch: 14 }, { wch: 15 }, { wch: 19 }, { wch: 17 },
      { wch: 20 }, { wch: 12 }, { wch: 44 }, { wch: 8 }, { wch: 13 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "ATO Travel Allowance Diary");
    XLSX.writeFile(wb, "ato_travel_allowance_diary_2025-26.xlsx");
  };

  // ── Recommendation engine ────────────────────────────────────────────────────
  const completed = trips.filter(tripHasDiary);
  const recDays = completed.flatMap(t => t.days.filter(d => dayActual(d) > 0).map(d => ({
    allowance: dayAllowance(d.expected, t.locType, tripPct(t), t.rateCap).allowance,
    actual: dayActual(d),
    nMeals: ["B","L","D"].filter(k => d.expected[k]).length,
  })));
  const fullDays    = recDays.filter(d => d.nMeals === 3);
  const partialDays = recDays.filter(d => d.nMeals > 0 && d.nMeals < 3);
  const avg = arr => arr.length ? arr.reduce((s, x) => s + x, 0) / arr.length : 0;

  // ── Derived form state ───────────────────────────────────────────────────────
  const pctVal   = Math.min(100, Math.max(1, parseFloat(spendPct) || DEFAULT_SPEND_PCT));
  const locRates = locType ? RATES[locType] : RATES.regional;
  const nDays    = startDate && endDate && endDate >= startDate ? datesBetween(startDate, endDate).length : 0;
  const showGrid = locType && nDays > 0 && nDays <= MAX_DAYS && expDays.length === nDays;
  const formTotal = showGrid ? expDays.reduce((s, d) => s + dayAllowance(d, locType, pctVal).allowance, 0) : 0;
  const ytdAllow  = trips.reduce((s, t) => s + tripAllowance(t), 0);

  return (
    <div style={{ minHeight: "100vh", background: "#F4F1EB", fontFamily: "'Inter', system-ui, sans-serif", fontSize: 14, color: "#1A2535" }}>

      {/* ── HEADER ── */}
      <header style={{ background: "#1A2535", color: "#fff", padding: "20px 24px 18px" }}>
        <div style={{ maxWidth: 1060, margin: "0 auto" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 3 }}>
                <span style={{ fontSize: 20, fontWeight: 700 }}>ATO Travel Allowance Diary</span>
                <span style={BDG("#2A7B5E", "#D8F3DC")}>ATO TD 2025/4</span>
              </div>
              <p style={{ margin: "0 0 14px", color: "#8FA3B8", fontSize: 12 }}>
                1. Determine allowance before travel (fixed) · 2. Record actual spend after (evidence) · Allowance is never adjusted by actuals
              </p>
              <div style={{ display: "flex", gap: 28, flexWrap: "wrap" }}>
                {[["Trips", trips.length], ["Diaries done", completed.length], ["YTD allowances", `$${ytdAllow.toFixed(2)}`]].map(([l, v]) => (
                  <div key={l}>
                    <div style={{ fontSize: 10, color: "#8FA3B8", textTransform: "uppercase", letterSpacing: "0.6px" }}>{l}</div>
                    <div style={{ fontSize: 24, fontWeight: 700, fontVariantNumeric: "tabular-nums" }}>{v}</div>
                  </div>
                ))}
              </div>
            </div>
            <button onClick={exportExcel}
              style={{ background: "#C97A3A", color: "#fff", border: "none", borderRadius: 8, padding: "10px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", whiteSpace: "nowrap" }}>
              ↓ Export to Excel
            </button>
          </div>
        </div>
      </header>

      <div style={{ maxWidth: 1060, margin: "0 auto", padding: "0 16px 48px" }}>

        {/* ── STEP 1: DETERMINATION FORM ── */}
        <div style={{ background: "#fff", borderRadius: "0 0 12px 12px", boxShadow: "0 2px 10px rgba(0,0,0,.08)", padding: "20px 20px 18px", marginBottom: 14 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
            <span style={BDG("#1A2535", "#fff")}>STEP 1</span>
            <span style={{ fontWeight: 700, fontSize: 15 }}>Pre-Trip Allowance Determination</span>
            <span style={{ fontSize: 12, color: "#8FA3B8" }}>— complete BEFORE travelling; this fixes the allowance</span>
          </div>

          {flash && (
            <div style={{ position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)",
              zIndex: 1000, maxWidth: "90vw", padding: "12px 22px", borderRadius: 8,
              fontSize: 14, fontWeight: 600, boxShadow: "0 4px 18px rgba(0,0,0,.3)",
              background: flash.type === "ok" ? "#1E5C35" : "#9B1C1C", color: "#fff" }}>
              {flash.msg}
            </div>
          )}

          <div style={{ display: "grid", gridTemplateColumns: "1fr 140px 140px 110px", gap: 12, marginBottom: 12 }}>
            <div style={{ position: "relative" }}>
              <label style={LBL}>Destination</label>
              <div style={{ position: "relative" }}>
                <input value={destInput}
                  onChange={e => { setDestInput(e.target.value); setDestLoc(null); setShowSugg(true); if (!e.target.value) setLocType(null); }}
                  onFocus={() => setShowSugg(true)}
                  onBlur={() => setTimeout(() => setShowSugg(false), 180)}
                  placeholder="Start typing a city or town…"
                  style={{ ...INP, paddingRight: locType ? 116 : 12 }} />
                {locType && (
                  <span style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)",
                    fontSize: 10, fontWeight: 700, borderRadius: 4, padding: "2px 6px",
                    background: locType === "capital" ? "#E8F5E9" : "#FFF3E0",
                    color:      locType === "capital" ? "#2A7B5E"  : "#C45200" }}>
                    {locType === "capital" ? "Capital/High-cost" : "Regional"}
                  </span>
                )}
              </div>
              {showSugg && sugg.length > 0 && (
                <div style={{ position: "absolute", top: "100%", left: 0, right: 0, zIndex: 200,
                  background: "#fff", border: "1.5px solid #DDD8D0", borderTop: "none",
                  borderRadius: "0 0 8px 8px", boxShadow: "0 6px 16px rgba(0,0,0,.12)", maxHeight: 240, overflowY: "auto" }}>
                  {sugg.map(loc => (
                    <div key={`${loc.name}-${loc.state}`} onMouseDown={() => selectLoc(loc)}
                      style={{ padding: "9px 12px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: "1px solid #F0EDE8" }}
                      onMouseEnter={e => e.currentTarget.style.background = "#F4F1EB"}
                      onMouseLeave={e => e.currentTarget.style.background = "#fff"}>
                      <div>
                        <span style={{ fontWeight: 600 }}>{loc.name}, {loc.state}</span>
                        {loc.note && <span style={{ marginLeft: 6, fontSize: 11, color: "#8FA3B8" }}>· {loc.note}</span>}
                      </div>
                      <span style={{ fontSize: 10, fontWeight: 700, borderRadius: 4, padding: "2px 6px", flexShrink: 0, marginLeft: 8,
                        background: loc.type === "capital" ? "#E8F5E9" : "#FFF3E0",
                        color:      loc.type === "capital" ? "#2A7B5E"  : "#C45200" }}>
                        {loc.type === "capital" ? "Capital" : "Regional"}
                      </span>
                    </div>
                  ))}
                </div>
              )}
            </div>
            <div>
              <label style={LBL}>Departure</label>
              <input type="date" value={startDate}
                onChange={e => { setStartDate(e.target.value); if (e.target.value > endDate) setEndDate(e.target.value); }}
                style={INP} />
            </div>
            <div>
              <label style={LBL}>Return</label>
              <input type="date" value={endDate} min={startDate} onChange={e => setEndDate(e.target.value)} style={INP} />
            </div>
            <div>
              <label style={LBL}>Spend % of ATO</label>
              <input type="text" inputMode="decimal" value={spendPct} onChange={e => setSpendPct(e.target.value)}
                style={{ ...INP, textAlign: "right", fontFamily: "monospace" }} />
            </div>
          </div>

          {destInput.trim() && !destLoc && (
            <div style={{ marginBottom: 12 }}>
              <label style={LBL}>Location type <span style={{ color: "#B0BAC5", fontWeight: 400, textTransform: "none", letterSpacing: 0 }}>— not in lookup, please select</span></label>
              <div style={{ display: "flex", gap: 8 }}>
                {[["capital", "Capital / High-cost"], ["regional", "Regional"]].map(([v, l]) => (
                  <button key={v} onClick={() => setLocType(v)}
                    style={{ padding: "7px 14px", borderRadius: 6, border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600,
                      background: locType === v ? "#1A2535" : "#F0EDE8", color: locType === v ? "#fff" : "#5A6A7A" }}>{l}</button>
                ))}
              </div>
            </div>
          )}

          <div style={{ marginBottom: 14 }}>
            <label style={LBL}>Work purpose</label>
            <input value={purpose} onChange={e => setPurpose(e.target.value)}
              placeholder="e.g. HSR training delivery — Robina, Gold Coast"
              style={{ ...INP, width: "100%", boxSizing: "border-box" }} />
          </div>

          {nDays > MAX_DAYS && (
            <div style={{ marginBottom: 12, padding: "9px 14px", borderRadius: 6, background: "#FDECEC", color: "#9B1C1C", fontSize: 13 }}>
              Date range exceeds {MAX_DAYS} days — please split into multiple determinations.
            </div>
          )}

          {showGrid && (
            <div style={{ marginBottom: 16 }}>
              <div style={{ fontSize: 12, color: "#8FA3B8", marginBottom: 8 }}>
                Tick meals <strong>reasonably expected</strong> each day (e.g. untick breakfast/lunch on a 6pm departure day) ·
                Incidentals {fmt$(locRates.I)}/day included automatically ·
                Day allowance = {pctVal}% of ATO rate for expected meals + incidentals, rounded down to nearest $5 — scales with location automatically
              </div>
              <div style={{ overflowX: "auto", borderRadius: 8, border: "1px solid #E8E4DC" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 560 }}>
                  <thead>
                    <tr style={{ background: "#1A2535" }}>
                      <th style={{ ...TH, textAlign: "left", paddingLeft: 12, color: "#8FA3B8", width: 110 }}>Date</th>
                      {MEALS.map(({ key, label }) => (
                        <th key={key} style={{ ...TH, color: "#8FA3B8", textAlign: "center", width: 90 }}>
                          {label}<div style={{ fontSize: 9, color: "#5A6A7A", fontWeight: 400 }}>{fmt$(locRates[key])}</div>
                        </th>
                      ))}
                      <th style={{ ...TH, color: "#8FA3B8", width: 92 }}>ATO rate<div style={{ fontSize: 9, color: "#5A6A7A", fontWeight: 400 }}>incl. incid.</div></th>
                      <th style={{ ...TH, textAlign: "right", paddingRight: 12, color: "#D8F3DC", width: 110 }}>Day allowance</th>
                    </tr>
                  </thead>
                  <tbody>
                    {expDays.map((day, i) => {
                      const det = dayAllowance(day, locType, pctVal);
                      return (
                        <tr key={day.date} style={{ background: i % 2 === 0 ? "#fff" : "#FAFAF8" }}>
                          <td style={{ padding: "8px 8px 8px 12px", fontSize: 12, color: "#5A6A7A", whiteSpace: "nowrap" }}>{fmtShort(day.date)}</td>
                          {MEALS.map(({ key }) => (
                            <td key={key} style={{ padding: "6px", textAlign: "center" }}>
                              <input type="checkbox" checked={day[key]} onChange={() => toggleExp(day.date, key)}
                                style={{ width: 18, height: 18, accentColor: "#2A7B5E", cursor: "pointer" }} />
                            </td>
                          ))}
                          <td style={{ padding: "8px 6px", textAlign: "right", fontFamily: "monospace", fontSize: 12, color: "#5A6A7A" }}>
                            {fmt$(det.atoFull)}
                          </td>
                          <td style={{ padding: "8px 12px 8px 6px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, fontSize: 14, color: "#2A7B5E" }}>
                            {fmt$(det.allowance)}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot>
                    <tr style={{ background: "#1A2535", borderTop: "2px solid #2A7B5E" }}>
                      <td colSpan={5} style={{ padding: "11px 8px 11px 12px", color: "#8FA3B8", fontSize: 12, fontWeight: 600 }}>
                        Trip allowance — pay this in MYOB · {nDays} {nDays === 1 ? "day" : "days"}
                      </td>
                      <td style={{ padding: "11px 12px 11px 6px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, fontSize: 17, color: "#fff" }}>{fmt$(formTotal)}</td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>
          )}

          <button onClick={saveDetermination}
            style={{ background: "#2A7B5E", color: "#fff", border: "none", borderRadius: 8, padding: "11px 28px", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>
            Save determination
          </button>
        </div>

        {/* ── ATO NOTE ── */}
        <div style={{ background: "#EEF7F3", border: "1px solid #BDE4D0", borderRadius: 8, padding: "10px 14px", marginBottom: 14, fontSize: 12, color: "#2D5E42" }}>
          <strong>Why two steps:</strong> The allowance must be determined <em>before</em> travel based on reasonably expected costs, and is never adjusted afterwards — adjusting to actuals makes it a reimbursement, which requires receipts.
          The diary (Step 2, on each trip card) is the employee's evidence the allowance was expended on travel.
          Max claimable caps each item at its ATO rate — if it falls below the allowance paid (e.g. one oversized meal), the diary flags the gap, which may be assessable income. Retain records 5 years.
        </div>

        {/* ── RECOMMENDATION PANEL ── */}
        {recDays.length >= 3 && (
          <div style={{ background: "#fff", border: "1px solid #E8E4DC", borderRadius: 10, padding: "14px 16px", marginBottom: 18 }}>
            <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 8, color: "#1A2535" }}>📊 Reasonable Allowance Benchmark <span style={{ fontWeight: 400, color: "#8FA3B8" }}>— based on {recDays.length} diarised days across {completed.length} {completed.length === 1 ? "trip" : "trips"}</span></div>
            <div style={{ display: "flex", gap: 28, flexWrap: "wrap", fontSize: 13 }}>
              {fullDays.length > 0 && (
                <div>
                  <div style={{ fontSize: 10, color: "#8FA3B8", textTransform: "uppercase", letterSpacing: "0.5px" }}>Full days (3 meals) · {fullDays.length}</div>
                  <div>avg allowance <strong style={{ fontFamily: "monospace" }}>{fmt$(avg(fullDays.map(d => d.allowance)))}</strong> · avg spend <strong style={{ fontFamily: "monospace" }}>{fmt$(avg(fullDays.map(d => d.actual)))}</strong></div>
                </div>
              )}
              {partialDays.length > 0 && (
                <div>
                  <div style={{ fontSize: 10, color: "#8FA3B8", textTransform: "uppercase", letterSpacing: "0.5px" }}>Partial days · {partialDays.length}</div>
                  <div>avg allowance <strong style={{ fontFamily: "monospace" }}>{fmt$(avg(partialDays.map(d => d.allowance)))}</strong> · avg spend <strong style={{ fontFamily: "monospace" }}>{fmt$(avg(partialDays.map(d => d.actual)))}</strong></div>
                </div>
              )}
            </div>
            <div style={{ fontSize: 12, color: "#5A6A7A", marginTop: 8 }}>
              {avg(recDays.map(d => d.actual)) >= avg(recDays.map(d => d.allowance)) * 0.9
                ? "✓ Spend patterns support the current allowance settings as reasonable to expect."
                : "⚠ Average spend is running well below the determined allowances — consider whether the expected-meal selections or daily cap remain reasonable, and discuss with your accountant."}
            </div>
          </div>
        )}

        {/* ── TRIP CARDS ── */}
        {!ready ? (
          <div style={{ textAlign: "center", padding: 40, color: "#8FA3B8" }}>Loading…</div>
        ) : trips.length === 0 ? (
          <div style={{ textAlign: "center", padding: 48, background: "#fff", borderRadius: 10, color: "#8FA3B8" }}>
            <div style={{ fontSize: 36, marginBottom: 8 }}>📋</div>
            <div style={{ fontWeight: 600, marginBottom: 4 }}>No trips yet</div>
            <div style={{ fontSize: 13 }}>Create a pre-trip allowance determination above</div>
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {[...trips].sort((a, b) => b.startDate.localeCompare(a.startDate)).map(t => {
              const allow = tripAllowance(t);
              const hasDiary = tripHasDiary(t);
              const actual = tripActual(t);
              const isExp = expanded[t.id];
              const editing = !!diaryEdit[t.id];
              const nd = t.days.length;
              const claimable = tripClaimable(t);
              const shortfall = hasDiary && claimable < allow - 0.005;
              return (
                <div key={t.id} style={{ background: "#fff", borderRadius: 10, boxShadow: "0 2px 6px rgba(0,0,0,.06)", overflow: "hidden" }}>
                  <div style={{ display: "flex", alignItems: "center", padding: "14px 16px", gap: 12, cursor: "pointer" }}
                    onClick={() => setExpanded(p => ({ ...p, [t.id]: !p[t.id] }))}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 3, flexWrap: "wrap" }}>
                        <span style={{ fontWeight: 700, fontSize: 15 }}>{t.dest}</span>
                        <span style={{ fontSize: 10, fontWeight: 700, borderRadius: 4, padding: "2px 6px",
                          background: t.locType === "capital" ? "#E8F5E9" : "#FFF3E0",
                          color:      t.locType === "capital" ? "#2A7B5E"  : "#C45200" }}>
                          {t.locType === "capital" ? "Capital" : "Regional"}
                        </span>
                        <span style={hasDiary ? BDG("#E8F5E9", "#2A7B5E") : BDG("#FFF2CC", "#7F6000")}>
                          {hasDiary ? "✓ Diary recorded" : "◷ Diary pending"}
                        </span>
                        <span style={{ fontSize: 10, color: "#B0BAC5" }}>{nd} {nd === 1 ? "day" : "days"} · determined {fmtLong(t.determinedOn)}</span>
                      </div>
                      <div style={{ fontSize: 12, color: "#8FA3B8", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                        {nd > 1 ? `${fmtLong(t.startDate)} → ${fmtLong(t.endDate)}` : fmtLong(t.startDate)} · {t.purpose}
                      </div>
                    </div>
                    <div style={{ textAlign: "right", flexShrink: 0 }}>
                      <div style={{ fontFamily: "monospace", fontWeight: 700, fontSize: 18, color: "#2A7B5E" }}>{fmt$(allow)}</div>
                      <div style={{ fontSize: 10, color: "#8FA3B8" }}>allowance (fixed)</div>
                      {hasDiary && (
                        <div style={{ fontSize: 11, fontFamily: "monospace", color: shortfall ? "#C45200" : "#8FA3B8" }}>
                          spent {fmt$(actual)} · claim {fmt$(claimable)}{shortfall && " ⚠"}
                        </div>
                      )}
                    </div>
                    <button onClick={ev => { ev.stopPropagation(); setDeleteId(t.id); }}
                      style={{ background: "none", border: "none", cursor: "pointer", color: "#C7B5A0", fontSize: 18, padding: "0 0 0 4px", flexShrink: 0 }}>✕</button>
                  </div>

                  {deleteId === t.id && (
                    <div style={{ padding: "12px 16px", background: "#FFF5F5", display: "flex", alignItems: "center", gap: 10, borderTop: "1px solid #F0EDE8" }}>
                      <span style={{ flex: 1, fontSize: 13, color: "#9B1C1C" }}>Delete {t.dest} trip ({fmtLong(t.startDate)})? This removes the determination record.</span>
                      <button onClick={() => confirmDelete(t.id)} style={DBTN}>Delete</button>
                      <button onClick={() => setDeleteId(null)} style={CBTN}>Cancel</button>
                    </div>
                  )}

                  {isExp && (
                    <div style={{ borderTop: "1px solid #F0EDE8", padding: "12px 16px" }}>
                      {shortfall && (
                        <div style={{ marginBottom: 10, padding: "8px 12px", borderRadius: 6, background: "#FFF2CC", color: "#7F6000", fontSize: 12 }}>
                          ⚠ <strong>Max claimable {fmt$(claimable)} is below the {fmt$(allow)} allowance:</strong> spend within the ATO per-meal caps does not cover the allowance paid (e.g. one large meal exceeding its cap). The {fmt$(allow - claimable)} gap may be assessable income — flag with your accountant.
                        </div>
                      )}

                      {!editing && (
                        <button onClick={() => startDiary(t)}
                          style={{ marginBottom: 12, background: hasDiary ? "#F0EDE8" : "#C97A3A", color: hasDiary ? "#5A6A7A" : "#fff", border: "none", borderRadius: 7, padding: "8px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                          {hasDiary ? "Edit diary" : "STEP 2 — Record actual spend"}
                        </button>
                      )}

                      <div style={{ overflowX: "auto", borderRadius: 8, border: "1px solid #E8E4DC" }}>
                        <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 720 }}>
                          <thead>
                            <tr style={{ background: "#F4F1EB" }}>
                              <th style={{ ...TH, textAlign: "left", paddingLeft: 12 }}>Date</th>
                              <th style={{ ...TH, textAlign: "center" }}>Expected</th>
                              <th style={TH}>Allowance</th>
                              {ITEMS.map(({ key, label }) => (
                                <th key={key} style={TH}>{label}<div style={{ fontSize: 9, color: "#B0BAC5", fontWeight: 400 }}>actual $</div></th>
                              ))}
                              <th style={TH}>Spent</th>
                              <th style={{ ...TH, textAlign: "right", paddingRight: 12 }}>Max claimable</th>
                            </tr>
                          </thead>
                          <tbody>
                            {t.days.map((d, i) => {
                              const det = dayAllowance(d.expected, t.locType, tripPct(t), t.rateCap);
                              const edit = diaryEdit[t.id]?.find(x => x.date === d.date);
                              const act = editing
                                ? ITEMS.reduce((s, { key }) => s + (parseFloat(edit?.[key]) || 0), 0)
                                : dayActual(d);
                              const claim = editing
                                ? ITEMS.reduce((s, { key }) => s + Math.min(parseFloat(edit?.[key]) || 0, RATES[t.locType][key]), 0)
                                : dayClaimable(d, t.locType);
                              const expStr = ["B","L","D"].filter(k => d.expected[k]).join(" ") || "—";
                              return (
                                <tr key={d.date} style={{ background: i % 2 === 0 ? "#fff" : "#FAFAF8" }}>
                                  <td style={{ padding: "8px 8px 8px 12px", fontSize: 12, color: "#5A6A7A", whiteSpace: "nowrap" }}>{fmtShort(d.date)}</td>
                                  <td style={{ padding: "8px 6px", textAlign: "center", fontSize: 11, fontFamily: "monospace", color: "#5A6A7A" }}>{expStr} +I</td>
                                  <td style={{ padding: "8px 6px", textAlign: "right", fontFamily: "monospace", fontSize: 12, fontWeight: 700, color: "#2A7B5E" }}>{fmt$(det.allowance)}</td>
                                  {ITEMS.map(({ key }) => (
                                    <td key={key} style={{ padding: "5px" }}>
                                      {editing ? (
                                        <input type="text" inputMode="decimal" value={edit?.[key] ?? ""}
                                          onChange={e => setDiaryVal(t.id, d.date, key, e.target.value)}
                                          placeholder="—"
                                          style={{ width: "100%", boxSizing: "border-box", padding: "5px 7px", border: "1.5px solid #DDD8D0", borderRadius: 6, fontSize: 12, fontFamily: "monospace", textAlign: "right", outline: "none", minWidth: 62 }} />
                                      ) : (
                                        <div style={{ textAlign: "right", fontFamily: "monospace", fontSize: 12, color: d.actual[key] ? "#1A2535" : "#DDD8D0", paddingRight: 7 }}>
                                          {d.actual[key] ? fmt$(d.actual[key]) : "—"}
                                        </div>
                                      )}
                                    </td>
                                  ))}
                                  <td style={{ padding: "8px 6px", textAlign: "right", fontFamily: "monospace", fontSize: 12, color: act > 0 ? "#5A6A7A" : "#C7B5A0" }}>
                                    {act > 0 ? fmt$(act) : "—"}
                                  </td>
                                  <td style={{ padding: "8px 12px 8px 6px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, fontSize: 13, color: act > 0 ? (claim < det.allowance - 0.005 ? "#C45200" : "#2A7B5E") : "#C7B5A0" }}>
                                    {act > 0 ? (
                                      <span>
                                        {fmt$(claim)}
                                        {claim < act - 0.005 && <span style={{ display: "block", fontSize: 9, fontWeight: 600, color: "#C45200" }}>capped from {fmt$(act)}</span>}
                                      </span>
                                    ) : "—"}
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                          <tfoot>
                            <tr style={{ background: "#1A2535" }}>
                              <td colSpan={2} style={{ padding: "10px 8px 10px 12px", color: "#8FA3B8", fontSize: 12, fontWeight: 600 }}>Totals</td>
                              <td style={{ padding: "10px 6px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, fontSize: 13, color: "#D8F3DC" }}>{fmt$(allow)}</td>
                              <td colSpan={4} style={{ padding: "10px 6px", textAlign: "right", color: "#5A6A7A", fontSize: 11 }}>{hasDiary || editing ? "spent · claimable →" : ""}</td>
                              <td style={{ padding: "10px 6px", textAlign: "right", fontFamily: "monospace", fontSize: 13, color: "#8FA3B8" }}>
                                {editing
                                  ? fmt$(diaryEdit[t.id].reduce((s, e) => s + ITEMS.reduce((x, { key }) => x + (parseFloat(e[key]) || 0), 0), 0))
                                  : hasDiary ? fmt$(actual) : "—"}
                              </td>
                              <td style={{ padding: "10px 12px 10px 6px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, fontSize: 14, color: "#fff" }}>
                                {editing
                                  ? fmt$(diaryEdit[t.id].reduce((s, e) => s + ITEMS.reduce((x, { key }) => x + Math.min(parseFloat(e[key]) || 0, RATES[t.locType][key]), 0), 0))
                                  : hasDiary ? fmt$(claimable) : "—"}
                              </td>
                            </tr>
                          </tfoot>
                        </table>
                      </div>

                      {editing && (
                        <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                          <button onClick={() => saveDiary(t.id)}
                            style={{ background: "#2A7B5E", color: "#fff", border: "none", borderRadius: 7, padding: "9px 20px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                            Save diary
                          </button>
                          <button onClick={() => setDiaryEdit(p => { const q = { ...p }; delete q[t.id]; return q; })}
                            style={CBTN}>Cancel</button>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}

        <p style={{ textAlign: "center", marginTop: 22, fontSize: 11, color: "#B0BAC5" }}>
          ATO TD 2025/4 · Capital/High-cost: B $34.75 / L $39.10 / D $66.65 · Regional: B $31.15 / L $35.55 / D $61.30 · Incidentals $24.50 · Default allowance {DEFAULT_SPEND_PCT}% of ATO rate, rounded down to $5
        </p>
      </div>
    </div>
  );
}

const LBL  = { display: "block", fontSize: 10, fontWeight: 700, color: "#8FA3B8", textTransform: "uppercase", letterSpacing: "0.5px", marginBottom: 4 };
const INP  = { width: "100%", boxSizing: "border-box", padding: "9px 11px", border: "1.5px solid #DDD8D0", borderRadius: 7, fontSize: 14, color: "#1A2535", background: "#FAFAF8", outline: "none" };
const TH   = { padding: "8px 6px", fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.4px", color: "#8FA3B8", textAlign: "right" };
const DBTN = { background: "#C0392B", color: "#fff", border: "none", borderRadius: 6, padding: "6px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" };
const CBTN = { background: "#F0EDE8", color: "#5A6A7A", border: "none", borderRadius: 6, padding: "6px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" };
const BDG  = (bg, color) => ({ fontSize: 10, background: bg, color, borderRadius: 4, padding: "2px 7px", fontWeight: 700, letterSpacing: "0.4px" });
