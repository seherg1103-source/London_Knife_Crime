/* Knife Crime Strategic Command Center (concept simulation)
 * - Source crime file (M1045): ./data/M1045_MonthlyCrimeDashboard_KnifeCrimeData_with_coords.csv
 *   Copy also kept at ./data/knife_crime_london.csv. All offence counts use sum(Count) per M1045 row.
 * - Rebuild aggregates: python3 tools/rebuild_crime_aggregates_from_m1045.py [--copy-to-knife]
 *   Then refresh offline bundle: python3 tools/sync_data_bundle_crime.py
 * - Other data: ./data/neets-*.csv, ./data/Police_Force_Strength.csv
 *
 * M1045 Count semantics (Measure = Offences only):
 * - Pre-aggregated CSV `offences` / `totalOffences` = sum of raw Count for matching rows.
 * - Incidents (selected) = Σ Count over filtered areas and date window (= Σ offences).
 * - Subtype composition = for each Crime Subtype, Σ Count in scope (same as Σ offences per subtype).
 * - Personal Robbery Share = Σ Count(Personal Robbery subtype) ÷ Incidents.
 */

// Pre-aggregated data (generated via tools/rebuild_crime_aggregates_from_m1045.py)
const CRIME_CSV_PATH = "./data/knife_crime_london.csv"; // fallback (unused when aggregates / bundle exist)
const NEETS_XLSX_PATH = "./data/neets-borough-region.xlsx"; // fallback (unused when NEET csv exists)
const POLICE_CSV_PATH = "./data/Police_Force_Strength.csv";

const NEETS_MONTH_CSV_PATH = "./data/neets_london_month.csv";
const NEETS_YEAR_CSV_PATH = "./data/neets_london_year.csv";

const CRIME_BOROUGH_COORDS_CSV_PATH = "./data/crime_borough_coords.csv";
const CRIME_BOROUGH_META_CSV_PATH = "./data/crime_borough_meta.csv";
const CRIME_AREA_CODE_MAP_CSV_PATH = "./data/crime_area_code_map.csv";
const CRIME_MONTH_TOTAL_CSV_PATH = "./data/crime_month_total.csv";
const CRIME_YEAR_TOTAL_CSV_PATH = "./data/crime_year_total.csv";
const CRIME_MONTH_BOROUGH_CSV_PATH = "./data/crime_month_borough.csv";
const CRIME_YEAR_BOROUGH_CSV_PATH = "./data/crime_year_borough.csv";
const CRIME_MONTH_BOROUGH_BOROUGH_CSV_PATH = "./data/crime_month_borough_borough.csv";
const CRIME_YEAR_BOROUGH_BOROUGH_CSV_PATH = "./data/crime_year_borough_borough.csv";
const CRIME_MONTH_BOROUGH_SNT_CSV_PATH = "./data/crime_month_borough_snt.csv";
const CRIME_YEAR_BOROUGH_SNT_CSV_PATH = "./data/crime_year_borough_snt.csv";
const CRIME_MONTH_BOROUGH_SUBTYPE_BOROUGH_CSV_PATH = "./data/crime_month_borough_subtype_borough.csv";
const CRIME_YEAR_BOROUGH_SUBTYPE_BOROUGH_CSV_PATH = "./data/crime_year_borough_subtype_borough.csv";
const CRIME_MONTH_BOROUGH_SUBTYPE_SNT_CSV_PATH = "./data/crime_month_borough_subtype_snt.csv";
const CRIME_YEAR_BOROUGH_SUBTYPE_SNT_CSV_PATH = "./data/crime_year_borough_subtype_snt.csv";
const CRIME_MONTH_AREA_CODE_CSV_PATH = "./data/crime_month_area_code.csv";
const CRIME_YEAR_AREA_CODE_CSV_PATH = "./data/crime_year_area_code.csv";
const CRIME_MONTH_SUBTYPE_CSV_PATH = "./data/crime_month_subtype.csv";
const CRIME_YEAR_SUBTYPE_CSV_PATH = "./data/crime_year_subtype.csv";
const CRIME_MONTH_BOROUGH_SUBTYPE_CSV_PATH = "./data/crime_month_borough_subtype.csv";
const CRIME_YEAR_BOROUGH_SUBTYPE_CSV_PATH = "./data/crime_year_borough_subtype.csv";

/** Borough unemployment % — 2025 Q3 series from Trust for London dataset #1563 (“2024 Q3 and 2025 Q3” chart); London mean 5.6%; City of London 5.6%. */
const UNEMPLOYMENT_CSV_PATH = "./data/T72_rYkAqhQ.csv";
/** Half-width (percentage points) around London mean for the "Average" band. */
const UNEMPLOYMENT_MEAN_BAND_PP = 0.35;

const UNEMPLOYMENT_STYLE = {
  better: { fill: "#0d9488", stroke: "#0f172a" },
  average: { fill: "#f5edd8", stroke: "#1e293b" },
  worse: { fill: "#eab4c8", stroke: "#1e293b" },
  none: { fill: "#cbd5e1", stroke: "#475569" }
};

/** Inline copy of T72 CSV for file:// or failed fetch (Papa.parse). Synced with Trust for London /data/dataset/1563 (2025 Q3). */
const UNEMPLOYMENT_CSV_FALLBACK = `borough,unemployment_rate
London,5.6
Newham,8.7
Barking and Dagenham,7.9
Brent,7.9
Haringey,7.1
Redbridge,6.6
Lewisham,6.5
Enfield,6.3
Southwark,6.2
Waltham Forest,6.2
Tower Hamlets,6.1
Ealing,6
Greenwich,5.9
Hounslow,5.9
Barnet,5.7
Lambeth,5.6
Hillingdon,5.6
Hackney,5.4
Croydon,5.4
Merton,5.1
Hammersmith and Fulham,4.9
Harrow,4.8
Havering,4.8
Sutton,4.8
Kensington and Chelsea,4.7
Islington,4.6
Richmond upon Thames,4.6
Bexley,4.4
Camden,4.2
Bromley,4.2
Kingston upon Thames,4
Wandsworth,3.6
Westminster,3.3
City of London,5.6`;

const PALETTE = {
  base: "#132843",
  text: "#F8F6F6",
  blue1: "#4F92E0",
  blue2: "#6B90DE",
  blue3: "#76B6E4",
  warm1: "#E39F5A",
  warm2: "#DE8438",
  red1: "#C14C3B",
  green1: "#639A72",
  green2: "#ABD46E",
  violet: "#A368A9",
  /** Right-column / accent set: distinct for forecast bar, line, pie */
  forecast: ["#E8D248", "#EBA281", "#ABD46E", "#5EEAD4", "#FBBF24", "#F472B6"],
  /** High-luminance map overlays (avoid risk red / orange / green) */
  hubMarker: "#00F5FF",
  hubMarkerStroke: "#CCFFFF",
  electronicsMarker: "#E040FB",
  electronicsMarkerStroke: "#F5D0FF"
};

/** Full palette for forecast charts: unique hues, expanded if more series than length */
const ALL_CHART_COLORS = (() => {
  const seen = new Set();
  const out = [];
  const add = (c) => {
    if (!c || typeof c !== "string" || seen.has(c)) return;
    seen.add(c);
    out.push(c);
  };
  PALETTE.forecast.forEach(add);
  add(PALETTE.blue1);
  add(PALETTE.blue2);
  add(PALETTE.blue3);
  add(PALETTE.warm1);
  add(PALETTE.warm2);
  add(PALETTE.red1);
  add(PALETTE.green1);
  add(PALETTE.green2);
  add(PALETTE.violet);
  add(PALETTE.hubMarker);
  add(PALETTE.electronicsMarker);
  add(PALETTE.hubMarkerStroke);
  add(PALETTE.electronicsMarkerStroke);
  ["#0EA5E9", "#10B981", "#6366F1", "#F43F5E", "#84CC16", "#F59E0B", "#14B8A6", "#EC4899", "#22C55E", "#A855F7"].forEach(add);
  return out;
})();

function colorsForSeriesCount(n) {
  const pool = ALL_CHART_COLORS;
  if (n <= pool.length) return pool.slice(0, n);
  const out = pool.slice();
  for (let i = pool.length; i < n; i++) {
    const hue = Math.round((i * 47.3 + (i % 7) * 31) % 360);
    out.push(`hsl(${hue}, 68%, 54%)`);
  }
  return out;
}

/** Choropleth: deepen fill vs previous build (~+30% perceived weight) */
const CHORO_BASE_FILL_OPACITY = 0.47;
const CHORO_FALLBACK_FILL_OPACITY = 0.34;
const CHORO_ACTIVE_FILL_OPACITY = 0.29;
const CHORO_FILL_DARKEN = 0.7;

const SOCIAL_RADAR_LABELS = [
  "Geospatial Anchoring",
  "Aggression & Demand Velocity",
  "Network Symbiosis",
  "Hub Influence & Mobilization",
  "Commodity-to-Violence"
];

/** Crime Subtype string for Personal Robbery Share KPI (Count-weighted; Offences only). */
const PERSONAL_ROBBERY_SUBTYPE = "Knife Crime with Injury (Personal Robbery)";

const ui = {
  loadingOverlay: document.getElementById("loadingOverlay"),
  loadingMsg: document.getElementById("loadingMsg"),

  recordCount: document.getElementById("recordCount"),
  dateRange: document.getElementById("dateRange"),
  mapSubtitle: document.getElementById("mapSubtitle"),
  mapLegend: document.getElementById("mapLegend"),
  layerStatusText: document.getElementById("layerStatusText"),

  granularityMonthly: document.getElementById("granularityMonthly"),
  granularityYearly: document.getElementById("granularityYearly"),

  startTimeSelect: document.getElementById("startTimeSelect"),
  endTimeSelect: document.getElementById("endTimeSelect"),
  areaTypeFilter: document.getElementById("areaTypeFilter"),
  boroughSntFilter: document.getElementById("boroughSntFilter"),
  areaNameFilter: document.getElementById("areaNameFilter"),

  toggleCrimeCSV: document.getElementById("toggleCrimeCSV"),
  toggleHubs: document.getElementById("toggleHubs"),
  toggleElectronics: document.getElementById("toggleElectronics"),
  toggleGgphiSim: document.getElementById("toggleGgphiSim"),
  ggphiSimPanel: document.getElementById("ggphiSimPanel"),
  ggphiChatterInput: document.getElementById("ggphiChatterInput"),
  ggphiMarketInput: document.getElementById("ggphiMarketInput"),
  ggphiChatterLabel: document.getElementById("ggphiChatterLabel"),
  ggphiMarketLabel: document.getElementById("ggphiMarketLabel"),
  ggphiCrimeRate: document.getElementById("ggphiCrimeRate"),
  ggphiTrustRate: document.getElementById("ggphiTrustRate"),
  ggphiSustainRate: document.getElementById("ggphiSustainRate"),
  ggphiBtnPatrol: document.getElementById("ggphiBtnPatrol"),
  ggphiBtnCpted: document.getElementById("ggphiBtnCpted"),
  ggphiBtnVru: document.getElementById("ggphiBtnVru"),
  ggphiResetBtn: document.getElementById("ggphiResetBtn"),
  ggphiLegPatrol: document.getElementById("ggphiLegPatrol"),
  ggphiLegCpted: document.getElementById("ggphiLegCpted"),
  ggphiLegVru: document.getElementById("ggphiLegVru"),

  // KPIs
  kpiTotal: document.getElementById("kpiTotal"),
  kpiDelta: document.getElementById("kpiDelta"),
  kpiRobbery: document.getElementById("kpiRobbery"),
  kpiDisplacement: document.getElementById("kpiDisplacement"),
  trendSubtitle: document.getElementById("trendSubtitle"),
  kpiLineChartCanvas: document.getElementById("kpiLineChart"),
  kpiPieChartCanvas: document.getElementById("kpiPieChart"),

  // Predicted Intelligence
  predModeBar: document.getElementById("predModeBar"),
  predModePie: document.getElementById("predModePie"),
  predModeLine: document.getElementById("predModeLine"),
  predictionSubtitle: document.getElementById("predictionSubtitle"),
  predictionChartCanvas: document.getElementById("predictionChart"),
  socialRadarChartCanvas: document.getElementById("socialRadarChart"),

  // Police Deployment
  policeDonutCanvas: document.getElementById("policeDonut"),
  policeMeta: document.getElementById("policeMeta")
};

function darkenHexColor(hex, factor = CHORO_FILL_DARKEN) {
  if (!hex || typeof hex !== "string") return hex;
  let h = hex.trim().replace("#", "");
  if (h.length === 3) h = [...h].map((c) => c + c).join("");
  if (h.length !== 6) return hex;
  const scale = (x) => Math.min(255, Math.max(0, Math.round(parseInt(x, 16) * factor)));
  return `rgb(${scale(h.slice(0, 2))},${scale(h.slice(2, 4))},${scale(h.slice(4, 6))})`;
}

function hexToRgba(hex, a) {
  if (!hex || typeof hex !== "string") return `rgba(0,0,0,${a})`;
  let h = hex.trim().replace("#", "");
  if (h.length === 3) h = [...h].map((c) => c + c).join("");
  if (h.length !== 6) return `rgba(0,0,0,${a})`;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return `rgba(${r},${g},${b},${a})`;
}

const state = {
  granularity: "monthly", // monthly | yearly
  selectedStartKey: null,
  selectedEndKey: null,

  predictionMode: "bar", // bar | pie | line

  toggles: {
    crime: true,
    hubs: false,
    electronics: false,
    ggphiSim: false
  },

  /** G-GPHI strategic flow (aligned with 模拟器/<!DOCTYPE html>.html) */
  ggphiSim: {
    trust: 60,
    sustain: 40,
    chatter: 20,
    market: 10,
    /** Cumulative response inputs for suppression model (0–1+ scale, clamped on click) */
    patrolForce: 0,
    vruIntervention: 0,
    cptedForce: 0
  },

  filters: {
    areaType: "Borough",
    boroughSnt: "ALL",
    areaName: "ALL"
  },

  // Crime pre-aggregations
  crime: {
    boroughCoords: new Map(), // borough -> { lat, lon, areaCode }
    boroughMeta: new Map(), // borough -> { areaType, boroughSnt, areaName }
    boroughBoundaries: new Map(), // normalized borough -> geometry
    areaCodeToBorough: new Map(), // areaCode -> borough

    monthKeysSorted: [],
    yearKeysSorted: [],

    // periodKey -> total
    monthTotal: new Map(),
    yearTotal: new Map(),

    // periodKey -> Map(borough -> offences)
    monthByBorough: new Map(),
    yearByBorough: new Map(),
    // Split by Area Type (Borough vs Safer Neighbourhood Teams) — keys are Area Name, matching coords/meta
    monthByBoroughBorough: new Map(),
    yearByBoroughBorough: new Map(),
    monthByBoroughSnt: new Map(),
    yearByBoroughSnt: new Map(),

    // periodKey -> Map(areaCode -> offences)
    monthByAreaCode: new Map(),
    yearByAreaCode: new Map(),

    // periodKey -> Map(subtype -> offences)
    monthBySubtype: new Map(),
    yearBySubtype: new Map(),

    // periodKey -> Map(borough -> Map(subtype -> offences))
    monthByBoroughSubtype: new Map(),
    yearByBoroughSubtype: new Map(),
    monthByBoroughSubtypeBorough: new Map(),
    yearByBoroughSubtypeBorough: new Map(),
    monthByBoroughSubtypeSnt: new Map(),
    yearByBoroughSubtypeSnt: new Map()
  },

  // NEET factors (London series)
  neets: {
    monthToPercent: new Map(), // monthKey -> percent (e.g. 14.5)
    yearToPercent: new Map(), // year -> percent
    lastKnownMonthKey: null
  },

  // Police strengths
  police: {
    monthToStrength: new Map(), // monthKey -> { officer, staff, pcso }
    yearToStrength: new Map() // year -> { officer, staff, pcso } (averaged)
  },

  /** Borough unemployment (%): Σ Count semantics N/A — one value per borough from T72 CSV. */
  unemployment: {
    loaded: false,
    ratesByKey: new Map(), // normalizeBoroughName key -> rate %
    londonAvg: 0,
    miniMap: null,
    geoLayer: null,
    resizeObserver: null
  },

  // Map
  map: null,
  choroplethLayer: null,
  choroplethByBorough: new Map(),
  markersByBorough: new Map(),
  hubsLayer: null,
  electronicsLayer: null,
  /** ILP / outcome-driven risk hotspots (🔪 pulses in simulation) */
  ggphiIntelPulseLayer: null,
  /** Response actions: 🚓 / 🔦 / 🏥 pulses (2 per click) */
  ggphiResponsePulseLayer: null,
  /** Cache signature to avoid redundant intel hotspot rebuilds */
  ggphiPulseSignature: "",

  // Charts
  charts: {
    kpiLine: null,
    kpiPie: null,
    prediction: null,
    socialRadar: null,
    policeDonut: null
  }
};

init();

async function init() {
  buildMap();
  wireUI();
  // If a self-contained data bundle is available, we can render offline.
  if (initFromBundleIfAvailable()) return;
  if (window.location.protocol === "file:") {
    // Browsers block XHR/fetch for local file pages (origin = null),
    // which prevents loading ./data/*.csv from this UI.
    hideLoading();
    // Extra safety: some browsers may ignore attribute changes when CSS applies.
    if (ui.loadingOverlay) ui.loadingOverlay.style.display = "none";
    ui.layerStatusText.textContent =
      "Local file mode detected (file://). Please start a local server (e.g., python3 -m http.server) and open http://…/index.html.";
    ui.mapSubtitle.textContent = "Start a local server to load CSV data (CORS restrictions apply on file://).";
    ui.mapLegend.innerHTML = `<div class="legend-title">CORS blocked</div><div style="color: rgba(231,240,255,0.78); font-size: 0.9rem;">file:// pages cannot load relative CSV via fetch/XHR.</div>`;
    return;
  }
  showLoading("Loading crime data…");

  try {
    await Promise.all([loadCrimeCSV(), loadNeetsXLSX(), loadPoliceCSV(), loadUnemploymentCsv()]);
    hideLoading();

    buildTimeSelectors();
    setDefaultTimeRange();
    refreshAll();
  } catch (err) {
    console.error(err);
    hideLoading();
    ui.mapLegend.innerHTML = `<div class="legend-title">Legend</div><div style="color: rgba(231,240,255,0.8); font-size: 0.9rem;">Data load failed. Run via a local server so relative paths work.</div>`;
    ui.layerStatusText.textContent = "Layer status: failed to load data.";
  }
}

// Offline / no-network mode: if data bundle is present, load from it without fetch/XHR.
function initFromBundleIfAvailable() {
  const bundle = window.__GPHI_DATA_BUNDLE__;
  if (!bundle) return false;

  // Populate state from bundle
  const c = state.crime;
  c.boroughCoords.clear();
  c.boroughMeta.clear();
  c.boroughBoundaries.clear();
  c.areaCodeToBorough.clear();
  c.monthKeysSorted = bundle.crime.monthKeysSorted || [];
  c.yearKeysSorted = bundle.crime.yearKeysSorted || [];
  c.monthTotal.clear();
  c.yearTotal.clear();
  c.monthByBorough.clear();
  c.yearByBorough.clear();
  c.monthByBoroughBorough.clear();
  c.yearByBoroughBorough.clear();
  c.monthByBoroughSnt.clear();
  c.yearByBoroughSnt.clear();
  c.monthByAreaCode.clear();
  c.yearByAreaCode.clear();
  c.monthBySubtype.clear();
  c.yearBySubtype.clear();
  c.monthByBoroughSubtype.clear();
  c.yearByBoroughSubtype.clear();
  c.monthByBoroughSubtypeBorough.clear();
  c.yearByBoroughSubtypeBorough.clear();
  c.monthByBoroughSubtypeSnt.clear();
  c.yearByBoroughSubtypeSnt.clear();

  for (const [b, v] of Object.entries(bundle.crime.boroughCoords || {})) {
    c.boroughCoords.set(b, { lat: Number(v.lat), lon: Number(v.lon) });
  }
  for (const [b, v] of Object.entries(bundle.crime.boroughMeta || {})) {
    c.boroughMeta.set(b, {
      areaType: v.areaType || "Borough",
      boroughSnt: v.boroughSnt || b,
      areaName: v.areaName || b
    });
  }
  for (const [name, geom] of Object.entries(bundle.crime.boroughBoundaries || {})) {
    c.boroughBoundaries.set(normalizeBoroughName(name), geom);
  }
  for (const [code, borough] of Object.entries(bundle.crime.areaCodeToBorough || {})) {
    c.areaCodeToBorough.set(code, borough);
  }
  for (const [k, v] of Object.entries(bundle.crime.monthTotal || {})) c.monthTotal.set(k, Number(v));
  for (const [k, v] of Object.entries(bundle.crime.yearTotal || {})) c.yearTotal.set(k, Number(v));

  const hydrateNested = (outerSrc, targetOuterMap) => {
    for (const [outerKey, innerObj] of Object.entries(outerSrc || {})) {
      const innerMap = new Map();
      for (const [innerKey, innerVal] of Object.entries(innerObj || {})) {
        const num = Number(innerVal);
        if (!Number.isFinite(num)) continue;
        innerMap.set(innerKey, num);
      }
      targetOuterMap.set(outerKey, innerMap);
    }
  };

  hydrateNested(bundle.crime.monthByBorough, c.monthByBorough);
  hydrateNested(bundle.crime.yearByBorough, c.yearByBorough);
  hydrateNested(bundle.crime.monthByBoroughBorough || bundle.crime.monthByBorough, c.monthByBoroughBorough);
  hydrateNested(bundle.crime.yearByBoroughBorough || bundle.crime.yearByBorough, c.yearByBoroughBorough);
  hydrateNested(bundle.crime.monthByBoroughSnt || {}, c.monthByBoroughSnt);
  hydrateNested(bundle.crime.yearByBoroughSnt || {}, c.yearByBoroughSnt);
  hydrateNested(bundle.crime.monthByAreaCode, c.monthByAreaCode);
  hydrateNested(bundle.crime.yearByAreaCode, c.yearByAreaCode);
  hydrateNested(bundle.crime.monthBySubtype, c.monthBySubtype);
  hydrateNested(bundle.crime.yearBySubtype, c.yearBySubtype);

  const hydrateBoroughSubtype = (outerSrc, targetOuterMap) => {
    for (const [periodKey, boroughObj] of Object.entries(outerSrc || {})) {
      const boroughMap = new Map();
      for (const [borough, subtypeObj] of Object.entries(boroughObj || {})) {
        const subtypeMap = new Map();
        for (const [subtype, val] of Object.entries(subtypeObj || {})) {
          const num = Number(val);
          if (!Number.isFinite(num)) continue;
          subtypeMap.set(subtype, num);
        }
        boroughMap.set(borough, subtypeMap);
      }
      targetOuterMap.set(periodKey, boroughMap);
    }
  };
  hydrateBoroughSubtype(bundle.crime.monthByBoroughSubtype, c.monthByBoroughSubtype);
  hydrateBoroughSubtype(bundle.crime.yearByBoroughSubtype, c.yearByBoroughSubtype);
  hydrateBoroughSubtype(
    bundle.crime.monthByBoroughSubtypeBorough || bundle.crime.monthByBoroughSubtype,
    c.monthByBoroughSubtypeBorough
  );
  hydrateBoroughSubtype(
    bundle.crime.yearByBoroughSubtypeBorough || bundle.crime.yearByBoroughSubtype,
    c.yearByBoroughSubtypeBorough
  );
  hydrateBoroughSubtype(bundle.crime.monthByBoroughSubtypeSnt || {}, c.monthByBoroughSubtypeSnt);
  hydrateBoroughSubtype(bundle.crime.yearByBoroughSubtypeSnt || {}, c.yearByBoroughSubtypeSnt);

  state.neets.monthToPercent.clear();
  state.neets.yearToPercent.clear();
  for (const [k, v] of Object.entries(bundle.neets.monthToPercent || {})) state.neets.monthToPercent.set(k, Number(v));
  for (const [k, v] of Object.entries(bundle.neets.yearToPercent || {})) state.neets.yearToPercent.set(k, Number(v));
  state.neets.lastKnownMonthKey = bundle.neets.lastKnownMonthKey || null;

  state.police.monthToStrength.clear();
  state.police.yearToStrength.clear();
  const p = bundle.police || {};
  for (const [mk, v] of Object.entries(p.monthToStrength || {})) {
    state.police.monthToStrength.set(mk, { officer: Number(v.officer), staff: Number(v.staff), pcso: Number(v.pcso) });
  }
  for (const [yk, v] of Object.entries(p.yearToStrength || {})) {
    state.police.yearToStrength.set(yk, { officer: Number(v.officer), staff: Number(v.staff), pcso: Number(v.pcso) });
  }

  patchMissingBoroughsFromBoundaries();

  try {
    const parsed = Papa.parse(UNEMPLOYMENT_CSV_FALLBACK, {
      header: true,
      dynamicTyping: true,
      skipEmptyLines: true
    });
    ingestUnemploymentRows(parsed.data || []);
  } catch {
    state.unemployment.loaded = false;
  }

  // Now render
  hideLoading();
  buildSpatialFilters();
  buildTimeSelectors();
  setDefaultTimeRange();
  refreshAll();
  if (window.location.protocol !== "file:") {
    queueMicrotask(() => {
      loadUnemploymentCsv()
        .then(() => refreshUnemploymentMiniMap())
        .catch(() => {});
    });
  }
  return true;
}

function wireUI() {
  ui.granularityMonthly?.addEventListener("click", () => setGranularity("monthly"));
  ui.granularityYearly?.addEventListener("click", () => setGranularity("yearly"));

  ui.startTimeSelect.addEventListener("change", () => onTimeChanged());
  ui.endTimeSelect.addEventListener("change", () => onTimeChanged());

  ui.toggleCrimeCSV.addEventListener("change", () => {
    state.toggles.crime = ui.toggleCrimeCSV.checked;
    refreshAll();
  });
  ui.toggleHubs?.addEventListener("change", () => {
    state.toggles.hubs = ui.toggleHubs.checked;
    updateContextLayers();
    if (state.crime.monthKeysSorted.length || state.crime.yearKeysSorted.length) refreshAll();
  });
  ui.toggleElectronics?.addEventListener("change", () => {
    state.toggles.electronics = ui.toggleElectronics.checked;
    updateContextLayers();
    if (state.crime.monthKeysSorted.length || state.crime.yearKeysSorted.length) refreshAll();
  });

  ui.toggleGgphiSim?.addEventListener("change", () => {
    state.toggles.ggphiSim = ui.toggleGgphiSim.checked;
    if (ui.ggphiSimPanel) ui.ggphiSimPanel.hidden = !state.toggles.ggphiSim;
    if (!state.toggles.ggphiSim) clearGgphiSimMapOverlays();
    if (state.toggles.ggphiSim) {
      syncGgphiRangeLabels();
      updateGgphiOutcomeUI();
    }
    if (state.crime.monthKeysSorted.length || state.crime.yearKeysSorted.length) refreshAll();
    else if (state.toggles.ggphiSim) updateGgphiPulseHotspots();
  });

  ui.ggphiChatterInput?.addEventListener("input", () => {
    if (!state.toggles.ggphiSim) return;
    const v = Number(ui.ggphiChatterInput.value);
    state.ggphiSim.chatter = v;
    if (ui.ggphiChatterLabel) ui.ggphiChatterLabel.textContent = v > 70 ? "High" : v > 30 ? "Medium" : "Low";
    if (state.crime.monthKeysSorted.length || state.crime.yearKeysSorted.length) refreshAll();
  });

  ui.ggphiMarketInput?.addEventListener("input", () => {
    if (!state.toggles.ggphiSim) return;
    const v = Number(ui.ggphiMarketInput.value);
    state.ggphiSim.market = v;
    if (ui.ggphiMarketLabel) ui.ggphiMarketLabel.textContent = v > 70 ? "High" : v > 30 ? "Medium" : "Low";
    if (state.crime.monthKeysSorted.length || state.crime.yearKeysSorted.length) refreshAll();
  });

  ui.ggphiBtnPatrol?.addEventListener("click", () => ggphiApplyAction("patrol"));
  ui.ggphiBtnCpted?.addEventListener("click", () => ggphiApplyAction("cpted"));
  ui.ggphiBtnVru?.addEventListener("click", () => ggphiApplyAction("vru"));
  ui.ggphiResetBtn?.addEventListener("click", () => resetGgphiSimulation());

  ui.areaTypeFilter?.addEventListener("change", () => {
    state.filters.areaType = ui.areaTypeFilter.value || "ALL";
    refreshAll();
  });
  ui.boroughSntFilter?.addEventListener("change", () => {
    state.filters.boroughSnt = ui.boroughSntFilter.value || "ALL";
    refreshAll();
  });
  ui.areaNameFilter?.addEventListener("change", () => {
    state.filters.areaName = ui.areaNameFilter.value || "ALL";
    refreshAll();
  });

  ui.predModeBar.addEventListener("click", () => setPredictionMode("bar"));
  ui.predModePie.addEventListener("click", () => setPredictionMode("pie"));
  ui.predModeLine.addEventListener("click", () => setPredictionMode("line"));
}

function setPredictionMode(mode) {
  state.predictionMode = mode;
  ui.predModeBar.classList.toggle("is-active", mode === "bar");
  ui.predModePie.classList.toggle("is-active", mode === "pie");
  ui.predModeLine.classList.toggle("is-active", mode === "line");
  refreshAll();
}

function setGranularity(granularity) {
  if (state.granularity === granularity) return;
  state.granularity = granularity;

  if (ui.granularityMonthly) ui.granularityMonthly.classList.toggle("is-active", granularity === "monthly");
  if (ui.granularityYearly) ui.granularityYearly.classList.toggle("is-active", granularity === "yearly");

  buildTimeSelectors();
  setDefaultTimeRange();
  refreshAll();
}

function onTimeChanged() {
  const start = ui.startTimeSelect.value || null;
  const end = ui.endTimeSelect.value || null;
  if (!start || !end) return;

  const keys = state.granularity === "monthly" ? state.crime.monthKeysSorted : state.crime.yearKeysSorted;
  const startIdx = keys.indexOf(start);
  const endIdx = keys.indexOf(end);
  if (startIdx === -1 || endIdx === -1) return;
  if (startIdx > endIdx) {
    ui.endTimeSelect.value = start; // keep valid range
    state.selectedStartKey = start;
    state.selectedEndKey = start;
  } else {
    state.selectedStartKey = start;
    state.selectedEndKey = end;
  }

  refreshAll();
}

function buildTimeSelectors() {
  const keys = state.granularity === "monthly" ? state.crime.monthKeysSorted : state.crime.yearKeysSorted;
  ui.startTimeSelect.innerHTML = "";
  ui.endTimeSelect.innerHTML = "";

  for (const k of keys) {
    const labelMonth = k.slice(0, 7); // YYYY-MM
    addOption(ui.startTimeSelect, k, state.granularity === "monthly" ? labelMonth : String(k));
    addOption(ui.endTimeSelect, k, state.granularity === "monthly" ? labelMonth : String(k));
  }
}

function setDefaultTimeRange() {
  const keys = state.granularity === "monthly" ? state.crime.monthKeysSorted : state.crime.yearKeysSorted;
  if (!keys.length) return;
  state.selectedStartKey = keys[0];
  state.selectedEndKey = keys[keys.length - 1];
  ui.startTimeSelect.value = state.selectedStartKey;
  ui.endTimeSelect.value = state.selectedEndKey;

  ui.dateRange.textContent = `Date range: ${formatKey(state.selectedStartKey)} → ${formatKey(state.selectedEndKey)}`;
}

function addOption(selectEl, value, label) {
  const op = document.createElement("option");
  op.value = value;
  op.textContent = label;
  selectEl.appendChild(op);
}

function buildSpatialFilters() {
  if (!ui.areaTypeFilter || !ui.boroughSntFilter || !ui.areaNameFilter) return;
  ui.areaTypeFilter.innerHTML = "";
  ui.boroughSntFilter.innerHTML = "";
  ui.areaNameFilter.innerHTML = "";

  addOption(ui.areaTypeFilter, "ALL", "All Area Types");
  addOption(ui.boroughSntFilter, "ALL", "All Borough_SNT");
  addOption(ui.areaNameFilter, "ALL", "All Area Names");

  const areaTypes = new Set();
  const boroughSnts = new Set();
  const areaNames = new Set();
  for (const [borough, meta] of state.crime.boroughMeta.entries()) {
    areaTypes.add(meta.areaType || "Borough");
    boroughSnts.add(meta.boroughSnt || borough);
    areaNames.add(meta.areaName || borough);
  }

  [...areaTypes].sort().forEach((v) => addOption(ui.areaTypeFilter, v, v));
  [...boroughSnts].sort().forEach((v) => addOption(ui.boroughSntFilter, v, v));
  [...areaNames].sort().forEach((v) => addOption(ui.areaNameFilter, v, v));

  if ([...areaTypes].includes(state.filters.areaType)) {
    ui.areaTypeFilter.value = state.filters.areaType;
  } else {
    state.filters.areaType = "ALL";
    ui.areaTypeFilter.value = "ALL";
  }
  ui.boroughSntFilter.value = state.filters.boroughSnt;
  ui.areaNameFilter.value = state.filters.areaName;
}

function getAllowedBoroughSet() {
  const out = new Set();
  for (const [borough, meta] of state.crime.boroughMeta.entries()) {
    if (state.filters.areaType !== "ALL" && (meta.areaType || "Borough") !== state.filters.areaType) continue;
    if (state.filters.boroughSnt !== "ALL" && (meta.boroughSnt || borough) !== state.filters.boroughSnt) continue;
    if (state.filters.areaName !== "ALL" && (meta.areaName || borough) !== state.filters.areaName) continue;
    out.add(borough);
  }
  // Fallback if meta is empty
  if (!out.size) {
    for (const b of state.crime.boroughCoords.keys()) out.add(b);
  }
  return out;
}

function buildMap() {
  state.map = L.map("map", { zoomControl: true }).setView([51.5072, -0.1276], 10);
  if (window.__GPHI_STANDALONE__) {
    state.map.getContainer().style.background = "#d8e4f2";
    const ms = document.getElementById("mapSubtitle");
    if (ms && !String(ms.dataset.offlineNote || "").includes("1")) {
      ms.textContent = `${ms.textContent} • Offline: no basemap tiles`;
      ms.dataset.offlineNote = "1";
    }
  } else {
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      maxZoom: 19,
      attribution: "&copy; OpenStreetMap contributors"
    }).addTo(state.map);
  }

  // Ensure layer order: choropleth polygons below symbol circles.
  state.map.createPane("choroplethPane");
  state.map.getPane("choroplethPane").style.zIndex = 380;
  state.map.createPane("symbolPane");
  state.map.getPane("symbolPane").style.zIndex = 430;
  state.map.createPane("ggphiPulsePane");
  const pulsePane = state.map.getPane("ggphiPulsePane");
  pulsePane.style.zIndex = 680;
  pulsePane.style.pointerEvents = "none";
  state.choroplethLayer = L.layerGroup().addTo(state.map);

  // Contextual point layers (created once; visibility controlled by toggles)
  state.hubsLayer = L.layerGroup();
  state.electronicsLayer = L.layerGroup();
  state.ggphiIntelPulseLayer = L.layerGroup();
  state.ggphiResponsePulseLayer = L.layerGroup();
  updateContextLayers();

  // Stable markers: create once, update styles only (prevents flashing).
  const boroughCoords = state.crime.boroughCoords; // empty until data load
  void boroughCoords;
}

const LAYER_TYPE_HUBS = "High-traffic Transportation Hubs";
const LAYER_TYPE_ELECTRONICS = "Electronics Market";

const CONTEXT_HUBS = [
  { name: "King's Cross / St Pancras", lat: 51.5308, lon: -0.1238, level: "High" },
  { name: "Oxford Circus", lat: 51.5154, lon: -0.1410, level: "High" },
  { name: "Victoria Station", lat: 51.4965, lon: -0.1447, level: "Medium" },
  { name: "Stratford Interchange", lat: 51.5413, lon: -0.0030, level: "High" }
];

const CONTEXT_ELECTRONICS = [
  { name: "Tottenham Court Road Cluster", lat: 51.5201, lon: -0.1324 },
  { name: "Edgware Road Cluster", lat: 51.5199, lon: -0.1700 },
  { name: "Westfield Stratford Electronics", lat: 51.5440, lon: -0.0090 }
];

/** Faster animation when incident pressure is higher (ILP − response). */
function makeGgphiIntelPulseIcon(durationSec) {
  const d = clamp(durationSec, 0.28, 1.2);
  return L.divIcon({
    className: "ggphi-pulse-divicon",
    html: `<div class="ggphi-pulse-emoji ggphi-pulse-emoji--intel" style="animation-duration:${d}s"><span class="ggphi-emoji-char ggphi-emoji-char--intel">🔪</span></div>`,
    iconSize: [40, 40],
    iconAnchor: [20, 20]
  });
}

const GGPHI_EMOJI_BY_KIND = { patrol: "🚓", cpted: "🔦", vru: "🏥" };
const ggphiEmojiIconByKind = {};

function makeGgphiEmojiPulseIcon(kind) {
  if (!ggphiEmojiIconByKind[kind]) {
    const emoji = GGPHI_EMOJI_BY_KIND[kind] || "•";
    ggphiEmojiIconByKind[kind] = L.divIcon({
      className: "ggphi-pulse-divicon",
      html: `<div class="ggphi-pulse-emoji ggphi-pulse-emoji--${kind}"><span class="ggphi-emoji-char">${emoji}</span></div>`,
      iconSize: [40, 40],
      iconAnchor: [20, 20]
    });
  }
  return ggphiEmojiIconByKind[kind];
}

/** Aligned with range labels: Low = ≤30, Medium = 31–70, High = >70 */
const GGPHI_ILP_LOW_MAX = 30;

/**
 * G-GPHI pressure model (0–1 scale): intelligence = social × resale; suppression from cumulative responses.
 * Rule 1: patrol, CPTED, and PCSO/victim concern all reduce residual incident pressure (equal weight per unit force).
 */
function computeGgphiPressureModel() {
  const s = state.ggphiSim;
  const socialSignal = clamp(s.chatter / 100, 0, 1);
  const illicitResale = clamp(s.market / 100, 0, 1);
  const intelligencePressure = socialSignal * illicitResale;
  const totalForce = s.patrolForce + s.vruIntervention + s.cptedForce;
  const responseSuppression = totalForce * 0.58;
  const incidentPressure = Math.max(0, intelligencePressure - responseSuppression);
  return { intelligencePressure, responseSuppression, incidentPressure };
}

function isGgphiIlpBothLow() {
  return state.ggphiSim.chatter <= GGPHI_ILP_LOW_MAX && state.ggphiSim.market <= GGPHI_ILP_LOW_MAX;
}

/**
 * Map pulse count: when both ILP sliders are Low, baseline is 6 (scaled down as suppression removes residual pressure).
 * Otherwise scales with incident pressure.
 */
function getGgphiIntelHotspotCount() {
  const { intelligencePressure, incidentPressure } = computeGgphiPressureModel();
  if (intelligencePressure < 1e-9) return 0;

  if (isGgphiIlpBothLow()) {
    const ratio = clamp(incidentPressure / intelligencePressure, 0, 1);
    const n = Math.round(6 * ratio);
    return clamp(n, 0, 6);
  }

  if (intelligencePressure < 0.015) return 0;
  return clamp(Math.round(2 + incidentPressure * 44), 1, 24);
}

/** Rule 2: patrol force reduces trust; PCSO concern increases trust; CPTED neutral. */
function computeGgphiDerivedTrust() {
  const s = state.ggphiSim;
  const base = 60;
  return clamp(base - 40 * s.patrolForce + 34 * s.vruIntervention, 10, 100);
}

/**
 * Rule 3: CPTED and PCSO concern raise sustainability; extra boost when both are used together.
 * Patrol does not enter this formula.
 */
function computeGgphiDerivedSustain() {
  const s = state.ggphiSim;
  const c = s.cptedForce;
  const v = s.vruIntervention;
  const base = 40;
  return clamp(base + 22 * c + 22 * v + 44 * c * v, 10, 100);
}

function ggphiIntelPulseDurationSec() {
  const { intelligencePressure, incidentPressure } = computeGgphiPressureModel();
  const stress = intelligencePressure > 1e-6 ? clamp(incidentPressure / intelligencePressure, 0, 1) : 1;
  return 0.32 + (1 - stress) * 0.82;
}

function clearGgphiSimMapOverlays() {
  if (!state.map) return;
  if (state.ggphiIntelPulseLayer) {
    state.ggphiIntelPulseLayer.clearLayers();
    if (state.map.hasLayer(state.ggphiIntelPulseLayer)) state.map.removeLayer(state.ggphiIntelPulseLayer);
  }
  if (state.ggphiResponsePulseLayer) {
    state.ggphiResponsePulseLayer.clearLayers();
    if (state.map.hasLayer(state.ggphiResponsePulseLayer)) state.map.removeLayer(state.ggphiResponsePulseLayer);
  }
  state.ggphiPulseSignature = "";
}

function pulseGgphiResponseLegend(kind) {
  const map = { patrol: ui.ggphiLegPatrol, cpted: ui.ggphiLegCpted, vru: ui.ggphiLegVru };
  const el = map[kind];
  if (!el) return;
  el.classList.remove("is-pulsing");
  void el.offsetWidth;
  el.classList.add("is-pulsing");
  window.setTimeout(() => el.classList.remove("is-pulsing"), 1400);
}

/** Two map pulses per click; legend icon pulses in sync. */
function spawnGgphiResponseMarkers(kind) {
  if (!state.map || !state.toggles.ggphiSim || !state.ggphiResponsePulseLayer) return;
  const icon = makeGgphiEmojiPulseIcon(kind);
  if (!state.map.hasLayer(state.ggphiResponsePulseLayer)) state.ggphiResponsePulseLayer.addTo(state.map);
  for (let i = 0; i < 2; i++) {
    const lat = 51.45 + Math.random() * 0.1;
    const lng = -0.2 + Math.random() * 0.2;
    L.marker([lat, lng], {
      icon,
      interactive: false,
      pane: "ggphiPulsePane",
      keyboard: false,
      zIndexOffset: 950
    }).addTo(state.ggphiResponsePulseLayer);
  }
}

/** ILP / outcome hotspots (🔪) — only when simulation switch is ON. Does not clear response markers. */
function updateGgphiPulseHotspots() {
  if (!state.map || !state.ggphiIntelPulseLayer) return;

  if (!state.toggles.ggphiSim) {
    clearGgphiSimMapOverlays();
    return;
  }

  const s = state.ggphiSim;
  const count = getGgphiIntelHotspotCount();
  const dur = ggphiIntelPulseDurationSec();
  const sig = `${count}|${s.chatter}|${s.market}|${s.patrolForce}|${s.vruIntervention}|${s.cptedForce}|${dur.toFixed(3)}`;

  if (sig === state.ggphiPulseSignature && state.ggphiIntelPulseLayer.getLayers().length === count) {
    if (state.ggphiResponsePulseLayer.getLayers().length && !state.map.hasLayer(state.ggphiResponsePulseLayer)) {
      state.ggphiResponsePulseLayer.addTo(state.map);
    }
    return;
  }

  state.ggphiIntelPulseLayer.clearLayers();
  state.ggphiPulseSignature = sig;
  if (!state.map.hasLayer(state.ggphiIntelPulseLayer)) state.ggphiIntelPulseLayer.addTo(state.map);

  const icon = makeGgphiIntelPulseIcon(dur);
  for (let i = 0; i < count; i++) {
    const lat = 51.45 + Math.random() * 0.1;
    const lng = -0.2 + Math.random() * 0.2;
    L.marker([lat, lng], {
      icon,
      interactive: false,
      pane: "ggphiPulsePane",
      keyboard: false,
      zIndexOffset: 800
    }).addTo(state.ggphiIntelPulseLayer);
  }

  if (state.ggphiResponsePulseLayer.getLayers().length && !state.map.hasLayer(state.ggphiResponsePulseLayer)) {
    state.ggphiResponsePulseLayer.addTo(state.map);
  }
}

function updateContextLayers() {
  if (!state.map || !state.hubsLayer || !state.electronicsLayer) return;

  state.hubsLayer.clearLayers();
  state.electronicsLayer.clearLayers();

  if (state.toggles.hubs) {
    for (const h of CONTEXT_HUBS) {
      const marker = L.circleMarker([h.lat, h.lon], {
        radius: 7,
        color: PALETTE.hubMarkerStroke,
        fillColor: PALETTE.hubMarker,
        fillOpacity: 0.95,
        weight: 2,
        pane: "symbolPane"
      }).bindTooltip(
        `<div style="font-size:0.74rem; font-weight:700; color: rgba(118,182,228,0.98); margin-bottom:0.2rem;">Layer: ${escapeHtml(LAYER_TYPE_HUBS)}</div>
         <div style="font-weight:700; margin-bottom:0.15rem;">${escapeHtml(h.name)}</div>
         <div style="font-size:0.82rem; color: rgba(248,246,246,0.9);">Throughput: ${h.level}</div>`,
        { direction: "top" }
      );
      marker.addTo(state.hubsLayer);
    }
    if (!state.map.hasLayer(state.hubsLayer)) state.map.addLayer(state.hubsLayer);
  } else if (state.map.hasLayer(state.hubsLayer)) {
    state.map.removeLayer(state.hubsLayer);
  }

  if (state.toggles.electronics) {
    for (const e of CONTEXT_ELECTRONICS) {
      const marker = L.circleMarker([e.lat, e.lon], {
        radius: 7,
        color: PALETTE.electronicsMarkerStroke,
        fillColor: PALETTE.electronicsMarker,
        fillOpacity: 0.95,
        weight: 2,
        pane: "symbolPane"
      }).bindTooltip(
        `<div style="font-size:0.74rem; font-weight:700; color: rgba(224,64,251,0.98); margin-bottom:0.2rem;">Layer: ${escapeHtml(LAYER_TYPE_ELECTRONICS)}</div>
         <div style="font-weight:700; margin-bottom:0.15rem;">${escapeHtml(e.name)}</div>
         <div style="font-size:0.82rem; color: rgba(248,246,246,0.9);">Electronics market hotspot</div>`,
        { direction: "top" }
      );
      marker.addTo(state.electronicsLayer);
    }
    if (!state.map.hasLayer(state.electronicsLayer)) state.map.addLayer(state.electronicsLayer);
  } else if (state.map.hasLayer(state.electronicsLayer)) {
    state.map.removeLayer(state.electronicsLayer);
  }
}

function ensureMarkersCreated() {
  if (!state.map) return;
  if (state.markersByBorough.size) return;

  for (const [borough, coord] of state.crime.boroughCoords.entries()) {
    const m = L.circleMarker([coord.lat, coord.lon], {
      radius: 10,
      color: "rgba(255,255,255,0.7)",
      weight: 1,
      fillColor: "rgba(34,197,94,0.9)",
      fillOpacity: 0.0, // start hidden until crime toggle refresh
      pane: "symbolPane"
    });
    m.addTo(state.map);
    state.markersByBorough.set(borough, m);
  }

  // add popup content lazily via click
  for (const [borough, marker] of state.markersByBorough.entries()) {
    marker.on("click", () => {
      const info = state._lastBoroughAgg?.get(borough);
      if (!info) {
        marker.bindPopup(`<strong>${escapeHtml(borough)}</strong>`).openPopup();
        return;
      }
      const html = `
        <div style="font-weight:900; margin-bottom: 0.25rem;">${escapeHtml(borough)}</div>
        <div style="color: rgba(231,240,255,0.85); font-size: 0.88rem;">Incidents: ${formatCompactNumber(info.total)}</div>
        <div style="color: rgba(231,240,255,0.78); font-size: 0.86rem;">Dominant type: ${escapeHtml(info.dominantSubtype || "N/A")}</div>
        <div style="color: rgba(231,240,255,0.78); font-size: 0.86rem;">Area Type: ${escapeHtml(info.areaType || "N/A")} • Borough_SNT: ${escapeHtml(info.boroughSnt || borough)}</div>
        <div style="color: rgba(231,240,255,0.78); font-size: 0.86rem;">Risk level: ${info.riskLevel || "N/A"}</div>
        <div style="color: rgba(231,240,255,0.78); font-size: 0.86rem;">NEET index: ${info.neetsIndex || "N/A"}</div>
        <div style="color: rgba(231,240,255,0.7); font-size: 0.86rem; margin-top: 0.2rem;">Δ vs prior: ${formatDeltaPercent(info.delta)}</div>
      `;
      marker.bindPopup(html).openPopup();
    });
  }
}

function ensureChoroplethCreated() {
  if (!state.map || !state.choroplethLayer) return;
  if (state.choroplethByBorough.size) return;
  // Prefer real borough polygons.
  if (state.crime.boroughBoundaries.size) {
    for (const [borough] of state.crime.boroughCoords.entries()) {
      const geom = state.crime.boroughBoundaries.get(normalizeBoroughName(borough));
      if (!geom) continue;
      const layer = L.geoJSON({ type: "Feature", properties: { name: borough }, geometry: geom }, {
        pane: "choroplethPane",
        style: {
          color: "rgba(248,246,246,0.45)",
          weight: 1,
          fillColor: "rgba(34,197,94,0.2)",
          fillOpacity: CHORO_BASE_FILL_OPACITY
        }
      });
      layer.addTo(state.choroplethLayer);
      // geoJSON returns group layer, get first polygon layer
      layer.eachLayer((sub) => {
        state.choroplethByBorough.set(borough, sub);
      });
    }
    return;
  }

  // Fallback: voronoi if polygons unavailable
  if (typeof d3 === "undefined" || !d3.Delaunay) return;
  const points = [];
  const boroughs = [];
  for (const [borough, coord] of state.crime.boroughCoords.entries()) {
    points.push([coord.lon, coord.lat]);
    boroughs.push(borough);
  }
  if (points.length < 3) return;
  const bbox = [-0.55, 51.28, 0.35, 51.72];
  const delaunay = d3.Delaunay.from(points);
  const vor = delaunay.voronoi(bbox);
  for (let i = 0; i < points.length; i++) {
    const poly = vor.cellPolygon(i);
    if (!poly || poly.length < 3) continue;
    const latLngs = poly.map(([x, y]) => [y, x]);
    const borough = boroughs[i];
    const p = L.polygon(latLngs, {
      pane: "choroplethPane",
      color: "rgba(248,246,246,0.45)",
      weight: 1,
      fillColor: "rgba(34,197,94,0.2)",
      fillOpacity: CHORO_FALLBACK_FILL_OPACITY
    }).addTo(state.choroplethLayer);
    state.choroplethByBorough.set(borough, p);
  }
}

async function loadCrimeCSV() {
  showLoading("Loading crime aggregates…");

  // Reset structures
  const c = state.crime;
  c.boroughCoords.clear();
  c.boroughMeta.clear();
  c.areaCodeToBorough.clear();
  c.monthTotal.clear();
  c.yearTotal.clear();
  c.monthByBorough.clear();
  c.yearByBorough.clear();
  c.monthByBoroughBorough.clear();
  c.yearByBoroughBorough.clear();
  c.monthByBoroughSnt.clear();
  c.yearByBoroughSnt.clear();
  c.monthByAreaCode.clear();
  c.yearByAreaCode.clear();
  c.monthBySubtype.clear();
  c.yearBySubtype.clear();
  c.monthByBoroughSubtype.clear();
  c.yearByBoroughSubtype.clear();
  c.monthByBoroughSubtypeBorough.clear();
  c.yearByBoroughSubtypeBorough.clear();
  c.monthByBoroughSubtypeSnt.clear();
  c.yearByBoroughSubtypeSnt.clear();
  c.monthKeysSorted = [];
  c.yearKeysSorted = [];

  // Load pre-aggregated CSVs
  const [
    boroughCoordRows,
    boroughMetaRows,
    areaCodeMapRows,
    monthTotalRows,
    yearTotalRows,
    monthBoroughRows,
    yearBoroughRows,
    monthBoroughBoroughRows,
    yearBoroughBoroughRows,
    monthBoroughSntRows,
    yearBoroughSntRows,
    monthAreaCodeRows,
    yearAreaCodeRows,
    monthSubtypeRows,
    yearSubtypeRows,
    monthBoroughSubtypeRows,
    yearBoroughSubtypeRows,
    monthBoroughSubtypeBoroughRows,
    yearBoroughSubtypeBoroughRows,
    monthBoroughSubtypeSntRows,
    yearBoroughSubtypeSntRows
  ] = await Promise.all([
    parseCsv(CRIME_BOROUGH_COORDS_CSV_PATH),
    parseCsv(CRIME_BOROUGH_META_CSV_PATH),
    parseCsv(CRIME_AREA_CODE_MAP_CSV_PATH),
    parseCsv(CRIME_MONTH_TOTAL_CSV_PATH),
    parseCsv(CRIME_YEAR_TOTAL_CSV_PATH),
    parseCsv(CRIME_MONTH_BOROUGH_CSV_PATH),
    parseCsv(CRIME_YEAR_BOROUGH_CSV_PATH),
    parseCsv(CRIME_MONTH_BOROUGH_BOROUGH_CSV_PATH),
    parseCsv(CRIME_YEAR_BOROUGH_BOROUGH_CSV_PATH),
    parseCsv(CRIME_MONTH_BOROUGH_SNT_CSV_PATH),
    parseCsv(CRIME_YEAR_BOROUGH_SNT_CSV_PATH),
    parseCsv(CRIME_MONTH_AREA_CODE_CSV_PATH),
    parseCsv(CRIME_YEAR_AREA_CODE_CSV_PATH),
    parseCsv(CRIME_MONTH_SUBTYPE_CSV_PATH),
    parseCsv(CRIME_YEAR_SUBTYPE_CSV_PATH),
    parseCsv(CRIME_MONTH_BOROUGH_SUBTYPE_CSV_PATH),
    parseCsv(CRIME_YEAR_BOROUGH_SUBTYPE_CSV_PATH),
    parseCsv(CRIME_MONTH_BOROUGH_SUBTYPE_BOROUGH_CSV_PATH),
    parseCsv(CRIME_YEAR_BOROUGH_SUBTYPE_BOROUGH_CSV_PATH),
    parseCsv(CRIME_MONTH_BOROUGH_SUBTYPE_SNT_CSV_PATH),
    parseCsv(CRIME_YEAR_BOROUGH_SUBTYPE_SNT_CSV_PATH)
  ]);

  let totalIncidentsAll = 0;

  for (const r of boroughCoordRows) {
    const borough = r.borough != null ? String(r.borough) : "";
    const lat = Number(r.latitude);
    const lon = Number(r.longitude);
    if (!borough || !Number.isFinite(lat) || !Number.isFinite(lon)) continue;
    c.boroughCoords.set(borough, { lat, lon });
  }

  for (const r of boroughMetaRows) {
    const borough = r.areaName != null ? String(r.areaName) : "";
    if (!borough) continue;
    c.boroughMeta.set(borough, {
      areaType: r.areaType != null ? String(r.areaType) : "Borough",
      boroughSnt: r.boroughSnt != null ? String(r.boroughSnt) : borough,
      areaName: borough
    });
  }

  for (const r of areaCodeMapRows) {
    const code = r.areaCode != null ? String(r.areaCode) : "";
    const borough = r.borough != null ? String(r.borough) : "";
    if (!code || !borough) continue;
    c.areaCodeToBorough.set(code, borough);
  }

  for (const r of monthTotalRows) {
    const key = String(r.monthKey);
    const v = Number(r.totalOffences);
    if (!key || !Number.isFinite(v)) continue;
    c.monthTotal.set(key, v);
    totalIncidentsAll += v;
  }
  for (const r of yearTotalRows) {
    const key = String(r.yearKey);
    const v = Number(r.totalOffences);
    if (!key || !Number.isFinite(v)) continue;
    c.yearTotal.set(key, v);
  }

  const ensureMap = (outerMap, pk) => {
    if (!outerMap.has(pk)) outerMap.set(pk, new Map());
    return outerMap.get(pk);
  };

  for (const r of monthBoroughRows) {
    const pk = String(r.monthKey);
    const borough = r.item != null ? String(r.item) : "";
    const v = Number(r.offences);
    if (!pk || !borough || !Number.isFinite(v)) continue;
    addToMapSum(ensureMap(c.monthByBorough, pk), borough, v);
  }

  for (const r of yearBoroughRows) {
    const pk = String(r.yearKey);
    const borough = r.item != null ? String(r.item) : "";
    const v = Number(r.offences);
    if (!pk || !borough || !Number.isFinite(v)) continue;
    addToMapSum(ensureMap(c.yearByBorough, pk), borough, v);
  }

  for (const r of monthBoroughBoroughRows) {
    const pk = String(r.monthKey);
    const borough = r.item != null ? String(r.item) : "";
    const v = Number(r.offences);
    if (!pk || !borough || !Number.isFinite(v)) continue;
    addToMapSum(ensureMap(c.monthByBoroughBorough, pk), borough, v);
  }
  for (const r of yearBoroughBoroughRows) {
    const pk = String(r.yearKey);
    const borough = r.item != null ? String(r.item) : "";
    const v = Number(r.offences);
    if (!pk || !borough || !Number.isFinite(v)) continue;
    addToMapSum(ensureMap(c.yearByBoroughBorough, pk), borough, v);
  }
  for (const r of monthBoroughSntRows) {
    const pk = String(r.monthKey);
    const borough = r.item != null ? String(r.item) : "";
    const v = Number(r.offences);
    if (!pk || !borough || !Number.isFinite(v)) continue;
    addToMapSum(ensureMap(c.monthByBoroughSnt, pk), borough, v);
  }
  for (const r of yearBoroughSntRows) {
    const pk = String(r.yearKey);
    const borough = r.item != null ? String(r.item) : "";
    const v = Number(r.offences);
    if (!pk || !borough || !Number.isFinite(v)) continue;
    addToMapSum(ensureMap(c.yearByBoroughSnt, pk), borough, v);
  }

  for (const r of monthAreaCodeRows) {
    const pk = String(r.monthKey);
    const code = r.item != null ? String(r.item) : "";
    const v = Number(r.offences);
    if (!pk || !code || !Number.isFinite(v)) continue;
    addToMapSum(ensureMap(c.monthByAreaCode, pk), code, v);
  }

  for (const r of yearAreaCodeRows) {
    const pk = String(r.yearKey);
    const code = r.item != null ? String(r.item) : "";
    const v = Number(r.offences);
    if (!pk || !code || !Number.isFinite(v)) continue;
    addToMapSum(ensureMap(c.yearByAreaCode, pk), code, v);
  }

  for (const r of monthSubtypeRows) {
    const pk = String(r.monthKey);
    const subtype = r.item != null ? String(r.item) : "";
    const v = Number(r.offences);
    if (!pk || !subtype || !Number.isFinite(v)) continue;
    addToMapSum(ensureMap(c.monthBySubtype, pk), subtype, v);
  }

  for (const r of yearSubtypeRows) {
    const pk = String(r.yearKey);
    const subtype = r.item != null ? String(r.item) : "";
    const v = Number(r.offences);
    if (!pk || !subtype || !Number.isFinite(v)) continue;
    addToMapSum(ensureMap(c.yearBySubtype, pk), subtype, v);
  }

  const ensureNestedBoroughSubtype = (outerMap, periodKey, borough) => {
    if (!outerMap.has(periodKey)) outerMap.set(periodKey, new Map());
    const boroughMap = outerMap.get(periodKey);
    if (!boroughMap.has(borough)) boroughMap.set(borough, new Map());
    return boroughMap.get(borough);
  };

  for (const r of monthBoroughSubtypeRows) {
    const pk = String(r.monthKey);
    const borough = r.borough != null ? String(r.borough) : "";
    const subtype = r.subtype != null ? String(r.subtype) : "";
    const v = Number(r.offences);
    if (!pk || !borough || !subtype || !Number.isFinite(v)) continue;
    addToMapSum(ensureNestedBoroughSubtype(c.monthByBoroughSubtype, pk, borough), subtype, v);
  }

  for (const r of yearBoroughSubtypeRows) {
    const pk = String(r.yearKey);
    const borough = r.borough != null ? String(r.borough) : "";
    const subtype = r.subtype != null ? String(r.subtype) : "";
    const v = Number(r.offences);
    if (!pk || !borough || !subtype || !Number.isFinite(v)) continue;
    addToMapSum(ensureNestedBoroughSubtype(c.yearByBoroughSubtype, pk, borough), subtype, v);
  }

  for (const r of monthBoroughSubtypeBoroughRows) {
    const pk = String(r.monthKey);
    const borough = r.borough != null ? String(r.borough) : "";
    const subtype = r.subtype != null ? String(r.subtype) : "";
    const v = Number(r.offences);
    if (!pk || !borough || !subtype || !Number.isFinite(v)) continue;
    addToMapSum(ensureNestedBoroughSubtype(c.monthByBoroughSubtypeBorough, pk, borough), subtype, v);
  }
  for (const r of yearBoroughSubtypeBoroughRows) {
    const pk = String(r.yearKey);
    const borough = r.borough != null ? String(r.borough) : "";
    const subtype = r.subtype != null ? String(r.subtype) : "";
    const v = Number(r.offences);
    if (!pk || !borough || !subtype || !Number.isFinite(v)) continue;
    addToMapSum(ensureNestedBoroughSubtype(c.yearByBoroughSubtypeBorough, pk, borough), subtype, v);
  }
  for (const r of monthBoroughSubtypeSntRows) {
    const pk = String(r.monthKey);
    const borough = r.borough != null ? String(r.borough) : "";
    const subtype = r.subtype != null ? String(r.subtype) : "";
    const v = Number(r.offences);
    if (!pk || !borough || !subtype || !Number.isFinite(v)) continue;
    addToMapSum(ensureNestedBoroughSubtype(c.monthByBoroughSubtypeSnt, pk, borough), subtype, v);
  }
  for (const r of yearBoroughSubtypeSntRows) {
    const pk = String(r.yearKey);
    const borough = r.borough != null ? String(r.borough) : "";
    const subtype = r.subtype != null ? String(r.subtype) : "";
    const v = Number(r.offences);
    if (!pk || !borough || !subtype || !Number.isFinite(v)) continue;
    addToMapSum(ensureNestedBoroughSubtype(c.yearByBoroughSubtypeSnt, pk, borough), subtype, v);
  }

  c.monthKeysSorted = [...c.monthTotal.keys()].sort();
  c.yearKeysSorted = [...c.yearTotal.keys()].sort();

  ui.recordCount.textContent = `Incidents (Σ Count): ${Math.round(totalIncidentsAll).toLocaleString()}`;

  // Clear any previously created markers so styling refresh can proceed cleanly.
  state.markersByBorough.clear();
  ensureMarkersCreated();
  buildSpatialFilters();
}

function addToMapSum(map, key, value) {
  map.set(key, (map.get(key) || 0) + value);
}

async function loadNeetsXLSX() {
  showLoading("Loading NEET signals…");

  state.neets.monthToPercent = new Map();
  state.neets.yearToPercent = new Map();
  state.neets.lastKnownMonthKey = null;

  // Load pre-generated NEET csv (London only)
  const [monthRows, yearRows] = await Promise.all([parseCsv(NEETS_MONTH_CSV_PATH), parseCsv(NEETS_YEAR_CSV_PATH)]);

  for (const r of monthRows) {
    const mk = String(r.monthKey);
    const v = Number(r.neetPercent);
    if (!mk || !Number.isFinite(v)) continue;
    state.neets.monthToPercent.set(mk, v);
    if (!state.neets.lastKnownMonthKey || mk > state.neets.lastKnownMonthKey) state.neets.lastKnownMonthKey = mk;
  }

  for (const r of yearRows) {
    const y = String(r.yearKey);
    const v = Number(r.neetPercent);
    if (!y || !Number.isFinite(v)) continue;
    state.neets.yearToPercent.set(y, v);
  }
}

function toNumber(v) {
  if (v == null) return NaN;
  if (typeof v === "number") return v;
  const s = String(v).trim().replace(/,/g, "");
  if (!s) return NaN;
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
}

function quarterLabelToMonthKeys(quarterLabel) {
  // e.g. "2000 Q2"
  const m = String(quarterLabel).match(/(\d{4})\s*Q([1-4])/i);
  if (!m) return [];
  const year = Number(m[1]);
  const q = Number(m[2]);
  const startMonth = (q - 1) * 3 + 1; // 1..12

  const keys = [];
  for (let k = 0; k < 3; k++) {
    const mm = String(startMonth + k).padStart(2, "0");
    keys.push(`${year}-${mm}-01`);
  }
  return keys;
}

function getNeetsFactorForRangeEnd() {
  if (state.granularity === "monthly") {
    const end = state.selectedEndKey;
    if (!end) return null;
    if (state.neets.monthToPercent.has(end)) return state.neets.monthToPercent.get(end);

    // Carry-forward: find closest month <= end
    const keys = [...state.neets.monthToPercent.keys()].sort();
    let best = null;
    for (const k of keys) {
      if (k <= end) best = k;
    }
    if (best != null) return state.neets.monthToPercent.get(best);
    // fallback to last known
    if (state.neets.lastKnownMonthKey && state.neets.monthToPercent.has(state.neets.lastKnownMonthKey)) {
      return state.neets.monthToPercent.get(state.neets.lastKnownMonthKey);
    }
    return null;
  }

  const endYear = state.selectedEndKey;
  if (!endYear) return null;
  if (state.neets.yearToPercent.has(endYear)) return state.neets.yearToPercent.get(endYear);
  // fallback: last known
  const years = [...state.neets.yearToPercent.keys()].sort();
  let best = null;
  for (const y of years) {
    if (y <= endYear) best = y;
  }
  if (best != null) return state.neets.yearToPercent.get(best);
  const last = years[years.length - 1];
  return last != null ? state.neets.yearToPercent.get(last) : null;
}

async function loadPoliceCSV() {
  showLoading("Loading Police Force Strength…");
  const rows = await parseCsv(POLICE_CSV_PATH);

  const monthToStrength = new Map();

  for (const r of rows) {
    const dateStr = r.Date ? String(r.Date).trim() : "";
    const officer = Number(r["Police Officer Strength"]) || 0;
    const staff = Number(r["Police Staff Strength"]) || 0;
    const pcso = Number(r["PCSO Strength"]) || 0;
    const mk = policeDateToMonthKey(dateStr);
    if (!mk) continue;
    monthToStrength.set(mk, { officer, staff, pcso });
  }

  const yearToStrength = new Map();
  const yearToCount = new Map();
  for (const [mk, v] of monthToStrength.entries()) {
    const y = mk.slice(0, 4);
    if (!yearToStrength.has(y)) yearToStrength.set(y, { officer: 0, staff: 0, pcso: 0 });
    const agg = yearToStrength.get(y);
    agg.officer += v.officer;
    agg.staff += v.staff;
    agg.pcso += v.pcso;
    yearToCount.set(y, (yearToCount.get(y) || 0) + 1);
  }
  for (const [y, agg] of yearToStrength.entries()) {
    const n = Math.max(1, yearToCount.get(y) || 1);
    agg.officer /= n;
    agg.staff /= n;
    agg.pcso /= n;
  }

  state.police.monthToStrength = monthToStrength;
  state.police.yearToStrength = yearToStrength;
}

function policeDateToMonthKey(dateStr) {
  // e.g. "May-13"
  const m = String(dateStr).match(/^([A-Za-z]{3})-(\d{2})$/);
  if (!m) return null;
  const mon = m[1].toLowerCase();
  const yy = Number(m[2]);
  const year = 2000 + yy;
  const monthNames = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const idx = monthNames.indexOf(mon.slice(0, 3));
  if (idx === -1) return null;
  const mm = String(idx + 1).padStart(2, "0");
  return `${year}-${mm}-01`;
}

function parseCsv(path) {
  return new Promise((resolve, reject) => {
    Papa.parse(path, {
      download: true,
      header: true,
      dynamicTyping: true,
      skipEmptyLines: true,
      complete: ({ data }) => resolve(data),
      error: (err) => reject(err)
    });
  });
}

function stripUnemploymentBoroughLabel(s) {
  return String(s || "")
    .trim()
    .replace(/^["']|["']$/g, "")
    .replace(/^London Borough of /i, "")
    .replace(/^The Royal Borough of /i, "")
    .replace(/^Royal Borough of /i, "");
}

function unemploymentRateLookupKey(name) {
  return normalizeBoroughName(stripUnemploymentBoroughLabel(name));
}

function ingestUnemploymentRows(rows) {
  const u = state.unemployment;
  u.ratesByKey.clear();
  let londonExplicit = null;
  for (const r of rows) {
    const b = stripUnemploymentBoroughLabel(
      r.borough ?? r.Borough ?? r["Area name"] ?? r["Area Name"] ?? r.areaName ?? ""
    );
    const raw = r.unemployment_rate ?? r["Unemployment rate"] ?? r.rate ?? r.Value ?? r.value;
    if (!b) continue;
    const lk = b.toLowerCase();
    if (lk === "london" || lk.includes("london average") || lk === "greater london" || lk.includes("london mean")) {
      const v = Number(raw);
      if (Number.isFinite(v)) londonExplicit = v;
      continue;
    }
    const v = Number(raw);
    if (!Number.isFinite(v)) continue;
    u.ratesByKey.set(unemploymentRateLookupKey(b), v);
  }
  let sum = 0;
  let n = 0;
  for (const v of u.ratesByKey.values()) {
    sum += v;
    n++;
  }
  u.londonAvg = londonExplicit != null && Number.isFinite(londonExplicit) ? londonExplicit : n ? sum / n : 0;
  u.loaded = n > 0;
}

function classifyUnemploymentVsMean(rate, mean) {
  if (!Number.isFinite(rate) || !Number.isFinite(mean)) return "none";
  const d = rate - mean;
  if (d < -UNEMPLOYMENT_MEAN_BAND_PP) return "better";
  if (d > UNEMPLOYMENT_MEAN_BAND_PP) return "worse";
  return "average";
}

async function loadUnemploymentCsv() {
  const u = state.unemployment;
  try {
    const rows = await parseCsv(UNEMPLOYMENT_CSV_PATH);
    ingestUnemploymentRows(rows);
  } catch (e) {
    try {
      const parsed = Papa.parse(UNEMPLOYMENT_CSV_FALLBACK, {
        header: true,
        dynamicTyping: true,
        skipEmptyLines: true
      });
      ingestUnemploymentRows(parsed.data || []);
    } catch (e2) {
      u.loaded = false;
      throw e2;
    }
  }
}

function refreshUnemploymentMiniMap() {
  const el = document.getElementById("unemploymentMiniMap");
  const hint = document.getElementById("unemploymentMapHint");
  const leg = document.getElementById("unemploymentLegendMean");
  if (!el || typeof L === "undefined") return;

  if (!state.unemployment.loaded) {
    if (hint) hint.textContent = "No unemployment rows loaded.";
    return;
  }

  const mean = state.unemployment.londonAvg;
  if (leg) leg.textContent = `vs London mean (${mean.toFixed(1)}%)`;

  const u = state.unemployment;
  if (!u.miniMap) {
    u.miniMap = L.map("unemploymentMiniMap", {
      zoomControl: false,
      attributionControl: false,
      dragging: true,
      scrollWheelZoom: true,
      doubleClickZoom: true,
      boxZoom: true,
      keyboard: true,
      touchZoom: true,
      minZoom: 8,
      maxZoom: 18
    });
    L.control.zoom({ position: "topright" }).addTo(u.miniMap);
    if (!u.miniMap.getPane("unemploymentMiniPane")) {
      u.miniMap.createPane("unemploymentMiniPane");
      u.miniMap.getPane("unemploymentMiniPane").style.zIndex = 400;
    }
    el.classList.add("unemploymentMiniMap--ready");

    const wrap = document.getElementById("unemploymentMapWrap");
    if (wrap && typeof ResizeObserver !== "undefined" && !u.resizeObserver) {
      u.resizeObserver = new ResizeObserver(() => {
        if (state.unemployment.miniMap) state.unemployment.miniMap.invalidateSize();
      });
      u.resizeObserver.observe(wrap);
    }
  }

  if (u.geoLayer) {
    u.miniMap.removeLayer(u.geoLayer);
    u.geoLayer = null;
  }

  const features = [];
  for (const borough of state.crime.boroughCoords.keys()) {
    const meta = state.crime.boroughMeta.get(borough);
    if (meta && isSntAreaMeta(meta)) continue;

    const norm = normalizeBoroughName(borough);
    const geom = state.crime.boroughBoundaries.get(norm);
    if (!geom) continue;

    const key = unemploymentRateLookupKey(borough);
    const rate = u.ratesByKey.has(key) ? u.ratesByKey.get(key) : null;
    const cls = classifyUnemploymentVsMean(rate != null ? rate : NaN, mean);
    features.push({
      type: "Feature",
      properties: { name: borough, rate, cls },
      geometry: geom
    });
  }

  if (!features.length) {
    if (hint) hint.textContent = "Borough boundaries not loaded — cannot draw unemployment map.";
    return;
  }

  const fc = { type: "FeatureCollection", features };
  u.geoLayer = L.geoJSON(fc, {
    pane: "unemploymentMiniPane",
    style(feat) {
      const st = UNEMPLOYMENT_STYLE[feat.properties.cls] || UNEMPLOYMENT_STYLE.none;
      return { fillColor: st.fill, color: st.stroke, weight: 1, fillOpacity: 0.93 };
    },
    onEachFeature(feat, layer) {
      const b = feat.properties.name;
      const rt = feat.properties.rate;
      const pct = Number.isFinite(rt) ? `${Number(rt).toFixed(1)}%` : "No data";
      layer.bindTooltip(
        `<div style="font-weight:800;margin-bottom:0.2rem;">${escapeHtml(b)}</div><div style="font-size:0.9rem;">Unemployment rate: ${escapeHtml(
          pct
        )}</div>`,
        { direction: "top", sticky: true, opacity: 1, className: "unemploymentBoroughTooltip" }
      );
      layer.on("mouseover", () => {
        layer.setStyle({ weight: 2, color: "#2563eb" });
      });
      layer.on("mouseout", () => {
        if (state.unemployment.geoLayer) state.unemployment.geoLayer.resetStyle(layer);
      });
    }
  });
  u.geoLayer.addTo(u.miniMap);

  const b = u.geoLayer.getBounds();
  if (b.isValid()) u.miniMap.fitBounds(b, { padding: [14, 14], maxZoom: 13 });
  else u.miniMap.setView([51.507, -0.12], 10);

  window.setTimeout(() => {
    if (u.miniMap) u.miniMap.invalidateSize();
  }, 80);
}

function showLoading(msg) {
  if (!ui.loadingOverlay) return;
  ui.loadingOverlay.hidden = false;
  ui.loadingOverlay.style.display = "flex";
  if (ui.loadingMsg) ui.loadingMsg.textContent = msg;
}

function hideLoading() {
  if (!ui.loadingOverlay) return;
  ui.loadingOverlay.hidden = true;
  ui.loadingOverlay.style.display = "none";
}

function refreshAll() {
  if (!state.crime.monthKeysSorted.length && !state.crime.yearKeysSorted.length) return;

  const crimeEnabled = state.toggles.crime;

  const keys = state.granularity === "monthly" ? state.crime.monthKeysSorted : state.crime.yearKeysSorted;
  if (!keys.length || !state.selectedStartKey || !state.selectedEndKey) return;

  const startIdx = keys.indexOf(state.selectedStartKey);
  const endIdx = keys.indexOf(state.selectedEndKey);
  if (startIdx === -1 || endIdx === -1) return;

  const selectedKeys = keys.slice(startIdx, endIdx + 1);
  const n = selectedKeys.length;

  const priorEndIdx = startIdx - 1;
  const priorStartIdx = Math.max(0, priorEndIdx - n + 1);
  const priorKeys = priorEndIdx >= 0 ? keys.slice(priorStartIdx, priorEndIdx + 1) : [];

  // Update layer status (simple)
  const hubsLabel = state.toggles.hubs ? "Hubs ON" : "Hubs OFF";
  const electronicsLabel = state.toggles.electronics ? "Electronics ON" : "Electronics OFF";
  const ggphiLabel = state.toggles.ggphiSim ? "G-GPHI sim ON" : "G-GPHI sim OFF";
  ui.layerStatusText.textContent = `Layer status: ${crimeEnabled ? "Knife Crime" : "Knife Crime OFF"} • ${hubsLabel} • ${electronicsLabel} • ${ggphiLabel} • ${state.granularity}`;

  ui.dateRange.textContent = `Date range: ${formatKey(state.selectedStartKey)} → ${formatKey(state.selectedEndKey)}`;

  // Neets factor (only affects simulations; map still primarily driven by incident rates)
  const neetsFactor = getNeetsFactorForRangeEnd();
  const allowedBoroughs = getAllowedBoroughSet();

  // KPIs / tables respect spatial filters
  const boroughAggSelected = crimeEnabled ? aggregateByGeoKey(selectedKeys, allowedBoroughs) : new Map();
  const boroughAggPrior = crimeEnabled ? aggregateByGeoKey(priorKeys, allowedBoroughs) : new Map();
  const boroughSubtypeSelected = crimeEnabled ? aggregateBoroughSubtypeGeo(selectedKeys, allowedBoroughs) : new Map();

  // Map choropleth: full geo key set (all coords) so filtered KPIs can still show London-wide rate context
  const boroughAggMap = crimeEnabled ? aggregateByGeoKey(selectedKeys, null) : new Map();
  const boroughAggPriorMap = crimeEnabled ? aggregateByGeoKey(priorKeys, null) : new Map();
  const boroughSubtypeMap = crimeEnabled ? aggregateBoroughSubtypeGeo(selectedKeys, null) : new Map();

  state._lastBoroughAgg = computeBoroughRisk(boroughAggMap, boroughAggPriorMap);

  updateMapMarkers(state._lastBoroughAgg, boroughSubtypeMap, selectedKeys, neetsFactor);

  // KPIs
  updateKPIs(boroughAggSelected, boroughAggPrior, boroughSubtypeSelected, selectedKeys, priorKeys);

  // Charts (kpi charts) — same geography as KPI row when filters are applied
  updateKpiCharts(selectedKeys, priorKeys, allowedBoroughs);

  // Predicted Intelligence
  updatePredictedIntelligence(selectedKeys, priorKeys, neetsFactor, allowedBoroughs);

  // Police deployment
  updatePoliceDeployment(selectedKeys);

  updateGgphiPulseHotspots();
  if (state.toggles.ggphiSim) updateGgphiOutcomeUI();

  refreshUnemploymentMiniMap();
}

function formatKey(k) {
  if (state.granularity === "monthly") return String(k).slice(0, 7);
  return String(k);
}

function isSntAreaMeta(meta) {
  const t = (meta?.areaType || "").toLowerCase();
  return t.includes("safer") && t.includes("neighbourhood");
}

/**
 * Sum incidents (Σ Count / offences) per map point using Borough-series or SNT-series to match Area Type.
 */
function aggregateByGeoKey(periodKeys, allowedBoroughs = null) {
  const monthly = state.granularity === "monthly";
  const byPeriodB = monthly ? state.crime.monthByBoroughBorough : state.crime.yearByBoroughBorough;
  const byPeriodS = monthly ? state.crime.monthByBoroughSnt : state.crime.yearByBoroughSnt;
  const legacy = monthly ? state.crime.monthByBorough : state.crime.yearByBorough;
  const hasSplit =
    (byPeriodB && byPeriodB.size > 0) || (byPeriodS && byPeriodS.size > 0);
  const fbB = hasSplit && byPeriodB && byPeriodB.size > 0 ? byPeriodB : legacy;
  const fbS = hasSplit && byPeriodS && byPeriodS.size > 0 ? byPeriodS : new Map();

  const out = new Map();
  for (const borough of state.crime.boroughCoords.keys()) {
    if (allowedBoroughs && !allowedBoroughs.has(borough)) continue;
    const meta = state.crime.boroughMeta.get(borough);
    const byPeriod = isSntAreaMeta(meta) ? fbS : fbB;
    let sum = 0;
    for (const pk of periodKeys) {
      const m = byPeriod.get(pk);
      if (m) sum += Number(m.get(borough)) || 0;
    }
    out.set(borough, sum);
  }
  return out;
}

function aggregateBoroughSubtypeGeo(periodKeys, allowedBoroughs = null) {
  const monthly = state.granularity === "monthly";
  const byPeriodB = monthly ? state.crime.monthByBoroughSubtypeBorough : state.crime.yearByBoroughSubtypeBorough;
  const byPeriodS = monthly ? state.crime.monthByBoroughSubtypeSnt : state.crime.yearByBoroughSubtypeSnt;
  const legacy = monthly ? state.crime.monthByBoroughSubtype : state.crime.yearByBoroughSubtype;
  const hasSplit =
    (byPeriodB && byPeriodB.size > 0) || (byPeriodS && byPeriodS.size > 0);
  const fbB = hasSplit && byPeriodB && byPeriodB.size > 0 ? byPeriodB : legacy;
  const fbS = hasSplit && byPeriodS && byPeriodS.size > 0 ? byPeriodS : new Map();

  const out = new Map();
  for (const borough of state.crime.boroughCoords.keys()) {
    if (allowedBoroughs && !allowedBoroughs.has(borough)) continue;
    const meta = state.crime.boroughMeta.get(borough);
    const byPeriod = isSntAreaMeta(meta) ? fbS : fbB;
    const targetSubtypeMap = new Map();
    for (const pk of periodKeys) {
      const boroughMap = byPeriod.get(pk);
      if (!boroughMap) continue;
      const subtypeMap = boroughMap.get(borough);
      if (!subtypeMap) continue;
      for (const [subtype, v] of subtypeMap.entries()) {
        addToMapSum(targetSubtypeMap, subtype, v);
      }
    }
    out.set(borough, targetSubtypeMap);
  }
  return out;
}

function computeBoroughRisk(boroughAggSelected, boroughAggPrior) {
  const riskMap = new Map();
  const selectedEntries = [...boroughAggSelected.entries()];
  for (const [borough, total] of selectedEntries) {
    const prev = boroughAggPrior.get(borough) || 0;
    const delta = (total - prev) / Math.max(1, prev);
    riskMap.set(borough, { total, prev, delta });
  }
  return riskMap;
}

function updateMapMarkers(boroughRiskMap, boroughSubtypeSelected, selectedKeys, neetsFactor) {
  ensureChoroplethCreated();
  ensureMarkersCreated();
  const markers = state.markersByBorough;

  const crimeEnabled = state.toggles.crime;
  const values = [...boroughRiskMap.values()].map((x) => x.total);
  const min = values.length ? Math.min(...values) : 0;
  const max = values.length ? Math.max(...values) : 1;
  const denom = Math.max(1, max - min);

  const positiveVals = values.filter((v) => v > 0);
  const pLow = positiveVals.length ? percentile(positiveVals, 0.33) : min;
  const pHigh = positiveVals.length ? percentile(positiveVals, 0.66) : max;

  // optional multiplier to slightly adjust intensity when NEETs enabled
  const neetsMult = neetsFactor == null ? 1 : clamp(0.85 + (neetFactorScaled(neetsFactor) * 0.3), 0.7, 1.15);

  // Update styles without recreating markers (prevents flicker)
  for (const [borough, marker] of markers.entries()) {
    const info = boroughRiskMap.get(borough);
    const poly = state.choroplethByBorough.get(borough);
    if (!crimeEnabled || !info) {
      marker.setStyle({ fillOpacity: 0.0 });
      marker.setRadius(6);
      if (poly) poly.setStyle({ fillOpacity: 0.0, opacity: 0.0 });
      continue;
    }

    // classify risk level by total incidents
    let riskLevel = "Low";
    if (info.total >= pHigh) riskLevel = "High";
    else if (info.total >= pLow) riskLevel = "Medium";
    info.riskLevel = riskLevel;
    const meta = state.crime.boroughMeta.get(borough) || { areaType: "Borough", boroughSnt: borough, areaName: borough };
    info.areaType = meta.areaType || "Borough";
    info.boroughSnt = meta.boroughSnt || borough;

    // dominant subtype per borough in selected window
    const subtypeMap = boroughSubtypeSelected.get(borough) || new Map();
    let dominantSubtype = "N/A";
    let dominantSubtypeShare = 0;
    if (subtypeMap.size && info.total > 0) {
      const topSubtype = [...subtypeMap.entries()].sort((a, b) => b[1] - a[1])[0];
      if (topSubtype) {
        dominantSubtype = compressSubtypeLabel(topSubtype[0]);
        dominantSubtypeShare = topSubtype[1] / info.total;
      }
    }
    info.dominantSubtype = dominantSubtype;
    info.dominantSubtypeShare = dominantSubtypeShare;

    // NEET index is an auxiliary contextual indicator (global London factor)
    const neetsIdx = neetsFactor == null ? "N/A" : neetIndexLabel(neetsFactor);
    info.neetsIndex = neetsIdx;

    const score = info.total * neetsMult;
    const t = (score - min * neetsMult) / (denom * neetsMult);
    const tt = clamp(t, 0, 1);
    const areaTypeFactor = info.areaType === "Borough" ? 1.0 : 0.75;
    const radius = (6 + tt * 24) * areaTypeFactor;

    // stepped color palette for clearer distinction
    let fill;
    if (riskLevel === "High") fill = PALETTE.red1;
    else if (riskLevel === "Medium") fill = PALETTE.warm2;
    else fill = PALETTE.green1;

    marker.setStyle({
      fillColor: fill,
      color: fill,
      opacity: 1,
      fillOpacity: 0.62
    });
    marker.setRadius(radius);

    if (poly) {
      poly.setStyle({
        fillColor: darkenHexColor(fill, CHORO_FILL_DARKEN),
        fillOpacity: CHORO_ACTIVE_FILL_OPACITY,
        color: "rgba(248,246,246,0.42)",
        opacity: 0.9,
        weight: 1
      });
    }

    const hoverHtml = `
      <div style="font-weight:900; margin-bottom:0.25rem;">${escapeHtml(borough)}</div>
      <div style="font-size:0.86rem; color: rgba(30,30,30,0.88);">Incidents: ${formatCompactNumber(info.total)}</div>
      <div style="font-size:0.86rem; color: rgba(30,30,30,0.88);">Dominant type: ${escapeHtml(info.dominantSubtype)}${
      info.dominantSubtypeShare ? ` (${Math.round(info.dominantSubtypeShare * 100)}%)` : ""
    }</div>
      <div style="font-size:0.86rem; color: rgba(30,30,30,0.88);">Area Type: ${escapeHtml(info.areaType)} • Borough_SNT: ${escapeHtml(info.boroughSnt)}</div>
      <div style="font-size:0.86rem; color: rgba(30,30,30,0.88);">Risk level: ${info.riskLevel}</div>
      <div style="font-size:0.86rem; color: rgba(30,30,30,0.88);">NEET index: ${info.neetsIndex}</div>
      <div style="font-size:0.84rem; color: rgba(30,30,30,0.7);">Δ vs prior: ${formatDeltaPercent(info.delta)}</div>
    `;
    marker.bindTooltip(hoverHtml, {
      direction: "top",
      sticky: true,
      opacity: 0.96
    });
    if (poly) {
      poly.bindTooltip(hoverHtml, {
        direction: "top",
        sticky: true,
        opacity: 0.96
      });
    }
  }

  // Legend: stepped risk levels
  ui.mapLegend.innerHTML = `
    <div class="legend-title">Rate Map (Area fill + Symbol size)</div>
    <div style="font-size:0.82rem; color: rgba(248,246,246,0.82); margin-bottom:0.35rem;">Knife-crime intensity by area (borough vs Safer Neighbourhood Team rows)</div>
    <div style="display:flex; flex-direction:column; gap:0.3rem; font-size:0.85rem; color: rgba(231,240,255,0.85);">
      <div style="display:flex; align-items:center; gap:0.45rem;">
        <span style="width:14px;height:14px;border-radius:999px;background:rgba(220,38,38,0.95);border:1px solid rgba(255,255,255,0.35);"></span>
        <span>High (≥ ${formatCompactNumber(pHigh)})</span>
      </div>
      <div style="display:flex; align-items:center; gap:0.45rem;">
        <span style="width:14px;height:14px;border-radius:999px;background:rgba(249,115,22,0.95);border:1px solid rgba(255,255,255,0.35);"></span>
        <span>Medium (≈ ${formatCompactNumber(pLow)} – ${formatCompactNumber(pHigh)})</span>
      </div>
      <div style="display:flex; align-items:center; gap:0.45rem;">
        <span style="width:14px;height:14px;border-radius:999px;background:rgba(34,197,94,0.95);border:1px solid rgba(255,255,255,0.35);"></span>
        <span>Low (&lt; ${formatCompactNumber(pLow)})</span>
      </div>
    </div>
  `;
}

function neetFactorScaled(neetsFactorPercent) {
  // Normalize around plausible range; NEET% in the dataset is ~8%~23%.
  // returns 0..1
  return clamp((neetsFactorPercent - 8) / (23 - 8), 0, 1);
}

function neetIndexLabel(neetsFactorPercent) {
  if (!Number.isFinite(neetsFactorPercent)) return "N/A";
  const n = neetFactorScaled(neetsFactorPercent);
  if (n >= 0.66) return "High";
  if (n >= 0.33) return "Medium";
  return "Low";
}

function normalizeBoroughName(name) {
  const n = String(name || "").trim().toLowerCase();
  const map = new Map([
    ["westminster", "westminster"],
    ["city of westminster", "westminster"],
    ["kensington and chelsea", "kensington and chelsea"],
    ["hammersmith and fulham", "hammersmith and fulham"],
    ["city of london", "city of london"]
  ]);
  return map.get(n) || n;
}

/** Rough centroid for Polygon / MultiPolygon (GeoJSON). */
function centroidFromGeometry(geom) {
  if (!geom || !geom.coordinates) return null;
  const pts = [];
  const ring = (arr) => {
    if (!Array.isArray(arr)) return;
    for (const p of arr) {
      if (Array.isArray(p) && typeof p[0] === "number" && typeof p[1] === "number") pts.push([p[1], p[0]]);
    }
  };
  try {
    if (geom.type === "Polygon" && geom.coordinates[0]) ring(geom.coordinates[0]);
    else if (geom.type === "MultiPolygon" && geom.coordinates[0]?.[0]) ring(geom.coordinates[0][0]);
  } catch {
    return null;
  }
  if (!pts.length) return null;
  const lat = pts.reduce((s, p) => s + p[0], 0) / pts.length;
  const lon = pts.reduce((s, p) => s + p[1], 0) / pts.length;
  return { lat, lon };
}

/**
 * GeoJSON can include boroughs missing from crime_borough_coords (e.g. City of London).
 * Add centroid + default meta so choropleth and markers align with boundaries.
 */
function patchMissingBoroughsFromBoundaries() {
  const c = state.crime;
  if (!c.boroughBoundaries.size) return;

  const displayNameForNormKey = (normKey) => {
    if (normKey === "city of london") return "City of London";
    return String(normKey)
      .split(/\s+/)
      .map((w) => w.charAt(0).toUpperCase() + w.slice(1))
      .join(" ");
  };

  for (const normKey of c.boroughBoundaries.keys()) {
    let matched = false;
    for (const b of c.boroughCoords.keys()) {
      if (normalizeBoroughName(b) === normKey) {
        matched = true;
        break;
      }
    }
    if (matched) continue;

    const display = displayNameForNormKey(normKey);
    const geom = c.boroughBoundaries.get(normKey);
    const center = centroidFromGeometry(geom) || { lat: 51.5155, lon: -0.0922 };
    c.boroughCoords.set(display, center);
    if (!c.boroughMeta.has(display)) {
      c.boroughMeta.set(display, { areaType: "Borough", boroughSnt: display, areaName: display });
    }
  }
}

function sumSubtypeAcrossFilteredBoroughs(boroughSubtypeMap, subtype) {
  let sum = 0;
  for (const subMap of boroughSubtypeMap.values()) {
    sum += Number(subMap.get(subtype)) || 0;
  }
  return sum;
}

function updateKPIs(boroughAggSelected, boroughAggPrior, boroughSubtypeSelected, selectedKeys, priorKeys) {
  if (!state.toggles.crime || !selectedKeys.length) {
    ui.kpiTotal.textContent = "0";
    ui.kpiDelta.textContent = "—";
    ui.kpiRobbery.textContent = "0%";
    ui.kpiDisplacement.textContent = "—";

    if (state.charts.kpiLine) {
      state.charts.kpiLine.destroy();
      state.charts.kpiLine = null;
    }
    if (state.charts.kpiPie) {
      state.charts.kpiPie.destroy();
      state.charts.kpiPie = null;
    }
    return;
  }

  // Incidents (selected) = Σ Count over filtered boroughs/SNTs and selected period keys (offences from aggregates).
  const incidentsSumCount = [...boroughAggSelected.values()].reduce((s, v) => s + v, 0);
  const incidentsPriorSumCount = [...boroughAggPrior.values()].reduce((s, v) => s + v, 0);

  ui.kpiTotal.textContent = formatCompactNumber(incidentsSumCount);
  const delta = (incidentsSumCount - incidentsPriorSumCount) / Math.max(1, incidentsPriorSumCount);
  ui.kpiDelta.textContent = formatDeltaPercent(delta);

  // Personal Robbery Share = Σ Count(Personal Robbery subtype) ÷ Incidents (same filters & date range).
  const personalRobberySumCount = sumSubtypeAcrossFilteredBoroughs(boroughSubtypeSelected, PERSONAL_ROBBERY_SUBTYPE);
  const share = incidentsSumCount > 0 ? personalRobberySumCount / incidentsSumCount : 0;
  ui.kpiRobbery.textContent = `${(share * 100).toFixed(1)}%`;

  // displacement risk:
  // pick top borough by delta, then compute weighted displacement to other boroughs using distance and their deltas
  const topA = [...boroughAggSelected.entries()]
    .map(([b, total]) => {
      const prev = boroughAggPrior.get(b) || 0;
      const deltaB = (total - prev) / Math.max(1, prev);
      return { borough: b, total, prev, delta: deltaB };
    })
    .sort((x, y) => y.delta - x.delta)[0];

  if (!topA) {
    ui.kpiDisplacement.textContent = "—";
    return;
  }

  const coords = state.crime.boroughCoords;
  const distWeighted = [];
  for (const [b, total] of boroughAggSelected.entries()) {
    if (b === topA.borough) continue;
    const prev = boroughAggPrior.get(b) || 0;
    const deltaB = (total - prev) / Math.max(1, prev);
    const cA = coords.get(topA.borough);
    const cB = coords.get(b);
    if (!cA || !cB) continue;
    const dist = haversineKm(cA.lat, cA.lon, cB.lat, cB.lon);
    const weight = 1 / Math.max(0.3, dist);
    distWeighted.push({ borough: b, delta: deltaB, score: Math.max(0, deltaB) * weight });
  }

  distWeighted.sort((a, b) => b.score - a.score);
  const top2 = distWeighted.slice(0, 2);
  if (!top2.length || top2[0].score <= 0) {
    ui.kpiDisplacement.textContent = "—";
    return;
  }
  const sumScore = top2.reduce((s, x) => s + x.score, 0) || 1;
  const p1 = top2[0].score / sumScore;
  const p2 = top2[1] ? top2[1].score / sumScore : 0;
  ui.kpiDisplacement.textContent = `${topA.borough} → ${top2[0].borough} ${Math.round(p1 * 100)}%${
    top2[1] ? ` • alt ${top2[1].borough} ${Math.round(p2 * 100)}%` : ""
  }`;
}

function isGeoFilterActive() {
  return (
    state.filters.areaType !== "ALL" ||
    state.filters.boroughSnt !== "ALL" ||
    state.filters.areaName !== "ALL"
  );
}

/** Σ Count (offences) in allowed areas for one period key — matches KPI / map filters. */
function getPeriodTotalForGeoFilter(periodKey, allowedBoroughs) {
  const m = aggregateByGeoKey([periodKey], allowedBoroughs);
  let s = 0;
  for (const v of m.values()) s += v;
  return s;
}

/** Merge per-borough subtype maps into one Crime Subtype → Σ Count map (for pie & radar when filtered). */
function aggregateSubtypeForGeoFilter(periodKeys, allowedBoroughs) {
  const boroughMap = aggregateBoroughSubtypeGeo(periodKeys, allowedBoroughs);
  const out = new Map();
  for (const subMap of boroughMap.values()) {
    for (const [sub, v] of subMap.entries()) {
      addToMapSum(out, sub, v);
    }
  }
  return out;
}

function updateKpiCharts(selectedKeys, priorKeys, allowedBoroughs) {
  if (!state.toggles.crime || !selectedKeys.length) return;

  const labels = selectedKeys.map((k) => (state.granularity === "monthly" ? String(k).slice(0, 7) : String(k)));
  const curValues = selectedKeys.map((k) => getPeriodTotalForGeoFilter(k, allowedBoroughs));
  const prevValues =
    priorKeys.length === selectedKeys.length
      ? priorKeys.map((k) => getPeriodTotalForGeoFilter(k, allowedBoroughs))
      : // align by length: pad at left
        selectedKeys.map((_, i) => (priorKeys[i] ? getPeriodTotalForGeoFilter(priorKeys[i], allowedBoroughs) : null));

  const geoNote = isGeoFilterActive() ? " • Filtered geography" : "";
  ui.trendSubtitle.textContent =
    (state.granularity === "monthly"
      ? "Current range vs previous equal-length period"
      : "Current years vs previous equal-length window") + geoNote;

  const min = Math.min(...curValues, ...prevValues.filter((v) => v != null));
  const max = Math.max(...curValues, ...prevValues.filter((v) => v != null));
  const { suggestedMin, suggestedMax } = computeSuggestedScale(min, max);

  const commonTick = "rgba(248,246,246,0.9)";
  const commonGrid = "rgba(160,190,255,0.18)";

  if (!state.charts.kpiLine) {
    state.charts.kpiLine = new Chart(ui.kpiLineChartCanvas, {
      type: "line",
      data: {
        labels,
        datasets: [
          {
            label: "Incidents (Σ Count)",
            data: curValues,
            borderColor: PALETTE.blue1,
            backgroundColor: "rgba(79,146,224,0.22)",
            borderWidth: 2,
            tension: 0.25,
            fill: true,
            pointRadius: 2
          },
          {
            label: "Prior window (Σ Count)",
            data: prevValues.map((v) => (v == null ? 0 : v)),
            borderColor: PALETTE.warm1,
            backgroundColor: "rgba(227,159,90,0.16)",
            borderWidth: 2,
            tension: 0.25,
            fill: false,
            pointRadius: 2
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { labels: { color: "rgba(248,246,246,0.92)", font: { size: 11, weight: "700" } } }
        },
        scales: {
          x: { ticks: { color: commonTick }, grid: { color: "transparent" } },
          y: { suggestedMin, suggestedMax, ticks: { color: commonTick }, grid: { color: commonGrid } }
        }
      }
    });
  } else {
    state.charts.kpiLine.data.labels = labels;
    state.charts.kpiLine.data.datasets[0].data = curValues;
    state.charts.kpiLine.data.datasets[1].data = prevValues.map((v) => (v == null ? 0 : v));
    state.charts.kpiLine.options.scales.y.suggestedMin = suggestedMin;
    state.charts.kpiLine.options.scales.y.suggestedMax = suggestedMax;
    state.charts.kpiLine.update();
  }

  // Pie: Crime Subtype × Σ Count at range end (same geography as KPI row when filtered)
  const endKey = state.selectedEndKey;
  const subtypeAgg = aggregateSubtypeForGeoFilter([endKey], allowedBoroughs);
  updatePieComposition(subtypeAgg);
}

function updatePieComposition(subtypeAgg) {
  const entries = [...subtypeAgg.entries()].sort((a, b) => b[1] - a[1]);
  const top = entries.slice(0, 6);
  const rest = entries.slice(6);

  const labels = top.map((x) => compressSubtypeLabel(x[0]));
  const values = top.map((x) => x[1]);
  if (rest.length) {
    labels.push("Other");
    values.push(rest.reduce((s, x) => s + x[1], 0));
  }

  const colors = [PALETTE.blue1, PALETTE.violet, PALETTE.warm1, PALETTE.red1, PALETTE.green1, "#728191"];

  if (!state.charts.kpiPie) {
    state.charts.kpiPie = new Chart(ui.kpiPieChartCanvas, {
      type: "pie",
      data: {
        labels,
        datasets: [
          {
            label: "Σ Count by subtype",
            data: values,
            backgroundColor: colors.slice(0, labels.length),
            borderWidth: 0
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { labels: { color: "rgba(248,246,246,0.92)", font: { size: 11, weight: "700" } } },
          tooltip: { callbacks: { label: (ctx) => `${ctx.label}: ${formatCompactNumber(ctx.raw)}` } }
        }
      }
    });
  } else {
    state.charts.kpiPie.data.labels = labels;
    state.charts.kpiPie.data.datasets[0].data = values;
    state.charts.kpiPie.update();
  }
}

function compressSubtypeLabel(s) {
  const t = String(s || "");
  return t
    .replace(/Knife Crime with Injury/gi, "Injury")
    .replace(/Personal Robbery/gi, "Personal Robbery")
    .replace(/Knife Injury Victims/gi, "Injury Victims")
    .replace(/Positive Outcomes/gi, "Positive Outcomes")
    .replace(/\s+/g, " ")
    .trim();
}

function updatePredictedIntelligence(selectedKeys, priorKeys, neetsFactor, allowedBoroughs = null) {
  if (!state.toggles.crime || !selectedKeys.length) {
    // Keep radar visible, but dim values and show empty prediction chart.
    const empty = [18, 18, 18, 18, 18];
    updateRadar(state.toggles.ggphiSim ? applyGgphiToRadar(empty) : empty);
    updatePredictionChart([], [], "bar");
    ui.predictionSubtitle.textContent = "Enable Knife Crime to compute area-code hotspot prediction.";
    return;
  }

  // Area code forecast from frequency + momentum
  const currentByCode = aggregateByAreaCode(selectedKeys, allowedBoroughs);
  const priorByCode = aggregateByAreaCode(priorKeys, allowedBoroughs);

  const scores = [];
  let totalCur = 0;
  let totalPrev = 0;
  for (const v of currentByCode.values()) totalCur += v;
  for (const v of priorByCode.values()) totalPrev += v;

  for (const [code, cur] of currentByCode.entries()) {
    const prev = priorByCode.get(code) || 0;
    const freq = cur;
    const momentum = (cur - prev) / Math.max(1, prev);
    scores.push({ code, cur, prev, momentum });
  }

  scores.sort((a, b) => {
    // predicted score: weighted momentum + frequency
    return predictedScore(b, scores, totalCur) - predictedScore(a, scores, totalCur);
  });

  const topN = 10;
  const top = scores.slice(0, topN);
  const labels = top.map((x) => formatAreaCodeLabel(x.code));
  const values = top.map((x) => x.cur);

  // Prediction subtitle
  const hasPrev = priorKeys.length > 0;
  ui.predictionSubtitle.textContent = `Computed from Area Code frequency${hasPrev ? " + momentum vs prior window" : ""}. Top ${Math.min(
    topN,
    top.length
  )} shown.`;

  if (state.predictionMode === "bar") {
    updatePredictionChart(labels, values, "bar");
  } else if (state.predictionMode === "pie") {
    updatePredictionChart(labels, values, "pie");
  } else {
    updatePredictionLineChart(top.slice(0, 4).map((x) => x.code));
  }

  // Social radar: simulated from borough risk + Personal Robbery share (Σ Count) + NEET factor
  const neetsPercent = neetsFactor == null ? null : neetsFactor;
  const boroughRiskMap = state._lastBoroughAgg || new Map();
  const topBorough = [...boroughRiskMap.entries()]
    .map(([b, info]) => ({ borough: b, total: info.total, delta: info.delta }))
    .sort((a, b) => b.delta - a.delta)[0];

  // Same Count-weighting & geography as KPIs: Σ Count by subtype over selected window + filters
  const subtypeAggGeo = aggregateSubtypeForGeoFilter(selectedKeys, allowedBoroughs);
  const totalCountSubtype = valuesSum(subtypeAggGeo);
  const personalRobberyCount = Number(subtypeAggGeo.get(PERSONAL_ROBBERY_SUBTYPE)) || 0;
  const personalRobberyShare = totalCountSubtype > 0 ? personalRobberyCount / totalCountSubtype : 0;
  const knifeSum = [...subtypeAggGeo.entries()].reduce((s, [sub, v]) => s + (/knife/i.test(sub) ? v : 0), 0);
  const knifeShare = totalCountSubtype > 0 ? knifeSum / totalCountSubtype : 0;

  const metrics = simulateSocialRadar({
    neetsPercent,
    personalRobberyShare,
    knifeShare,
    topMomentum: top[0]?.momentum ?? 0,
    displacement: extractDisplacementSignal()
  });
  updateRadar(state.toggles.ggphiSim ? applyGgphiToRadar(metrics) : metrics);
}

function valuesSum(map) {
  let sum = 0;
  for (const v of map.values()) sum += v;
  return sum;
}

function predictedScore(item, scores, totalCur) {
  // normalize momentum and frequency among current scores
  const moms = scores.map((x) => x.momentum);
  const freqs = scores.map((x) => x.cur);
  const momMin = Math.min(...moms);
  const momMax = Math.max(...moms);
  const freqMin = Math.min(...freqs);
  const freqMax = Math.max(...freqs);
  const momN = momMax === momMin ? 0.5 : (item.momentum - momMin) / (momMax - momMin);
  const freqN = freqMax === freqMin ? 0.5 : (item.cur - freqMin) / (freqMax - freqMin);
  return 0.7 * momN + 0.3 * freqN;
}

function formatAreaCodeLabel(code) {
  const borough = state.crime.areaCodeToBorough.get(code);
  return borough ? `${code} • ${borough}` : String(code);
}

function aggregateByAreaCode(periodKeys, allowedBoroughs = null) {
  const byArea = state.granularity === "monthly" ? state.crime.monthByAreaCode : state.crime.yearByAreaCode;
  const out = new Map();
  for (const pk of periodKeys) {
    const m = byArea.get(pk);
    if (!m) continue;
    for (const [code, v] of m.entries()) {
      if (allowedBoroughs) {
        const b = state.crime.areaCodeToBorough.get(code);
        if (!b || !allowedBoroughs.has(b)) continue;
      }
      addToMapSum(out, code, v);
    }
  }
  return out;
}

function updatePredictionChart(labels, values, type) {
  if (!state.charts.prediction) {
    state.charts.prediction = buildPredictionChart(ui.predictionChartCanvas, labels, values, type);
    return;
  }
  // Rebuild if type changed
  if (state.charts.prediction.config.type !== type) {
    state.charts.prediction.destroy();
    state.charts.prediction = buildPredictionChart(ui.predictionChartCanvas, labels, values, type);
    return;
  }
  state.charts.prediction.data.labels = labels;
  const ds0 = state.charts.prediction.data.datasets[0];
  ds0.data = values;
  const n = values.length;
  if (type === "bar" || type === "pie") {
    const cols = colorsForSeriesCount(Math.max(n, 1));
    ds0.backgroundColor = values.map((_, i) => cols[i]);
  }
  state.charts.prediction.update();
}

function buildPredictionChart(canvas, labels, values, type) {
  const ctx = canvas;
  const commonTick = "rgba(248,246,246,0.9)";
  const commonGrid = "rgba(160,190,255,0.18)";
  const n = Math.max(values.length, 1);
  const seriesColors = colorsForSeriesCount(n);
  const accentLine = seriesColors[0];

  if (type === "pie") {
    return new Chart(ctx, {
      type: "pie",
      data: {
        labels: labels,
        datasets: [
          {
            data: values,
            backgroundColor: values.map((_, i) => seriesColors[i]),
            borderWidth: 0
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { labels: { color: "rgba(248,246,246,0.92)", font: { size: 11, weight: "700" } } },
          tooltip: { callbacks: { label: (c) => `${c.label}: ${formatCompactNumber(c.raw)}` } }
        }
      }
    });
  }

  return new Chart(ctx, {
    type: type === "bar" ? "bar" : "line",
    data: {
      labels,
      datasets: [
        {
          label: "Predicted incidents frequency",
          data: values,
          backgroundColor:
            type === "bar"
              ? values.map((_, idx) => seriesColors[idx])
              : hexToRgba(accentLine, 0.2),
          borderColor: type === "bar" ? "transparent" : accentLine,
          borderWidth: type === "bar" ? 0 : 2,
          fill: type === "line"
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false }
      },
      scales: {
        x: { ticks: { color: commonTick, autoSkip: false }, grid: { color: "transparent" } },
        y: { ticks: { color: commonTick }, grid: { color: commonGrid }, beginAtZero: true }
      }
    }
  });
}

function updatePredictionLineChart(topCodes) {
  if (!topCodes.length) return;
  const keys = state.granularity === "monthly" ? state.crime.monthKeysSorted : state.crime.yearKeysSorted;
  const startIdx = keys.indexOf(state.selectedStartKey);
  const endIdx = keys.indexOf(state.selectedEndKey);
  const selectedKeys = keys.slice(startIdx, endIdx + 1);

  const labels = selectedKeys.map((k) => (state.granularity === "monthly" ? String(k).slice(0, 7) : String(k)));

  const lineColors = colorsForSeriesCount(topCodes.length);
  const datasets = topCodes.map((code, idx) => {
    const values = selectedKeys.map((k) => {
      const map = state.granularity === "monthly" ? state.crime.monthByAreaCode.get(k) : state.crime.yearByAreaCode.get(k);
      return map?.get(code) || 0;
    });
    const borough = state.crime.areaCodeToBorough.get(code) || code;
    const color = lineColors[idx];
    return {
      label: borough,
      data: values,
      borderColor: color,
      backgroundColor: hexToRgba(color, 0.18),
      borderWidth: 2,
      tension: 0.25,
      fill: true,
      pointRadius: 2
    };
  });

  // Recreate
  if (state.charts.prediction) state.charts.prediction.destroy();
  state.charts.prediction = new Chart(ui.predictionChartCanvas, {
    type: "line",
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: "rgba(231,240,255,0.85)", font: { size: 11, weight: "700" } } },
        tooltip: { mode: "index", intersect: false }
      },
      scales: {
        x: { ticks: { color: "rgba(248,246,246,0.9)" }, grid: { color: "transparent" } },
        y: { ticks: { color: "rgba(248,246,246,0.9)" }, grid: { color: "rgba(177,201,231,0.2)" }, beginAtZero: true }
      }
    }
  });
}

function pickHeatColor(v, values) {
  // green->red based on relative magnitude
  const min = Math.min(...values);
  const max = Math.max(...values);
  const t = max === min ? 0.5 : (v - min) / (max - min);
  const c = gradientColor(clamp(t, 0, 1), [34, 197, 94], [239, 68, 68]);
  return c;
}

function updateRadar(values) {
  const labels = SOCIAL_RADAR_LABELS;
  const data = values.map((v) => clamp(v, 0, 100));
  const datasetBase = {
    label: "Real-time Risk Index",
    fill: true,
    backgroundColor: "rgba(255, 99, 132, 0.2)",
    borderColor: "rgb(255, 99, 132)",
    pointBackgroundColor: "rgb(255, 99, 132)",
    pointBorderColor: "#fff",
    pointHoverBackgroundColor: "#fff",
    pointHoverBorderColor: "rgb(255, 99, 132)",
    borderWidth: 3,
    tension: 0.1
  };

  const radarOptions = {
    responsive: true,
    maintainAspectRatio: false,
    scales: {
      r: {
        min: 0,
        max: 100,
        angleLines: { color: "rgba(255, 255, 255, 0.2)" },
        grid: { color: "rgba(255, 255, 255, 0.2)" },
        pointLabels: {
          color: "#e0e0e0",
          font: { size: 11, family: "'Segoe UI', sans-serif", weight: "700" }
        },
        ticks: { display: false, stepSize: 20 }
      }
    },
    plugins: {
      legend: { display: false },
      tooltip: {
        backgroundColor: "rgba(0, 0, 0, 0.8)",
        titleFont: { size: 14 },
        bodyFont: { size: 13 },
        callbacks: {
          label: (ctx) => {
            const raw = ctx.parsed?.r ?? ctx.raw;
            return ` Risk Level: ${raw}%`;
          }
        }
      }
    }
  };

  if (!state.charts.socialRadar) {
    state.charts.socialRadar = new Chart(ui.socialRadarChartCanvas, {
      type: "radar",
      data: {
        labels,
        datasets: [{ ...datasetBase, data }]
      },
      options: radarOptions
    });
  } else {
    state.charts.socialRadar.data.labels = labels;
    const d0 = state.charts.socialRadar.data.datasets[0];
    d0.data = data;
    Object.assign(d0, datasetBase);
    state.charts.socialRadar.update();
  }
}

function syncGgphiRangeLabels() {
  if (ui.ggphiChatterInput) ui.ggphiChatterInput.value = String(state.ggphiSim.chatter);
  if (ui.ggphiMarketInput) ui.ggphiMarketInput.value = String(state.ggphiSim.market);
  const c = state.ggphiSim.chatter;
  const m = state.ggphiSim.market;
  if (ui.ggphiChatterLabel) ui.ggphiChatterLabel.textContent = c > 70 ? "High" : c > 30 ? "Medium" : "Low";
  if (ui.ggphiMarketLabel) ui.ggphiMarketLabel.textContent = m > 70 ? "High" : m > 30 ? "Medium" : "Low";
}

function updateGgphiOutcomeUI() {
  const s = state.ggphiSim;
  s.trust = computeGgphiDerivedTrust();
  s.sustain = computeGgphiDerivedSustain();
  const { incidentPressure } = computeGgphiPressureModel();
  if (ui.ggphiCrimeRate) {
    const idx = Math.round(incidentPressure * 100);
    ui.ggphiCrimeRate.textContent = String(idx);
    ui.ggphiCrimeRate.style.color =
      incidentPressure > 0.42 ? PALETTE.red1 : incidentPressure > 0.18 ? PALETTE.warm2 : PALETTE.green1;
  }
  if (ui.ggphiTrustRate) {
    ui.ggphiTrustRate.textContent = `${Math.floor(s.trust)}%`;
    if (s.trust < 42) ui.ggphiTrustRate.style.color = PALETTE.red1;
    else if (s.trust < 52) ui.ggphiTrustRate.style.color = PALETTE.warm2;
    else ui.ggphiTrustRate.style.color = PALETTE.blue3;
  }
  if (ui.ggphiSustainRate) {
    ui.ggphiSustainRate.textContent = s.sustain > 60 ? "High" : "Moderate";
    ui.ggphiSustainRate.style.color = s.sustain > 60 ? PALETTE.green1 : PALETTE.text;
  }
}

function resetGgphiSimulation() {
  state.ggphiSim.chatter = 20;
  state.ggphiSim.market = 10;
  state.ggphiSim.patrolForce = 0;
  state.ggphiSim.vruIntervention = 0;
  state.ggphiSim.cptedForce = 0;
  syncGgphiRangeLabels();
  updateGgphiOutcomeUI();
  clearGgphiSimMapOverlays();
  if (state.crime.monthKeysSorted.length || state.crime.yearKeysSorted.length) refreshAll();
  else if (state.toggles.ggphiSim) updateGgphiPulseHotspots();
}

function ggphiApplyAction(type) {
  const s = state.ggphiSim;
  const step = 0.12;
  if (type === "patrol") {
    s.patrolForce = clamp(s.patrolForce + step, 0, 1.28);
  } else if (type === "cpted") {
    s.cptedForce = clamp(s.cptedForce + step, 0, 1.28);
  } else if (type === "vru") {
    s.vruIntervention = clamp(s.vruIntervention + step, 0, 1.28);
  }
  updateGgphiOutcomeUI();
  spawnGgphiResponseMarkers(type);
  pulseGgphiResponseLegend(type);
  if (state.crime.monthKeysSorted.length || state.crime.yearKeysSorted.length) refreshAll();
  else if (state.toggles.ggphiSim) updateGgphiPulseHotspots();
}

function applyGgphiToRadar(base) {
  if (!state.toggles.ggphiSim) return base;
  const s = state.ggphiSim;
  const c = s.chatter / 100;
  const m = s.market / 100;
  const { intelligencePressure, incidentPressure } = computeGgphiPressureModel();
  const crimeN = clamp(0.18 + incidentPressure * 1.05, 0.12, 1);
  const trustN = s.trust / 100;

  return [
    clamp(base[0] * (0.82 + 0.18 * crimeN) + 26 * intelligencePressure + 8 * c, 0, 100),
    clamp(base[1] * crimeN + 18 * c + 10 * m + 14 * intelligencePressure, 0, 100),
    clamp(base[2] + 22 * Math.sqrt(c * m + 0.015) + 12 * (trustN - 0.5), 0, 100),
    clamp(base[3] + 20 * m + 8 * c * crimeN + 10 * intelligencePressure, 0, 100),
    clamp(base[4] * crimeN + 24 * m + 6 * c + 12 * incidentPressure, 0, 100)
  ];
}

function extractDisplacementSignal() {
  const txt = ui.kpiDisplacement.textContent || "";
  // map to a number: higher percent -> higher displacement signal
  const m = txt.match(/(\d+)%/);
  const pct = m ? Number(m[1]) : 25;
  return clamp(pct / 100, 0, 1);
}

function simulateSocialRadar({ neetsPercent, personalRobberyShare, knifeShare, topMomentum, displacement }) {
  const neetsN = neetsPercent == null ? 0.5 : clamp((neetsPercent - 8) / (23 - 8), 0, 1);
  const robN = clamp(personalRobberyShare / 0.45, 0, 1);
  const knifeN = clamp((knifeShare || 0) / 0.35, 0, 1);
  const momN = clamp((topMomentum + 0.5) / 2.0, 0, 1);
  const dispN = clamp(displacement || 0.3, 0, 1);
  const hubsOn = state.toggles.hubs ? 1 : 0;
  const elecOn = state.toggles.electronics ? 1 : 0;

  const geospatialAnchoring = clamp(36 + 44 * dispN + 30 * momN + 14 * neetsN * (1 - momN * 0.25), 0, 100);
  const aggressionDemand = clamp(20 + 50 * robN + 38 * momN + 20 * neetsN + 12 * knifeN, 0, 100);
  const networkSymbiosis = clamp(18 + 40 * Math.sqrt(robN * neetsN + 0.04) + 35 * momN + 22 * robN * neetsN, 0, 100);
  const hubInfluence = clamp(22 + 52 * momN + 32 * robN + 28 * hubsOn + 14 * elecOn, 0, 100);
  const commodityViolence = clamp(16 + 48 * knifeN + 42 * robN + 26 * elecOn + 18 * robN * knifeN, 0, 100);

  return [geospatialAnchoring, aggressionDemand, networkSymbiosis, hubInfluence, commodityViolence];
}

function updatePoliceDeployment(selectedKeys) {
  if (!state.toggles.crime || !selectedKeys.length) {
    ui.policeMeta.textContent = "Enable Knife Crime to compute deployment context.";
    if (state.charts.policeDonut) {
      state.charts.policeDonut.destroy();
      state.charts.policeDonut = null;
    }
    return;
  }

  // Average strengths over selected range
  let officer = 0;
  let staff = 0;
  let pcso = 0;
  let n = 0;

  if (state.granularity === "monthly") {
    for (const mk of selectedKeys) {
      const v = state.police.monthToStrength.get(mk);
      if (!v) continue;
      officer += v.officer;
      staff += v.staff;
      pcso += v.pcso;
      n++;
    }
  } else {
    for (const y of selectedKeys) {
      const v = state.police.yearToStrength.get(y);
      if (!v) continue;
      officer += v.officer;
      staff += v.staff;
      pcso += v.pcso;
      n++;
    }
  }

  if (n === 0) n = 1;
  officer /= n;
  staff /= n;
  pcso /= n;

  const total = officer + staff + pcso;
  const sOfficer = total > 0 ? officer / total : 0;
  const sStaff = total > 0 ? staff / total : 0;
  const sPcso = total > 0 ? pcso / total : 0;

  ui.policeMeta.textContent = `Avg strength: ${formatCompactNumber(total)} (Officer ${Math.round(sOfficer * 100)}% • Staff ${Math.round(
    sStaff * 100
  )}% • PCSO ${Math.round(sPcso * 100)}%)`;

  if (!state.charts.policeDonut) {
    state.charts.policeDonut = new Chart(ui.policeDonutCanvas, {
      type: "doughnut",
      data: {
        labels: ["Police Officer", "Police Staff", "PCSO"],
        datasets: [
          {
            data: [officer, staff, pcso],
            backgroundColor: [PALETTE.blue1, PALETTE.green1, PALETTE.warm1],
            borderWidth: 0,
            hoverOffset: 6
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: "68%",
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.label}: ${formatCompactNumber(ctx.raw)}`
            }
          }
        }
      }
    });
  } else {
    state.charts.policeDonut.data.datasets[0].data = [officer, staff, pcso];
    state.charts.policeDonut.update();
  }
}

function gradientColor(t, rgbA, rgbB) {
  const tt = clamp(t, 0, 1);
  const r = Math.round(rgbA[0] + (rgbB[0] - rgbA[0]) * tt);
  const g = Math.round(rgbA[1] + (rgbB[1] - rgbA[1]) * tt);
  const b = Math.round(rgbA[2] + (rgbB[2] - rgbA[2]) * tt);
  return `rgba(${r},${g},${b},0.95)`;
}

function clamp(v, min, max) {
  return Math.min(max, Math.max(min, v));
}

function formatCompactNumber(n) {
  if (!Number.isFinite(n)) return "0";
  if (n >= 1e6) return `${(n / 1e6).toFixed(1)}M`;
  if (n >= 1e3) return `${(n / 1e3).toFixed(1)}K`;
  return String(Math.round(n));
}

function formatDeltaPercent(deltaRatio) {
  if (!Number.isFinite(deltaRatio)) return "—";
  const pct = deltaRatio * 100;
  const sign = pct >= 0 ? "+" : "";
  return `${sign}${Math.round(pct)}%`;
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function computeSuggestedScale(min, max) {
  if (!Number.isFinite(min) || !Number.isFinite(max)) return { suggestedMin: 0, suggestedMax: 1 };
  if (min === max) return { suggestedMin: 0, suggestedMax: Math.max(1, max * 1.2) };
  const pad = (max - min) * 0.12;
  return { suggestedMin: Math.max(0, min - pad), suggestedMax: max + pad };
}

function percentile(arr, p) {
  const vals = (arr || []).filter((x) => Number.isFinite(x)).slice().sort((a, b) => a - b);
  if (!vals.length) return 0;
  const idx = Math.min(vals.length - 1, Math.max(0, Math.floor(p * vals.length)));
  return vals[idx];
}

function haversineKm(lat1, lon1, lat2, lon2) {
  const R = 6371;
  const dLat = ((lat2 - lat1) * Math.PI) / 180;
  const dLon = ((lon2 - lon1) * Math.PI) / 180;
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos((lat1 * Math.PI) / 180) * Math.cos((lat2 * Math.PI) / 180) * Math.sin(dLon / 2) ** 2;
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

