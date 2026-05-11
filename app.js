(function () {
  "use strict";

  var DATA = {
    es: "geo/ESBoundaries.json",
    ms: "geo/MSBoundaries.json",
    hs: "geo/HSBoundaries.json",
    schools: "geo/SchoolLocations.json",
    /** Single source for enrollment, capacity, utilization, From-To capture KPIs, facility stats, demographics (CSV). */
    masterCsv: "data/school_master.csv",
    sankeyEsMs: "data/processed/sankey_es_ms.json",
    /** Deduped bundle (see scripts/bundle-dedupe-student-hex.cjs) — one geometry per hex, all student rows. */
    studentHexagons: "geo/StudentHexagons.bundle.json",
    schoolParcels: "geo/SchoolParcels.geojson",
    schoolBoardDistricts: "geo/SchoolBoardDistricts.geojson",
    municipalBoundaries: "geo/MunicipalBoundaries.geojson",
    charterSchoolLocations: "geo/CharterSchoolLocations.geojson",
    privateSchoolLocations: "geo/PrivateSchools.json",
    /** Homeschool students joined to hex grid (GRID_ID); one polygon row per student in source export. */
    homeschoolStudentHexagons: "geo/HomeschoolStudentHexagons.geojson",
    /** Meadowlane Primary/Intermediate grade-band capture overrides (see notes inside file). */
    meadowlaneCaptureOverride: "data/processed/meadowlane_capture_override.json",
    /** K-12 ESE feeder matrix (columns = program destinations per school row). */
    eseFeederMatrix: "data/processed/ese_feeder_matrix.json",
    /** Johnson / McNair / Stone travel-distance scenario workbooks (see scripts/export_travel_impact_from_xlsx.py). */
    travelImpact: "data/processed/travel_impact.json",
    /** Network isochrones (1–10 mi by school); "Name" encodes MSID and ToBreak in feet. */
    schoolIsochrones: "geo/SchoolIsochrones.geojson",
    /** On-site BPS employee counts by MSID (see scripts/export_bps_employee_count_from_xlsx.py). */
    bpsEmployeeCount: "data/processed/bps_employee_count_by_msid.json",
  };

  var FEET_PER_MILE = 5280;
  /** US survey / international foot–based conversion (Turf geodesic area in m² → sq mi for student density). */
  var SQ_METERS_PER_SQ_MI = 2589988.110336;
  /** Histogram bucket width for travel-distance charts (miles). */
  var TRAVEL_BIN_MI = 0.25;
  /** Median line and label (flag) for travel histograms. */
  var TRAVEL_MEDIAN_COLOR = "#dc2626";
  /** Mean / average line and label. */
  var TRAVEL_MEAN_COLOR = "#2563eb";

  /** Mapbox access token (prefer env / serverless proxy for public deployments). */
  var MAPBOX_ACCESS_TOKEN =
    "pk.eyJ1IjoicGF0d2QwNSIsImEiOiJjbTZ2bGVhajIwMTlvMnFwc2owa3BxZHRoIn0.moDNfqMUolnHphdwsIF87w";
  var MAPBOX_STYLES = {
    light: "mapbox://styles/mapbox/light-v11",
    streets: "mapbox://styles/mapbox/streets-v12",
    satellite: "mapbox://styles/mapbox/satellite-v9",
  };
  /** Sentinel for municipal hover line layer filter when nothing is highlighted. */
  var MUN_HOVER_FILTER_OFF = "__mun_hover_off__";
  /** Cached Promise.all results so basemap style switches can re-add GeoJSON layers. */
  var geoJsonDataCache = null;
  /** Meadowlane 2031/2041 capture numerators/denominators; null until fetch completes. */
  var MEADOWLANE_CAPTURE_OVERRIDE = null;

  /**
   * @param {string} text
   * @returns {string[][]}
   */
  function parseCsvRows(text) {
    var rows = [];
    var row = [];
    var cur = "";
    var inQ = false;
    if (!text) return rows;
    if (text.charCodeAt(0) === 0xfeff) {
      text = text.slice(1);
    }
    for (var i = 0; i < text.length; i++) {
      var c = text[i];
      if (inQ) {
        if (c === '"') {
          if (text[i + 1] === '"') {
            cur += '"';
            i++;
          } else {
            inQ = false;
          }
        } else {
          cur += c;
        }
      } else {
        if (c === '"') {
          inQ = true;
        } else if (c === ",") {
          row.push(cur);
          cur = "";
        } else if (c === "\n") {
          row.push(cur);
          rows.push(row);
          row = [];
          cur = "";
        } else if (c === "\r") {
          /* ignore */
        } else {
          cur += c;
        }
      }
    }
    row.push(cur);
    if (row.length > 1 || (row.length === 1 && row[0] !== "")) {
      rows.push(row);
    }
    return rows;
  }

  /**
   * @param {string} text raw CSV
   * @returns {Object<string, Object>|null} keyed by MSID string
   */
  function parseSchoolMasterCsv(text) {
    var grid = parseCsvRows(text);
    if (!grid || grid.length < 2) return null;
    var headers = grid[0].map(function (h) {
      return String(h).trim();
    });
    var byMsid = {};
    for (var r = 1; r < grid.length; r++) {
      var cells = grid[r];
      if (!cells || !cells.length) continue;
      var obj = {};
      for (var c = 0; c < headers.length; c++) {
        obj[headers[c]] = cells[c] != null ? String(cells[c]).trim() : "";
      }
      var idRaw = obj.msid != null ? String(obj.msid).trim() : "";
      if (!idRaw) continue;
      var idNum = parseInt(idRaw, 10);
      if (isNaN(idNum)) continue;
      var idPadded = String(idNum).padStart(4, "0");
      var idUnpadded = String(idNum);
      obj.msid = idPadded;
      byMsid[idPadded] = obj;
      byMsid[idUnpadded] = obj;
    }
    return byMsid;
  }

  /** @returns {Object|null} */
  function masterRow(msid) {
    if (msid == null || isNaN(msid) || !MASTER_BY_MSID) return null;
    return MASTER_BY_MSID[String(msid)] || null;
  }

  /** Formatted count from data/processed/bps_employee_count_by_msid.json, or "—". */
  function bpsOnSiteEmployeeCountDisplay(msid) {
    if (msid == null || isNaN(Number(msid)) || !BPS_EMPLOYEE_COUNT_BY_MSID) {
      return "—";
    }
    var map = BPS_EMPLOYEE_COUNT_BY_MSID;
    var n = Number(msid);
    var keys = [String(n), String(n).padStart(4, "0")];
    var raw = null;
    for (var ki = 0; ki < keys.length; ki++) {
      if (Object.prototype.hasOwnProperty.call(map, keys[ki])) {
        raw = map[keys[ki]];
        break;
      }
    }
    if (raw == null || raw === "") return "—";
    var num = Number(raw);
    if (isNaN(num)) return "—";
    return num.toLocaleString();
  }

  function schoolLevelToTypeString(level) {
    var lv = String(level || "").toLowerCase();
    if (lv === "elementary") return "ELEMENTARY";
    if (lv === "middle") return "MIDDLE";
    if (lv === "high") return "HIGH";
    if (lv === "jr_sr_high") return "JR SR HIGH";
    return "";
  }

  /** Overlays TYPE from master CSV school_level when present (single source of truth). */
  function schoolPropsWithMasterType(p) {
    if (!p) return p;
    var msid = p.SCHOOLS_ID != null ? Number(p.SCHOOLS_ID) : NaN;
    var m = masterRow(msid);
    var t = m && schoolLevelToTypeString(m.school_level);
    if (t) return Object.assign({}, p, { TYPE: t });
    return p;
  }

  /** Applies master TYPE to every school feature so map layers and parcels match CSV (e.g. 7–12). */
  function enrichSchoolsFcWithMasterType(schoolsFc) {
    if (!schoolsFc || !schoolsFc.features || !schoolsFc.features.length) {
      return schoolsFc;
    }
    return {
      type: "FeatureCollection",
      features: schoolsFc.features.map(function (ft) {
        var p = ft.properties;
        var merged = schoolPropsWithMasterType(p);
        return Object.assign({}, ft, { properties: merged || p });
      }),
    };
  }

  /** @returns {{ ethnicity: Object<string, number>, lunchStatus: Object<string, number> }|null} */
  function demographicsObjectsFromMaster(m) {
    if (!m) return null;
    var eth = {};
    var lunch = {};
    for (var i = 0; i < DEMO_ETH_SLUGS.length; i++) {
      var d = DEMO_ETH_SLUGS[i];
      var v = m[d.slug];
      if (v !== "" && v != null && !isNaN(Number(v))) {
        var n = Number(v);
        if (n > 0) eth[d.label] = n;
      }
    }
    for (var j = 0; j < DEMO_LUNCH_SLUGS.length; j++) {
      var e = DEMO_LUNCH_SLUGS[j];
      var w = m[e.slug];
      if (w !== "" && w != null && !isNaN(Number(w))) {
        lunch[e.label] = Number(w);
      }
    }
    return { ethnicity: eth, lunchStatus: lunch };
  }

  function projectedColumnForSyLabel(label) {
    return "projected_" + String(label).replace(/-/g, "_");
  }

  /** Set after GeoJSON loads; used to zoom to assignment boundaries. */
  var GEO_CACHE = { es: null, ms: null, hs: null, schools: null };
  /** Enriched `SchoolIsochrones.geojson` (parsed Name → iso_msid, iso_miles, etc.); set from fetch. */
  var SCHOOL_ISOCHRONES_ENRICHED = null;
  /** Parsed rows from data/school_master.csv keyed by MSID string; null if missing or failed to load. */
  var MASTER_BY_MSID = null;
  /** From data/processed/ese_feeder_matrix.json; null if missing or failed to load. */
  var ESE_FEEDER_MATRIX = null;
  /** From data/processed/bps_employee_count_by_msid.json; null if missing or failed to load. */
  var BPS_EMPLOYEE_COUNT_BY_MSID = null;
  /** SCHOOLS_ID keys for `SchAB_Type === "CHOICE"` from SchoolLocations (capture KPI + CSV download). */
  var CHOICE_SCHOOL_MSIDS = null;
  /** SCHOOLS_ID keys for charter schools (TYPE/SchAB_Type CHARTER on boundary + charter location layers). */
  var CHARTER_SCHOOL_MSIDS = null;
  /** Projected school-year column labels (matches CSV projected_* headers). */
  var MASTER_PROJECTION_LABELS = ["2026-27", "2027-28", "2028-29", "2029-30", "2030-31"];
  /** Slugs and display labels for ethnicity count columns in the master CSV. */
  var DEMO_ETH_SLUGS = [
    { slug: "eth_hawaiian_native_pacific_islander", label: "Hawaiian Native/Pacific Islander" },
    { slug: "eth_asian", label: "Asian" },
    { slug: "eth_black_non_hispanic", label: "Black, Non-Hispanic" },
    { slug: "eth_hispanic", label: "Hispanic" },
    { slug: "eth_amer_indian_or_alaskan_native", label: "Amer. Indian or Alaskan Native" },
    { slug: "eth_multi_racial", label: "Multi-Racial" },
    { slug: "eth_white_non_hispanic", label: "White, Non-Hispanic" },
  ];
  var DEMO_LUNCH_SLUGS = [
    { slug: "lunch_not_free_reduced", label: "Not free/reduced" },
    { slug: "lunch_free", label: "Free" },
    { slug: "lunch_reduced", label: "Reduced" },
  ];
  /** ES→MS flows from SankeyFlowHelper export; null if missing. */
  var SANKEY_CACHE = null;
  /** Travel impact triples [attendance_msid, scenario_msid, ft] per middle workbook; see DATA.travelImpact. */
  var TRAVEL_IMPACT_ALL = null;
  /** Map string MSID -> true where GeoJSON(+master) TYPE is middle school (non–Jr/Sr high). */
  var MIDDLE_SCHOOL_MSID_SET = null;
  /**
   * Student hex overlay index: counts + geometry by hex key, per-student detail rows
   * (Grade, MSID attendance, zoned ELEM_/MID_/INT_/HIGH_), and districtwide charter hex counts
   * (attendance MSID 6500–6699).
   */
  var STUDENT_HEX_INDEX = null;
  /** Per-hex homeschool student counts (`studentHexKey` → count), from homeschool GeoJSON. */
  var HOMESCHOOL_HEX_COUNTS = null;
  /**
   * Hex geometries from homeschool export for IDs missing from `STUDENT_HEX_INDEX.geometryByHexKey`
   * (bundle-first resolution in `homeschoolHexGeometry`).
   */
  var HOMESCHOOL_HEX_GEOMETRY_FALLBACK = null;
  /**
   * Per-hex arrays of sandbox detail rows for homeschool students (`studentHexKey` → rows).
   * Built from homeschool GeoJSON when layers refresh.
   */
  var HOMESCHOOL_DETAILS_BY_HEX_KEY = null;
  /** Canonical attendance MSID for homeschool (district lookup / exports). */
  var HOMESCHOOL_ATTENDANCE_MSID = 9998;
  /** Lazily filled: assignment MSID string → homeschool student count in that polygon (centroid-in-boundary). */
  var homeschoolInBoundaryByMsidCache = Object.create(null);

  function clearHomeschoolInBoundaryCountCache() {
    homeschoolInBoundaryByMsidCache = Object.create(null);
  }
  /**
   * All student residence rows (any MSID): per-hex grade tallies + hex centroids
   * for travel-shed tooltips (centroid-in-isochrone, districtwide).
   */
  var TRAVEL_SHED_RESIDENCE_INDEX = null;
  /** @type {number|undefined} */
  var travelShedResidenceDebounceId = null;
  /** Incremented to drop stale travel-shed count results after rapid cursor moves. */
  var travelShedResidenceHoverGen = 0;
  /** Dropdown- or map-driven selection; kept in sync with #school-select. */
  var selectedSchoolMsid = null;
  /**
   * When a map click applies #school-select, the next `applyExistingSchoolFromSelectValue` run uses
   * this: "centerOnSchool" = pan only; "assignment" = fit assignment (dropdown default, boundary picks).
   */
  var pendingMapSelectFrame = null;
  /** { source, id } for assignment outline emphasis when a school is chosen from the dropdown. */
  var selectedAssignmentBoundary = null;

  /** Scenario Testing: merged K–8 tool state (middle MS + feeder checkboxes). */
  var scenarioSchoolByMsid = null;
  var scenarioMiddleMsid = null;
  var scenarioLastFeederRows = [];
  var scenarioFeederChecked = {};
  /** MSIDs last given map feature-state `scenarioFeeder`; cleared before each update. */
  var lastScenarioFeederHighlightMsids = [];
  /** `{ source, id }` for assignment polygon `feature-state: scenarioRelevant`; cleared in scenario or when leaving the view. */
  var lastScenarioBoundaryRelevant = [];
  /** When true, each selected elementary counts at 100%; when false (default), use flow proportion × enrollment. */
  var scenarioCompleteMerger = false;
  /** Set to false to restore the single aggregated bar chart on the Scenario page. */
  var SCENARIO_USE_STACKED_ENROLLMENT_CHART = true;
  /**
   * Boundary Sandbox: user-selected student hex keys (string) → true. Survives tab changes.
   * `confirmed` reserved for lasso/confirm flow (Milestone 1+).
   */
  var BOUNDARY_SANDBOX = {
    selectedHexKeys: Object.create(null),
    selectionConfirmed: false,
    /** Copy of `selectedHexKeys` at last Confirm; drives sidebar stats while the map selection is edited. */
    confirmedHexKeysSnapshot: Object.create(null),
    /** @type {Object<string, boolean|undefined>} canonical grade key (e.g. "K", "07") → include in att/zoned/demographics */
    gradeToggles: Object.create(null),
    /** @type {Object<string, boolean|undefined>} `zonedTraditional` | `otherTraditional` | `charter` | `choice` → include in att/zoned/demographics (after grade filter) */
    attendanceTypeToggles: Object.create(null),
    /** @type {{ attendance: boolean, zoned: boolean }} */
    schoolListExpanded: { attendance: false, zoned: false },
    /** Single Polygon/MultiPolygon for lasso tint/outline; grows via turf.union (select), shrinks via turf.difference (erase). */
    lassoRegionFootprintFeature: null,
  };

  /**
   * Paint: drag uses Select (add) / Erase (remove); click-to-toggle when pointer did not drag.
   * @type {{ active: boolean, lastKey: string|null, startX: number, startY: number, clickKey: string|null, isDrag: boolean }}
   */
  var BOUNDARY_SANDBOX_PAINT = {
    active: false,
    lastKey: null,
    startX: 0,
    startY: 0,
    clickKey: null,
    isDrag: false,
  };
  /** @const Compare squared distance to 5px drag threshold. */
  var BOUNDARY_SANDBOX_BRUSH_DRAG_THRESH2 = 25;
  /** @type {{ active: boolean, points: [number, number][]|null }} */
  var BOUNDARY_SANDBOX_LASSO = { active: false, points: null };

  /** No-op until setupMapInteractions wires the density tooltip popup. */
  var dismissStudentHexDensityTooltip = function () {};

  /** Show/disable the density-tooltip control when either student or charter residence density is on. */
  function syncStudentHexTooltipCheckboxVisibility() {
    var row = document.getElementById("student-hex-tooltip-row");
    var main = document.getElementById("toggle-student-hex");
    var ch = document.getElementById("toggle-charter-student-hex");
    var hm = document.getElementById("toggle-homeschool-student-hex");
    var tt = document.getElementById("toggle-student-hex-density-tooltip");
    var modeWrap = document.getElementById("student-hex-residence-modes");
    if (!row) return;
    var anyDensityOn =
      (!!main && main.checked) || (!!ch && ch.checked) || (!!hm && hm.checked);
    row.hidden = !anyDensityOn;
    if (tt) {
      tt.disabled = !anyDensityOn;
    }
    if (modeWrap) {
      modeWrap.classList.toggle("student-hex-residence-modes--inactive", !main || !main.checked);
    }
    if (!anyDensityOn) dismissStudentHexDensityTooltip();
    syncMapDensityLegend();
  }

  var mapDensityLegendValueRefreshHandle = null;
  var mapDensityLegendViewListenersSet = false;

  function getMapDensityLegendVisibility() {
    var stuInp = document.getElementById("toggle-student-hex");
    var chInp = document.getElementById("toggle-charter-student-hex");
    var hmInp = document.getElementById("toggle-homeschool-student-hex");
    var stuOn = !!(stuInp && stuInp.checked);
    var chOn = !!(chInp && chInp.checked);
    var hmOn = !!(hmInp && hmInp.checked);
    var stuVis = stuOn;
    var chVis = chOn;
    var hmVis = hmOn;
    if (map && map.getLayer) {
      try {
        if (map.getLayer("student-hex-heatmap")) {
          stuVis =
            stuOn && map.getLayoutProperty("student-hex-heatmap", "visibility") === "visible";
        }
      } catch (e0) {
        /* ignore */
      }
      try {
        if (map.getLayer("charter-student-hex-heatmap")) {
          chVis =
            chOn && map.getLayoutProperty("charter-student-hex-heatmap", "visibility") === "visible";
        }
      } catch (e1) {
        /* ignore */
      }
      try {
        if (map.getLayer("homeschool-student-hex-heatmap")) {
          hmVis =
            hmOn && map.getLayoutProperty("homeschool-student-hex-heatmap", "visibility") === "visible";
        }
      } catch (e2) {
        /* ignore */
      }
    }
    return { stu: stuVis, ch: chVis, hm: hmVis };
  }

  function formatMapLegendStudentsPerSqMi(n) {
    if (n == null || !isFinite(n)) {
      return "—";
    }
    return Math.round(Number(n)).toLocaleString();
  }

  /**
   * Min/max of neighborhood-mean students/sq mi (center hex + adjacents) for each hex
   * centroid in the viewport — matches tooltip / smoothed treatment.
   */
  function minMaxNeighborhoodSchoolDensitiesInViewForLegend() {
    if (!map || !map.getSource || !map.getSource("student-hex")) {
      return { min: null, max: null };
    }
    var b;
    try {
      b = map.getBounds();
    } catch (e) {
      return { min: null, max: null };
    }
    if (!b) {
      return { min: null, max: null };
    }
    var features;
    try {
      features = map.querySourceFeatures("student-hex", {});
    } catch (e2) {
      return { min: null, max: null };
    }
    if (!features || !features.length) {
      return { min: null, max: null };
    }
    var preIdx = buildStudentHexDisplayCountsByHex();
    if (preIdx == null) {
      preIdx = Object.create(null);
    }
    var minC = null;
    var maxC = null;
    for (var i = 0; i < features.length; i++) {
      var f = features[i];
      if (!f || !f.properties) continue;
      var g = f.geometry;
      if (!g || g.type !== "Point" || !g.coordinates) continue;
      var lng = g.coordinates[0];
      var lat = g.coordinates[1];
      if (lng == null || lat == null) continue;
      var ll;
      try {
        ll = new mapboxgl.LngLat(lng, lat);
      } catch (e3) {
        continue;
      }
      if (!b.contains(ll)) {
        continue;
      }
      var c = null;
      var hk = f.properties._hexKey != null ? String(f.properties._hexKey) : null;
      if (hk) {
        c = neighborhoodAverageSchoolResidenceStudentsPerSqMi(hk, preIdx);
      }
      if (c == null || !isFinite(c)) {
        if (f.properties.students_per_sq_mi == null) {
          continue;
        }
        c = Number(f.properties.students_per_sq_mi);
        if (!isFinite(c)) continue;
      }
      if (minC == null || c < minC) {
        minC = c;
      }
      if (maxC == null || c > maxC) {
        maxC = c;
      }
    }
    return { min: minC, max: maxC };
  }

  function minMaxNeighborhoodCharterDensitiesInViewForLegend() {
    if (!map || !map.getSource || !map.getSource("charter-student-hex")) {
      return { min: null, max: null };
    }
    var b;
    try {
      b = map.getBounds();
    } catch (e) {
      return { min: null, max: null };
    }
    if (!b) {
      return { min: null, max: null };
    }
    var features;
    try {
      features = map.querySourceFeatures("charter-student-hex", {});
    } catch (e2) {
      return { min: null, max: null };
    }
    if (!features || !features.length) {
      return { min: null, max: null };
    }
    var preCh =
      (STUDENT_HEX_INDEX && STUDENT_HEX_INDEX.charterDistrictHexCounts) ||
      Object.create(null);
    var minC = null;
    var maxC = null;
    for (var j = 0; j < features.length; j++) {
      var f2 = features[j];
      if (!f2 || !f2.properties) continue;
      var g2 = f2.geometry;
      if (!g2 || g2.type !== "Point" || !g2.coordinates) continue;
      var lng2 = g2.coordinates[0];
      var lat2 = g2.coordinates[1];
      if (lng2 == null || lat2 == null) continue;
      var ll2;
      try {
        ll2 = new mapboxgl.LngLat(lng2, lat2);
      } catch (e3b) {
        continue;
      }
      if (!b.contains(ll2)) {
        continue;
      }
      var c2 = null;
      var hk2 = f2.properties._hexKey != null ? String(f2.properties._hexKey) : null;
      if (hk2) {
        c2 = neighborhoodAverageCharterResidenceStudentsPerSqMi(hk2, preCh);
      }
      if (c2 == null || !isFinite(c2)) {
        if (f2.properties.students_per_sq_mi == null) {
          continue;
        }
        c2 = Number(f2.properties.students_per_sq_mi);
        if (!isFinite(c2)) continue;
      }
      if (minC == null || c2 < minC) {
        minC = c2;
      }
      if (maxC == null || c2 > maxC) {
        maxC = c2;
      }
    }
    return { min: minC, max: maxC };
  }

  function minMaxNeighborhoodHomeschoolDensitiesInViewForLegend() {
    if (!map || !map.getSource || !map.getSource("homeschool-student-hex")) {
      return { min: null, max: null };
    }
    var b;
    try {
      b = map.getBounds();
    } catch (e) {
      return { min: null, max: null };
    }
    if (!b) {
      return { min: null, max: null };
    }
    var features;
    try {
      features = map.querySourceFeatures("homeschool-student-hex", {});
    } catch (e2) {
      return { min: null, max: null };
    }
    if (!features || !features.length) {
      return { min: null, max: null };
    }
    var preHm = HOMESCHOOL_HEX_COUNTS || Object.create(null);
    var minC = null;
    var maxC = null;
    for (var j = 0; j < features.length; j++) {
      var f2 = features[j];
      if (!f2 || !f2.properties) continue;
      var g2 = f2.geometry;
      if (!g2 || g2.type !== "Point" || !g2.coordinates) continue;
      var lng2 = g2.coordinates[0];
      var lat2 = g2.coordinates[1];
      if (lng2 == null || lat2 == null) continue;
      var ll2;
      try {
        ll2 = new mapboxgl.LngLat(lng2, lat2);
      } catch (e3b) {
        continue;
      }
      if (!b.contains(ll2)) {
        continue;
      }
      var c2 = null;
      var hk2 = f2.properties._hexKey != null ? String(f2.properties._hexKey) : null;
      if (hk2) {
        c2 = neighborhoodAverageHomeschoolResidenceStudentsPerSqMi(hk2, preHm);
      }
      if (c2 == null || !isFinite(c2)) {
        if (f2.properties.students_per_sq_mi == null) {
          continue;
        }
        c2 = Number(f2.properties.students_per_sq_mi);
        if (!isFinite(c2)) continue;
      }
      if (minC == null || c2 < minC) {
        minC = c2;
      }
      if (maxC == null || c2 > maxC) {
        maxC = c2;
      }
    }
    return { min: minC, max: maxC };
  }

  function scheduleRefreshMapDensityLegendValueRanges() {
    if (mapDensityLegendValueRefreshHandle) {
      clearTimeout(mapDensityLegendValueRefreshHandle);
    }
    mapDensityLegendValueRefreshHandle = setTimeout(function () {
      mapDensityLegendValueRefreshHandle = null;
      refreshMapDensityLegendValueRanges();
    }, 100);
  }

  function refreshMapDensityLegendValueRanges() {
    var stuMin = document.getElementById("map-density-legend-student-min");
    var stuMax = document.getElementById("map-density-legend-student-max");
    var chMin = document.getElementById("map-density-legend-charter-min");
    var chMax = document.getElementById("map-density-legend-charter-max");
    var hmMin = document.getElementById("map-density-legend-homeschool-min");
    var hmMax = document.getElementById("map-density-legend-homeschool-max");
    if (!stuMin && !chMin && !hmMin) {
      return;
    }
    var v = getMapDensityLegendVisibility();
    if (v.stu && stuMin && stuMax) {
      var r1 = minMaxNeighborhoodSchoolDensitiesInViewForLegend();
      stuMin.textContent = formatMapLegendStudentsPerSqMi(r1.min);
      stuMax.textContent = formatMapLegendStudentsPerSqMi(r1.max);
      var bar = document.getElementById("map-density-legend-student-scale");
      if (bar) {
        bar.setAttribute(
          "aria-label",
          "Color scale: student residences; in current view, neighborhood-mean students per square mile, minimum " +
            (r1.min == null ? "—" : formatMapLegendStudentsPerSqMi(r1.min)) +
            " to maximum " +
            (r1.max == null ? "—" : formatMapLegendStudentsPerSqMi(r1.max))
        );
      }
    } else {
      if (stuMin) stuMin.textContent = "—";
      if (stuMax) stuMax.textContent = "—";
    }
    if (v.ch && chMin && chMax) {
      var r2 = minMaxNeighborhoodCharterDensitiesInViewForLegend();
      chMin.textContent = formatMapLegendStudentsPerSqMi(r2.min);
      chMax.textContent = formatMapLegendStudentsPerSqMi(r2.max);
      var bar2 = document.getElementById("map-density-legend-charter-scale");
      if (bar2) {
        bar2.setAttribute(
          "aria-label",
          "Color scale: charter student residences; in current view, neighborhood-mean students per square mile, minimum " +
            (r2.min == null ? "—" : formatMapLegendStudentsPerSqMi(r2.min)) +
            " to maximum " +
            (r2.max == null ? "—" : formatMapLegendStudentsPerSqMi(r2.max))
        );
      }
    } else {
      if (chMin) chMin.textContent = "—";
      if (chMax) chMax.textContent = "—";
    }
    if (v.hm && hmMin && hmMax) {
      var r3 = minMaxNeighborhoodHomeschoolDensitiesInViewForLegend();
      hmMin.textContent = formatMapLegendStudentsPerSqMi(r3.min);
      hmMax.textContent = formatMapLegendStudentsPerSqMi(r3.max);
      var bar3 = document.getElementById("map-density-legend-homeschool-scale");
      if (bar3) {
        bar3.setAttribute(
          "aria-label",
          "Color scale: homeschool student residences; in current view, neighborhood-mean students per square mile, minimum " +
            (r3.min == null ? "—" : formatMapLegendStudentsPerSqMi(r3.min)) +
            " to maximum " +
            (r3.max == null ? "—" : formatMapLegendStudentsPerSqMi(r3.max))
        );
      }
    } else {
      if (hmMin) hmMin.textContent = "—";
      if (hmMax) hmMax.textContent = "—";
    }
  }

  function setupMapDensityLegendViewListeners() {
    if (mapDensityLegendViewListenersSet || !map) {
      return;
    }
    mapDensityLegendViewListenersSet = true;
    function onView() {
      syncResidenceDensityHeatmapZoomVisibility();
      scheduleRefreshMapDensityLegendValueRanges();
    }
    map.on("moveend", onView);
    map.on("zoomend", onView);
    map.on("resize", onView);
  }

  /**
   * Bottom-right map legend for student / charter residence heatmap scales.
   * Shown when the corresponding layer toggle is on and the heatmap layer is visible.
   */
  function syncMapDensityLegend() {
    var leg = document.getElementById("map-density-legend");
    if (!leg) {
      return;
    }
    var v = getMapDensityLegendVisibility();
    var rowStu = document.getElementById("map-density-legend-student");
    var rowCh = document.getElementById("map-density-legend-charter");
    var rowHm = document.getElementById("map-density-legend-homeschool");
    if (rowStu) {
      rowStu.hidden = !v.stu;
    }
    if (rowCh) {
      rowCh.hidden = !v.ch;
    }
    if (rowHm) {
      rowHm.hidden = !v.hm;
    }
    leg.hidden = !v.stu && !v.ch && !v.hm;
    scheduleRefreshMapDensityLegendValueRanges();
  }

  var ENCHART_COLORS = { calendar: "#94a3b8", projected: "#93c5fd" };

  /** Matches school location dot colors (elementary / middle / high). */
  var PALETTE = {
    /** `highlightStroke`: light tint for the thick selection / hover / scenario ring (same hue family as `fill`). */
    elementary: { fill: "#16a34a", line: "#15803d", highlightStroke: "#4ade80" },
    middle: { fill: "#2563eb", line: "#1d4ed8", highlightStroke: "#93c5fd" },
    high: { fill: "#9333ea", line: "#7e22ce", highlightStroke: "#d8b4fe" },
    /** 7–12 / Jr–Sr schools (distinct from 9–12 high on map and Sankey). */
    jrSr: { fill: "#ea580c", line: "#c2410c", highlightStroke: "#fb923c" },
    charter: { fill: "#ec4899", line: "#be185d", highlightStroke: "#fbcfe8" },
    /** Private schools (non-BPS): golden yellow dot / hover ring. */
    privateSchool: { fill: "#eab308", line: "#ca8a04", highlightStroke: "#fde047" },
  };
  var schoolMapCircleStrokeColorDefault = "#ffffff";

  /** More transparent assignment zone fills */
  var BOUNDARY_FILL_OPACITY = 0.1;

  /** @param {GeoJSON.FeatureCollection} fc */
  function computeBbox(fc) {
    var minX = Infinity;
    var minY = Infinity;
    var maxX = -Infinity;
    var maxY = -Infinity;

    function walk(coords) {
      if (typeof coords[0] === "number") {
        var x = coords[0];
        var y = coords[1];
        if (x < minX) minX = x;
        if (x > maxX) maxX = x;
        if (y < minY) minY = y;
        if (y > maxY) maxY = y;
        return;
      }
      for (var i = 0; i < coords.length; i++) walk(coords[i]);
    }

    if (!fc || !fc.features) return null;
    for (var f = 0; f < fc.features.length; f++) {
      var g = fc.features[f].geometry;
      if (g) walk(g.coordinates);
    }
    if (!isFinite(minX)) return null;
    return [minX, minY, maxX, maxY];
  }

  function mergeBbox(a, b) {
    if (!a) return b;
    if (!b) return a;
    return [
      Math.min(a[0], b[0]),
      Math.min(a[1], b[1]),
      Math.max(a[2], b[2]),
      Math.max(a[3], b[3]),
    ];
  }

  mapboxgl.accessToken = MAPBOX_ACCESS_TOKEN;

  var map = new mapboxgl.Map({
    container: "map",
    style: MAPBOX_STYLES.light,
    center: [-80.7, 28.2],
    zoom: 8,
    maxZoom: 19,
  });

  map.addControl(new mapboxgl.NavigationControl(), "top-left");
  map.addControl(
    new mapboxgl.ScaleControl({
      maxWidth: 120,
      unit: "imperial",
    }),
    "bottom-left"
  );
  map.addControl(new mapboxgl.AttributionControl({ compact: true }), "bottom-right");

  function setMapboxBasemap(mode) {
    if (!MAPBOX_STYLES[mode]) return;
    map.setStyle(MAPBOX_STYLES[mode]);
    var root = document.getElementById("basemap-toggle");
    if (root) {
      root.querySelectorAll("[data-basemap]").forEach(function (btn) {
        var active = btn.getAttribute("data-basemap") === mode;
        btn.classList.toggle("is-active", active);
        btn.setAttribute("aria-pressed", active ? "true" : "false");
      });
    }
  }

  /** After first fetch; GeoJSON layers are re-added on each Mapbox `style.load` (basemap switch). */
  var mapLayersInitialized = false;

  var outlinePaintBase = {
    "line-width": [
      "case",
      [
        "any",
        ["==", ["feature-state", "highlight"], true],
        ["==", ["feature-state", "selectedAssignment"], true],
      ],
      4,
      1,
    ],
    "line-opacity": [
      "case",
      [
        "any",
        ["==", ["feature-state", "highlight"], true],
        ["==", ["feature-state", "selectedAssignment"], true],
      ],
      1,
      0.75,
    ],
  };

  /**
   * For single–school (sparse) views, `heatmap-density` is often 0 – ~0.25 across most of the view while
   * the warm part of the ramp (red → yellow) is concentrated at 0.5 – 1.0. A sublinear power stretches the
   * 0 – 1 range so local peaks use more of the full blue → yellow palette (perceived auto–scaling).
   */
  var HEAT_SCHOOL_DENSITY = [
    "^",
    ["max", 0, ["min", 1, ["heatmap-density"]]],
    0.42,
  ];
  /**
   * District / all–schools view: ^0.45 `heatmap-density` remapping. Single–school (or scenario middle /
   * sandbox) selected: pre–exponent linear stop keys, with HEAT_SCHOOL_DENSITY on the input.
   */
  var HEAT_STUDENT_RAMP_SCHOOL = [
    "interpolate",
    ["linear"],
    HEAT_SCHOOL_DENSITY,
    0,
    "rgba(34, 211, 238, 0)",
    0.018,
    "rgba(34, 211, 238, 0.088)",
    0.045,
    "rgba(20, 198, 225, 0.229)",
    0.08,
    "rgba(8, 172, 198, 0.334)",
    0.12,
    "rgba(6, 155, 182, 0.422)",
    0.16,
    "rgba(6, 182, 212, 0.484)",
    0.2,
    "rgba(56, 189, 248, 0.528)",
    0.213,
    "rgba(70, 150, 244, 0.525)",
    0.225,
    "rgba(85, 120, 238, 0.533)",
    0.238,
    "rgba(98, 88, 230, 0.473)",
    0.25,
    "rgba(105, 58, 220, 0.476)",
    0.26,
    "rgba(109, 40, 217, 0.48)",
    0.29,
    "rgba(128, 46, 225, 0.495)",
    0.32,
    "rgba(147, 51, 234, 0.51)",
    0.35,
    "rgba(158, 68, 240, 0.525)",
    0.38,
    "rgba(168, 85, 247, 0.54)",
    0.41,
    "rgba(180, 62, 230, 0.555)",
    0.44,
    "rgba(192, 38, 211, 0.57)",
    0.47,
    "rgba(205, 38, 180, 0.585)",
    0.485,
    "rgba(212, 38, 150, 0.593)",
    0.5,
    "rgba(219, 39, 119, 0.8)",
    0.56,
    "rgba(225, 29, 72, 0.83)",
    0.62,
    "rgba(220, 38, 38, 0.86)",
    0.68,
    "rgba(234, 88, 12, 0.88)",
    0.74,
    "rgba(245, 101, 20, 0.9)",
    0.8,
    "rgba(251, 146, 60, 0.92)",
    0.86,
    "rgba(253, 186, 55, 0.94)",
    0.91,
    "rgba(253, 224, 71, 0.96)",
    0.95,
    "rgba(254, 240, 138, 0.98)",
    0.98,
    "rgba(255, 251, 200, 0.99)",
    1,
    "rgba(255, 255, 230, 1)",
  ];
  var HEAT_STUDENT_RAMP_UNIFORM = [
    "interpolate",
    ["linear"],
    ["heatmap-density"],
    0,
    "rgba(34, 211, 238, 0)",
    0.164,
    "rgba(34, 211, 238, 0.088)",
    0.2477,
    "rgba(20, 198, 225, 0.229)",
    0.3209,
    "rgba(8, 172, 198, 0.334)",
    0.3852,
    "rgba(6, 155, 182, 0.422)",
    0.4384,
    "rgba(6, 182, 212, 0.484)",
    0.4847,
    "rgba(56, 189, 248, 0.528)",
    0.4986,
    "rgba(70, 150, 244, 0.525)",
    0.5111,
    "rgba(85, 120, 238, 0.533)",
    0.5242,
    "rgba(98, 88, 230, 0.473)",
    0.5359,
    "rgba(105, 58, 220, 0.476)",
    0.5454,
    "rgba(109, 40, 217, 0.48)",
    0.5729,
    "rgba(128, 46, 225, 0.495)",
    0.5988,
    "rgba(147, 51, 234, 0.51)",
    0.6235,
    "rgba(158, 68, 240, 0.525)",
    0.647,
    "rgba(168, 85, 247, 0.54)",
    0.6695,
    "rgba(180, 62, 230, 0.555)",
    0.6911,
    "rgba(192, 38, 211, 0.57)",
    0.7119,
    "rgba(205, 38, 180, 0.585)",
    0.7221,
    "rgba(212, 38, 150, 0.593)",
    0.732,
    "rgba(219, 39, 119, 0.8)",
    0.7703,
    "rgba(225, 29, 72, 0.83)",
    0.8064,
    "rgba(220, 38, 38, 0.86)",
    0.8407,
    "rgba(234, 88, 12, 0.88)",
    0.8733,
    "rgba(245, 101, 20, 0.9)",
    0.9045,
    "rgba(251, 146, 60, 0.92)",
    0.9344,
    "rgba(253, 186, 55, 0.94)",
    0.9584,
    "rgba(253, 224, 71, 0.96)",
    0.9772,
    "rgba(254, 240, 138, 0.98)",
    0.9909,
    "rgba(255, 251, 200, 0.99)",
    1,
    "rgba(255, 255, 230, 1)",
  ];
  var HEAT_CHARTER_RAMP_SCHOOL = [
    "interpolate",
    ["linear"],
    HEAT_SCHOOL_DENSITY,
    0,
    "rgba(255, 255, 255, 0)",
    0.04,
    "rgba(252, 197, 231, 0.14)",
    0.1,
    "rgba(252, 185, 227, 0.26)",
    0.18,
    "rgba(252, 171, 222, 0.38)",
    0.27,
    "rgba(253, 156, 216, 0.48)",
    0.37,
    "rgba(253, 142, 211, 0.58)",
    0.48,
    "rgba(254, 128, 206, 0.68)",
    0.6,
    "rgba(254, 112, 200, 0.78)",
    0.72,
    "rgba(254, 112, 200, 0.78)",
    0.88,
    "rgba(255, 92, 192, 0.86)",
    0.95,
    "rgba(255, 81, 218, 0.9)",
    1,
    "rgba(255, 64, 255, 0.94)",
  ];
  var HEAT_CHARTER_RAMP_UNIFORM = [
    "interpolate",
    ["linear"],
    ["heatmap-density"],
    0,
    "rgba(255, 255, 255, 0)",
    0.2349,
    "rgba(252, 197, 231, 0.14)",
    0.3548,
    "rgba(252, 185, 227, 0.26)",
    0.4622,
    "rgba(252, 171, 222, 0.38)",
    0.5548,
    "rgba(253, 156, 216, 0.48)",
    0.6393,
    "rgba(253, 142, 211, 0.58)",
    0.7187,
    "rgba(254, 128, 206, 0.68)",
    0.7946,
    "rgba(254, 112, 200, 0.78)",
    0.8626,
    "rgba(254, 112, 200, 0.78)",
    0.9441,
    "rgba(255, 92, 192, 0.86)",
    0.9772,
    "rgba(255, 81, 218, 0.9)",
    1,
    "rgba(255, 64, 255, 0.94)",
  ];
  /** Same structure as charter ramp; red/orange family for homeschool residential density. */
  var HEAT_HOMESCHOOL_RAMP_SCHOOL = [
    "interpolate",
    ["linear"],
    HEAT_SCHOOL_DENSITY,
    0,
    "rgba(255, 255, 255, 0)",
    0.04,
    "rgba(254, 226, 226, 0.14)",
    0.1,
    "rgba(252, 165, 165, 0.26)",
    0.18,
    "rgba(248, 113, 113, 0.38)",
    0.27,
    "rgba(239, 68, 68, 0.48)",
    0.37,
    "rgba(220, 38, 38, 0.58)",
    0.48,
    "rgba(185, 28, 28, 0.68)",
    0.6,
    "rgba(153, 27, 27, 0.78)",
    0.72,
    "rgba(153, 27, 27, 0.78)",
    0.88,
    "rgba(127, 29, 29, 0.86)",
    0.95,
    "rgba(91, 17, 17, 0.9)",
    1,
    "rgba(69, 10, 10, 0.94)",
  ];
  var HEAT_HOMESCHOOL_RAMP_UNIFORM = [
    "interpolate",
    ["linear"],
    ["heatmap-density"],
    0,
    "rgba(255, 255, 255, 0)",
    0.2349,
    "rgba(254, 226, 226, 0.14)",
    0.3548,
    "rgba(252, 165, 165, 0.26)",
    0.4622,
    "rgba(248, 113, 113, 0.38)",
    0.5548,
    "rgba(239, 68, 68, 0.48)",
    0.6393,
    "rgba(220, 38, 38, 0.58)",
    0.7187,
    "rgba(185, 28, 28, 0.68)",
    0.7946,
    "rgba(153, 27, 27, 0.78)",
    0.8626,
    "rgba(153, 27, 27, 0.78)",
    0.9441,
    "rgba(127, 29, 29, 0.86)",
    0.9772,
    "rgba(91, 17, 17, 0.9)",
    1,
    "rgba(69, 10, 10, 0.94)",
  ];

  /** Default zoom–scaled heat for student + charter residence heatmaps; restored when leaving school context. */
  var HEAT_RESIDENCE_INTENSITY = [
    "interpolate",
    ["linear"],
    ["zoom"],
    8,
    0.05,
    10,
    0.07,
    12,
    0.1,
    14,
    0.16,
    16,
    0.24,
    17,
    0.3,
  ];

  /**
   * Hide student / charter residence-density heatmaps (and density hover tooltips) at neighborhood scale
   * and closer. Higher zoom level number = more zoomed in; this threshold is one step further zoomed out than z14.
   * Hex hit-fill layers stay visible for hover tooltips on the map (without density popup when zoomed in).
   */
  var RESIDENCE_HEATMAP_HIDE_ZOOM = 13;

  function residenceDensityHeatmapHiddenAtCurrentZoom() {
    if (!map || typeof map.getZoom !== "function") {
      return false;
    }
    try {
      return map.getZoom() >= RESIDENCE_HEATMAP_HIDE_ZOOM;
    } catch (eZ) {
      return false;
    }
  }

  /**
   * Match heatmap visibility to hit-fill visibility, except heatmaps are hidden when zoomed in past
   * `RESIDENCE_HEATMAP_HIDE_ZOOM`.
   */
  function syncResidenceDensityHeatmapZoomVisibility() {
    if (!map || !map.getLayer) {
      return;
    }
    var hideHeat = residenceDensityHeatmapHiddenAtCurrentZoom();
    var stuHitOk = false;
    var chHitOk = false;
    var hmHitOk = false;
    try {
      if (map.getLayer("student-hex-hit-fill")) {
        stuHitOk = map.getLayoutProperty("student-hex-hit-fill", "visibility") === "visible";
      }
    } catch (e0) {
      /* ignore */
    }
    try {
      if (map.getLayer("charter-student-hex-hit-fill")) {
        chHitOk = map.getLayoutProperty("charter-student-hex-hit-fill", "visibility") === "visible";
      }
    } catch (e1) {
      /* ignore */
    }
    try {
      if (map.getLayer("homeschool-student-hex-hit-fill")) {
        hmHitOk =
          map.getLayoutProperty("homeschool-student-hex-hit-fill", "visibility") === "visible";
      }
    } catch (e1b) {
      /* ignore */
    }
    var stuHm = stuHitOk && !hideHeat ? "visible" : "none";
    var chHm = chHitOk && !hideHeat ? "visible" : "none";
    var hmHm = hmHitOk && !hideHeat ? "visible" : "none";
    try {
      if (map.getLayer("student-hex-heatmap")) {
        map.setLayoutProperty("student-hex-heatmap", "visibility", stuHm);
      }
      if (map.getLayer("charter-student-hex-heatmap")) {
        map.setLayoutProperty("charter-student-hex-heatmap", "visibility", chHm);
      }
      if (map.getLayer("homeschool-student-hex-heatmap")) {
        map.setLayoutProperty("homeschool-student-hex-heatmap", "visibility", hmHm);
      }
    } catch (eL) {
      /* ignore */
    }
    if (hideHeat && typeof dismissStudentHexDensityTooltip === "function") {
      dismissStudentHexDensityTooltip();
    }
    syncMapDensityLegend();
  }

  function applyResidenceHeatmapSymbology() {
    if (!map || !map.getLayer) {
      return;
    }
    var m = getActiveDashboardSchoolMsid();
    var useOriginalRamp = m != null && !isNaN(m);
    var intExpr = useOriginalRamp
      ? ["*", 1.4, HEAT_RESIDENCE_INTENSITY]
      : HEAT_RESIDENCE_INTENSITY;
    try {
      if (map.getLayer("student-hex-heatmap")) {
        map.setPaintProperty(
          "student-hex-heatmap",
          "heatmap-color",
          useOriginalRamp ? HEAT_STUDENT_RAMP_SCHOOL : HEAT_STUDENT_RAMP_UNIFORM
        );
        map.setPaintProperty("student-hex-heatmap", "heatmap-intensity", intExpr);
      }
      if (map.getLayer("charter-student-hex-heatmap")) {
        map.setPaintProperty(
          "charter-student-hex-heatmap",
          "heatmap-color",
          useOriginalRamp ? HEAT_CHARTER_RAMP_SCHOOL : HEAT_CHARTER_RAMP_UNIFORM
        );
        map.setPaintProperty("charter-student-hex-heatmap", "heatmap-intensity", intExpr);
      }
      if (map.getLayer("homeschool-student-hex-heatmap")) {
        map.setPaintProperty(
          "homeschool-student-hex-heatmap",
          "heatmap-color",
          useOriginalRamp ? HEAT_HOMESCHOOL_RAMP_SCHOOL : HEAT_HOMESCHOOL_RAMP_UNIFORM
        );
        map.setPaintProperty("homeschool-student-hex-heatmap", "heatmap-intensity", intExpr);
      }
    } catch (eHmap) {
      /* ignore */
    }
  }

  /**
   * When `style.load` runs more than once without `setStyle` (can happen during init),
   * sources already exist — update data in place instead of `addSource` (which throws).
   */
  function refreshGeoJsonSourcesAfterStyleReload(results, opts) {
    var fitBounds = !opts || opts.fitBounds !== false;
    var es = results[0];
    var ms = results[1];
    var hs = results[2];
    var schools = enrichSchoolsFcWithMasterType(results[3]);
    CHOICE_SCHOOL_MSIDS = buildChoiceSchoolMsidSet(schools);
    var studentHexFc = results[6];
    var schoolParcelsRaw = results[7];
    var schoolBoardFc = results[8];
    var charterFc = results[9];
    var municipalFc = results[11];
    var privateFc = filterZeroEnrollmentPrivateSchoolsFc(results[16]);
    var homeschoolFc = results[17];
    CHARTER_SCHOOL_MSIDS = buildCharterSchoolMsidSet(schools, charterFc);

    if (studentHexFc && studentHexFc.features && studentHexFc.features.length) {
      STUDENT_HEX_INDEX = buildStudentHexIndex(studentHexFc);
      TRAVEL_SHED_RESIDENCE_INDEX = buildTravelShedResidenceIndex(studentHexFc);
    } else {
      STUDENT_HEX_INDEX = null;
      TRAVEL_SHED_RESIDENCE_INDEX = null;
    }
    HOMESCHOOL_HEX_COUNTS = buildHomeschoolHexCounts(
      homeschoolFc && homeschoolFc.features ? homeschoolFc : null
    );
    HOMESCHOOL_HEX_GEOMETRY_FALLBACK = buildHomeschoolHexGeometryFallback(
      homeschoolFc && homeschoolFc.features ? homeschoolFc : null
    );
    HOMESCHOOL_DETAILS_BY_HEX_KEY = buildHomeschoolDetailsByHexKey(
      homeschoolFc && homeschoolFc.features ? homeschoolFc : null
    );
    clearHomeschoolInBoundaryCountCache();

    GEO_CACHE.es = es;
    GEO_CACHE.ms = ms;
    GEO_CACHE.hs = hs;
    GEO_CACHE.schools = schools;

    var schoolParcelsFc = buildFilteredSchoolParcelsFc(schools, schoolParcelsRaw);

    map.getSource("es-boundaries").setData(es);
    map.getSource("ms-boundaries").setData(ms);
    map.getSource("hs-boundaries").setData(hs);
    map.getSource("schools").setData(schools);
    map.getSource("school-board-districts").setData(
      schoolBoardFc || { type: "FeatureCollection", features: [] }
    );
    map.getSource("municipal-boundaries").setData(
      municipalFc || { type: "FeatureCollection", features: [] }
    );
    map.getSource("school-parcels").setData(schoolParcelsFc);
    map.getSource("charter-schools").setData(
      charterFc || { type: "FeatureCollection", features: [] }
    );
    if (map.getSource("private-schools")) {
      map.getSource("private-schools").setData(
        privateFc || { type: "FeatureCollection", features: [] }
      );
    }
    SCHOOL_ISOCHRONES_ENRICHED = buildSchoolIsochronesEnriched(
      results[14] || { type: "FeatureCollection", features: [] }
    );
    if (map.getSource("school-isochrones")) {
      map.getSource("school-isochrones").setData(
        SCHOOL_ISOCHRONES_ENRICHED || {
          type: "FeatureCollection",
          features: [],
        }
      );
    }
    map.getSource("student-hex").setData({
      type: "FeatureCollection",
      features: [],
    });
    if (map.getSource("student-hex-hit")) {
      map.getSource("student-hex-hit").setData({
        type: "FeatureCollection",
        features: [],
      });
    }
    if (map.getSource("charter-student-hex")) {
      map.getSource("charter-student-hex").setData({
        type: "FeatureCollection",
        features: [],
      });
    }
    if (map.getSource("charter-student-hex-hit")) {
      map.getSource("charter-student-hex-hit").setData({
        type: "FeatureCollection",
        features: [],
      });
    }
    if (map.getSource("homeschool-student-hex")) {
      map.getSource("homeschool-student-hex").setData({
        type: "FeatureCollection",
        features: [],
      });
    }
    if (map.getSource("homeschool-student-hex-hit")) {
      map.getSource("homeschool-student-hex-hit").setData({
        type: "FeatureCollection",
        features: [],
      });
    }

    map.resize();
    var combined = null;
    combined = mergeBbox(combined, computeBbox(es));
    combined = mergeBbox(combined, computeBbox(ms));
    combined = mergeBbox(combined, computeBbox(hs));
    combined = mergeBbox(combined, computeBbox(schools));
    combined = mergeBbox(combined, computeBbox(schoolParcelsFc));
    combined = mergeBbox(combined, computeBbox(charterFc));
    combined = mergeBbox(combined, computeBbox(privateFc));
    if (fitBounds && combined) {
      map.fitBounds(combined, { padding: 48, maxZoom: 12, duration: 0 });
    }
    requestAnimationFrame(function () {
      map.resize();
    });

    if (!mapLayersInitialized) {
      mapLayersInitialized = true;
      var schoolByMsid = buildSchoolLookup(schools);
      populateSchoolSelect(schools);
      populateScenarioSchoolSelect(schools);
      setupToggles();
      setupMapInteractions(schoolByMsid);
      setupSchoolSelection(schoolByMsid);
      setupScenarioSchoolSelection(schoolByMsid, schools);
      initDashboardResizer(map);
      clearSelectedSchoolHighlight();
      syncStudentHexLayer();
      renderEnrollmentChart(null);
      renderDemographicsCharts(null);
    } else {
      syncStudentHexLayer();
      refreshAssignmentBoundaryHighlight();
      if (selectedSchoolMsid != null) {
        try {
          map.setFeatureState(
            { source: "schools", id: selectedSchoolMsid },
            { selected: true }
          );
        } catch (e) {
          /* ignore */
        }
      }
    }
    applyScenarioFeederMapHighlights();
    syncTravelShedLayerFilter();
    rebuildBoundarySandboxHexSourceFromIndex();
    syncBoundarySandboxMapLayers();
    applyResidenceHeatmapSymbology();
  }

  /**
   * @param {string|undefined} name e.g. "2191 : 0 - 47520" → MSID 2191, 47520 ft
   * @returns {{ msid: number, toBreakFt: number }|null}
   */
  function parseSchoolIsochroneName(name) {
    if (name == null || name === "") return null;
    var m = String(name).trim().match(/^(\d+)\s*:\s*0\s*-\s*(\d+)\s*$/i);
    if (!m) return null;
    var msid = parseInt(m[1], 10);
    var toBreakFt = parseInt(m[2], 10);
    if (isNaN(msid) || isNaN(toBreakFt)) return null;
    return { msid: msid, toBreakFt: toBreakFt };
  }

  /**
   * Enriches Esri export: ToBreak in feet, Name encodes same; adds iso_msid, iso_miles (1–10) for filter/paint.
   * Feature order: larger network distance first so smaller (inner) rings draw on top within one fill layer.
   * @param {Object|null} fc
   * @returns {Object}
   */
  function buildSchoolIsochronesEnriched(fc) {
    if (!fc || !fc.features || !fc.features.length) {
      return { type: "FeatureCollection", features: [] };
    }
    var out = [];
    for (var i = 0; i < fc.features.length; i++) {
      var f = fc.features[i];
      if (!f) continue;
      var p = f.properties || {};
      var rawName = p.Name != null ? p.Name : p.name;
      var parsed = parseSchoolIsochroneName(rawName);
      if (!parsed) continue;
      var toBreak =
        p.ToBreak != null && p.ToBreak !== ""
          ? Number(p.ToBreak)
          : parsed.toBreakFt;
      if (isNaN(toBreak) || toBreak < 0) toBreak = parsed.toBreakFt;
      var miles = Math.round(toBreak / FEET_PER_MILE);
      if (miles < 1) {
        miles = 1;
      } else if (miles > 10) {
        miles = 10;
      }
      var pr = Object.assign({}, p, {
        iso_msid: parsed.msid,
        iso_break_ft: toBreak,
        iso_miles: miles,
      });
      out.push({ type: "Feature", geometry: f.geometry, properties: pr });
    }
    out.sort(function (a, b) {
      return (b.properties.iso_break_ft || 0) - (a.properties.iso_break_ft || 0);
    });
    return { type: "FeatureCollection", features: out };
  }

  /** Middle school in scenario panel, else #school-select MSID. */
  function getActiveTravelShedMsid() {
    if (isBoundarySandboxViewActive()) {
      return getSandboxBaseSchoolMsid();
    }
    var panel = document.getElementById("page-scenario");
    if (panel && !panel.hidden) {
      if (scenarioMiddleMsid != null && !isNaN(scenarioMiddleMsid)) {
        return scenarioMiddleMsid;
      }
      return null;
    }
    return selectedSchoolMsid;
  }

  /** Upper bound in miles (1–10) for isochrones shown when Travel sheds is on; controlled by #travel-shed-max-miles. */
  var travelShedMaxMiles = 10;

  function syncTravelShedMaxMilesRowVisibility() {
    var row = document.getElementById("travel-shed-max-miles-row");
    var tgl = document.getElementById("toggle-travel-sheds");
    if (!row) {
      return;
    }
    if (!tgl) {
      row.setAttribute("hidden", "");
      return;
    }
    if (tgl.checked) {
      row.removeAttribute("hidden");
    } else {
      row.setAttribute("hidden", "");
    }
  }

  function formatTravelShedMilesOutput(miles) {
    var m = Math.round(miles);
    if (m === 1) return "1 mi";
    return m + " mi";
  }

  function updateTravelShedMilesFromRangeControl() {
    var range = document.getElementById("travel-shed-max-miles");
    var out = document.getElementById("travel-shed-max-miles-output");
    if (!range) {
      return;
    }
    var v = Number(range.value);
    if (isNaN(v) || v < 1) v = 1;
    if (v > 10) v = 10;
    travelShedMaxMiles = v;
    range.setAttribute("aria-valuenow", String(v));
    if (out) {
      out.textContent = formatTravelShedMilesOutput(v);
      out.value = out.textContent;
    }
    range.setAttribute("aria-valuetext", v === 1 ? "1 mile" : v + " miles");
  }

  function setupTravelShedMaxMilesControl() {
    var range = document.getElementById("travel-shed-max-miles");
    if (!range) {
      return;
    }
    updateTravelShedMilesFromRangeControl();
    range.addEventListener("input", function () {
      updateTravelShedMilesFromRangeControl();
      syncTravelShedLayerFilter();
    });
  }

  function syncTravelShedLayerFilter() {
    if (!map || !map.getSource || !map.getSource("school-isochrones")) {
      return;
    }
    var full = SCHOOL_ISOCHRONES_ENRICHED;
    var outFc = { type: "FeatureCollection", features: [] };
    if (full && full.features && full.features.length) {
      var ms = getActiveTravelShedMsid();
      if (ms != null && !isNaN(ms)) {
        var mNum = Number(ms);
        var maxM = travelShedMaxMiles;
        if (isNaN(maxM) || maxM < 1) maxM = 10;
        if (maxM > 10) maxM = 10;
        var matched = [];
        for (var i = 0; i < full.features.length; i++) {
          var f0 = full.features[i];
          if (!f0 || !f0.properties) continue;
          if (Number(f0.properties.iso_msid) !== mNum) continue;
          var ringMi0 = Number(f0.properties.iso_miles);
          if (isNaN(ringMi0)) continue;
          if (ringMi0 <= maxM) {
            matched.push(f0);
          }
        }
        var maxRing = -1;
        for (var j = 0; j < matched.length; j++) {
          var rj = Number(
            matched[j].properties != null
              ? matched[j].properties.iso_miles
              : NaN
          );
          if (!isNaN(rj) && rj > maxRing) {
            maxRing = rj;
          }
        }
        for (var k = 0; k < matched.length; k++) {
          var fk = matched[k];
          var pk = fk.properties || {};
          var rk = Number(pk.iso_miles);
          var isO =
            !isNaN(rk) && maxRing >= 0 && rk === maxRing;
          outFc.features.push({
            type: "Feature",
            geometry: fk.geometry,
            properties: Object.assign({}, pk, { iso_outer: isO ? "yes" : "no" }),
          });
        }
      }
    }
    try {
      map.getSource("school-isochrones").setData(outFc);
    } catch (e) {
      /* ignore */
    }
    /* Avoid layer filters: Mapbox v3 is unreliable for iso_msid matching on GeoJSON. Show all in source. */
    if (map.getLayer("school-isochrones-fill")) {
      try {
        map.setFilter("school-isochrones-fill", null);
      } catch (e2) {
        /* ignore */
      }
    }
  }

  function applyGeoJsonLayersFromFetchResults(results, opts) {
    var fitBounds = !opts || opts.fitBounds !== false;
    var es = results[0];
    var ms = results[1];
    var hs = results[2];
    var schools = enrichSchoolsFcWithMasterType(results[3]);
    CHOICE_SCHOOL_MSIDS = buildChoiceSchoolMsidSet(schools);
    var studentHexFc = results[6];
    var schoolParcelsRaw = results[7];
    var schoolBoardFc = results[8];
    var charterFc = results[9];
    var municipalFc = results[11];
    var privateFc = filterZeroEnrollmentPrivateSchoolsFc(results[16]);
    var homeschoolFc = results[17];
    CHARTER_SCHOOL_MSIDS = buildCharterSchoolMsidSet(schools, charterFc);
    SCHOOL_ISOCHRONES_ENRICHED = buildSchoolIsochronesEnriched(
      results[14] || { type: "FeatureCollection", features: [] }
    );

    if (map.getSource("es-boundaries")) {
      refreshGeoJsonSourcesAfterStyleReload(results, {
        fitBounds: fitBounds,
      });
      return;
    }

    var boundarySourceOpts = { type: "geojson", promoteId: "MSID" };

        map.addSource("es-boundaries", Object.assign({ data: es }, boundarySourceOpts));
        map.addSource("ms-boundaries", Object.assign({ data: ms }, boundarySourceOpts));
        map.addSource("hs-boundaries", Object.assign({ data: hs }, boundarySourceOpts));
        map.addSource("schools", {
          type: "geojson",
          data: schools,
          promoteId: "SCHOOLS_ID",
        });
        map.addSource("municipal-boundaries", {
          type: "geojson",
          data: municipalFc || { type: "FeatureCollection", features: [] },
        });

        map.addSource("school-board-districts", {
          type: "geojson",
          data: schoolBoardFc || { type: "FeatureCollection", features: [] },
          promoteId: "OBJECTID",
        });
        map.addLayer({
          id: "school-board-districts-fill",
          type: "fill",
          source: "school-board-districts",
          paint: {
            "fill-color": "#000000",
            "fill-opacity": 0,
          },
          layout: { visibility: "none" },
        });
        map.addLayer({
          id: "school-board-districts-outline",
          type: "line",
          source: "school-board-districts",
          paint: {
            "line-color": "#374151",
            "line-width": [
              "interpolate",
              ["linear"],
              ["zoom"],
              8,
              2,
              12,
              2.5,
              16,
              3.5,
            ],
            "line-opacity": 0.95,
          },
          layout: { visibility: "none" },
        });

        map.addLayer({
          id: "municipal-boundaries-fill",
          type: "fill",
          source: "municipal-boundaries",
          paint: {
            "fill-color": "#000000",
            "fill-opacity": 0,
          },
          layout: { visibility: "none" },
        });
        map.addLayer({
          id: "municipal-boundaries-outline",
          type: "line",
          source: "municipal-boundaries",
          paint: {
            "line-color": "#9ca3af",
            "line-width": [
              "interpolate",
              ["linear"],
              ["zoom"],
              8,
              1.35,
              12,
              1.75,
              16,
              2.5,
            ],
            "line-opacity": 0.95,
          },
          layout: { visibility: "none" },
        });
        /** Hover stroke only: filter toggled on mousemove. Placed above assignment fills after `moveLayer` below. */
        map.addLayer({
          id: "municipal-boundaries-hover",
          type: "line",
          source: "municipal-boundaries",
          filter: ["==", ["to-string", ["get", "OBJECTID"]], MUN_HOVER_FILTER_OFF],
          paint: {
            "line-color": "#374151",
            "line-width": [
              "interpolate",
              ["linear"],
              ["zoom"],
              8,
              2,
              12,
              2.5,
              16,
              3.5,
            ],
            "line-opacity": 1,
          },
          layout: { visibility: "none" },
        });

        map.addLayer({
          id: "hs-fill",
          type: "fill",
          source: "hs-boundaries",
          paint: {
            "fill-color": PALETTE.high.fill,
            "fill-opacity": BOUNDARY_FILL_OPACITY,
          },
        });
        map.addLayer({
          id: "hs-outline",
          type: "line",
          source: "hs-boundaries",
          paint: Object.assign({}, outlinePaintBase, {
            "line-color": PALETTE.high.line,
          }),
        });
        map.addLayer({
          id: "ms-fill",
          type: "fill",
          source: "ms-boundaries",
          paint: {
            "fill-color": PALETTE.middle.fill,
            "fill-opacity": BOUNDARY_FILL_OPACITY,
          },
        });
        map.addLayer({
          id: "ms-outline",
          type: "line",
          source: "ms-boundaries",
          paint: Object.assign({}, outlinePaintBase, {
            "line-color": PALETTE.middle.line,
          }),
        });
        map.addLayer({
          id: "es-fill",
          type: "fill",
          source: "es-boundaries",
          paint: {
            "fill-color": PALETTE.elementary.fill,
            "fill-opacity": BOUNDARY_FILL_OPACITY,
          },
        });
        map.addLayer({
          id: "es-outline",
          type: "line",
          source: "es-boundaries",
          paint: Object.assign({}, outlinePaintBase, {
            "line-color": PALETTE.elementary.line,
          }),
        });

        var schoolParcelsFc = buildFilteredSchoolParcelsFc(
          schools,
          schoolParcelsRaw
        );
        map.addSource("school-parcels", {
          type: "geojson",
          data: schoolParcelsFc,
        });
        var schoolParcelLineLayout = { visibility: "visible" };
        var schoolParcelLinePaintBase = {
          "line-width": 1.5,
          "line-opacity": 0.9,
          "line-dasharray": [4, 3],
        };
        map.addLayer({
          id: "school-parcels-high",
          type: "line",
          source: "school-parcels",
          filter: ["==", ["get", "_parcelLevel"], "high"],
          layout: schoolParcelLineLayout,
          paint: Object.assign({}, schoolParcelLinePaintBase, {
            "line-color": PALETTE.high.line,
          }),
        });
        map.addLayer({
          id: "school-parcels-middle",
          type: "line",
          source: "school-parcels",
          filter: ["==", ["get", "_parcelLevel"], "middle"],
          layout: schoolParcelLineLayout,
          paint: Object.assign({}, schoolParcelLinePaintBase, {
            "line-color": PALETTE.middle.line,
          }),
        });
        map.addLayer({
          id: "school-parcels-jr-sr",
          type: "line",
          source: "school-parcels",
          filter: ["==", ["get", "_parcelLevel"], "jr_sr"],
          layout: schoolParcelLineLayout,
          paint: Object.assign({}, schoolParcelLinePaintBase, {
            "line-color": PALETTE.jrSr.line,
          }),
        });
        map.addLayer({
          id: "school-parcels-elementary",
          type: "line",
          source: "school-parcels",
          filter: ["==", ["get", "_parcelLevel"], "elementary"],
          layout: schoolParcelLineLayout,
          paint: Object.assign({}, schoolParcelLinePaintBase, {
            "line-color": PALETTE.elementary.line,
          }),
        });

        map.addSource("school-isochrones", {
          type: "geojson",
          /* Empty until `syncTravelShedLayerFilter` (Mapbox v3 rejects legacy filter ["==", 1, 0]). */
          data: { type: "FeatureCollection", features: [] },
        });
        map.addLayer({
          id: "school-isochrones-fill",
          type: "fill",
          source: "school-isochrones",
          paint: {
            "fill-color": [
              "match",
              ["to-number", ["get", "iso_miles"]],
              1,
              "#fffbeb",
              2,
              "#fef3c7",
              3,
              "#fde68a",
              4,
              "#fcd34d",
              5,
              "#fbbf24",
              6,
              "#d97706",
              7,
              "#b45309",
              8,
              "#92400e",
              9,
              "#78350f",
              10,
              "#451a03",
              "rgba(212, 212, 216, 0.35)",
            ],
            "fill-opacity": [
              "match",
              ["to-number", ["get", "iso_miles"]],
              1,
              0.52,
              2,
              0.46,
              3,
              0.4,
              4,
              0.35,
              5,
              0.3,
              6,
              0.25,
              7,
              0.2,
              8,
              0.16,
              9,
              0.12,
              10,
              0.1,
              0.2,
            ],
          },
          layout: { visibility: "none" },
        });
        map.addLayer({
          id: "school-isochrones-outline",
          type: "line",
          source: "school-isochrones",
          filter: ["==", ["get", "iso_outer"], "yes"],
          paint: {
            "line-color": "#5c2e0e",
            "line-width": [
              "interpolate",
              ["linear"],
              ["zoom"],
              9,
              1.75,
              12,
              2.5,
              16,
              3.5,
            ],
            "line-opacity": 0.95,
          },
          layout: { visibility: "none" },
        });

        map.addSource("student-hex", {
          type: "geojson",
          data: { type: "FeatureCollection", features: [] },
        });
        map.addLayer({
          id: "student-hex-heatmap",
          type: "heatmap",
          source: "student-hex",
          paint: {
            /**
             * Sqrt of count shrinks the gap between large weights so a few very high–count hexes do not
             * wash the whole district into the high end of `heatmap-density` after Mapbox’s normalization.
             */
            "heatmap-weight": [
              "max",
              0,
              [
                "sqrt",
                [
                  "max",
                  0,
                  ["to-number", ["get", "count"]],
                ],
              ],
            ],
            "heatmap-intensity": HEAT_RESIDENCE_INTENSITY,
            /** Tighter kernel at z16+ = sharper local peaks when zoomed in. */
            "heatmap-radius": [
              "interpolate",
              ["linear"],
              ["zoom"],
              8,
              16,
              11,
              30,
              14,
              42,
              16,
              32,
              17,
              28,
            ],
            "heatmap-opacity": 0.88,
            /* Default: district 0.45 remapped keys; per-school view overrides in applyResidenceHeatmapSymbology. */
            "heatmap-color": HEAT_STUDENT_RAMP_UNIFORM,
          },
          layout: { visibility: "none" },
        });

        map.addSource("student-hex-hit", {
          type: "geojson",
          data: { type: "FeatureCollection", features: [] },
        });
        map.addLayer({
          id: "student-hex-hit-fill",
          type: "fill",
          source: "student-hex-hit",
          paint: {
            "fill-opacity": 0,
            "fill-color": "#000000",
          },
          layout: { visibility: "none" },
        });

        map.addSource("charter-student-hex", {
          type: "geojson",
          data: { type: "FeatureCollection", features: [] },
        });
        map.addLayer({
          id: "charter-student-hex-heatmap",
          type: "heatmap",
          source: "charter-student-hex",
          paint: {
            "heatmap-weight": [
              "max",
              0,
              [
                "sqrt",
                [
                  "max",
                  0,
                  ["to-number", ["get", "count"]],
                ],
              ],
            ],
            "heatmap-intensity": HEAT_RESIDENCE_INTENSITY,
            "heatmap-radius": [
              "interpolate",
              ["linear"],
              ["zoom"],
              8,
              16,
              11,
              30,
              14,
              42,
              16,
              32,
              17,
              28,
            ],
            "heatmap-opacity": 0.88,
            "heatmap-color": HEAT_CHARTER_RAMP_UNIFORM,
          },
          layout: { visibility: "none" },
        });
        applyResidenceHeatmapSymbology();
        map.addSource("charter-student-hex-hit", {
          type: "geojson",
          data: { type: "FeatureCollection", features: [] },
        });
        map.addLayer({
          id: "charter-student-hex-hit-fill",
          type: "fill",
          source: "charter-student-hex-hit",
          paint: {
            "fill-opacity": 0,
            "fill-color": "#000000",
          },
          layout: { visibility: "none" },
        });

        map.addSource("homeschool-student-hex", {
          type: "geojson",
          data: { type: "FeatureCollection", features: [] },
        });
        map.addLayer({
          id: "homeschool-student-hex-heatmap",
          type: "heatmap",
          source: "homeschool-student-hex",
          paint: {
            "heatmap-weight": [
              "max",
              0,
              [
                "sqrt",
                [
                  "max",
                  0,
                  ["to-number", ["get", "count"]],
                ],
              ],
            ],
            "heatmap-intensity": HEAT_RESIDENCE_INTENSITY,
            "heatmap-radius": [
              "interpolate",
              ["linear"],
              ["zoom"],
              8,
              16,
              11,
              30,
              14,
              42,
              16,
              32,
              17,
              28,
            ],
            "heatmap-opacity": 0.88,
            "heatmap-color": HEAT_HOMESCHOOL_RAMP_UNIFORM,
          },
          layout: { visibility: "none" },
        });
        map.addSource("homeschool-student-hex-hit", {
          type: "geojson",
          data: { type: "FeatureCollection", features: [] },
        });
        map.addLayer({
          id: "homeschool-student-hex-hit-fill",
          type: "fill",
          source: "homeschool-student-hex-hit",
          paint: {
            "fill-opacity": 0,
            "fill-color": "#000000",
          },
          layout: { visibility: "none" },
        });

        map.addSource("boundary-sandbox-lasso-region-fill", {
          type: "geojson",
          data: { type: "FeatureCollection", features: [] },
        });
        map.addLayer({
          id: "boundary-sandbox-lasso-region-fill",
          type: "fill",
          source: "boundary-sandbox-lasso-region-fill",
          paint: {
            "fill-color": "#84cc16",
            "fill-opacity": 0.14,
          },
          layout: { visibility: "none" },
        });
        map.addSource("boundary-sandbox-lasso-region-outline", {
          type: "geojson",
          data: { type: "FeatureCollection", features: [] },
        });
        map.addLayer({
          id: "boundary-sandbox-lasso-region-outline",
          type: "line",
          source: "boundary-sandbox-lasso-region-outline",
          paint: {
            "line-color": "#65a30d",
            "line-width": [
              "interpolate",
              ["linear"],
              ["zoom"],
              8,
              1,
              12,
              1.25,
              16,
              1.5,
            ],
            "line-opacity": 0.88,
          },
          layout: { visibility: "none" },
        });
        map.addSource("boundary-sandbox-hex", {
          type: "geojson",
          data: { type: "FeatureCollection", features: [] },
          promoteId: "_hexKey",
        });
        map.addLayer({
          id: "boundary-sandbox-hex-fill",
          type: "fill",
          source: "boundary-sandbox-hex",
          paint: {
            "fill-color": "#84cc16",
            "fill-opacity": [
              "case",
              ["boolean", ["feature-state", "selected"], false],
              0.4,
              0,
            ],
          },
          layout: { visibility: "none" },
        });
        map.addSource("boundary-sandbox-selection-outline", {
          type: "geojson",
          data: { type: "FeatureCollection", features: [] },
        });
        map.addLayer({
          id: "boundary-sandbox-selection-outline-line",
          type: "line",
          source: "boundary-sandbox-selection-outline",
          paint: {
            "line-color": "#65a30d",
            "line-width": [
              "interpolate",
              ["linear"],
              ["zoom"],
              8,
              2,
              12,
              2.75,
              16,
              3.5,
            ],
            "line-opacity": 0.92,
          },
          layout: { visibility: "none" },
        });
        map.addSource("boundary-sandbox-lasso-trace", {
          type: "geojson",
          data: { type: "FeatureCollection", features: [] },
        });
        map.addLayer({
          id: "boundary-sandbox-lasso-line",
          type: "line",
          source: "boundary-sandbox-lasso-trace",
          paint: {
            "line-color": "#64748b",
            "line-width": 2,
            "line-opacity": 0.9,
          },
          layout: { visibility: "none" },
        });

        var schoolMapHighlightStateAny = [
          "any",
          ["==", ["feature-state", "ring"], true],
          ["==", ["feature-state", "selected"], true],
          ["==", ["feature-state", "scenarioFeeder"], true],
        ];
        var schoolMapCircleBasePaint = {
          "circle-pitch-alignment": "map",
          "circle-radius": [
            "interpolate",
            ["linear"],
            ["zoom"],
            8,
            3,
            12,
            6,
            16,
            10,
          ],
          "circle-stroke-width": [
            "case",
            schoolMapHighlightStateAny,
            5.5,
            1,
          ],
          "circle-stroke-opacity": 1,
          "circle-opacity": 0.92,
        };

        map.addLayer({
          id: "schools-elementary",
          type: "circle",
          source: "schools",
          filter: ["==", ["get", "TYPE"], "ELEMENTARY"],
          paint: Object.assign({}, schoolMapCircleBasePaint, {
            "circle-color": PALETTE.elementary.fill,
            "circle-stroke-color": [
              "case",
              schoolMapHighlightStateAny,
              PALETTE.elementary.highlightStroke,
              schoolMapCircleStrokeColorDefault,
            ],
          }),
        });
        map.addLayer({
          id: "schools-middle",
          type: "circle",
          source: "schools",
          filter: ["==", ["get", "TYPE"], "MIDDLE"],
          paint: Object.assign({}, schoolMapCircleBasePaint, {
            "circle-color": PALETTE.middle.fill,
            "circle-stroke-color": [
              "case",
              schoolMapHighlightStateAny,
              PALETTE.middle.highlightStroke,
              schoolMapCircleStrokeColorDefault,
            ],
          }),
        });
        map.addLayer({
          id: "schools-high",
          type: "circle",
          source: "schools",
          filter: [
            "any",
            ["==", ["get", "TYPE"], "HIGH"],
            ["==", ["get", "TYPE"], "JR SR HIGH"],
          ],
          paint: Object.assign({}, schoolMapCircleBasePaint, {
            "circle-color": [
              "match",
              ["get", "TYPE"],
              "HIGH",
              PALETTE.high.fill,
              "JR SR HIGH",
              PALETTE.jrSr.fill,
              PALETTE.high.fill,
            ],
            "circle-stroke-color": [
              "case",
              schoolMapHighlightStateAny,
              [
                "match",
                ["get", "TYPE"],
                "HIGH",
                PALETTE.high.highlightStroke,
                "JR SR HIGH",
                PALETTE.jrSr.highlightStroke,
                PALETTE.high.highlightStroke,
              ],
              schoolMapCircleStrokeColorDefault,
            ],
          }),
        });
        map.addSource("charter-schools", {
          type: "geojson",
          data: charterFc || { type: "FeatureCollection", features: [] },
          promoteId: "OBJECTID",
        });
        map.addLayer({
          id: "schools-charter",
          type: "circle",
          source: "charter-schools",
          filter: ["==", ["get", "TYPE"], "CHARTER"],
          paint: Object.assign({}, schoolMapCircleBasePaint, {
            "circle-color": PALETTE.charter.fill,
            "circle-stroke-color": [
              "case",
              schoolMapHighlightStateAny,
              PALETTE.charter.highlightStroke,
              schoolMapCircleStrokeColorDefault,
            ],
          }),
        });
        map.addSource("private-schools", {
          type: "geojson",
          data: privateFc || { type: "FeatureCollection", features: [] },
          promoteId: "FID",
        });
        map.addLayer({
          id: "schools-private",
          type: "circle",
          source: "private-schools",
          paint: Object.assign({}, schoolMapCircleBasePaint, {
            "circle-color": PALETTE.privateSchool.fill,
            "circle-stroke-color": [
              "case",
              schoolMapHighlightStateAny,
              PALETTE.privateSchool.highlightStroke,
              schoolMapCircleStrokeColorDefault,
            ],
          }),
        });

        ["schools-elementary", "schools-middle", "schools-high", "schools-charter", "schools-private"].forEach(
          function (lid) {
            if (map.getLayer(lid)) {
              map.moveLayer(lid);
            }
          }
        );

        /**
         * Default stack draws municipal hover under HS/MS/ES fills — the stroke is invisible. Place it above
         * assignment boundaries (before parcels) so the hover line is actually visible.
         */
        if (map.getLayer("municipal-boundaries-hover") && map.getLayer("school-parcels-high")) {
          try {
            map.moveLayer("municipal-boundaries-hover", "school-parcels-high");
          } catch (errMh) {
            /* ignore */
          }
        }

        var combined = null;
        combined = mergeBbox(combined, computeBbox(es));
        combined = mergeBbox(combined, computeBbox(ms));
        combined = mergeBbox(combined, computeBbox(hs));
        combined = mergeBbox(combined, computeBbox(schools));
        combined = mergeBbox(combined, computeBbox(schoolParcelsFc));
        combined = mergeBbox(combined, computeBbox(charterFc));
        combined = mergeBbox(combined, computeBbox(privateFc));

        map.resize();
        if (fitBounds && combined) {
          map.fitBounds(combined, { padding: 48, maxZoom: 12, duration: 0 });
        }
        requestAnimationFrame(function () {
          map.resize();
        });

        GEO_CACHE.es = es;
        GEO_CACHE.ms = ms;
        GEO_CACHE.hs = hs;
        GEO_CACHE.schools = schools;

        if (studentHexFc && studentHexFc.features && studentHexFc.features.length) {
          STUDENT_HEX_INDEX = buildStudentHexIndex(studentHexFc);
          TRAVEL_SHED_RESIDENCE_INDEX = buildTravelShedResidenceIndex(studentHexFc);
        } else {
          STUDENT_HEX_INDEX = null;
          TRAVEL_SHED_RESIDENCE_INDEX = null;
        }
        HOMESCHOOL_HEX_COUNTS = buildHomeschoolHexCounts(
          homeschoolFc && homeschoolFc.features ? homeschoolFc : null
        );
        HOMESCHOOL_HEX_GEOMETRY_FALLBACK = buildHomeschoolHexGeometryFallback(
          homeschoolFc && homeschoolFc.features ? homeschoolFc : null
        );
        HOMESCHOOL_DETAILS_BY_HEX_KEY = buildHomeschoolDetailsByHexKey(
          homeschoolFc && homeschoolFc.features ? homeschoolFc : null
        );
        clearHomeschoolInBoundaryCountCache();

        if (!mapLayersInitialized) {
          mapLayersInitialized = true;
          var schoolByMsid = buildSchoolLookup(schools);
          populateSchoolSelect(schools);
          populateScenarioSchoolSelect(schools);
          setupToggles();
          setupMapInteractions(schoolByMsid);
          setupSchoolSelection(schoolByMsid);
          setupScenarioSchoolSelection(schoolByMsid, schools);
          initDashboardResizer(map);
          clearSelectedSchoolHighlight();
          syncStudentHexLayer();
          renderEnrollmentChart(null);
          renderDemographicsCharts(null);
        } else {
          syncStudentHexLayer();
          refreshAssignmentBoundaryHighlight();
          if (selectedSchoolMsid != null) {
            try {
              map.setFeatureState(
                { source: "schools", id: selectedSchoolMsid },
                { selected: true }
              );
            } catch (e) {
              /* ignore */
            }
          }
          resyncToolbarLayerToggleVisibility();
        }
        applyScenarioFeederMapHighlights();
        syncTravelShedLayerFilter();
        rebuildBoundarySandboxHexSourceFromIndex();
        syncBoundarySandboxMapLayers();
  }

  map.on("style.load", function () {
    if (!geoJsonDataCache) return;
    applyGeoJsonLayersFromFetchResults(geoJsonDataCache, { fitBounds: false });
  });

  map.on("load", function () {
    setupMapDensityLegendViewListeners();
    var basemapRoot = document.getElementById("basemap-toggle");
    if (basemapRoot) {
      basemapRoot.querySelectorAll("[data-basemap]").forEach(function (btn) {
        btn.addEventListener("click", function () {
          var mode = btn.getAttribute("data-basemap");
          if (MAPBOX_STYLES[mode]) setMapboxBasemap(mode);
        });
      });
    }

    Promise.all([
      fetch(DATA.es).then(function (r) {
        return r.json();
      }),
      fetch(DATA.ms).then(function (r) {
        return r.json();
      }),
      fetch(DATA.hs).then(function (r) {
        return r.json();
      }),
      fetch(DATA.schools).then(function (r) {
        return r.json();
      }),
      fetch(DATA.masterCsv)
        .then(function (r) {
          return r.ok ? r.text() : "";
        })
        .catch(function () {
          return "";
        }),
      fetch(DATA.sankeyEsMs)
        .then(function (r) {
          return r.ok ? r.json() : null;
        })
        .catch(function () {
          return null;
        }),
      fetch(DATA.studentHexagons)
        .then(function (r) {
          return r.ok ? r.json() : null;
        })
        .then(function (data) {
          if (!data) {
            return null;
          }
          if (data.v === 2) {
            return expandStudentHexBundleToFeatureCollection(data);
          }
          if (data.type === "FeatureCollection") {
            return data;
          }
          return null;
        })
        .catch(function () {
          return null;
        }),
      fetch(DATA.schoolParcels)
        .then(function (r) {
          return r.ok ? r.json() : null;
        })
        .catch(function () {
          return null;
        }),
      fetch(DATA.schoolBoardDistricts)
        .then(function (r) {
          return r.ok ? r.json() : { type: "FeatureCollection", features: [] };
        })
        .catch(function () {
          return { type: "FeatureCollection", features: [] };
        }),
      fetch(DATA.charterSchoolLocations)
        .then(function (r) {
          return r.ok ? r.json() : { type: "FeatureCollection", features: [] };
        })
        .catch(function () {
          return { type: "FeatureCollection", features: [] };
        }),
      fetch(DATA.meadowlaneCaptureOverride)
        .then(function (r) {
          return r.ok ? r.json() : null;
        })
        .catch(function () {
          return null;
        }),
      fetch(DATA.municipalBoundaries)
        .then(function (r) {
          return r.ok ? r.json() : { type: "FeatureCollection", features: [] };
        })
        .catch(function () {
          return { type: "FeatureCollection", features: [] };
        }),
      fetch(DATA.eseFeederMatrix)
        .then(function (r) {
          return r.ok ? r.json() : null;
        })
        .catch(function () {
          return null;
        }),
      fetch(DATA.travelImpact)
        .then(function (r) {
          return r.ok ? r.json() : null;
        })
        .catch(function () {
          return null;
        }),
      fetch(DATA.schoolIsochrones)
        .then(function (r) {
          return r.ok ? r.json() : { type: "FeatureCollection", features: [] };
        })
        .catch(function () {
          return { type: "FeatureCollection", features: [] };
        }),
      fetch(DATA.bpsEmployeeCount)
        .then(function (r) {
          return r.ok ? r.json() : null;
        })
        .catch(function () {
          return null;
        }),
      fetch(DATA.privateSchoolLocations)
        .then(function (r) {
          return r.ok ? r.json() : { type: "FeatureCollection", features: [] };
        })
        .catch(function () {
          return { type: "FeatureCollection", features: [] };
        }),
      fetch(DATA.homeschoolStudentHexagons)
        .then(function (r) {
          return r.ok ? r.json() : { type: "FeatureCollection", features: [] };
        })
        .catch(function () {
          return { type: "FeatureCollection", features: [] };
        }),
    ])
      .then(function (results) {
        MASTER_BY_MSID = parseSchoolMasterCsv(results[4] || "");
        SANKEY_CACHE = results[5];
        geoJsonDataCache = results;
        MEADOWLANE_CAPTURE_OVERRIDE = results[10];
        ESE_FEEDER_MATRIX = results[12] || null;
        TRAVEL_IMPACT_ALL = results[13] || null;
        BPS_EMPLOYEE_COUNT_BY_MSID =
          results[15] && results[15].byMsid ? results[15].byMsid : null;
        MIDDLE_SCHOOL_MSID_SET = buildMiddleSchoolMsidSetFromSchoolsFc(
          enrichSchoolsFcWithMasterType(results[3])
        );
        if (MEADOWLANE_CAPTURE_OVERRIDE && MEADOWLANE_CAPTURE_OVERRIDE.zoning_audit) {
          var za =
            MEADOWLANE_CAPTURE_OVERRIDE.zoning_audit
              .student_count_with_zoned_msid_2031_in_any_column;
          if (za != null && !isNaN(Number(za)) && Number(za) > 0) {
            console.warn(
              "[Meadowlane] zoning_audit: non-zero count of students with zoned MSID 2031 in a zoning column:",
              za
            );
          }
        }
        applyGeoJsonLayersFromFetchResults(results, { fitBounds: true });
        var selAfter = document.getElementById("school-select");
        if (selAfter && selAfter.value) {
          renderEseFeederFlowsTable(Number(selAfter.value));
        } else {
          renderEseFeederFlowsTable(null);
        }
      })
      .catch(function (err) {
        console.error(err);
        alert(
          "Could not load GeoJSON data. Use Live Server (or any local web server) from this project folder so files under /geo can be fetched."
        );
      });
  });

  function buildSchoolLookup(schoolsFc) {
    var byMsid = {};
    if (!schoolsFc || !schoolsFc.features) return byMsid;
    schoolsFc.features.forEach(function (ft) {
      var p = ft.properties;
      if (p && p.SCHOOLS_ID != null) byMsid[p.SCHOOLS_ID] = p;
    });
    return byMsid;
  }

  function buildChoiceSchoolMsidSet(schoolsFc) {
    var o = {};
    if (!schoolsFc || !schoolsFc.features) return o;
    for (var i = 0; i < schoolsFc.features.length; i++) {
      var p = schoolsFc.features[i].properties;
      if (!p || p.SCHOOLS_ID == null || p.SCHOOLS_ID === "") continue;
      if (String(p.SchAB_Type || "").toUpperCase() !== "CHOICE") continue;
      o[String(p.SCHOOLS_ID)] = true;
    }
    return o;
  }

  /** @returns {Object<string, true>} */
  function buildCharterSchoolMsidSet(schoolsFc, charterFc) {
    var o = {};
    function addProps(p) {
      if (!p || p.SCHOOLS_ID == null || p.SCHOOLS_ID === "") return;
      var t = String(p.TYPE || "").toUpperCase();
      var ab = String(p.SchAB_Type || "").toUpperCase();
      if (t !== "CHARTER" && ab !== "CHARTER") return;
      var id = parseInt(String(p.SCHOOLS_ID).trim(), 10);
      if (isNaN(id)) return;
      o[String(id)] = true;
    }
    if (schoolsFc && schoolsFc.features) {
      for (var i = 0; i < schoolsFc.features.length; i++) {
        addProps(schoolsFc.features[i].properties);
      }
    }
    if (charterFc && charterFc.features) {
      for (var j = 0; j < charterFc.features.length; j++) {
        addProps(charterFc.features[j].properties);
      }
    }
    return o;
  }

  /** Choice or charter schools have no boundary-based "zoned" cohort for student hex overlay. */
  function selectedSchoolDisallowsZonedStudentHex(msid) {
    if (msid == null || isNaN(msid)) return true;
    var k = String(parseInt(String(msid), 10));
    if (CHOICE_SCHOOL_MSIDS && CHOICE_SCHOOL_MSIDS[k]) return true;
    if (CHARTER_SCHOOL_MSIDS && CHARTER_SCHOOL_MSIDS[k]) return true;
    return false;
  }

  /** @returns {boolean} */
  function schoolIsChoiceFromProps(p) {
    return !!p && String(p.SchAB_Type || "").toUpperCase() === "CHOICE";
  }

  function joinCsvQuotedRow(cells) {
    return cells
      .map(function (cell) {
        var s = cell != null ? String(cell) : "";
        return '"' + s.replace(/"/g, '""') + '"';
      })
      .join(",");
  }

  function applyChoiceSchoolCaptureToCsvGrid(grid) {
    if (
      !grid ||
      grid.length < 2 ||
      !CHOICE_SCHOOL_MSIDS ||
      !Object.keys(CHOICE_SCHOOL_MSIDS).length
    ) {
      return grid;
    }
    var headers = grid[0].map(function (h) {
      return String(h).trim();
    });
    var capSlugs = [
      "assignment_capture_rate",
      "other_district_capture_rate",
      "choice_capture_rate",
      "charter_capture_rate",
      "fromto_resident_denominator",
      "assignment_capture_students",
      "other_district_capture_students",
      "choice_capture_students",
      "charter_capture_students",
    ];
    var capIdxs = [];
    for (var si = 0; si < capSlugs.length; si++) {
      var j = headers.indexOf(capSlugs[si]);
      if (j >= 0) {
        capIdxs.push(j);
      }
    }
    if (!capIdxs.length) return grid;
    var na = "N/A (Choice School)";
    for (var r = 1; r < grid.length; r++) {
      var row = grid[r];
      if (!row || !row.length) continue;
      var idRaw = row[0] != null ? String(row[0]).trim() : "";
      if (!idRaw) continue;
      var idNum = parseInt(idRaw, 10);
      if (isNaN(idNum)) continue;
      if (CHOICE_SCHOOL_MSIDS[String(idNum)] || CHOICE_SCHOOL_MSIDS[idRaw]) {
        for (var ci = 0; ci < capIdxs.length; ci++) {
          row[capIdxs[ci]] = na;
        }
      }
    }
    return grid;
  }

  /** Appends or fills `bps_employee_count` from BPS_EMPLOYEE_COUNT_BY_MSID (MSID in column 0). */
  function applyBpsEmployeeCountToCsvGrid(grid) {
    if (!grid || grid.length < 2) return grid;
    var colName = "bps_employee_count";
    var headers = grid[0].map(function (h) {
      return String(h).trim();
    });
    var existingIdx = headers.indexOf(colName);
    var msidCol = 0;

    function lookupCount(msidNum) {
      if (!BPS_EMPLOYEE_COUNT_BY_MSID || msidNum == null || isNaN(msidNum)) {
        return "";
      }
      var map = BPS_EMPLOYEE_COUNT_BY_MSID;
      var k = String(Number(msidNum));
      var v =
        map[k] != null
          ? map[k]
          : map[String(Number(msidNum)).padStart(4, "0")];
      if (v == null || v === "") return "";
      return String(v);
    }

    if (existingIdx >= 0) {
      for (var r = 1; r < grid.length; r++) {
        var row = grid[r];
        if (!row) continue;
        var idNum = parseInt(
          String(row[msidCol] != null ? row[msidCol] : "").trim(),
          10
        );
        row[existingIdx] = lookupCount(idNum);
      }
      return grid;
    }

    grid[0] = grid[0].concat([colName]);
    for (var r2 = 1; r2 < grid.length; r2++) {
      var row2 = grid[r2];
      if (!row2) continue;
      var idNum2 = parseInt(
        String(row2[msidCol] != null ? row2[msidCol] : "").trim(),
        10
      );
      grid[r2] = row2.concat([lookupCount(idNum2)]);
    }
    return grid;
  }

  /** MSIDs currently listed in #school-select (excludes placeholder and separator). */
  function getSchoolDropdownMsidSet() {
    var sel = document.getElementById("school-select");
    var o = {};
    if (!sel || !sel.options) return o;
    for (var i = 0; i < sel.options.length; i++) {
      var v = sel.options[i].value;
      if (v === "" || v == null) continue;
      var n = parseInt(String(v).trim(), 10);
      if (!isNaN(n)) o[String(n)] = true;
    }
    return o;
  }

  function isMsidInSchoolSelectDropdown(msid) {
    if (msid == null || isNaN(msid)) return false;
    var a = getSchoolDropdownMsidSet();
    return !!(a[String(msid)] || a[String(Number(msid))]);
  }

  function isExistingConditionsViewActive() {
    var p = document.getElementById("page-existing");
    return !!(p && !p.hidden);
  }

  function isBoundarySandboxViewActive() {
    var p = document.getElementById("page-sandbox");
    return !!(p && !p.hidden);
  }

  function shallowCopyHexKeyBag(from) {
    var o = Object.create(null);
    if (!from) {
      return o;
    }
    for (var k in from) {
      if (Object.prototype.hasOwnProperty.call(from, k) && from[k]) {
        o[k] = true;
      }
    }
    return o;
  }

  function countSandboxHexKeys(bag) {
    var n = 0;
    if (!bag) {
      return 0;
    }
    for (var kb in bag) {
      if (Object.prototype.hasOwnProperty.call(bag, kb) && bag[kb]) {
        n++;
      }
    }
    return n;
  }

  /**
   * Keys used for sidebar aggregates: live selection when confirmed; otherwise last confirmed snapshot.
   * @returns {Object<string, boolean>|null}
   */
  function getHexKeysForSandboxStatistics() {
    if (BOUNDARY_SANDBOX.selectionConfirmed) {
      return BOUNDARY_SANDBOX.selectedHexKeys;
    }
    if (countSandboxHexKeys(BOUNDARY_SANDBOX.confirmedHexKeysSnapshot) > 0) {
      return BOUNDARY_SANDBOX.confirmedHexKeysSnapshot;
    }
    return null;
  }

  /**
   * Turf v7 `polygonToLine` returns a Feature *or* a FeatureCollection (MultiPolygon → multiple lines).
   * Mapbox sources must receive a proper FeatureCollection of Feature objects, not a nested FC.
   * @param {GeoJSON.Feature<GeoJSON.Polygon|GeoJSON.MultiPolygon>|null} polyFeature
   * @returns {GeoJSON.FeatureCollection|null}
   */
  function turfPolygonToLineAsFeatureCollection(polyFeature) {
    if (!polyFeature || !polyFeature.geometry) {
      return null;
    }
    var gt = polyFeature.geometry.type;
    if (gt !== "Polygon" && gt !== "MultiPolygon") {
      return null;
    }
    if (typeof turf === "undefined" || !turf || typeof turf.polygonToLine !== "function") {
      return null;
    }
    try {
      var r = turf.polygonToLine(polyFeature);
      if (!r) {
        return null;
      }
      if (r.type === "FeatureCollection") {
        return r.features && r.features.length ? r : null;
      }
      if (r.type === "Feature") {
        return { type: "FeatureCollection", features: [r] };
      }
      return null;
    } catch (err) {
      return null;
    }
  }

  /**
   * Turf v7 `union` takes one FeatureCollection of polygons; two-arg `union(a,b)` does not merge correctly.
   * @param {GeoJSON.Feature<GeoJSON.Polygon|GeoJSON.MultiPolygon>} polyA
   * @param {GeoJSON.Feature<GeoJSON.Polygon|GeoJSON.MultiPolygon>} polyB
   * @returns {GeoJSON.Feature<GeoJSON.Polygon|GeoJSON.MultiPolygon>|null}
   */
  function turfUnionPolygonFeatures(polyA, polyB) {
    if (!polyA || !polyB || typeof turf === "undefined" || !turf || typeof turf.union !== "function") {
      return null;
    }
    try {
      var fc =
        typeof turf.featureCollection === "function"
          ? turf.featureCollection([polyA, polyB])
          : { type: "FeatureCollection", features: [polyA, polyB] };
      return turf.union(fc);
    } catch (err) {
      return null;
    }
  }

  /**
   * Turf v7 `difference`: FeatureCollection where result is first polygon minus overlap with the rest.
   * @returns {GeoJSON.Feature<GeoJSON.Polygon|GeoJSON.MultiPolygon>|null}
   */
  function turfDifferencePolygonFeatures(polyA, polyB) {
    if (!polyA || !polyB || !polyA.geometry || !polyB.geometry) {
      return null;
    }
    var ta = polyA.geometry.type;
    var tb = polyB.geometry.type;
    if ((ta !== "Polygon" && ta !== "MultiPolygon") || (tb !== "Polygon" && tb !== "MultiPolygon")) {
      return null;
    }
    if (typeof turf === "undefined" || !turf || typeof turf.difference !== "function") {
      return null;
    }
    try {
      var fc =
        typeof turf.featureCollection === "function"
          ? turf.featureCollection([polyA, polyB])
          : { type: "FeatureCollection", features: [polyA, polyB] };
      return turf.difference(fc);
    } catch (err) {
      return null;
    }
  }

  /**
   * Merge multiple Polygon/MultiPolygon features into one (used for zoned-hex footprint union).
   * @param {GeoJSON.Feature[]} feats
   * @returns {GeoJSON.Feature|null}
   */
  function mergePolygonFeatureArrayToOne(feats) {
    if (!feats || !feats.length) {
      return null;
    }
    if (feats.length === 1) {
      return feats[0];
    }
    if (typeof turf === "undefined" || !turf || typeof turf.union !== "function") {
      return null;
    }
    try {
      if (typeof turf.featureCollection === "function" && feats.length > 2) {
        var bulk = turf.union(turf.featureCollection(feats));
        if (bulk && bulk.geometry) {
          return bulk;
        }
      }
    } catch (eBulk) {
      /* pairwise fallback below */
    }
    var merged = feats[0];
    for (var i = 1; i < feats.length; i++) {
      var unn = turfUnionPolygonFeatures(merged, feats[i]);
      if (unn && unn.geometry) {
        merged = unn;
      }
    }
    return merged && merged.geometry ? merged : null;
  }

  /**
   * Sets `lassoRegionFootprintFeature` to the union of all currently selected hex geometries (light green tint
   * between hexes, including base-school zoned loads). Does not change hex selection state.
   */
  function syncSandboxLassoFootprintFromSelectedHexGeometries() {
    if (!STUDENT_HEX_INDEX || !STUDENT_HEX_INDEX.geometryByHexKey) {
      BOUNDARY_SANDBOX.lassoRegionFootprintFeature = null;
      syncBoundarySandboxLassoRegionSourcesFromAccumulator();
      return;
    }
    var bag = BOUNDARY_SANDBOX.selectedHexKeys;
    var feats = [];
    for (var sk in bag) {
      if (!Object.prototype.hasOwnProperty.call(bag, sk) || !bag[sk]) {
        continue;
      }
      var g = homeschoolHexGeometry(sk);
      if (!g) {
        continue;
      }
      feats.push({ type: "Feature", properties: {}, geometry: g });
    }
    if (!feats.length) {
      BOUNDARY_SANDBOX.lassoRegionFootprintFeature = null;
    } else {
      var merged = mergePolygonFeatureArrayToOne(feats);
      BOUNDARY_SANDBOX.lassoRegionFootprintFeature = merged && merged.geometry ? merged : null;
    }
    syncBoundarySandboxLassoRegionSourcesFromAccumulator();
  }

  /**
   * Shared boundary of confirmed hex polygons only (no convex hull — avoids enclosing unselected gaps).
   * @param {GeoJSON.Feature[]} feats Each feature’s geometry should be a Polygon or MultiPolygon.
   * @returns {GeoJSON.FeatureCollection|null} Line features for stroke (may be multiple lines after union).
   */
  function sandboxConfirmedHexUnionToOutlineLineFeature(feats) {
    if (!feats || !feats.length || typeof turf === "undefined") {
      return null;
    }
    try {
      var merged = feats[0];
      if (feats.length > 1) {
        for (var i = 1; i < feats.length; i++) {
          var unn = turfUnionPolygonFeatures(merged, feats[i]);
          if (unn && unn.geometry) {
            merged = unn;
          }
        }
      }
      return turfPolygonToLineAsFeatureCollection(merged);
    } catch (err) {
      return null;
    }
  }

  /** Outer perimeter of the last confirmed selection only (`confirmedHexKeysSnapshot`), via polygon union. */
  function updateBoundarySandboxSelectionOutline() {
    if (!map || !map.getSource("boundary-sandbox-selection-outline")) {
      return;
    }
    var empty = { type: "FeatureCollection", features: [] };
    var gk = STUDENT_HEX_INDEX && STUDENT_HEX_INDEX.geometryByHexKey;
    if (!gk || typeof turf === "undefined") {
      try {
        map.getSource("boundary-sandbox-selection-outline").setData(empty);
      } catch (e0) {
        /* ignore */
      }
      return;
    }
    var snap = BOUNDARY_SANDBOX.confirmedHexKeysSnapshot;
    var feats = [];
    for (var ks in snap) {
      if (!Object.prototype.hasOwnProperty.call(snap, ks) || !snap[ks]) {
        continue;
      }
      var g = homeschoolHexGeometry(ks);
      if (!g) {
        continue;
      }
      feats.push({ type: "Feature", properties: {}, geometry: g });
    }
    if (!feats.length) {
      try {
        map.getSource("boundary-sandbox-selection-outline").setData(empty);
      } catch (e1) {
        /* ignore */
      }
      return;
    }
    var outlineFcResult = sandboxConfirmedHexUnionToOutlineLineFeature(feats);
    if (!outlineFcResult || !outlineFcResult.features || !outlineFcResult.features.length) {
      try {
        map.getSource("boundary-sandbox-selection-outline").setData(empty);
      } catch (e2) {
        /* ignore */
      }
      return;
    }
    try {
      map.getSource("boundary-sandbox-selection-outline").setData(outlineFcResult);
    } catch (e3) {
      /* ignore */
    }
  }

  /** @returns {number|null} MSID from #sandbox-base-school, or null. */
  function getSandboxBaseSchoolMsid() {
    var sel = document.getElementById("sandbox-base-school");
    if (!sel || !sel.value) {
      return null;
    }
    var v = Number(sel.value);
    return isNaN(v) ? null : v;
  }

  function getBoundarySandboxSelectTool() {
    var g = document.querySelector('input[name="sandbox-hex-tool"]:checked');
    var v = g && g.value ? String(g.value) : "lasso";
    return v === "lasso" ? "lasso" : "brush";
  }

  /** @returns {"select"|"erase"} */
  function getBoundarySandboxHexMode() {
    var m = document.querySelector('input[name="sandbox-hex-mode"]:checked');
    var v = m && m.value ? String(m.value) : "select";
    return v === "erase" ? "erase" : "select";
  }

  /** @returns {string|null} */
  function querySandboxHexKeyAtPoint(pixelPoint) {
    if (!map) {
      return null;
    }
    var hits;
    try {
      hits = map.queryRenderedFeatures(pixelPoint, { layers: ["boundary-sandbox-hex-fill"] });
    } catch (eQ) {
      return null;
    }
    if (!hits || !hits.length) {
      return null;
    }
    var f0 = hits[0];
    var key = f0.properties && f0.properties._hexKey != null ? String(f0.properties._hexKey) : null;
    return key;
  }

  function setBoundarySandboxLassoSource(geojson) {
    if (!map || !map.getSource("boundary-sandbox-lasso-trace")) {
      return;
    }
    try {
      map.getSource("boundary-sandbox-lasso-trace").setData(geojson || { type: "FeatureCollection", features: [] });
    } catch (e0) {
      /* ignore */
    }
  }

  function clearBoundarySandboxLassoLine() {
    BOUNDARY_SANDBOX_LASSO.active = false;
    BOUNDARY_SANDBOX_LASSO.points = null;
    setBoundarySandboxLassoSource({ type: "FeatureCollection", features: [] });
  }

  function boundarySandboxSetHexSelected(key, selected) {
    if (!key || !map) {
      return;
    }
    try {
      if (selected) {
        BOUNDARY_SANDBOX.selectedHexKeys[key] = true;
        map.setFeatureState({ source: "boundary-sandbox-hex", id: key }, { selected: true });
      } else {
        delete BOUNDARY_SANDBOX.selectedHexKeys[key];
        map.setFeatureState({ source: "boundary-sandbox-hex", id: key }, { selected: false });
      }
    } catch (e0) {
      /* ignore */
    }
  }

  function clearBoundarySandboxGeographicSelection() {
    if (map) {
      for (var k in BOUNDARY_SANDBOX.selectedHexKeys) {
        if (Object.prototype.hasOwnProperty.call(BOUNDARY_SANDBOX.selectedHexKeys, k) && BOUNDARY_SANDBOX.selectedHexKeys[k]) {
          try {
            map.setFeatureState(
              { source: "boundary-sandbox-hex", id: k },
              { selected: false }
            );
          } catch (e1) {
            /* ignore */
          }
        }
      }
    }
    BOUNDARY_SANDBOX.selectedHexKeys = Object.create(null);
    BOUNDARY_SANDBOX.selectionConfirmed = false;
    BOUNDARY_SANDBOX.confirmedHexKeysSnapshot = Object.create(null);
    clearBoundarySandboxLassoLine();
    clearBoundarySandboxLassoRegionFill();
    BOUNDARY_SANDBOX_PAINT = {
      active: false,
      lastKey: null,
      startX: 0,
      startY: 0,
      clickKey: null,
      isDrag: false,
    };
    resetBoundarySandboxFilterState();
    updateSandboxSelectedHexCountUi();
  }

  /**
   * Point-in-polygon (odd–even) for a ring. Ring may be open or closed.
   * @param {number} lng
   * @param {number} lat
   * @param {number[][]} ring
   */
  function pointInRingLngLat(lng, lat, ring) {
    if (!ring || ring.length < 3) {
      return false;
    }
    var ins = false;
    var n = ring.length;
    for (var i = 0, j = n - 1; i < n; j = i++) {
      var xi = ring[i][0];
      var yi = ring[i][1];
      var xj = ring[j][0];
      var yj = ring[j][1];
      if (Math.abs(yj - yi) < 1e-12) {
        continue;
      }
      if ((yi > lat) !== (yj > lat)) {
        var xInt = (xj - xi) * (lat - yi) / (yj - yi) + xi;
        if (lng < xInt) {
          ins = !ins;
        }
      }
    }
    return ins;
  }

  function closeRingIfNeeded(pts) {
    if (!pts || pts.length < 1) {
      return [];
    }
    var a = pts.slice();
    if (a.length < 2) {
      return a;
    }
    if (a[0][0] !== a[a.length - 1][0] || a[0][1] !== a[a.length - 1][1]) {
      a.push([a[0][0], a[0][1]]);
    }
    return a;
  }

  /**
   * Closed lng/lat ring → GeoJSON Polygon for lasso “footprint” fill (hex gaps with no student data).
   * @param {number[][]} ring
   * @returns {Object|null} Feature or null
   */
  function closedLngLatRingToSandboxPolygonFeature(ring) {
    if (!ring || ring.length < 3) {
      return null;
    }
    var r = closeRingIfNeeded(ring);
    if (r.length < 4) {
      return null;
    }
    return {
      type: "Feature",
      properties: {},
      geometry: { type: "Polygon", coordinates: [r] },
    };
  }

  /**
   * Outer rings of polygon features → LineString features (fallback outline when Turf is unavailable).
   * @param {Object[]} polyFeatures
   * @returns {Object[]}
   */
  function sandboxPolygonFeaturesToOutlineLineFeatures(polyFeatures) {
    var out = [];
    if (!polyFeatures || !polyFeatures.length) {
      return out;
    }
    for (var i = 0; i < polyFeatures.length; i++) {
      var f = polyFeatures[i];
      if (!f || !f.geometry) {
        continue;
      }
      if (f.geometry.type === "Polygon") {
        var rings = f.geometry.coordinates;
        if (!rings || !rings[0] || rings[0].length < 4) {
          continue;
        }
        out.push({
          type: "Feature",
          properties: {},
          geometry: { type: "LineString", coordinates: rings[0] },
        });
        continue;
      }
      if (f.geometry.type === "MultiPolygon") {
        var mp = f.geometry.coordinates;
        for (var pi = 0; pi < mp.length; pi++) {
          var polyRing = mp[pi] && mp[pi][0];
          if (!polyRing || polyRing.length < 4) {
            continue;
          }
          out.push({
            type: "Feature",
            properties: {},
            geometry: { type: "LineString", coordinates: polyRing },
          });
        }
      }
    }
    return out;
  }

  function syncBoundarySandboxLassoRegionSourcesFromAccumulator() {
    if (!map) {
      return;
    }
    var footprint = BOUNDARY_SANDBOX.lassoRegionFootprintFeature;
    var emptyFc = { type: "FeatureCollection", features: [] };
    if (!footprint || !footprint.geometry) {
      if (map.getSource("boundary-sandbox-lasso-region-fill")) {
        try {
          map.getSource("boundary-sandbox-lasso-region-fill").setData(emptyFc);
        } catch (eE0) {
          /* ignore */
        }
      }
      if (map.getSource("boundary-sandbox-lasso-region-outline")) {
        try {
          map.getSource("boundary-sandbox-lasso-region-outline").setData(emptyFc);
        } catch (eE1) {
          /* ignore */
        }
      }
      return;
    }
    var gt = footprint.geometry.type;
    if (gt !== "Polygon" && gt !== "MultiPolygon") {
      if (map.getSource("boundary-sandbox-lasso-region-fill")) {
        try {
          map.getSource("boundary-sandbox-lasso-region-fill").setData(emptyFc);
        } catch (eE2) {
          /* ignore */
        }
      }
      if (map.getSource("boundary-sandbox-lasso-region-outline")) {
        try {
          map.getSource("boundary-sandbox-lasso-region-outline").setData(emptyFc);
        } catch (eE3) {
          /* ignore */
        }
      }
      return;
    }
    var fillFc = { type: "FeatureCollection", features: [footprint] };
    var outlineFc = turfPolygonToLineAsFeatureCollection(footprint) || emptyFc;
    if (!outlineFc.features || outlineFc.features.length === 0) {
      outlineFc = {
        type: "FeatureCollection",
        features: sandboxPolygonFeaturesToOutlineLineFeatures([footprint]),
      };
    }
    if (map.getSource("boundary-sandbox-lasso-region-fill")) {
      try {
        map.getSource("boundary-sandbox-lasso-region-fill").setData(fillFc);
      } catch (eFill) {
        /* ignore */
      }
    }
    if (map.getSource("boundary-sandbox-lasso-region-outline")) {
      try {
        map.getSource("boundary-sandbox-lasso-region-outline").setData(outlineFc);
      } catch (eOut) {
        /* ignore */
      }
    }
  }

  function clearBoundarySandboxLassoRegionFill() {
    BOUNDARY_SANDBOX.lassoRegionFootprintFeature = null;
    if (!map || !map.getSource("boundary-sandbox-lasso-region-fill")) {
      return;
    }
    try {
      map.getSource("boundary-sandbox-lasso-region-fill").setData({
        type: "FeatureCollection",
        features: [],
      });
    } catch (eClr) {
      /* ignore */
    }
    if (map.getSource("boundary-sandbox-lasso-region-outline")) {
      try {
        map.getSource("boundary-sandbox-lasso-region-outline").setData({
          type: "FeatureCollection",
          features: [],
        });
      } catch (eClr2) {
        /* ignore */
      }
    }
  }

  /**
   * Select-mode lasso: union the new ring into `lassoRegionFootprintFeature`.
   * @param {number[][]|null} closedRing
   */
  function applySelectLassoToLassoRegionFootprint(closedRing) {
    var feat = closedLngLatRingToSandboxPolygonFeature(closedRing || []);
    if (!feat) {
      return;
    }
    var fp = BOUNDARY_SANDBOX.lassoRegionFootprintFeature;
    if (!fp || !fp.geometry) {
      BOUNDARY_SANDBOX.lassoRegionFootprintFeature = feat;
    } else {
      var u = turfUnionPolygonFeatures(fp, feat);
      if (u && u.geometry) {
        BOUNDARY_SANDBOX.lassoRegionFootprintFeature = u;
      }
    }
    syncBoundarySandboxLassoRegionSourcesFromAccumulator();
  }

  /**
   * Erase-mode lasso: subtract the ring from the green footprint (turf.difference); clears tint if nothing left.
   * @param {number[][]|null} closedRing
   */
  function applyEraseLassoToLassoRegionFootprint(closedRing) {
    var fp = BOUNDARY_SANDBOX.lassoRegionFootprintFeature;
    if (!fp || !fp.geometry) {
      return;
    }
    var eraseFeat = closedLngLatRingToSandboxPolygonFeature(closedRing || []);
    if (!eraseFeat) {
      return;
    }
    var d = turfDifferencePolygonFeatures(fp, eraseFeat);
    if (d && d.geometry) {
      BOUNDARY_SANDBOX.lassoRegionFootprintFeature = d;
    } else {
      BOUNDARY_SANDBOX.lassoRegionFootprintFeature = null;
    }
    syncBoundarySandboxLassoRegionSourcesFromAccumulator();
  }

  function applyLassoToHexSelection(closedRing) {
    if (!STUDENT_HEX_INDEX || !STUDENT_HEX_INDEX.geometryByHexKey) {
      return 0;
    }
    if (!map) {
      return 0;
    }
    if (!closedRing || closedRing.length < 4) {
      return 0;
    }
    var gk = STUDENT_HEX_INDEX.geometryByHexKey;
    var any = 0;
    for (var hk in gk) {
      if (!Object.prototype.hasOwnProperty.call(gk, hk)) {
        continue;
      }
      var geom = gk[hk];
      if (!geom) {
        continue;
      }
      var c = polygonCentroid(geom);
      if (!c) {
        continue;
      }
      if (pointInRingLngLat(c[0], c[1], closedRing)) {
        var inSel = !!BOUNDARY_SANDBOX.selectedHexKeys[hk];
        var lMode = getBoundarySandboxHexMode();
        if (lMode === "select" && !inSel) {
          boundarySandboxSetHexSelected(hk, true);
          any++;
        } else if (lMode === "erase" && inSel) {
          boundarySandboxSetHexSelected(hk, false);
          any++;
        }
      }
    }
    if (any > 0) {
      BOUNDARY_SANDBOX.selectionConfirmed = false;
    }
    return any;
  }

  function endBoundarySandboxPaintOrLassoFromWindow() {
    if (!map) {
      return;
    }
    var needUi = false;
    if (BOUNDARY_SANDBOX_PAINT.active) {
      if (!BOUNDARY_SANDBOX_PAINT.isDrag && BOUNDARY_SANDBOX_PAINT.clickKey) {
        clearBoundarySandboxLassoRegionFill();
        var ck = BOUNDARY_SANDBOX_PAINT.clickKey;
        var w = !!BOUNDARY_SANDBOX.selectedHexKeys[ck];
        boundarySandboxSetHexSelected(ck, !w);
        BOUNDARY_SANDBOX.selectionConfirmed = false;
      }
      BOUNDARY_SANDBOX_PAINT.active = false;
      BOUNDARY_SANDBOX_PAINT.lastKey = null;
      BOUNDARY_SANDBOX_PAINT.clickKey = null;
      BOUNDARY_SANDBOX_PAINT.isDrag = false;
      needUi = true;
      try {
        map.dragPan.enable();
        map.getCanvas().style.cursor = "";
      } catch (e0) {
        /* ignore */
      }
    }
    if (BOUNDARY_SANDBOX_LASSO.active) {
      var raw = BOUNDARY_SANDBOX_LASSO.points;
      BOUNDARY_SANDBOX_LASSO.active = false;
      BOUNDARY_SANDBOX_LASSO.points = null;
      setBoundarySandboxLassoSource({ type: "FeatureCollection", features: [] });
      needUi = true;
      try {
        map.dragPan.enable();
        map.getCanvas().style.cursor = "";
      } catch (e1) {
        /* ignore */
      }
      if (raw && raw.length >= 3) {
        var closed = closeRingIfNeeded(raw);
        applyLassoToHexSelection(closed);
        if (getBoundarySandboxHexMode() === "select") {
          applySelectLassoToLassoRegionFootprint(closed);
        } else {
          applyEraseLassoToLassoRegionFootprint(closed);
        }
      }
    }
    if (needUi) {
      updateSandboxSelectedHexCountUi();
    }
  }

  function tryBrushDragAtPoint(pixelPoint) {
    if (!map) {
      return;
    }
    var key = querySandboxHexKeyAtPoint(pixelPoint);
    if (key == null) {
      BOUNDARY_SANDBOX_PAINT.lastKey = null;
      return;
    }
    if (key === BOUNDARY_SANDBOX_PAINT.lastKey) {
      return;
    }
    BOUNDARY_SANDBOX_PAINT.lastKey = key;
    var bMode = getBoundarySandboxHexMode();
    var inHex = !!BOUNDARY_SANDBOX.selectedHexKeys[key];
    if (bMode === "select") {
      if (!inHex) {
        boundarySandboxSetHexSelected(key, true);
        BOUNDARY_SANDBOX.selectionConfirmed = false;
      }
    } else {
      if (inHex) {
        boundarySandboxSetHexSelected(key, false);
        BOUNDARY_SANDBOX.selectionConfirmed = false;
      }
    }
  }

  function pruneBoundarySandboxSelectedKeysToGeometry() {
    if (!STUDENT_HEX_INDEX || !STUDENT_HEX_INDEX.geometryByHexKey) {
      BOUNDARY_SANDBOX.selectedHexKeys = Object.create(null);
      resetBoundarySandboxFilterState();
      return;
    }
    for (var ks in BOUNDARY_SANDBOX.selectedHexKeys) {
      if (!Object.prototype.hasOwnProperty.call(BOUNDARY_SANDBOX.selectedHexKeys, ks)) {
        continue;
      }
      if (!homeschoolHexGeometry(ks)) {
        delete BOUNDARY_SANDBOX.selectedHexKeys[ks];
      }
    }
  }

  function applyBoundarySandboxSelectionFeatureStates() {
    if (!map || !map.getSource("boundary-sandbox-hex") || !map.getLayer("boundary-sandbox-hex-fill")) {
      return;
    }
    for (var sk in BOUNDARY_SANDBOX.selectedHexKeys) {
      if (!Object.prototype.hasOwnProperty.call(BOUNDARY_SANDBOX.selectedHexKeys, sk) || !BOUNDARY_SANDBOX.selectedHexKeys[sk]) {
        continue;
      }
      try {
        map.setFeatureState(
          { source: "boundary-sandbox-hex", id: sk },
          { selected: true }
        );
      } catch (eSet) {
        /* ignore */
      }
    }
  }

  function requestApplyBoundarySandboxSelectionOnIdle() {
    if (!map) return;
    if (map.isStyleLoaded && !map.isStyleLoaded()) return;
    map.once("idle", function () {
      applyBoundarySandboxSelectionFeatureStates();
    });
  }

  function syncSandboxConfirmEditButtonStates() {
    var conf = document.getElementById("sandbox-confirm-btn");
    if (!conf) {
      return;
    }
    var n = 0;
    for (var kb in BOUNDARY_SANDBOX.selectedHexKeys) {
      if (Object.prototype.hasOwnProperty.call(BOUNDARY_SANDBOX.selectedHexKeys, kb) && BOUNDARY_SANDBOX.selectedHexKeys[kb]) {
        n++;
      }
    }
    if (n === 0) {
      BOUNDARY_SANDBOX.selectionConfirmed = false;
    }
    var canConfirm = n > 0;
    conf.setAttribute("aria-disabled", canConfirm ? "false" : "true");
    conf.classList.toggle("is-inert", !canConfirm);
    var hasUnconfirmedChanges = canConfirm && !BOUNDARY_SANDBOX.selectionConfirmed;
    conf.classList.toggle("sandbox-confirm-btn--pending", hasUnconfirmedChanges);
  }

  function updateSandboxSelectedHexCountUi() {
    var el = document.getElementById("sandbox-hex-count");
    if (!el) return;
    var n = 0;
    for (var kc in BOUNDARY_SANDBOX.selectedHexKeys) {
      if (Object.prototype.hasOwnProperty.call(BOUNDARY_SANDBOX.selectedHexKeys, kc) && BOUNDARY_SANDBOX.selectedHexKeys[kc]) {
        n++;
      }
    }
    if (n === 0) {
      BOUNDARY_SANDBOX.selectionConfirmed = false;
      BOUNDARY_SANDBOX.confirmedHexKeysSnapshot = Object.create(null);
    }
    el.textContent =
      n === 0
        ? "No hexes selected — use a tool and the map"
        : n === 1
          ? "1 hex selected"
          : n + " hexes selected";
    syncSandboxConfirmEditButtonStates();
    updateSandboxStatsPanelSummary();
    if (countSandboxHexKeys(BOUNDARY_SANDBOX.confirmedHexKeysSnapshot) === 0) {
      updateBoundarySandboxSelectionOutline();
    }
  }

  function resetBoundarySandboxFilterState() {
    BOUNDARY_SANDBOX.gradeToggles = Object.create(null);
    BOUNDARY_SANDBOX.attendanceTypeToggles = Object.create(null);
    BOUNDARY_SANDBOX.schoolListExpanded = { attendance: false, zoned: false };
  }

  function clearSandboxStatsAndDemographicsDisplays() {
    resetBoundarySandboxFilterState();
    var g = document.getElementById("sandbox-card-body-grade");
    var a = document.getElementById("sandbox-card-body-attendance");
    var z = document.getElementById("sandbox-card-body-zoned");
    if (g) g.textContent = "—";
    if (a) a.textContent = "—";
    if (z) z.textContent = "—";
    var attBody = document.getElementById("sandbox-card-body-attendance-type");
    if (attBody) attBody.textContent = "—";
    var ethEl = document.getElementById("sandbox-demographics-ethnicity");
    var lunchEl = document.getElementById("sandbox-demographics-lunch");
    if (ethEl) {
      ethEl.innerHTML = '<p class="demographics-pie-empty">No selection confirmed yet.</p>';
    }
    if (lunchEl) {
      lunchEl.innerHTML = '<p class="demographics-pie-empty">No selection confirmed yet.</p>';
    }
  }

  /**
   * Grade bucket key for boundary sandbox charts/toggles. Homeschool export uses grade code 13 for “no grade”.
   */
  function sandboxGradeCanonicalForDetail(d) {
    if (d && d.__homeschool) {
      var tr = String(d.Grade != null ? d.Grade : "").trim();
      if (tr !== "") {
        var n13 = parseInt(tr.replace(/^0+/, "") || tr, 10);
        if (!isNaN(n13) && n13 === 13) {
          return "__NOGRADE__";
        }
      }
    }
    return canonicalStudentGradeCode(d.Grade) || "__UNK__";
  }

  function detailIncludedBySandboxGradeToggle(d) {
    if (!d) {
      return true;
    }
    var gC = sandboxGradeCanonicalForDetail(d);
    var t = BOUNDARY_SANDBOX.gradeToggles;
    if (t && t[gC] === false) {
      return false;
    }
    return true;
  }

  function detailIncludedBySandboxAttendanceTypeToggle(d) {
    if (!d) {
      return true;
    }
    var cat = sandboxAttendanceCategoryForDetail(d);
    var t2 = BOUNDARY_SANDBOX.attendanceTypeToggles;
    if (t2 && t2[cat] === false) {
      return false;
    }
    return true;
  }

  /**
   * Buckets current enrollment (MSID) for charting: charter / choice, else zoned traditional vs other traditional.
   * Uses `zonedMsidForDetailForAggregate` vs attendance MSID for the public-assignment “zoned” match.
   */
  function sandboxAttendanceCategoryForDetail(d) {
    if (!d) {
      return "otherTraditional";
    }
    if (d.__homeschool) {
      return "homeschool";
    }
    var att = parseInt(String(d.MSID != null ? d.MSID : "").trim(), 10);
    if (isNaN(att) || att <= 0) {
      return "otherTraditional";
    }
    var attK = String(att);
    if (CHARTER_SCHOOL_MSIDS && CHARTER_SCHOOL_MSIDS[attK]) {
      return "charter";
    }
    if (CHOICE_SCHOOL_MSIDS && CHOICE_SCHOOL_MSIDS[attK]) {
      return "choice";
    }
    var zoned = zonedMsidForDetailForAggregate(d);
    if (zoned != null && !isNaN(zoned) && Number(zoned) === att) {
      return "zonedTraditional";
    }
    return "otherTraditional";
  }

  function isSandboxAttendanceTypeKeyIncludedForFilter(atype) {
    var t = BOUNDARY_SANDBOX.attendanceTypeToggles;
    if (t && t[atype] === false) {
      return false;
    }
    return true;
  }

  function syncSandboxAttendanceTypeTogglesFromFull(fullByType) {
    var t = BOUNDARY_SANDBOX.attendanceTypeToggles;
    if (!t) {
      t = Object.create(null);
      BOUNDARY_SANDBOX.attendanceTypeToggles = t;
    }
    var allKeys = [
      "zonedTraditional",
      "otherTraditional",
      "charter",
      "choice",
      "homeschool",
    ];
    for (var i = 0; i < allKeys.length; i++) {
      var gk = allKeys[i];
      if (
        Object.prototype.hasOwnProperty.call(t, gk) &&
        (fullByType[gk] == null || fullByType[gk] === 0)
      ) {
        delete t[gk];
      }
    }
    for (var j = 0; j < allKeys.length; j++) {
      var fk = allKeys[j];
      if ((fullByType[fk] || 0) > 0) {
        if (t[fk] === undefined) {
          t[fk] = true;
        }
      }
    }
  }

  /**
   * Renders the same control layout as the grade bar chart, with a checkbox per type and colored bars
   * when included (dimmed + `is-excluded` when unchecked, same as grade).
   * @param {Object<string, number>|undefined} byType unfiltered row counts in the full hex set
   * @param {number|undefined} selectionTotalAll students in hex selection (ignores checkboxes); footer total line
   * @param {number|undefined} includedInDetails cohort passing grade + attendance toggles (lists / demographics)
   */
  function formatSandboxAttendanceTypeBarHtml(byType, selectionTotalAll, includedInDetails) {
    var defRows = [
      { key: "zonedTraditional", label: "Zoned Traditional School", mod: "zoned" },
      { key: "otherTraditional", label: "Other Traditional School", mod: "other" },
      { key: "charter", label: "Charter School", mod: "charter" },
      { key: "choice", label: "Choice School", mod: "choice" },
      { key: "homeschool", label: "Homeschool", mod: "homeschool" },
    ];
    var rows = [];
    for (var d = 0; d < defRows.length; d++) {
      var cPre = (byType && byType[defRows[d].key]) || 0;
      if (cPre > 0) {
        rows.push(defRows[d]);
      }
    }
    if (!rows.length) {
      return "<p class=\"sandbox-stat-line\">—</p>";
    }
    var maxC = 0;
    var rowSum = 0;
    for (var t = 0; t < rows.length; t++) {
      var c0 = (byType && byType[rows[t].key]) || 0;
      rowSum += c0;
      if (c0 > maxC) {
        maxC = c0;
      }
    }
    var mapTotal =
      selectionTotalAll != null && !isNaN(Number(selectionTotalAll))
        ? Number(selectionTotalAll)
        : rowSum;
    var includedTotal =
      includedInDetails != null && !isNaN(Number(includedInDetails))
        ? Number(includedInDetails)
        : rowSum;
    if (maxC <= 0) {
      return "<p class=\"sandbox-stat-line\">—</p>";
    }
    var parts = [
      '<div class="sandbox-grade-chart sandbox-grade-chart--attendance" role="group" aria-label="Students by school type in this selection">',
    ];
    for (var k = 0; k < rows.length; k++) {
      var row = rows[k];
      var c = (byType && byType[row.key]) || 0;
      var inc = isSandboxAttendanceTypeKeyIncludedForFilter(row.key);
      var wPct = Math.max(0, Math.min(100, Math.round((c / maxC) * 100)));
      var title = c + " student" + (c === 1 ? "" : "s") + " — " + row.label;
      var aLab = row.label + (inc ? " — include in details below" : " — exclude from details below");
      var chk = inc ? " checked" : "";
      var innerCls = "sandbox-grade-bar-inner sandbox-atype--" + row.mod;
      if (!inc) {
        innerCls += " is-excluded";
      }
      parts.push(
        '<div class="sandbox-grade-row" data-sandbox-atype-row="' +
          escapeHtml(String(row.key)) +
          '">' +
          '<div class="sandbox-grade-check" title="Include in lists and demographics below">' +
          "<input" +
          chk +
          ' type="checkbox" class="sandbox-attendance-type-toggle" data-atype="' +
          escapeHtml(String(row.key)) +
          '" aria-label="' +
          escapeHtml(aLab) +
          '" title="' +
          escapeHtml("Include in details below: " + row.label) +
          '" />' +
          "</div>" +
          '<div class="sandbox-grade-label-col">' +
          escapeHtml(row.label) +
          "</div>" +
          '<div class="sandbox-grade-bar-area"><div class="sandbox-grade-bar-outer" title="' +
          escapeHtml(title) +
          '"><div class="' +
          innerCls +
          '" style="width:' +
          wPct +
          '%"></div></div></div>' +
          '<div class="sandbox-grade-count-col">' +
          c.toLocaleString() +
          "</div></div>"
      );
    }
    if (mapTotal > 0) {
      var totLine = "In selection (all types): " + mapTotal.toLocaleString() + " students";
      if (includedTotal !== mapTotal) {
        totLine +=
          " · included in details below: " + includedTotal.toLocaleString() + " students";
      } else {
        totLine += " (all included in details below)";
      }
      parts.push('<p class="sandbox-grade-total">' + escapeHtml(totLine) + "</p>");
    }
    parts.push("</div>");
    return parts.join("");
  }

  function syncSandboxGradeTogglesFromFullByGrade(fullByGrade) {
    var t = BOUNDARY_SANDBOX.gradeToggles;
    if (!t) {
      t = Object.create(null);
      BOUNDARY_SANDBOX.gradeToggles = t;
    }
    for (var gk in t) {
      if (Object.prototype.hasOwnProperty.call(t, gk)) {
        var cLeft = fullByGrade[gk];
        if (cLeft == null || cLeft === 0) {
          delete t[gk];
        }
      }
    }
    for (var fk in fullByGrade) {
      if (Object.prototype.hasOwnProperty.call(fullByGrade, fk) && (fullByGrade[fk] || 0) > 0) {
        if (t[fk] === undefined) {
          t[fk] = true;
        }
      }
    }
  }

  /**
   * @param {Object<string, boolean>|undefined} hexKeyBag Hex keys to aggregate (defaults to current map selection).
   */
  function aggregateBoundarySandboxSelectionFromIndex(hexKeyBag) {
    var out = {
      totalStudents: 0,
      /** All students in the hex selection (ignores grade / attendance checkboxes). */
      selectionTotalAllStudents: 0,
      byGrade: {},
      byAttendance: {},
      byAttendanceTypeFull: {},
      byZoned: {},
      ethnicity: {},
      lunch: {},
    };
    var hasTrad = STUDENT_HEX_INDEX && STUDENT_HEX_INDEX.detailsByMsid;
    var hmByHex = HOMESCHOOL_DETAILS_BY_HEX_KEY;
    var hasHm = !!(hmByHex && Object.keys(hmByHex).length);
    if (!hasTrad && !hasHm) {
      return out;
    }
    var keyBag = hexKeyBag || BOUNDARY_SANDBOX.selectedHexKeys;
    var byDet = hasTrad ? STUDENT_HEX_INDEX.detailsByMsid : null;
    var fullByGrade = {};
    /** Attendance-type histogram with grade toggles applied (symmetric to grade chart using attendance toggles). */
    var fullByAT = Object.create(null);
    /** Grade histogram with attendance-type toggles applied (symmetric to attendance chart using grade toggles). */
    var gradeByAttendanceFilter = {};
    if (hasTrad) {
      for (var attMs0 in byDet) {
        if (!Object.prototype.hasOwnProperty.call(byDet, attMs0)) {
          continue;
        }
        var hexMap0 = byDet[attMs0];
        for (var hk0 in keyBag) {
          if (!keyBag[hk0]) {
            continue;
          }
          var arr0 = hexMap0[hk0];
          if (!arr0 || !arr0.length) {
            continue;
          }
          for (var i0 = 0; i0 < arr0.length; i0++) {
            var d0 = arr0[i0];
            if (!d0) {
              continue;
            }
            var g0 = sandboxGradeCanonicalForDetail(d0);
            fullByGrade[g0] = (fullByGrade[g0] || 0) + 1;
            if (detailIncludedBySandboxAttendanceTypeToggle(d0)) {
              gradeByAttendanceFilter[g0] = (gradeByAttendanceFilter[g0] || 0) + 1;
            }
            if (detailIncludedBySandboxGradeToggle(d0)) {
              var aCat = sandboxAttendanceCategoryForDetail(d0);
              fullByAT[aCat] = (fullByAT[aCat] || 0) + 1;
            }
          }
        }
      }
    }
    if (hmByHex) {
      for (var hkHs in keyBag) {
        if (!keyBag[hkHs]) {
          continue;
        }
        var hmArr0 = hmByHex[hkHs];
        if (!hmArr0 || !hmArr0.length) {
          continue;
        }
        for (var ih0 = 0; ih0 < hmArr0.length; ih0++) {
          var dh0 = hmArr0[ih0];
          if (!dh0) {
            continue;
          }
          var gHs = sandboxGradeCanonicalForDetail(dh0);
          fullByGrade[gHs] = (fullByGrade[gHs] || 0) + 1;
          if (detailIncludedBySandboxAttendanceTypeToggle(dh0)) {
            gradeByAttendanceFilter[gHs] = (gradeByAttendanceFilter[gHs] || 0) + 1;
          }
          if (detailIncludedBySandboxGradeToggle(dh0)) {
            var aCatHs = sandboxAttendanceCategoryForDetail(dh0);
            fullByAT[aCatHs] = (fullByAT[aCatHs] || 0) + 1;
          }
        }
      }
    }
    for (var gFill in fullByGrade) {
      if (Object.prototype.hasOwnProperty.call(fullByGrade, gFill)) {
        if (gradeByAttendanceFilter[gFill] == null) {
          gradeByAttendanceFilter[gFill] = 0;
        }
      }
    }
    var selTot = 0;
    for (var stKey in fullByGrade) {
      if (Object.prototype.hasOwnProperty.call(fullByGrade, stKey)) {
        selTot += fullByGrade[stKey] || 0;
      }
    }
    out.selectionTotalAllStudents = selTot;
    out.byGrade = gradeByAttendanceFilter;
    out.byAttendanceTypeFull = fullByAT;
    syncSandboxGradeTogglesFromFullByGrade(fullByGrade);
    syncSandboxAttendanceTypeTogglesFromFull(fullByAT);

    if (hasTrad) {
      for (var attMs in byDet) {
        if (!Object.prototype.hasOwnProperty.call(byDet, attMs)) {
          continue;
        }
        var hexMap = byDet[attMs];
        for (var hk in keyBag) {
          if (!keyBag[hk]) {
            continue;
          }
          var arr = hexMap[hk];
          if (!arr || !arr.length) {
            continue;
          }
          for (var i = 0; i < arr.length; i++) {
            var d = arr[i];
            if (!d) {
              continue;
            }
            if (!detailIncludedBySandboxGradeToggle(d)) {
              continue;
            }
            if (!detailIncludedBySandboxAttendanceTypeToggle(d)) {
              continue;
            }
            out.totalStudents += 1;
            var am = parseInt(String(d.MSID).trim(), 10);
            if (!isNaN(am)) {
              var aKey = String(am);
              out.byAttendance[aKey] = (out.byAttendance[aKey] || 0) + 1;
            }
            var zm = zonedMsidForDetailForAggregate(d);
            var zKey = zm != null ? String(zm) : "__none__";
            out.byZoned[zKey] = (out.byZoned[zKey] || 0) + 1;
            var eth =
              d.ethnicity != null && String(d.ethnicity).trim() !== ""
                ? String(d.ethnicity).trim()
                : "Unspecified";
            out.ethnicity[eth] = (out.ethnicity[eth] || 0) + 1;
            var lNorm = normalizeSandboxLunchStatForPie(d.lunch_stat);
            out.lunch[lNorm] = (out.lunch[lNorm] || 0) + 1;
          }
        }
      }
    }
    if (hmByHex) {
      for (var hkHm in keyBag) {
        if (!keyBag[hkHm]) {
          continue;
        }
        var hmArr = hmByHex[hkHm];
        if (!hmArr || !hmArr.length) {
          continue;
        }
        for (var jh = 0; jh < hmArr.length; jh++) {
          var dh = hmArr[jh];
          if (!dh) {
            continue;
          }
          if (!detailIncludedBySandboxGradeToggle(dh)) {
            continue;
          }
          if (!detailIncludedBySandboxAttendanceTypeToggle(dh)) {
            continue;
          }
          out.totalStudents += 1;
          var amh = parseInt(String(dh.MSID).trim(), 10);
          if (!isNaN(amh)) {
            var aKeyh = String(amh);
            out.byAttendance[aKeyh] = (out.byAttendance[aKeyh] || 0) + 1;
          }
          var zmH = zonedMsidForDetailForAggregate(dh);
          var zKeyH;
          if (dh.__homeschool && sandboxGradeCanonicalForDetail(dh) === "__NOGRADE__") {
            zKeyH = "__homeschool_not_age_eligible__";
          } else {
            zKeyH = zmH != null ? String(zmH) : "__none__";
          }
          out.byZoned[zKeyH] = (out.byZoned[zKeyH] || 0) + 1;
        }
      }
    }
    if (out.lunch && out.lunch.Unspecified) {
      out.lunch["Not free/reduced"] = (out.lunch["Not free/reduced"] || 0) + out.lunch.Unspecified;
      delete out.lunch.Unspecified;
    }
    return out;
  }

  function findSchoolPropertiesFromGeoCacheByMsid(msid) {
    if (msid == null || isNaN(msid)) {
      return null;
    }
    var target = Number(msid);
    var fc = GEO_CACHE.schools;
    if (!fc || !fc.features) {
      return null;
    }
    for (var i = 0; i < fc.features.length; i++) {
      var p = fc.features[i].properties;
      if (p && Number(p.SCHOOLS_ID) === target) {
        return p;
      }
    }
    return null;
  }

  function sandboxDisplayNameForMsidKey(msidStr) {
    if (msidStr === "__homeschool_not_age_eligible__") {
      return "No Zoned School - Not Age Eligible";
    }
    if (msidStr === "__none__") {
      return "Zoning not set";
    }
    if (String(msidStr) === String(HOMESCHOOL_ATTENDANCE_MSID)) {
      return "Home Education (Homeschool)";
    }
    var n = parseInt(String(msidStr), 10);
    if (isNaN(n)) {
      return String(msidStr);
    }
    var props = findSchoolPropertiesFromGeoCacheByMsid(n);
    if (props) {
      return schoolDisplayNameFromProps(props);
    }
    var m = masterRow(n);
    if (m && m.school_name) {
      return formatSchoolDisplayName(standardCapitalization(expandElemSchoolName(m.school_name)));
    }
    if (m && m.CommonName) {
      return formatSchoolDisplayName(standardCapitalization(expandElemSchoolName(String(m.CommonName))));
    }
    return "Unlisted school (ID " + n + ")";
  }

  function isSandboxGradeKeyIncludedForFilter(gCanon) {
    var t = BOUNDARY_SANDBOX.gradeToggles;
    if (t && t[gCanon] === false) {
      return false;
    }
    return true;
  }

  /** Whether every grade row is included, every row excluded, or mixed (for select-all UI). */
  function sandboxGradeFilterAggregateState(byGrade) {
    var keys = Object.keys(byGrade || {});
    if (!keys.length) {
      return { allOn: false, allOff: false, keysCount: 0 };
    }
    var allOn = true;
    var allOff = true;
    for (var i = 0; i < keys.length; i++) {
      if (isSandboxGradeKeyIncludedForFilter(keys[i])) {
        allOff = false;
      } else {
        allOn = false;
      }
    }
    return { allOn: allOn, allOff: allOff, keysCount: keys.length };
  }

  /**
   * @param {number|undefined} selectionTotalAll students in hex selection (ignores checkboxes)
   * @param {number|undefined} includedInDetails cohort passing grade + attendance toggles
   */
  function formatSandboxGradeBarChartHtml(byGrade, selectionTotalAll, includedInDetails) {
    var keys = Object.keys(byGrade);
    if (!keys.length) {
      return "<p class=\"sandbox-stat-line\">—</p>";
    }
    keys.sort(function (a, b) {
      return travelShedGradeSortKey(a) - travelShedGradeSortKey(b);
    });
    var maxC = 0;
    var rowSum = 0;
    var allGradeFiltersOn = true;
    var allGradeFiltersOff = true;
    for (var t = 0; t < keys.length; t++) {
      var c0 = byGrade[keys[t]] || 0;
      rowSum += c0;
      var incRow = isSandboxGradeKeyIncludedForFilter(keys[t]);
      if (incRow) {
        allGradeFiltersOff = false;
      } else {
        allGradeFiltersOn = false;
      }
      if (c0 > maxC) {
        maxC = c0;
      }
    }
    var mapTotal =
      selectionTotalAll != null && !isNaN(Number(selectionTotalAll))
        ? Number(selectionTotalAll)
        : rowSum;
    var includedTotal =
      includedInDetails != null && !isNaN(Number(includedInDetails))
        ? Number(includedInDetails)
        : rowSum;
    if (maxC <= 0) {
      return "<p class=\"sandbox-stat-line\">—</p>";
    }
    var parts = [
      '<div class="sandbox-grade-chart" role="group" aria-label="Students by grade in this selection">',
    ];
    var selAllChecked = allGradeFiltersOn && keys.length > 0;
    parts.push(
      '<div class="sandbox-grade-row sandbox-grade-row--select-all">' +
        '<div class="sandbox-grade-check">' +
        "<input" +
        (selAllChecked ? " checked" : "") +
        ' type="checkbox" class="sandbox-grade-select-all" ' +
        'aria-label="Select or clear all grades in this list" ' +
        'title="Select or clear all grades" />' +
        "</div>" +
        '<div class="sandbox-grade-label-col sandbox-grade-label-col--select-all">All</div>' +
        '<div class="sandbox-grade-bar-area" aria-hidden="true"></div>' +
        '<div class="sandbox-grade-count-col" aria-hidden="true"></div>' +
        "</div>"
    );
    for (var k = 0; k < keys.length; k++) {
      var key = keys[k];
      var c = byGrade[key] || 0;
      var inc = isSandboxGradeKeyIncludedForFilter(key);
      var labFull = travelShedGradeDisplayLabel(key);
      var lab = key === "__NOGRADE__" ? "NG" : labFull;
      var wPct = Math.max(0, Math.min(100, Math.round((c / maxC) * 100)));
      var title =
        key === "__NOGRADE__"
          ? c + " student" + (c === 1 ? "" : "s") + " (no grade code)"
          : c + " student" + (c === 1 ? "" : "s") + ", grade " + labFull;
      var aLab =
        key === "__NOGRADE__"
          ? "No grade code" + (inc ? " — include in details below" : " — exclude from details below")
          : (labFull === "Unknown" ? "Unknown or unspecified" : "Grade " + labFull) +
            (inc ? " — include in details below" : " — exclude from details below");
      var chk = inc ? " checked" : "";
      var toggleTitleShort =
        key === "__NOGRADE__"
          ? "Include in attendance, zoned, and demographics: no grade code"
          : "Include in attendance, zoned, and demographics: " + labFull;
      parts.push(
        '<div class="sandbox-grade-row" data-sandbox-grade-row="' +
          escapeHtml(String(key)) +
          '">' +
          '<div class="sandbox-grade-check" title="Include in attendance, zoned, and demographic counts">' +
          "<input" +
          chk +
          ' type="checkbox" class="sandbox-grade-toggle" data-grade-canon="' +
          escapeHtml(String(key)) +
          '" aria-label="' +
          escapeHtml(aLab) +
          '" title="' +
          escapeHtml(toggleTitleShort) +
          '" />' +
          "</div>" +
          '<div class="sandbox-grade-label-col">' +
          escapeHtml(lab) +
          "</div>" +
          '<div class="sandbox-grade-bar-area"><div class="sandbox-grade-bar-outer" title="' +
          escapeHtml(title) +
          '"><div class="sandbox-grade-bar-inner' +
          (inc ? "" : " is-excluded") +
          '" style="width:' +
          wPct +
          '%"></div></div></div>' +
          '<div class="sandbox-grade-count-col">' +
          c.toLocaleString() +
          "</div></div>"
      );
    }
    if (mapTotal > 0) {
      var totLine = "In selection (all grades): " + mapTotal.toLocaleString() + " students";
      if (includedTotal !== mapTotal) {
        totLine +=
          " · included in details below: " + includedTotal.toLocaleString() + " students";
      } else {
        totLine += " (all included in details below)";
      }
      parts.push('<p class="sandbox-grade-total">' + escapeHtml(totLine) + "</p>");
    }
    parts.push("</div>");
    return parts.join("");
  }

  function oneSandboxSchoolLineRow(ms, countByMsid) {
    var nm = sandboxDisplayNameForMsidKey(ms);
    return (
      "<div class=\"sandbox-stat-line\"><span class=\"sandbox-stat-label\">" +
      escapeHtml(nm) +
      "</span> <span class=\"sandbox-stat-val\">" +
      (countByMsid[ms] || 0).toLocaleString() +
      "</span></div>"
    );
  }

  function formatSandboxSchoolListHtml(countByMsid, maxList, panelKey) {
    var key = panelKey != null && panelKey !== "" ? String(panelKey) : "attendance";
    var st = BOUNDARY_SANDBOX.schoolListExpanded;
    if (!st) {
      st = { attendance: false, zoned: false };
      BOUNDARY_SANDBOX.schoolListExpanded = st;
    }
    var keys = Object.keys(countByMsid);
    if (!keys.length) {
      return "<p class=\"sandbox-stat-line\">—</p>";
    }
    keys.sort(function (a, b) {
      return (countByMsid[b] || 0) - (countByMsid[a] || 0);
    });
    var max = maxList != null && maxList > 0 ? maxList : 5;
    var parts = ['<div class="sandbox-school-list">'];
    if (keys.length <= max) {
      for (var k0 = 0; k0 < keys.length; k0++) {
        parts.push(oneSandboxSchoolLineRow(keys[k0], countByMsid));
      }
      parts.push("</div>");
      return parts.join("");
    }
    var restC = 0;
    for (var r0 = max; r0 < keys.length; r0++) {
      restC += countByMsid[keys[r0]] || 0;
    }
    var expanded = !!st[key];
    if (!expanded) {
      for (var t = 0; t < max; t++) {
        parts.push(oneSandboxSchoolLineRow(keys[t], countByMsid));
      }
      var moreLine =
        "+" +
        (keys.length - max) +
        " more (includes " +
        restC.toLocaleString() +
        (restC === 1 ? " more student" : " more students") +
        ") — show all";
      parts.push(
        "<button " +
          'type="button" ' +
          'class="sandbox-school-expand" ' +
          'data-panel="' +
          escapeHtml(key) +
          '" ' +
          'aria-expanded="false" ' +
          ">" +
          escapeHtml(moreLine) +
          "</button>"
      );
    } else {
      parts.push(
        '<div class="sandbox-school-list-scroll" role="list" aria-label="Schools in this list">'
      );
      for (var r = 0; r < keys.length; r++) {
        parts.push(oneSandboxSchoolLineRow(keys[r], countByMsid));
      }
      parts.push(
        '</div><button type="button" class="sandbox-school-expand" data-panel="' +
          escapeHtml(key) +
          '" aria-expanded="true">' +
          escapeHtml("Show less") +
          "</button>"
      );
    }
    parts.push("</div>");
    return parts.join("");
  }

  function renderSandboxHexLayerDemographicPies(ethByLabel, lunchByLabel) {
    var ethEl = document.getElementById("sandbox-demographics-ethnicity");
    var lunchEl = document.getElementById("sandbox-demographics-lunch");
    if (!ethEl || !lunchEl) {
      return;
    }
    var emptyMsg =
      '<p class="demographics-pie-empty">No students with valid ethnicity in this layer for the selection.</p>';
    var emptyLunch = '<p class="demographics-pie-empty">No students with valid lunch in this layer for the selection.</p>';
    var ethRes = buildPieChartHtml(ethByLabel || {}, ethnicitySliceColor);
    var lunchRes = buildPieChartHtml(lunchByLabel || {}, function (label) {
      return lunchSliceColor(label);
    });
    ethEl.innerHTML = ethRes.total > 0 ? ethRes.html : emptyMsg;
    lunchEl.innerHTML = lunchRes.total > 0 ? lunchRes.html : emptyLunch;
  }

  function updateSandboxStatsPanelSummary() {
    var h = document.getElementById("sandbox-stats-heading");
    var lead = document.getElementById("sandbox-stats-lead");
    if (!h || !lead) {
      return;
    }
    var statsKeys = getHexKeysForSandboxStatistics();
    if (!statsKeys || countSandboxHexKeys(statsKeys) === 0) {
      h.textContent = "Students in selection";
      lead.innerHTML =
        "Choose hexes on the map (or a base school), then <strong>Confirm</strong> for grade, attendance, zoned, and demographics from the student hex layer.";
      clearSandboxStatsAndDemographicsDisplays();
      return;
    }
    var nHex = countSandboxHexKeys(statsKeys);
    var showPendingHint =
      !BOUNDARY_SANDBOX.selectionConfirmed &&
      countSandboxHexKeys(BOUNDARY_SANDBOX.confirmedHexKeysSnapshot) > 0;
    var agg = aggregateBoundarySandboxSelectionFromIndex(statsKeys);
    var totalInHex =
      agg.selectionTotalAllStudents != null && !isNaN(Number(agg.selectionTotalAllStudents))
        ? Number(agg.selectionTotalAllStudents)
        : 0;
    h.textContent = "Students in selection (" + nHex + (nHex === 1 ? " hex" : " hexes") + ")";
    var inHexStr = totalInHex.toLocaleString();
    var inHexNoun = totalInHex === 1 ? "student" : "students";
    var atB = document.getElementById("sandbox-card-body-attendance-type");
    var gB = document.getElementById("sandbox-card-body-grade");
    var aB = document.getElementById("sandbox-card-body-attendance");
    var zB = document.getElementById("sandbox-card-body-zoned");
    var suppressDetailedStats = totalInHex <= 10;
    var suppressionLead =
      "Detailed statistics are hidden when the filtered selection contains too few students.";
    if (suppressDetailedStats) {
      if (showPendingHint) {
        lead.innerHTML =
          "<strong>Statistics below reflect your last confirmed selection.</strong> The map may show unsaved edits — click <strong>Confirm selection</strong> to refresh these figures. " +
          suppressionLead;
      } else {
        lead.innerHTML = suppressionLead;
      }
      if (atB) atB.innerHTML = '<p class="sandbox-stat-line">—</p>';
      if (gB) gB.innerHTML = '<p class="sandbox-stat-line">—</p>';
      if (aB) aB.innerHTML = '<p class="sandbox-stat-line">—</p>';
      if (zB) zB.innerHTML = '<p class="sandbox-stat-line">—</p>';
      var ethSup = document.getElementById("sandbox-demographics-ethnicity");
      var lunchSup = document.getElementById("sandbox-demographics-lunch");
      if (ethSup) {
        ethSup.innerHTML =
          '<p class="demographics-pie-empty">Detailed demographics appear when more students are included.</p>';
      }
      if (lunchSup) {
        lunchSup.innerHTML =
          '<p class="demographics-pie-empty">Detailed demographics appear when more students are included.</p>';
      }
      return;
    }
    if (showPendingHint) {
      lead.innerHTML =
        "<strong>Statistics below reflect your last confirmed selection.</strong> The map may show unsaved edits — click <strong>Confirm selection</strong> to refresh these figures. " +
        inHexStr +
        " " +
        inHexNoun +
        " " +
        (totalInHex === 1 ? "lives" : "live") +
        " in that confirmed hex set. Toggle grades or school types to exclude them from statistics below.";
    } else {
      lead.textContent =
        inHexStr +
        " " +
        inHexNoun +
        " live in the selected hex cells. Toggle grades or school types to exclude them from statistics below.";
    }
    var detailIncluded =
      agg.totalStudents != null && !isNaN(Number(agg.totalStudents)) ? Number(agg.totalStudents) : 0;
    if (atB) {
      atB.innerHTML = formatSandboxAttendanceTypeBarHtml(
        agg.byAttendanceTypeFull || {},
        totalInHex,
        detailIncluded
      );
    }
    if (gB) {
      gB.innerHTML = formatSandboxGradeBarChartHtml(agg.byGrade, totalInHex, detailIncluded);
      var gSelAll = gB.querySelector(".sandbox-grade-select-all");
      if (gSelAll) {
        var gAgg = sandboxGradeFilterAggregateState(agg.byGrade);
        gSelAll.indeterminate =
          gAgg.keysCount > 0 && !gAgg.allOn && !gAgg.allOff;
      }
    }
    if (aB) aB.innerHTML = formatSandboxSchoolListHtml(agg.byAttendance, 5, "attendance");
    if (zB) zB.innerHTML = formatSandboxSchoolListHtml(agg.byZoned, 5, "zoned");
    /* Demographics use the grade + attendance-type filtered cohort only (`totalStudents`). */
    var filteredForDemographics = agg.totalStudents != null ? agg.totalStudents : 0;
    if (filteredForDemographics <= 5) {
      var ethDemo = document.getElementById("sandbox-demographics-ethnicity");
      var lunchDemo = document.getElementById("sandbox-demographics-lunch");
      if (ethDemo) {
        ethDemo.innerHTML =
          '<p class="demographics-pie-empty">Detailed demographics appear when more students are included.</p>';
      }
      if (lunchDemo) {
        lunchDemo.innerHTML =
          '<p class="demographics-pie-empty">Detailed demographics appear when more students are included.</p>';
      }
    } else {
      renderSandboxHexLayerDemographicPies(agg.ethnicity, agg.lunch);
    }
  }

  function downloadTextAsCsvFile(filename, text) {
    var out = "\uFEFF" + text;
    var blob = new Blob([out], { type: "text/csv;charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.setAttribute("aria-hidden", "true");
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  /**
   * Replaces sandbox hex selection with all hexes that have at least one student
   * zoned to the base school (same rules as the student-residence “zoned” view).
   */
  function prefillBoundarySandboxZonedHexesForBaseMsid(baseMsid) {
    BOUNDARY_SANDBOX.selectedHexKeys = Object.create(null);
    BOUNDARY_SANDBOX.selectionConfirmed = false;
    BOUNDARY_SANDBOX.confirmedHexKeysSnapshot = Object.create(null);
    resetBoundarySandboxFilterState();
    if (baseMsid == null || isNaN(baseMsid)) {
      return;
    }
    if (selectedSchoolDisallowsZonedStudentHex(baseMsid)) {
      return;
    }
    var m = masterRow(baseMsid);
    if (!m) {
      return;
    }
    var zMap = collectZonedDetailsByHex(baseMsid, m, false);
    for (var hk in zMap) {
      if (!Object.prototype.hasOwnProperty.call(zMap, hk)) {
        continue;
      }
      var arr = zMap[hk];
      if (arr && arr.length) {
        BOUNDARY_SANDBOX.selectedHexKeys[hk] = true;
      }
    }
    var hmInPoly = homeschoolHexKeysWithCentroidInAssignmentBoundary(baseMsid);
    for (var hmk in hmInPoly) {
      if (hmInPoly[hmk]) {
        BOUNDARY_SANDBOX.selectedHexKeys[hmk] = true;
      }
    }
    syncSandboxLassoFootprintFromSelectedHexGeometries();
    applyBoundarySandboxSelectionFeatureStates();
  }

  /** Invisible one-hex-per-feature layer; selection shown via feature-state. */
  function rebuildBoundarySandboxHexSourceFromIndex() {
    if (!map || !map.getSource("boundary-sandbox-hex")) return;
    if (!STUDENT_HEX_INDEX || !STUDENT_HEX_INDEX.geometryByHexKey) {
      BOUNDARY_SANDBOX.selectedHexKeys = Object.create(null);
      resetBoundarySandboxFilterState();
      try {
        map.getSource("boundary-sandbox-hex").setData({
          type: "FeatureCollection",
          features: [],
        });
      } catch (e0) {
        /* ignore */
      }
      requestApplyBoundarySandboxSelectionOnIdle();
      updateSandboxSelectedHexCountUi();
      return;
    }
    var gk = STUDENT_HEX_INDEX.geometryByHexKey;
    pruneBoundarySandboxSelectedKeysToGeometry();
    var feats = [];
    for (var k in gk) {
      if (!Object.prototype.hasOwnProperty.call(gk, k)) continue;
      var g = gk[k];
      if (!g) continue;
      feats.push({
        type: "Feature",
        properties: { _hexKey: k },
        geometry: g,
      });
    }
    if (HOMESCHOOL_HEX_GEOMETRY_FALLBACK) {
      for (var fk in HOMESCHOOL_HEX_GEOMETRY_FALLBACK) {
        if (!Object.prototype.hasOwnProperty.call(HOMESCHOOL_HEX_GEOMETRY_FALLBACK, fk)) {
          continue;
        }
        if (gk[fk]) {
          continue;
        }
        var gHm = HOMESCHOOL_HEX_GEOMETRY_FALLBACK[fk];
        if (!gHm) {
          continue;
        }
        feats.push({
          type: "Feature",
          properties: { _hexKey: fk },
          geometry: gHm,
        });
      }
    }
    try {
      map.getSource("boundary-sandbox-hex").setData({
        type: "FeatureCollection",
        features: feats,
      });
    } catch (e1) {
      /* ignore */
    }
    requestApplyBoundarySandboxSelectionOnIdle();
    updateSandboxSelectedHexCountUi();
  }

  function syncBoundarySandboxMapLayers() {
    if (!map || !map.getLayer("boundary-sandbox-hex-fill")) return;
    var vis = isBoundarySandboxViewActive() ? "visible" : "none";
    if (map.getLayer("boundary-sandbox-lasso-region-fill")) {
      try {
        map.setLayoutProperty("boundary-sandbox-lasso-region-fill", "visibility", vis);
      } catch (eLf) {
        /* ignore */
      }
    }
    if (map.getLayer("boundary-sandbox-lasso-region-outline")) {
      try {
        map.setLayoutProperty("boundary-sandbox-lasso-region-outline", "visibility", vis);
      } catch (eLo) {
        /* ignore */
      }
    }
    try {
      map.setLayoutProperty("boundary-sandbox-hex-fill", "visibility", vis);
    } catch (e) {
      /* ignore */
    }
    if (map.getLayer("boundary-sandbox-lasso-line")) {
      try {
        map.setLayoutProperty("boundary-sandbox-lasso-line", "visibility", vis);
      } catch (eL) {
        /* ignore */
      }
    }
    if (map.getLayer("boundary-sandbox-selection-outline-line")) {
      try {
        map.setLayoutProperty("boundary-sandbox-selection-outline-line", "visibility", vis);
      } catch (eO) {
        /* ignore */
      }
    }
    if (vis === "visible") {
      requestApplyBoundarySandboxSelectionOnIdle();
      updateBoundarySandboxSelectionOutline();
    }
  }

  /**
   * Keeps header + data rows whose msid (column 0) is in #school-select.
   * If the select has no school options yet, falls back to appears_in_dropdown=yes.
   */
  function filterCsvGridToDropdownSchools(grid) {
    if (!grid || grid.length < 2) return grid;
    var allowed = getSchoolDropdownMsidSet();
    var useSelect = Object.keys(allowed).length > 0;
    var headers = grid[0].map(function (h) {
      return String(h).trim();
    });
    var appearIdx = headers.indexOf("appears_in_dropdown");
    var out = [grid[0]];
    for (var r = 1; r < grid.length; r++) {
      var row = grid[r];
      if (!row || !row.length) continue;
      var idRaw = row[0] != null ? String(row[0]).trim() : "";
      var idNum = parseInt(idRaw, 10);
      if (isNaN(idNum)) continue;
      var include = false;
      if (useSelect) {
        include = !!(allowed[String(idNum)] || allowed[idRaw]);
      } else if (appearIdx >= 0) {
        var cell = row[appearIdx] != null ? String(row[appearIdx]).trim().toLowerCase() : "";
        include = cell === "yes";
      } else {
        include = true;
      }
      if (include) out.push(row);
    }
    return out;
  }

  /** Parcel GeoJSON may use SCHL_CODE with or without leading zeros; MSIDs match numerically. */
  function parcelPropertySchlCode(props) {
    if (!props) return null;
    var v =
      props.SCHL_CODE != null
        ? props.SCHL_CODE
        : props.Schl_Code != null
          ? props.Schl_Code
          : props.schl_code != null
            ? props.schl_code
            : null;
    if (v === null || v === "") return null;
    var n = Number(String(v).trim());
    return isNaN(n) ? null : n;
  }

  function schoolExcludedFromParcelOverlay(sp) {
    if (!sp) return true;
    var nm = String(sp.NAME || sp.CommonName || "").toUpperCase();
    if (nm.indexOf("CHARTER") >= 0) return true;
    return false;
  }

  /** @returns {"elementary"|"middle"|"high"|null} */
  function schoolParcelLevelFromType(sp) {
    if (!sp) return null;
    var t = String(sp.TYPE || "").toUpperCase();
    if (t === "ELEMENTARY") return "elementary";
    if (t === "MIDDLE") return "middle";
    if (t === "HIGH" || t === "JR SR HIGH") return "high";
    return null;
  }

  /**
   * Parcel styling level: Jr/Sr (7–12) is separate from 9–12 high so parcels can use orange.
   * Uses master CSV TYPE when present (same as school dots).
   * @returns {"elementary"|"middle"|"high"|"jr_sr"|null}
   */
  function schoolParcelStripeLevel(sp) {
    if (!sp) return null;
    var spM = schoolPropsWithMasterType(sp);
    var t = String(spM.TYPE || "").toUpperCase();
    if (t === "JR SR HIGH") return "jr_sr";
    if (t === "ELEMENTARY") return "elementary";
    if (t === "MIDDLE") return "middle";
    if (t === "HIGH") return "high";
    return null;
  }

  function buildFilteredSchoolParcelsFc(schoolsFc, parcelsFc) {
    var out = { type: "FeatureCollection", features: [] };
    if (!parcelsFc || !parcelsFc.features || !parcelsFc.features.length) {
      return out;
    }
    var byMsid = buildSchoolLookup(schoolsFc);
    for (var i = 0; i < parcelsFc.features.length; i++) {
      var ft = parcelsFc.features[i];
      var p = ft.properties || {};
      var msid = parcelPropertySchlCode(p);
      if (msid == null) continue;
      var sp = byMsid[msid];
      if (!sp) continue;
      if (schoolExcludedFromParcelOverlay(sp)) continue;
      var lvl = schoolParcelStripeLevel(sp);
      if (!lvl) continue;
      var geom = ft.geometry;
      if (!geom || (geom.type !== "Polygon" && geom.type !== "MultiPolygon")) {
        continue;
      }
      out.features.push({
        type: "Feature",
        geometry: geom,
        properties: { SCHOOLS_ID: msid, _parcelLevel: lvl },
      });
    }
    return out;
  }

  /** Re-apply toolbar layer checkbox visibility after map layers are recreated (e.g. basemap switch). */
  function resyncToolbarLayerToggleVisibility() {
    var panel = document.getElementById("toolbar-panel");
    if (!panel) return;
    panel.querySelectorAll('input[type="checkbox"][id^="toggle-"]').forEach(function (inp) {
      inp.dispatchEvent(new Event("change", { bubbles: true }));
    });
  }

  function appendToggleRow(container, def, onAfterChange) {
    var id = "toggle-" + def.id;
    var label = document.createElement("label");
    var input = document.createElement("input");
    input.type = "checkbox";
    input.id = id;
    input.checked =
      def.defaultChecked === undefined ? true : !!def.defaultChecked;
    function applyVisibilityToLayers() {
      var vis = input.checked ? "visible" : "none";
      def.layerIds.forEach(function (lid) {
        if (map.getLayer(lid)) map.setLayoutProperty(lid, "visibility", vis);
      });
    }
    applyVisibilityToLayers();
    input.addEventListener("change", function () {
      applyVisibilityToLayers();
      if (typeof onAfterChange === "function") onAfterChange();
    });
    label.appendChild(input);
    if (def.gradientStrip) {
      var gs = document.createElement("span");
      gs.className =
        "toggle-gradient-strip" +
        (def.gradientStripClass ? " " + def.gradientStripClass : "");
      gs.setAttribute("aria-hidden", "true");
      label.appendChild(gs);
    } else if (def.swatchVariant === "split-jr-sr") {
      var swSplit = document.createElement("span");
      swSplit.className = "swatch swatch--split-jr-sr";
      swSplit.setAttribute("aria-hidden", "true");
      label.appendChild(swSplit);
    } else if (def.swatchColor) {
      var sw = document.createElement("span");
      sw.className = "swatch";
      sw.style.background = def.swatchColor;
      sw.setAttribute("aria-hidden", "true");
      label.appendChild(sw);
    }
    if (def.sublabel) {
      var stack = document.createElement("span");
      stack.className = "toggle-label-stack";
      var main = document.createElement("span");
      main.className = "toggle-label-main";
      main.textContent = def.label;
      stack.appendChild(main);
      var sub = document.createElement("span");
      sub.className = "toggle-label-sub";
      sub.textContent = def.sublabel;
      stack.appendChild(sub);
      label.appendChild(stack);
    } else {
      var mainOnly = document.createElement("span");
      mainOnly.className = "toggle-label-main";
      mainOnly.textContent = def.label;
      label.appendChild(mainOnly);
    }
    container.appendChild(label);
  }

  function setupToggles() {
    var boundaryDefs = [
      {
        id: "es",
        label: "Elementary",
        layerIds: ["es-fill", "es-outline"],
        swatchColor: PALETTE.elementary.fill,
        defaultChecked: false,
      },
      {
        id: "ms",
        label: "Middle",
        layerIds: ["ms-fill", "ms-outline"],
        swatchColor: PALETTE.middle.fill,
        defaultChecked: false,
      },
      {
        id: "hs",
        label: "High",
        sublabel: "(incl. Jr/Sr)",
        layerIds: ["hs-fill", "hs-outline"],
        swatchColor: PALETTE.high.fill,
        defaultChecked: false,
      },
    ];
    var schoolDefs = [
      {
        id: "sch-es",
        label: "Elementary",
        layerIds: ["schools-elementary"],
        swatchColor: PALETTE.elementary.fill,
        defaultChecked: true,
      },
      {
        id: "sch-ms",
        label: "Middle",
        layerIds: ["schools-middle"],
        swatchColor: PALETTE.middle.fill,
        defaultChecked: true,
      },
      {
        id: "sch-hs",
        label: "High",
        sublabel: "(incl. Jr/Sr)",
        swatchVariant: "split-jr-sr",
        layerIds: ["schools-high"],
        defaultChecked: true,
      },
    ];

    var bEl = document.getElementById("boundary-toggles");
    var sEl = document.getElementById("school-toggles");
    boundaryDefs.forEach(function (def) {
      appendToggleRow(bEl, def, refreshAssignmentBoundaryHighlight);
    });
    schoolDefs.forEach(function (def) {
      appendToggleRow(sEl, def);
    });

    var parcelDefs = [
      {
        id: "parcel-es",
        label: "Elementary",
        layerIds: ["school-parcels-elementary"],
        swatchColor: PALETTE.elementary.fill,
        defaultChecked: false,
      },
      {
        id: "parcel-ms",
        label: "Middle",
        layerIds: ["school-parcels-middle"],
        swatchColor: PALETTE.middle.fill,
        defaultChecked: false,
      },
      {
        id: "parcel-hs",
        label: "High",
        sublabel: "(incl. Jr/Sr)",
        swatchVariant: "split-jr-sr",
        layerIds: ["school-parcels-jr-sr", "school-parcels-high"],
        defaultChecked: false,
      },
    ];
    var pEl = document.getElementById("school-parcel-toggles");
    if (pEl) {
      parcelDefs.forEach(function (def) {
        appendToggleRow(pEl, def);
      });
    }

    var hxEl = document.getElementById("student-hex-toggles");
    if (hxEl) {
      appendToggleRow(
        hxEl,
        {
          id: "student-hex",
          label: "Student residence density",
          layerIds: ["student-hex-heatmap", "student-hex-hit-fill"],
          gradientStrip: true,
          defaultChecked: false,
        },
        function () {
          var inp = document.getElementById("toggle-student-hex");
          if (inp && inp.checked) {
            syncStudentHexLayer();
          }
          syncStudentHexTooltipCheckboxVisibility();
          syncMapDensityLegend();
        }
      );
      var attMode = document.getElementById("toggle-student-hex-attending");
      var zonMode = document.getElementById("toggle-student-hex-zoned");
      if (attMode) {
        attMode.addEventListener("change", function () {
          syncStudentHexLayer();
        });
      }
      if (zonMode) {
        zonMode.addEventListener("change", function () {
          syncStudentHexLayer();
        });
      }
      syncStudentHexResidenceSubToggleAvailability();
    }

    var cHexEl = document.getElementById("charter-student-hex-toggles");
    if (cHexEl) {
      appendToggleRow(
        cHexEl,
        {
          id: "charter-student-hex",
          label: "Charter student residence density",
          layerIds: [
            "charter-student-hex-heatmap",
            "charter-student-hex-hit-fill",
          ],
          gradientStrip: true,
          gradientStripClass: "toggle-gradient-strip--charter-magenta",
          defaultChecked: false,
        },
        function () {
          syncCharterDistrictStudentHexLayer();
          syncStudentHexTooltipCheckboxVisibility();
          syncMapDensityLegend();
        }
      );
    }

    var hsHexEl = document.getElementById("homeschool-student-hex-toggles");
    if (hsHexEl) {
      appendToggleRow(
        hsHexEl,
        {
          id: "homeschool-student-hex",
          label: "Homeschool student residence density",
          layerIds: [
            "homeschool-student-hex-heatmap",
            "homeschool-student-hex-hit-fill",
          ],
          gradientStrip: true,
          gradientStripClass: "toggle-gradient-strip--homeschool-red",
          defaultChecked: false,
        },
        function () {
          syncHomeschoolStudentHexLayer();
          syncStudentHexTooltipCheckboxVisibility();
          syncMapDensityLegend();
        }
      );
    }
    if (document.getElementById("student-hex-tooltip-row") != null) {
      syncStudentHexTooltipCheckboxVisibility();
    }

    var travelShedEl = document.getElementById("travel-shed-toggles");
    if (travelShedEl) {
      appendToggleRow(
        travelShedEl,
        {
          id: "travel-sheds",
          label: "Travel sheds",
          layerIds: ["school-isochrones-fill", "school-isochrones-outline"],
          gradientStrip: true,
          gradientStripClass: "toggle-gradient-strip--travel-sheds",
          defaultChecked: false,
        },
        function () {
          syncTravelShedLayerFilter();
          syncTravelShedMaxMilesRowVisibility();
        }
      );
    }
    setupTravelShedMaxMilesControl();
    syncTravelShedMaxMilesRowVisibility();

    var sbdEl = document.getElementById("school-board-district-toggles");
    if (sbdEl) {
      appendToggleRow(sbdEl, {
        id: "school-board-districts",
        label: "School board districts",
        layerIds: ["school-board-districts-fill", "school-board-districts-outline"],
        swatchColor: "#374151",
        defaultChecked: false,
      });
    }

    var munEl = document.getElementById("municipal-boundary-toggles");
    if (munEl) {
      appendToggleRow(munEl, {
        id: "municipal-boundaries",
        label: "Municipal boundaries",
        layerIds: [
          "municipal-boundaries-fill",
          "municipal-boundaries-outline",
          "municipal-boundaries-hover",
        ],
        swatchColor: "#9ca3af",
        defaultChecked: false,
      });
    }

    var charterEl = document.getElementById("charter-school-toggles");
    if (charterEl) {
      appendToggleRow(charterEl, {
        id: "charter-schools",
        label: "Charter schools",
        layerIds: ["schools-charter"],
        swatchColor: PALETTE.charter.fill,
        defaultChecked: false,
      });
    }
    var privateSchoolTogglesEl = document.getElementById("private-school-toggles");
    if (privateSchoolTogglesEl) {
      appendToggleRow(privateSchoolTogglesEl, {
        id: "private-schools",
        label: "Private schools",
        layerIds: ["schools-private"],
        swatchColor: PALETTE.privateSchool.fill,
        defaultChecked: false,
      });
    }
    syncMapDensityLegend();
  }

  var BOUNDARY_FILL_LAYERS = ["es-fill", "ms-fill", "hs-fill"];
  var SCHOOL_LAYER_IDS = [
    "schools-elementary",
    "schools-middle",
    "schools-high",
    "schools-charter",
  ];
  /**
   * Map pick priority (top to bottom in stack) for each category.
   * Used with queryRenderedFeatures: first hit is the topmost visible in that set.
   */
  var SCHOOL_LAYERS_CLICK_TOP_FIRST = [
    "schools-private",
    "schools-charter",
    "schools-elementary",
    "schools-middle",
    "schools-high",
  ];
  var SCHOOL_PARCEL_LAYERS_CLICK_TOP_FIRST = [
    "school-parcels-elementary",
    "school-parcels-jr-sr",
    "school-parcels-middle",
    "school-parcels-high",
  ];
  var ASSIGNMENT_BOUNDARY_LAYERS_CLICK_TOP_FIRST = [
    "es-outline",
    "ms-outline",
    "hs-outline",
    "es-fill",
    "ms-fill",
    "hs-fill",
  ];

  /** Topmost paint order first: used so queryRenderedFeatures returns the visually top feature first. */
  var MAP_OVERLAY_HIT_LAYER_ORDER_TOP_FIRST = [
    "schools-private",
    "schools-charter",
    "schools-elementary",
    "schools-middle",
    "schools-high",
    "boundary-sandbox-hex-fill",
    "boundary-sandbox-lasso-region-outline",
    "boundary-sandbox-lasso-region-fill",
    "charter-student-hex-hit-fill",
    "homeschool-student-hex-hit-fill",
    "student-hex-hit-fill",
    "school-parcels-elementary",
    "school-parcels-jr-sr",
    "school-parcels-middle",
    "school-parcels-high",
    "school-isochrones-outline",
    "school-isochrones-fill",
    "es-outline",
    "ms-outline",
    "hs-outline",
    "es-fill",
    "ms-fill",
    "hs-fill",
    "school-board-districts-outline",
    "school-board-districts-fill",
    "municipal-boundaries-outline",
    "municipal-boundaries-fill",
  ];

  function boundaryLayerIdToSource(layerId) {
    if (layerId === "es-fill" || layerId === "es-outline") return "es-boundaries";
    if (layerId === "ms-fill" || layerId === "ms-outline") return "ms-boundaries";
    if (layerId === "hs-fill" || layerId === "hs-outline") return "hs-boundaries";
    return null;
  }

  /** Title-style capitalization for tooltip text (handles ALL CAPS source data). */
  function standardCapitalization(str) {
    if (str == null || str === "") return "";
    return String(str)
      .trim()
      .split(/\s+/)
      .map(function (word) {
        if (/^\d+$/.test(word)) return word;
        if (/^\d+[a-z]*$/i.test(word)) return word.charAt(0) + word.slice(1).toLowerCase();
        if (word.indexOf("-") !== -1) {
          return word
            .split("-")
            .map(function (part) {
              if (/^\d+$/.test(part)) return part;
              return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
            })
            .join("-");
        }
        return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
      })
      .join(" ");
  }

  /** GeoJSON sometimes uses "Elem" as shorthand; expand for display before title-casing. */
  function expandElemSchoolName(str) {
    if (str == null || str === "") return "";
    return String(str).replace(/\belem\b/gi, "elementary");
  }

  /**
   * Excel/Sheets often turn grade ranges like 9-12 or 7-8 into serial dates (e.g. 12-Sep, 8-Jul).
   * Normalizes those back to display ranges; pass through everything else.
   */
  function normalizeGradesServedForUi(raw) {
    if (raw == null || raw === "") return "";
    var t = String(raw).trim();
    /* Leading apostrophe = Excel “text” cell; strip before normalizing. */
    if (t.charAt(0) === "'") t = t.slice(1).trim();
    if (/^12-sep$/i.test(t)) return "9-12";
    if (/^12-jul$/i.test(t)) return "7-12";
    if (/^8-jul$/i.test(t)) return "7-8";
    if (/^6-apr$/i.test(t)) return "4-6";
    if (/^6-mar$/i.test(t)) return "3-6";
    return t;
  }

  /** Spell out W. Melbourne in city names (GeoJSON / CSV city lines). */
  function expandWestMelbourneCity(cityPart) {
    return String(cityPart).replace(/^W\.\s*Melbourne\b/i, "West Melbourne");
  }

  /** "CITY, ST 12345" → "City, ST 12345" */
  function formatCityStateZip(str) {
    if (!str) return "";
    var t = String(str).trim();
    var m = t.match(/^(.+),\s*([A-Za-z]{2})\s+(.+)$/);
    if (m) {
      var cityExpanded = expandWestMelbourneCity(m[1].trim());
      return (
        standardCapitalization(cityExpanded) +
        ", " +
        m[2].toUpperCase() +
        " " +
        m[3].trim()
      );
    }
    return standardCapitalization(expandWestMelbourneCity(t));
  }

  /**
   * UI polish for school display names (after standardCapitalization).
   * Covers Jr/Sr high labels, Turner/Creel/Williams/West Melbourne wording from mixed sources.
   */
  function formatSchoolDisplayName(str) {
    if (str == null || str === "") return "";
    var s = String(str);
    s = s.replace(/\bJunior\/Senior\b/gi, "Jr/Sr");
    s = s.replace(/\bJr\.?\s+Sr\.?\b/gi, "Jr/Sr");
    s = s.replace(/\bJR\s+SR\b/g, "Jr/Sr");
    s = s.replace(/\bJr\/sr\b/g, "Jr/Sr");
    s = s.replace(/,\s*Senior\b/gi, "");
    s = s.replace(/\bJohn F\.\s*Turner\s*,\s*Senior\b/gi, "John F. Turner");
    s = s.replace(/\bRalph M\s+Williams\b/gi, "Ralph M. Williams");
    s = s.replace(/\bW\.j\./gi, "W.J.");
    s = s.replace(/\bDr\.\s+W\.j\./gi, "Dr. W.J.");
    s = s.replace(/\bW\.\s*Melbourne\b/gi, "West Melbourne");
    s = s.replace(/\bMcnair\b/gi, "McNair");
    /* District CSV uses “… Elementary School For Science”; preferred public name drops the suffix. */
    s = s.replace(/\s+Elementary\s+School\s+For\s+Science$/i, " Elementary School");
    return s;
  }

  /** Prefer data/school_master.csv over GeoJSON NAME/CommonName (district GIS can have typos). */
  function schoolDisplayNamePreferMaster(p) {
    if (!p || p.SCHOOLS_ID == null || !MASTER_BY_MSID) return null;
    var sid = Number(p.SCHOOLS_ID);
    if (isNaN(sid)) return null;
    var m = masterRow(sid);
    if (!m || !m.school_name) return null;
    return formatSchoolDisplayName(
      standardCapitalization(expandElemSchoolName(m.school_name))
    );
  }

  /** Display name for map tooltips, dropdown, and sidebar (master CSV when present). */
  function schoolDisplayNameFromProps(p) {
    return (
      schoolDisplayNamePreferMaster(p) ||
      formatSchoolDisplayName(
        standardCapitalization(
          expandElemSchoolName(p.NAME || p.CommonName || "School")
        )
      )
    );
  }

  /** Shown first in the school dropdown; order is Johnson, McNair, Stone. MSIDs match SCHOOLS_ID in GeoJSON. */
  var PRIORITY_SCHOOL_MSIDS = [3031, 1081, 2071];

  /**
   * Preferred short names for named middle schools (avoids e.g. "Lyndon B. Johnson", "Ronald McNair" in UI).
   * Used for scenario travel chart titles, student-hex tooltips, ESE abbreviations, and capture KPI.
   */
  var SCENARIO_MIDDLE_SHORT_NAME = {
    3031: "Johnson MS",
    1081: "McNair MS",
    2071: "Stone MS",
  };

  /**
   * Short labels for ESE feeder table only: ES / MS / HS / Jr/Sr HS from school_master school_level + name.
   */
  function eseTableAbbreviatedSchoolName(m) {
    if (!m || !m.school_name) return "";
    var lv0 = String(m.school_level || "").toLowerCase().trim();
    if (lv0 === "middle" && m.msid != null) {
      var midNum = parseInt(String(m.msid), 10);
      if (!isNaN(midNum)) {
        var shortMid = SCENARIO_MIDDLE_SHORT_NAME[midNum];
        if (shortMid) {
          return shortMid;
        }
      }
    }
    var full = formatSchoolDisplayName(
      standardCapitalization(expandElemSchoolName(m.school_name))
    );
    var lv = String(m.school_level || "").toLowerCase().trim();
    if (!lv) return full;

    var base = full;

    if (lv === "elementary") {
      base = full
        .replace(/\s+Elementary\s+School\s+Of\s+International\s+Studies$/i, "")
        .replace(/\s+Elementary\s+Magnet\s+School$/i, "")
        .replace(/\s+Elementary\s+School$/i, "")
        .trim();
      return base ? base + " ES" : full;
    }
    if (lv === "middle") {
      base = full.replace(/\s+Magnet\s+Middle\s+School$/i, "").trim();
      base = base.replace(/\s+Middle\s+School$/i, "").trim();
      return base ? base + " MS" : full;
    }
    if (lv === "high") {
      base = full.replace(/\s+Magnet\s+Senior\s+High\s+School$/i, "").trim();
      base = base.replace(/\s+Senior\s+High\s+School$/i, "").trim();
      base = base.replace(/\s+High\s+School$/i, "").trim();
      return base ? base + " HS" : full;
    }
    if (lv === "jr_sr_high") {
      base = full.replace(/\s+Jr\s*\/\s*Sr\s+High\s+School$/i, "").trim();
      base = base.replace(/\s+Jr\.?\s*\/?\s*Sr\.?\s+High\s+School$/i, "").trim();
      base = base.replace(/\s+Magnet\s+Senior\s+High\s+School$/i, "").trim();
      base = base.replace(/\s+Senior\s+High\s+School$/i, "").trim();
      base = base.replace(/\s+High\s+School$/i, "").trim();
      return base ? base + " Jr/Sr HS" : full;
    }

    return full;
  }

  /**
   * Abbreviated school name for capture KPI row 1 (e.g. SHERWOOD ES), uppercase when a school is selected.
   * When nothing is selected, returns "Selected School" (displayed uppercase via .kpi-capture-card-label).
   */
  function captureRateAssignedSchoolLabelUpper(p) {
    if (!p || p.SCHOOLS_ID == null || p.SCHOOLS_ID === "") return "Selected School";
    var sid = Number(p.SCHOOLS_ID);
    if (isNaN(sid)) return "Selected School";
    var m = masterRow(sid);
    var s = "";
    if (m && m.school_name) {
      s = eseTableAbbreviatedSchoolName(m);
    }
    if (!s) {
      s = schoolDisplayNameFromProps(p) || "";
    }
    s = String(s).trim();
    return s ? s.toUpperCase() : "Selected School";
  }

  /** Display name from district MSID only (for feeder tables when GeoJSON props are unavailable). */
  function eseSchoolNameFromMsid(msidRaw) {
    var n = Number(msidRaw);
    if (isNaN(n)) return String(msidRaw);
    var m = masterRow(n);
    if (!m || !m.school_name) {
      return "MSID " + String(n);
    }
    return eseTableAbbreviatedSchoolName(m);
  }

  /**
   * Convert feeder MSID lists to sorted display names; drops the selected school's MSID (no self-loops).
   */
  function eseFilteredSortedSchoolNames(msidStrings, excludeMsid) {
    var ex = Number(excludeMsid);
    var seen = {};
    var pairs = [];
    for (var i = 0; i < (msidStrings || []).length; i++) {
      var raw = msidStrings[i];
      var n = Number(raw);
      if (isNaN(n)) continue;
      if (n === ex) continue;
      var idStr = String(n);
      if (seen[idStr]) continue;
      seen[idStr] = true;
      pairs.push({ id: idStr, name: eseSchoolNameFromMsid(raw) });
    }
    pairs.sort(function (a, b) {
      return a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
    });
    return pairs.map(function (p) {
      return p.name;
    });
  }

  /** Pre-K through Grade 12 enrollment columns on private-school GeoJSON features. */
  var PRIVATE_SCHOOL_GRADE_KEYS = [
    { key: "Pre_K", ord: -2 },
    { key: "Kindergart", ord: -1 },
    { key: "Grade_1", ord: 1 },
    { key: "Grade_2", ord: 2 },
    { key: "Grade_3", ord: 3 },
    { key: "Grade_4", ord: 4 },
    { key: "Grade_5", ord: 5 },
    { key: "Grade_6", ord: 6 },
    { key: "Grade_7", ord: 7 },
    { key: "Grade_8", ord: 8 },
    { key: "Grade_9", ord: 9 },
    { key: "Grade_10", ord: 10 },
    { key: "Grade_11", ord: 11 },
    { key: "Grade_12", ord: 12 },
  ];

  function privateSchoolGradeOrdinalLabel(ord) {
    if (ord === -2) return "Pre-K";
    if (ord === -1) return "K";
    return String(ord);
  }

  /**
   * Total enrollment (sum of grade columns) and display span from min–max grade with ≥1 student.
   * @returns {{ total: number, gradesLabel: string }}
   */
  function privateSchoolEnrollmentGradeSpan(props) {
    var total = 0;
    var minO = Infinity;
    var maxO = -Infinity;
    if (!props) {
      return { total: 0, gradesLabel: "" };
    }
    for (var i = 0; i < PRIVATE_SCHOOL_GRADE_KEYS.length; i++) {
      var g = PRIVATE_SCHOOL_GRADE_KEYS[i];
      var n = Number(props[g.key]);
      if (isNaN(n)) n = 0;
      total += n;
      if (n > 0) {
        if (g.ord < minO) minO = g.ord;
        if (g.ord > maxO) maxO = g.ord;
      }
    }
    if (!isFinite(minO)) {
      return { total: total, gradesLabel: "" };
    }
    var a = privateSchoolGradeOrdinalLabel(minO);
    var b = privateSchoolGradeOrdinalLabel(maxO);
    var gradesLabel = minO === maxO ? a : a + "–" + b;
    return { total: total, gradesLabel: gradesLabel };
  }

  /** Drop private-school points with no enrollment in any grade column. */
  function filterZeroEnrollmentPrivateSchoolsFc(fc) {
    if (!fc || fc.type !== "FeatureCollection" || !fc.features) {
      return { type: "FeatureCollection", features: [] };
    }
    var kept = [];
    for (var i = 0; i < fc.features.length; i++) {
      var eg = privateSchoolEnrollmentGradeSpan(fc.features[i].properties);
      if (eg.total > 0) kept.push(fc.features[i]);
    }
    return { type: "FeatureCollection", features: kept };
  }

  function formatPrivateSchoolZipFive(zipRaw) {
    if (zipRaw == null) return "";
    var d = String(zipRaw).replace(/\D/g, "");
    return d.length >= 5 ? d.slice(0, 5) : d;
  }

  /** Street, City, FL ZIP (ZIP trimmed to five digits). */
  function privateSchoolAddressLine(props) {
    if (!props) return "";
    var streetRaw = props.Address_1 != null ? String(props.Address_1).trim() : "";
    var cityRaw = props.City != null ? String(props.City).trim() : "";
    var zip5 = formatPrivateSchoolZipFive(props.Zip);
    var street = streetRaw ? standardCapitalization(streetRaw) : "";
    var city = cityRaw ? standardCapitalization(expandWestMelbourneCity(cityRaw)) : "";
    var parts = [];
    if (street) parts.push(street);
    if (city) parts.push(city);
    var head = parts.join(", ");
    if (!head) return zip5 ? "FL " + zip5 : "";
    return head + ", FL" + (zip5 ? " " + zip5 : "");
  }

  function privateSchoolDetailHtml(p) {
    var rawName = p && p.School_Nam != null ? String(p.School_Nam) : "";
    var name = formatSchoolDisplayName(
      standardCapitalization(expandElemSchoolName(rawName))
    );
    var eg = privateSchoolEnrollmentGradeSpan(p);
    var parts = [
      '<strong class="popup-school-name">' + escapeHtml(name) + "</strong>",
      '<div class="popup-detail">Grades Served: ' +
        escapeHtml(eg.gradesLabel || "—") +
        "</div>",
      '<div class="popup-detail">Total Enrollment: ' +
        escapeHtml(String(eg.total.toLocaleString())) +
        "</div>",
    ];
    var addr = privateSchoolAddressLine(p);
    if (addr) {
      parts.push('<div class="popup-detail">' + escapeHtml(addr) + "</div>");
    }
    return parts.join("");
  }

  function schoolDetailHtml(p) {
    var name = schoolDisplayNameFromProps(p);
    var sid = p.SCHOOLS_ID != null ? Number(p.SCHOOLS_ID) : NaN;
    var mRow = !isNaN(sid) ? masterRow(sid) : null;
    var grades = normalizeGradesServedForUi(
      (mRow && mRow.grades_served) || p.Grades || ""
    );
    var addr = p.ADDRESS || "";
    var city = p.CITY_ST_ZI || "";
    var parts = [
      '<strong class="popup-school-name">' + escapeHtml(name) + "</strong>",
    ];
    if (grades) {
      parts.push(
        '<div class="popup-detail">Grades: ' +
          escapeHtml(standardCapitalization(grades)) +
          "</div>"
      );
    }
    if (addr) {
      parts.push(
        '<div class="popup-detail">' +
          escapeHtml(standardCapitalization(addr)) +
          "</div>"
      );
    }
    if (city) {
      parts.push(
        '<div class="popup-detail">' +
          escapeHtml(formatCityStateZip(city)) +
          "</div>"
      );
    }
    return parts.join("");
  }

  function scenarioMiddleShortDisplayName(msid) {
    if (msid == null || isNaN(msid)) return null;
    var sh = SCENARIO_MIDDLE_SHORT_NAME[msid];
    return sh != null ? sh : null;
  }

  function schoolNameForSelect(p) {
    return schoolDisplayNameFromProps(p);
  }

  /** Fills #school-select; option values are SCHOOLS_ID (district MSID). */
  function populateSchoolSelect(schoolsFc) {
    var sel = document.getElementById("school-select");
    if (!sel || !schoolsFc || !schoolsFc.features) return;

    sel.innerHTML = "";

    var placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Select a school";
    sel.appendChild(placeholder);

    var byId = {};
    schoolsFc.features.forEach(function (ft) {
      var p = ft.properties;
      if (p && p.SCHOOLS_ID != null) byId[p.SCHOOLS_ID] = p;
    });

    var priorityUsed = {};
    PRIORITY_SCHOOL_MSIDS.forEach(function (msid) {
      var p = byId[msid];
      if (!p) return;
      priorityUsed[msid] = true;
      var opt = document.createElement("option");
      opt.value = String(msid);
      opt.textContent = schoolNameForSelect(p);
      sel.appendChild(opt);
    });

    var sep = document.createElement("option");
    sep.disabled = true;
    sep.value = "";
    sep.setAttribute("aria-hidden", "true");
    sep.textContent = "────────────────────────";
    sel.appendChild(sep);

    var rest = schoolsFc.features
      .map(function (ft) {
        return ft.properties;
      })
      .filter(function (p) {
        if (!p || p.SCHOOLS_ID == null) return false;
        return !priorityUsed[p.SCHOOLS_ID];
      })
      .sort(function (a, b) {
        var na = schoolDisplayNameFromProps(a).toLowerCase();
        var nb = schoolDisplayNameFromProps(b).toLowerCase();
        if (na < nb) return -1;
        if (na > nb) return 1;
        return 0;
      });

    rest.forEach(function (p) {
      var opt = document.createElement("option");
      opt.value = String(p.SCHOOLS_ID);
      opt.textContent = schoolNameForSelect(p);
      sel.appendChild(opt);
    });

    sel.value = "";
    sel.disabled = false;

    var sbox = document.getElementById("sandbox-base-school");
    if (sbox) {
      sbox.innerHTML = sel.innerHTML;
      if (sbox.options[0]) {
        sbox.options[0].textContent = "Start from school (optional)…";
      }
      sbox.value = "";
      sbox.disabled = false;
    }
  }

  function findBoundaryFeatureForMsid(msid) {
    var layers = [GEO_CACHE.es, GEO_CACHE.ms, GEO_CACHE.hs];
    for (var i = 0; i < layers.length; i++) {
      var fc = layers[i];
      if (!fc || !fc.features) continue;
      for (var j = 0; j < fc.features.length; j++) {
        var f = fc.features[j];
        var m =
          f.properties && f.properties.MSID != null
            ? Number(f.properties.MSID)
            : null;
        if (m === msid) return f;
      }
    }
    return null;
  }

  /** Map source id (e.g. "es-boundaries") for the assignment polygon containing this MSID, or null. */
  function findBoundarySourceForMsid(msid) {
    var layers = [
      { fc: GEO_CACHE.es, src: "es-boundaries" },
      { fc: GEO_CACHE.ms, src: "ms-boundaries" },
      { fc: GEO_CACHE.hs, src: "hs-boundaries" },
    ];
    for (var i = 0; i < layers.length; i++) {
      var fc = layers[i].fc;
      if (!fc || !fc.features) continue;
      for (var j = 0; j < fc.features.length; j++) {
        var f = fc.features[j];
        var m =
          f.properties && f.properties.MSID != null
            ? Number(f.properties.MSID)
            : null;
        if (m === msid) return layers[i].src;
      }
    }
    return null;
  }

  /**
   * Assignment MSIDs from elementary / middle / high boundary layers at a residence point
   * (same attendance-area polygons as capture KPIs and `countHomeschoolStudentsInAssignmentBoundary`).
   * @returns {{ elem: number|null, mid: number|null, high: number|null }}
   */
  function attendanceZoningTripletAtLngLat(lng, lat) {
    var out = { elem: null, mid: null, high: null };
    if (
      typeof turf === "undefined" ||
      !turf ||
      typeof turf.point !== "function" ||
      typeof turf.feature !== "function" ||
      typeof turf.booleanPointInPolygon !== "function"
    ) {
      return out;
    }
    var pt;
    try {
      pt = turf.point([lng, lat]);
    } catch (ePt) {
      return out;
    }
    function msidFromBoundaryFc(fc) {
      if (!fc || !fc.features) {
        return null;
      }
      for (var i = 0; i < fc.features.length; i++) {
        var f = fc.features[i];
        if (!f || !f.geometry) {
          continue;
        }
        try {
          var poly = turf.feature(f.geometry);
          if (turf.booleanPointInPolygon(pt, poly)) {
            var m =
              f.properties && f.properties.MSID != null ? Number(f.properties.MSID) : NaN;
            if (!isNaN(m) && m > 0) {
              return Math.round(m);
            }
          }
        } catch (eIn) {
          /* ignore */
        }
      }
      return null;
    }
    out.elem = msidFromBoundaryFc(GEO_CACHE.es);
    out.mid = msidFromBoundaryFc(GEO_CACHE.ms);
    out.high = msidFromBoundaryFc(GEO_CACHE.hs);
    return out;
  }

  /**
   * Zoning triplet for a homeschool hex: centroid of hex geometry vs es/ms/hs assignment layers.
   */
  function homeschoolAttendanceZoningTripletForHex(hexKey, feature) {
    var out = { elem: null, mid: null, high: null };
    var geom =
      feature &&
      feature.geometry &&
      (feature.geometry.type === "Polygon" || feature.geometry.type === "MultiPolygon")
        ? feature.geometry
        : homeschoolHexGeometry(hexKey);
    if (!geom) {
      return out;
    }
    var ctr = polygonCentroid(geom);
    if (!ctr || ctr.length < 2) {
      return out;
    }
    return attendanceZoningTripletAtLngLat(ctr[0], ctr[1]);
  }

  function boundaryFillVisibleForSource(src) {
    var fillId =
      src === "es-boundaries"
        ? "es-fill"
        : src === "ms-boundaries"
          ? "ms-fill"
          : src === "hs-boundaries"
            ? "hs-fill"
            : null;
    if (!fillId) return false;
    try {
      return map.getLayoutProperty(fillId, "visibility") !== "none";
    } catch (e) {
      return false;
    }
  }

  function clearSelectedAssignmentBoundary() {
    if (selectedAssignmentBoundary != null) {
      try {
        map.setFeatureState(
          {
            source: selectedAssignmentBoundary.source,
            id: selectedAssignmentBoundary.id,
          },
          { selectedAssignment: false }
        );
      } catch (e) {
        /* ignore */
      }
      selectedAssignmentBoundary = null;
    }
  }

  function applySelectedAssignmentBoundary(msid) {
    clearSelectedAssignmentBoundary();
    if (msid == null || isNaN(msid)) return;
    var src = findBoundarySourceForMsid(msid);
    if (!src) return;
    if (!boundaryFillVisibleForSource(src)) return;
    selectedAssignmentBoundary = { source: src, id: msid };
    try {
      map.setFeatureState({ source: src, id: msid }, { selectedAssignment: true });
    } catch (e) {
      /* ignore */
    }
  }

  function refreshAssignmentBoundaryHighlight() {
    if (selectedSchoolMsid == null) return;
    applySelectedAssignmentBoundary(selectedSchoolMsid);
  }

  function schoolPointLonLatForMsid(msid, schoolByMsid) {
    var p = schoolByMsid[msid];
    var lon;
    var lat;
    if (p && p.Longitude != null && p.Latitude != null) {
      lon = Number(p.Longitude);
      lat = Number(p.Latitude);
    } else if (GEO_CACHE.schools && GEO_CACHE.schools.features) {
      for (var i = 0; i < GEO_CACHE.schools.features.length; i++) {
        var ft = GEO_CACHE.schools.features[i];
        if (
          ft.properties &&
          Number(ft.properties.SCHOOLS_ID) === msid &&
          ft.geometry &&
          ft.geometry.coordinates
        ) {
          lon = ft.geometry.coordinates[0];
          lat = ft.geometry.coordinates[1];
          break;
        }
      }
    }
    if (lon == null || lat == null || isNaN(lon) || isNaN(lat)) {
      return null;
    }
    return [lon, lat];
  }

  /** Pans the map so the school location is centered; zoom level is unchanged. */
  function centerMapOnSchoolPoint(msid, schoolByMsid) {
    if (!map) return;
    var c = schoolPointLonLatForMsid(msid, schoolByMsid);
    if (!c) {
      zoomToSchoolAssignment(msid, schoolByMsid);
      return;
    }
    try {
      map.easeTo({
        center: c,
        duration: 750,
        essential: true,
      });
    } catch (e) {
      /* ignore */
    }
  }

  function zoomToSchoolAssignment(msid, schoolByMsid) {
    var boundaryFt = findBoundaryFeatureForMsid(msid);
    var bbox;
    if (boundaryFt) {
      bbox = computeBbox({
        type: "FeatureCollection",
        features: [boundaryFt],
      });
    } else {
      var c2 = schoolPointLonLatForMsid(msid, schoolByMsid);
      if (!c2) return;
      var lon = c2[0];
      var lat = c2[1];
      var d = 0.03;
      bbox = [lon - d, lat - d, lon + d, lat + d];
    }
    if (bbox) {
      map.fitBounds(bbox, { padding: 56, maxZoom: 15, duration: 750 });
    }
  }

  function clearSelectedSchoolHighlight() {
    clearSelectedAssignmentBoundary();
    if (selectedSchoolMsid != null) {
      try {
        map.setFeatureState(
          { source: "schools", id: selectedSchoolMsid },
          { selected: false }
        );
      } catch (e) {
        /* ignore */
      }
      selectedSchoolMsid = null;
    }
  }

  function applySelectedSchoolHighlight(msid) {
    clearSelectedSchoolHighlight();
    if (msid == null) return;
    selectedSchoolMsid = msid;
    try {
      map.setFeatureState({ source: "schools", id: msid }, { selected: true });
    } catch (e) {
      /* ignore */
    }
    applySelectedAssignmentBoundary(msid);
  }

  function isScenarioPlanningViewActive() {
    var panel = document.getElementById("page-scenario");
    return !!(panel && !panel.hidden);
  }

  function clearScenarioBoundaryRelevantFeatureStates() {
    if (!map) return;
    for (var i = 0; i < lastScenarioBoundaryRelevant.length; i++) {
      var b = lastScenarioBoundaryRelevant[i];
      if (!b || b.source == null || b.id == null) continue;
      try {
        map.setFeatureState(
          { source: b.source, id: b.id },
          { scenarioRelevant: false }
        );
      } catch (e) {
        /* ignore */
      }
    }
    lastScenarioBoundaryRelevant = [];
  }

  /**
   * - Scenario: fill for `highlight`, `selectedAssignment`, or `scenarioRelevant` (feeder + middle), else 0.
   * - Existing: fill only for `highlight` (hover) or `selectedAssignment` (dropdown), else 0.
   * - Boundary Sandbox: same as Existing (hover + selected assignment), not full-opacity on all zones.
   */
  function getAssignmentFillOpacityPaintValue() {
    if (isBoundarySandboxViewActive()) {
      return [
        "case",
        ["==", ["feature-state", "highlight"], true],
        BOUNDARY_FILL_OPACITY,
        ["==", ["feature-state", "selectedAssignment"], true],
        BOUNDARY_FILL_OPACITY,
        0,
      ];
    }
    if (isScenarioPlanningViewActive()) {
      return [
        "case",
        ["==", ["feature-state", "highlight"], true],
        BOUNDARY_FILL_OPACITY,
        ["==", ["feature-state", "selectedAssignment"], true],
        BOUNDARY_FILL_OPACITY,
        ["==", ["feature-state", "scenarioRelevant"], true],
        BOUNDARY_FILL_OPACITY,
        0,
      ];
    }
    return [
      "case",
      ["==", ["feature-state", "highlight"], true],
      BOUNDARY_FILL_OPACITY,
      ["==", ["feature-state", "selectedAssignment"], true],
      BOUNDARY_FILL_OPACITY,
      0,
    ];
  }

  function syncAssignmentFillPaintForView() {
    if (!map || !map.getLayer) {
      return;
    }
    var v = getAssignmentFillOpacityPaintValue();
    var lids = ["es-fill", "ms-fill", "hs-fill"];
    for (var l = 0; l < lids.length; l++) {
      if (!map.getLayer(lids[l])) continue;
      try {
        map.setPaintProperty(lids[l], "fill-opacity", v);
      } catch (e) {
        /* ignore */
      }
    }
  }

  /**
   * Marks the selected middle MS zone + checked feeder elementary zones. Other assignment fills stay at 0
   * in Scenario (see `getAssignmentFillOpacityPaintValue`) until hover `highlight` repopulates fill.
   */
  function applyScenarioBoundaryRelevantFeatureStates() {
    clearScenarioBoundaryRelevantFeatureStates();
    if (!map || !isScenarioPlanningViewActive()) {
      return;
    }
    if (
      scenarioMiddleMsid == null ||
      isNaN(scenarioMiddleMsid) ||
      !scenarioSchoolByMsid
    ) {
      return;
    }
    var sch = scenarioSchoolByMsid;
    var pushB = function (source, id) {
      if (id == null || isNaN(id)) return;
      lastScenarioBoundaryRelevant.push({ source: source, id: id });
      try {
        map.setFeatureState({ source: source, id: id }, { scenarioRelevant: true });
      } catch (e) {
        /* ignore */
      }
    };
    pushB("ms-boundaries", Number(scenarioMiddleMsid));
    for (var key in scenarioFeederChecked) {
      if (!Object.prototype.hasOwnProperty.call(scenarioFeederChecked, key)) {
        continue;
      }
      if (scenarioFeederChecked[key] === false) {
        continue;
      }
      var n = Number(key);
      if (isNaN(n)) continue;
      var p = sch[n];
      if (!p) continue;
      var t = (p.TYPE || "").toUpperCase();
      if (t.indexOf("ELEMENTARY") < 0) continue;
      pushB("es-boundaries", n);
    }
  }

  function applyScenarioFeederMapHighlights() {
    for (var i = 0; i < lastScenarioFeederHighlightMsids.length; i++) {
      try {
        map.setFeatureState(
          { source: "schools", id: lastScenarioFeederHighlightMsids[i] },
          { scenarioFeeder: false }
        );
      } catch (e) {
        /* ignore */
      }
    }
    lastScenarioFeederHighlightMsids = [];
    var panelScenario = document.getElementById("page-scenario");
    if (!panelScenario || panelScenario.hidden) {
      clearScenarioBoundaryRelevantFeatureStates();
      syncAssignmentFillPaintForView();
      return;
    }
    if (
      scenarioMiddleMsid == null ||
      isNaN(scenarioMiddleMsid) ||
      !scenarioSchoolByMsid
    ) {
      clearScenarioBoundaryRelevantFeatureStates();
      syncAssignmentFillPaintForView();
      return;
    }
    var sch = scenarioSchoolByMsid;
    for (var key in scenarioFeederChecked) {
      if (!Object.prototype.hasOwnProperty.call(scenarioFeederChecked, key)) {
        continue;
      }
      if (scenarioFeederChecked[key] === false) continue;
      var n = Number(key);
      if (isNaN(n)) continue;
      var p = sch[n];
      if (!p) continue;
      var t = (p.TYPE || "").toUpperCase();
      if (t.indexOf("ELEMENTARY") < 0) continue;
      try {
        map.setFeatureState(
          { source: "schools", id: n },
          { scenarioFeeder: true }
        );
        lastScenarioFeederHighlightMsids.push(n);
      } catch (e2) {
        /* ignore */
      }
    }
    applyScenarioBoundaryRelevantFeatureStates();
    syncAssignmentFillPaintForView();
  }

  function escapeXmlText(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  /** Aligns with export_facility_age_from_xls.ps1 Get-NameKey (source has no MSID). */
  function normalizeSchoolNameKey(str) {
    if (!str) return "";
    return String(str)
      .toUpperCase()
      .replace(/\//g, " ")
      .replace(/[.'’]/g, " ")
      .replace(/,/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function schoolPaletteKeyFromType(typeStr) {
    var t = (typeStr || "").toUpperCase();
    if (t.indexOf("ELEMENTARY") >= 0) return "elementary";
    if (t.indexOf("MIDDLE") >= 0) return "middle";
    if (t === "JR SR HIGH" || t.indexOf("HIGH") >= 0) return "high";
    return "middle";
  }

  function schoolTypeIsHigh(typeStr) {
    var t = (typeStr || "").toUpperCase();
    return t === "JR SR HIGH" || t.indexOf("HIGH") >= 0;
  }

  function schoolTypeIsElemOrMiddle(typeStr) {
    var t = (typeStr || "").toUpperCase();
    if (t.indexOf("ELEMENTARY") >= 0) return true;
    if (t.indexOf("MIDDLE") >= 0 && t.indexOf("HIGH") < 0) return true;
    return false;
  }

  /**
   * Match Sankey workbook labels (short names in the helper spreadsheet) to GeoJSON.
   * Uses compact CommonName equality and leading NAME tokens — not naive substring match —
   * so "Cocoa" does not match "Cocoa Beach" (both NAMEs start with COCOA).
   */
  function sankeyWorkbookLabelMatchesSchool(label, p) {
    var L = normalizeSchoolNameKey(label || "");
    if (!L) return false;

    var cn = normalizeSchoolNameKey(p.CommonName || "");
    var nm = normalizeSchoolNameKey(p.NAME || "");

    var Lc = L.replace(/\s+/g, "");
    var cnc = cn.replace(/\s+/g, "");
    if (cnc && Lc === cnc) return true;
    if (cn && L === cn) return true;

    var lTok = L.split(" ").filter(Boolean);
    var nmTok = nm.split(" ").filter(Boolean);
    if (!lTok.length || !nmTok.length || lTok.length > nmTok.length) return false;

    for (var ti = 0; ti < lTok.length; ti++) {
      if (lTok[ti] !== nmTok[ti]) return false;
    }

    if (lTok.length === 1 && nmTok.length >= 2) {
      if (lTok[0] === "COCOA" && nmTok[1] === "BEACH") return false;
    }

    return true;
  }

  /** Match Sankey workbook row/column labels (short names) to GeoJSON NAME/CommonName. */
  function sankeyElementaryLabelMatchesSchool(label, p) {
    var L = normalizeSchoolNameKey(label);
    var cn = normalizeSchoolNameKey(p.CommonName || "");
    var nm = normalizeSchoolNameKey(p.NAME || "");
    if (!L) return false;
    if (cn && L === cn) return true;
    if (nm.indexOf(L) !== -1) return true;
    var parts = nm.split(" ").filter(Boolean);
    if (parts.length && L === parts[0]) return true;
    var LnoElem = L.replace(/\s+ELEM$/, "");
    if (LnoElem.length >= 3 && nm.indexOf(LnoElem) !== -1) return true;
    return false;
  }

  function sankeyMiddleLabelMatchesSchool(label, p) {
    return sankeyWorkbookLabelMatchesSchool(label, p);
  }

  /**
   * @returns {{ elementary: string, middle: string, value: number, emphasis: boolean }[]}
   *   emphasis = flow is a primary focus for the selection (all ES→MS links when ES selected;
   *   ES→selected-MS when middle selected; other MS destinations from same feeders are emphasis:false).
   *   For Jr/Sr (7–12): ES→MS flows that share feeder elementaries with middle schools that feed this high
   *   (grades 6→7 transition); emphasis on links into those feeder middles.
   */
  function filterEsMsFlowsForSchool(flows, p) {
    if (!flows || !flows.length || !p) return [];
    var t = (p.TYPE || "").toUpperCase();
    if (t.indexOf("ELEMENTARY") >= 0) {
      return flows
        .filter(function (f) {
          return sankeyElementaryLabelMatchesSchool(f.elementary, p);
        })
        .map(function (f) {
          return {
            elementary: f.elementary,
            middle: f.middle,
            value: f.value,
            emphasis: true,
          };
        });
    }
    if (t.indexOf("MIDDLE") >= 0 && t.indexOf("HIGH") < 0) {
      var intoSelected = flows.filter(function (f) {
        return sankeyMiddleLabelMatchesSchool(f.middle, p);
      });
      if (!intoSelected.length) return [];
      var feederEs = {};
      intoSelected.forEach(function (f) {
        feederEs[f.elementary] = true;
      });
      return flows
        .filter(function (f) {
          return feederEs[f.elementary];
        })
        .map(function (f) {
          return {
            elementary: f.elementary,
            middle: f.middle,
            value: f.value,
            emphasis: sankeyMiddleLabelMatchesSchool(f.middle, p),
          };
        });
    }
    if (t === "JR SR HIGH" && SANKEY_CACHE && SANKEY_CACHE.msHsFlows) {
      var msHs = SANKEY_CACHE.msHsFlows;
      var intoJrSr = msHs.filter(function (hf) {
        return sankeyHighLabelMatchesSchool(hf.high, p);
      });
      if (!intoJrSr.length) return [];
      var feederMiddles = {};
      intoJrSr.forEach(function (hf) {
        feederMiddles[hf.middle] = true;
      });
      var feederEs = {};
      flows.forEach(function (f) {
        if (feederMiddles[f.middle]) feederEs[f.elementary] = true;
      });
      return flows
        .filter(function (f) {
          return feederEs[f.elementary];
        })
        .map(function (f) {
          return {
            elementary: f.elementary,
            middle: f.middle,
            value: f.value,
            emphasis: !!feederMiddles[f.middle],
          };
        });
    }
    return [];
  }

  function sankeyHighLabelMatchesSchool(label, p) {
    return sankeyWorkbookLabelMatchesSchool(label, p);
  }

  /**
   * @returns {{ middle: string, high: string, value: number, emphasis: boolean }[]}
   */
  function filterMsHsFlowsForSchool(flows, p) {
    if (!flows || !flows.length || !p) return [];
    var t = (p.TYPE || "").toUpperCase();
    if (t.indexOf("ELEMENTARY") >= 0) return [];
    if (t.indexOf("MIDDLE") >= 0 && t.indexOf("HIGH") < 0) {
      return flows
        .filter(function (f) {
          return sankeyMiddleLabelMatchesSchool(f.middle, p);
        })
        .map(function (f) {
          return {
            middle: f.middle,
            high: f.high,
            value: f.value,
            emphasis: true,
          };
        });
    }
    if (t === "JR SR HIGH" || t.indexOf("HIGH") >= 0) {
      return flows
        .filter(function (f) {
          return sankeyHighLabelMatchesSchool(f.high, p);
        })
        .map(function (f) {
          return {
            middle: f.middle,
            high: f.high,
            value: f.value,
            emphasis: true,
          };
        });
    }
    return [];
  }

  function findSchoolPropsForSankeyWorkbookLabel(label, schoolsFc, useMiddleMatcher) {
    if (!label || !schoolsFc || !schoolsFc.features) return null;
    var matchFn = useMiddleMatcher
      ? sankeyMiddleLabelMatchesSchool
      : sankeyHighLabelMatchesSchool;
    for (var i = 0; i < schoolsFc.features.length; i++) {
      var p = schoolsFc.features[i].properties;
      if (matchFn(label, p)) return p;
    }
    return null;
  }

  /** Middle → high Sankey: 7–12 / Jr–Sr nodes use orange; 9–12 high and middle use blue/purple. */
  function msHsSankeyNodeFill(name, isLeft) {
    var fc = GEO_CACHE.schools;
    var p = findSchoolPropsForSankeyWorkbookLabel(name, fc, isLeft);
    p = schoolPropsWithMasterType(p);
    if (p && (p.TYPE || "").toUpperCase() === "JR SR HIGH") {
      return PALETTE.jrSr.fill;
    }
    return isLeft ? PALETTE.middle.fill : PALETTE.high.fill;
  }

  /**
   * @param {{ from: string, to: string, value: number, emphasis: boolean }[]} normFlows
   * @param {{ leftFill: string, rightFill: string, emphStroke: string, ariaLabel: string, secondaryTooltip: string, leftNodeFill?: function(string): string, rightNodeFill?: function(string): string }} cfg
   */
  function renderBipartiteSankey(root, normFlows, cfg) {
    if (!normFlows.length) {
      root.innerHTML =
        '<p class="sankey-empty">No matching matriculation flows for this school selection.</p>';
      return;
    }
    if (typeof d3 === "undefined" || !d3.sankey || !d3.sankeyLinkHorizontal) {
      root.innerHTML =
        '<p class="sankey-empty">Sankey layout library failed to load.</p>';
      return;
    }

    /* Tighter horizontal flow band + side padding keeps viewBox width modest so SVG scales up larger in the sidebar. */
    var padL = 138;
    var padR = 138;
    var padY = 12;
    var cw = root.clientWidth || 400;
    var graphW = Math.max(96, Math.min(268, cw - 4));
    var totalW = padL + graphW + padR;

    var leftSet = {};
    var rightSet = {};
    normFlows.forEach(function (f) {
      leftSet[f.from] = true;
      rightSet[f.to] = true;
    });
    var leftList = Object.keys(leftSet);
    var rightList = Object.keys(rightSet);
    var h = Math.max(
      320,
      Math.min(580, leftList.length * 40 + rightList.length * 48 + 110)
    );
    var nodes = leftList
      .map(function (name) {
        return { name: name };
      })
      .concat(
        rightList.map(function (name) {
          return { name: name };
        })
      );
    var indexByLeft = {};
    var indexByRight = {};
    leftList.forEach(function (n, i) {
      indexByLeft[n] = i;
    });
    rightList.forEach(function (n, i) {
      indexByRight[n] = i + leftList.length;
    });
    var originTotal = {};
    var destTotal = {};
    normFlows.forEach(function (f) {
      originTotal[f.from] = (originTotal[f.from] || 0) + f.value;
      destTotal[f.to] = (destTotal[f.to] || 0) + f.value;
    });

    var emphasisByPair = {};
    normFlows.forEach(function (f) {
      emphasisByPair[f.from + "\u0000" + f.to] = f.emphasis !== false;
    });
    var links = normFlows.map(function (f) {
      return {
        source: indexByLeft[f.from],
        target: indexByRight[f.to],
        value: f.value,
      };
    });

    var sankeyLayout = d3
      .sankey()
      .nodeWidth(10)
      .nodePadding(8)
      .extent([
        [padL + 6, padY],
        [padL + graphW - 6, h - padY],
      ]);

    var graph = sankeyLayout({
      nodes: nodes.map(function (d) {
        return Object.assign({}, d);
      }),
      links: links.map(function (d) {
        return Object.assign({}, d);
      }),
    });

    var linkPath = d3.sankeyLinkHorizontal();
    var svgNs = "http://www.w3.org/2000/svg";
    var svg = document.createElementNS(svgNs, "svg");
    svg.setAttribute("viewBox", "0 0 " + totalW + " " + h);
    svg.setAttribute("width", "100%");
    svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
    svg.setAttribute("class", "sankey-svg");
    svg.setAttribute("role", "img");
    svg.setAttribute("aria-label", cfg.ariaLabel);

    var gLinks = document.createElementNS(svgNs, "g");
    gLinks.setAttribute("fill", "none");
    graph.links.forEach(function (d) {
      var path = document.createElementNS(svgNs, "path");
      path.setAttribute("d", linkPath(d));
      var srcN =
        d.source && d.source.name != null
          ? String(d.source.name)
          : "";
      var tgtN =
        d.target && d.target.name != null
          ? String(d.target.name)
          : "";
      var emph =
        emphasisByPair[srcN + "\u0000" + tgtN] !== false;
      path.setAttribute("stroke", emph ? cfg.emphStroke : "#94a3b8");
      path.setAttribute("stroke-opacity", emph ? "0.55" : "0.32");
      path.setAttribute(
        "class",
        "sankey-link" +
          (emph ? " sankey-link--emphasis" : " sankey-link--secondary")
      );
      var sw = d.width != null && !isNaN(Number(d.width)) ? Number(d.width) : 2;
      path.setAttribute("stroke-width", Math.max(1, sw));
      path.setAttribute("pointer-events", "stroke");
      var nv =
        d.value != null && !isNaN(Number(d.value)) ? Number(d.value) : 0;
      var tip = document.createElementNS(svgNs, "title");
      var line =
        srcN +
        " → " +
        tgtN +
        ": " +
        nv.toLocaleString() +
        " students";
      if (!emph && cfg.secondaryTooltip) {
        line += " " + cfg.secondaryTooltip;
      }
      tip.textContent = line;
      path.appendChild(tip);
      gLinks.appendChild(path);
    });
    svg.appendChild(gLinks);

    function truncLabel(s, maxLen) {
      if (!s) return "";
      if (s.length <= maxLen) return s;
      return s.slice(0, maxLen - 1) + "\u2026";
    }

    var gNodes = document.createElementNS(svgNs, "g");
    graph.nodes.forEach(function (d, i) {
      var nm = String(d.name);
      var rect = document.createElementNS(svgNs, "rect");
      rect.setAttribute("x", d.x0);
      rect.setAttribute("y", d.y0);
      rect.setAttribute("width", Math.max(1, d.x1 - d.x0));
      rect.setAttribute("height", Math.max(1, d.y1 - d.y0));
      var isLeft = i < leftList.length;
      var nodeFill;
      if (isLeft) {
        nodeFill =
          typeof cfg.leftNodeFill === "function"
            ? cfg.leftNodeFill(nm)
            : cfg.leftFill;
      } else {
        nodeFill =
          typeof cfg.rightNodeFill === "function"
            ? cfg.rightNodeFill(nm)
            : cfg.rightFill;
      }
      rect.setAttribute("fill", nodeFill);
      rect.setAttribute("rx", "2");
      rect.setAttribute("class", "sankey-node");
      gNodes.appendChild(rect);
      var tot = isLeft ? originTotal[nm] : destTotal[nm];
      var totStr =
        tot != null && !isNaN(Number(tot))
          ? Number(tot).toLocaleString()
          : "";
      var text = document.createElementNS(svgNs, "text");
      var tx = isLeft ? d.x0 - 8 : d.x1 + 8;
      var cy = (d.y0 + d.y1) / 2;
      text.setAttribute("x", tx);
      text.setAttribute("y", cy);
      text.setAttribute("class", "sankey-label");
      text.setAttribute("text-anchor", isLeft ? "end" : "start");
      var nameLine = truncLabel(nm, 22);
      var tName = document.createElementNS(svgNs, "tspan");
      tName.setAttribute("class", "sankey-label-name");
      tName.setAttribute("x", tx);
      tName.setAttribute("dy", "-0.5em");
      tName.textContent = nameLine;
      text.appendChild(tName);
      if (totStr) {
        var tTot = document.createElementNS(svgNs, "tspan");
        tTot.setAttribute("class", "sankey-label-total");
        tTot.setAttribute("x", tx);
        tTot.setAttribute("dy", "1.22em");
        tTot.textContent =
          (isLeft ? "Out: " : "In: ") + totStr;
        text.appendChild(tTot);
      }
      var tipFull = document.createElementNS(svgNs, "title");
      tipFull.textContent =
        nm +
        (totStr
          ? " — " + (isLeft ? "origin total" : "destination total") + ": " + totStr
          : "");
      text.appendChild(tipFull);
      gNodes.appendChild(text);
    });
    svg.appendChild(gNodes);

    root.innerHTML = "";
    root.appendChild(svg);
  }

  function renderEsMsChart(el, p) {
    if (!SANKEY_CACHE || !SANKEY_CACHE.flows) {
      el.innerHTML =
        '<p class="sankey-empty">Feeder flow data is not loaded.</p>';
      return;
    }
    var typeU = (p.TYPE || "").toUpperCase();
    if (!schoolTypeIsElemOrMiddle(p.TYPE) && typeU !== "JR SR HIGH") {
      el.innerHTML =
        '<p class="sankey-empty">No elementary–middle matrix for this school type.</p>';
      return;
    }
    var flows = filterEsMsFlowsForSchool(SANKEY_CACHE.flows, p);
    var norm = flows.map(function (f) {
      return {
        from: f.elementary,
        to: f.middle,
        value: f.value,
        emphasis: f.emphasis !== false,
      };
    });
    var selectedIsElem = typeU.indexOf("ELEMENTARY") >= 0;
    var emphStroke = selectedIsElem
      ? PALETTE.elementary.fill
      : PALETTE.middle.fill;
    renderBipartiteSankey(el, norm, {
      leftFill: PALETTE.elementary.fill,
      rightFill: PALETTE.middle.fill,
      emphStroke: emphStroke,
      ariaLabel:
        "Sankey diagram of student flows from elementary schools to middle schools",
      secondaryTooltip: "(other middle school destination)",
    });
  }

  function renderMsHsChart(el, p) {
    if (!SANKEY_CACHE || !SANKEY_CACHE.msHsFlows) {
      el.innerHTML =
        '<p class="sankey-empty">Middle–high feeder data is not loaded.</p>';
      return;
    }
    var t = (p.TYPE || "").toUpperCase();
    if (t.indexOf("ELEMENTARY") >= 0) {
      return;
    }
    var flows = filterMsHsFlowsForSchool(SANKEY_CACHE.msHsFlows, p);
    var norm = flows.map(function (f) {
      return {
        from: f.middle,
        to: f.high,
        value: f.value,
        emphasis: f.emphasis !== false,
      };
    });
    renderBipartiteSankey(el, norm, {
      leftFill: PALETTE.middle.fill,
      rightFill: PALETTE.high.fill,
      leftNodeFill: function (name) {
        return msHsSankeyNodeFill(name, true);
      },
      rightNodeFill: function (name) {
        return msHsSankeyNodeFill(name, false);
      },
      emphStroke: PALETTE.middle.fill,
      ariaLabel:
        "Sankey diagram of student flows from middle schools to high schools (grades 8 to 9 transition)",
      secondaryTooltip: "",
    });
  }

  function renderSankeyPanel(p) {
    var row = document.getElementById("sankey-row");
    var panel = document.getElementById("sankey-panel");
    var elEs = document.getElementById("sankey-es-ms");
    var elHs = document.getElementById("sankey-ms-hs");
    if (!elEs || !elHs || !row) return;

    function setSankeySplitLayout(isSplit) {
      if (panel) {
        if (isSplit) {
          panel.classList.add("sankey-panel--split");
        } else {
          panel.classList.remove("sankey-panel--split");
        }
      }
    }

    if (!SANKEY_CACHE) {
      var msg = '<p class="sankey-empty">Feeder flow data is not loaded.</p>';
      elEs.innerHTML = msg;
      elHs.innerHTML = msg;
      row.className = "sankey-row";
      setSankeySplitLayout(false);
      return;
    }

    if (!p) {
      elEs.innerHTML =
        '<p class="sankey-empty">Select a school to view feeder flows.</p>';
      elHs.innerHTML =
        '<p class="sankey-empty">Select a school to view feeder flows.</p>';
      row.className = "sankey-row";
      setSankeySplitLayout(false);
      return;
    }

    var t = (p.TYPE || "").toUpperCase();
    var isElem = t.indexOf("ELEMENTARY") >= 0;
    var isMid = t.indexOf("MIDDLE") >= 0 && t.indexOf("HIGH") < 0;
    var isHigh = schoolTypeIsHigh(p.TYPE);
    var isJrSr = t === "JR SR HIGH";

    row.className = "sankey-row";
    if (isMid || (isHigh && isJrSr)) {
      row.classList.add("sankey-row--split");
    } else if (isElem) {
      row.classList.add("sankey-row--es-only");
    } else if (isHigh) {
      row.classList.add("sankey-row--hs-only");
    }

    if (isElem) {
      setSankeySplitLayout(false);
      renderEsMsChart(elEs, p);
      elHs.innerHTML =
        '<p class="sankey-empty sankey-empty--muted">Middle → high transitions are not shown when an elementary school is selected.</p>';
    } else if (isHigh && isJrSr) {
      setSankeySplitLayout(true);
      renderEsMsChart(elEs, p);
      renderMsHsChart(elHs, p);
    } else if (isHigh) {
      setSankeySplitLayout(false);
      elEs.innerHTML =
        '<p class="sankey-empty sankey-empty--muted">Elementary → middle transitions are not shown when a high school is selected.</p>';
      renderMsHsChart(elHs, p);
    } else if (isMid) {
      setSankeySplitLayout(true);
      renderEsMsChart(elEs, p);
      renderMsHsChart(elHs, p);
    } else {
      elEs.innerHTML =
        '<p class="sankey-empty">No feeder matrix for this school type.</p>';
      elHs.innerHTML = "";
      row.className = "sankey-row";
      setSankeySplitLayout(false);
    }
  }

  /** Excel column year Y → school year label Y-(Y+1 mod 100), e.g. 2010→2010-11, 2025→2025-26. */
  function schoolYearLabelFromExcelYear(y) {
    var n = Number(y);
    if (isNaN(n)) return String(y);
    var end = (n + 1) % 100;
    var endStr = end < 10 ? "0" + end : String(end);
    return n + "-" + endStr;
  }

  /** Schools included in district-wide enrollment and demographics sums (matches dashboard scope). */
  function masterRowIncludedInDistrictAggregate(m) {
    if (!m) return false;
    return String(m.appears_in_dropdown || "").trim().toLowerCase() === "yes";
  }

  /** @returns {number[]} unique MSIDs with appears_in_dropdown=yes in school_master.csv */
  function getDistrictAggregateMsids() {
    var out = [];
    var seen = {};
    if (!MASTER_BY_MSID) return out;
    Object.keys(MASTER_BY_MSID).forEach(function (k) {
      var n = parseInt(k, 10);
      if (isNaN(n) || seen[n]) return;
      seen[n] = true;
      var m = MASTER_BY_MSID[k];
      if (!masterRowIncludedInDistrictAggregate(m)) return;
      out.push(n);
    });
    return out;
  }

  /** Sums enrollment calendar + projected series across all district schools (dropdown scope). */
  function buildDistrictEnrollmentSeries() {
    if (!MASTER_BY_MSID) return [];
    var msids = getDistrictAggregateMsids();
    if (!msids.length) return [];
    var merged = {};
    for (var i = 0; i < msids.length; i++) {
      var ser = buildEnrollmentSeries(msids[i]);
      for (var j = 0; j < ser.length; j++) {
        var s = ser[j];
        if (!merged[s.label]) {
          merged[s.label] = {
            label: s.label,
            value: 0,
            segment: s.segment,
          };
        }
        merged[s.label].value += Number(s.value) || 0;
      }
    }
    var labels = Object.keys(merged).sort(function (a, b) {
      return enrollmentLabelSortKey(a) - enrollmentLabelSortKey(b);
    });
    return labels.map(function (lb) {
      var pt = merged[lb];
      return {
        label: pt.label,
        value: Math.round(pt.value),
        segment: pt.segment,
      };
    });
  }

  function buildEnrollmentSeries(msid) {
    if (msid == null || isNaN(msid) || !MASTER_BY_MSID) return [];
    var m = masterRow(msid);
    if (!m) return [];
    var out = [];

    for (var y = 2010; y <= 2025; y++) {
      var col = "enrollment_" + y;
      var v = m[col];
      if (v !== "" && v != null && !isNaN(Number(v))) {
        out.push({
          label: schoolYearLabelFromExcelYear(y),
          value: Number(v),
          segment: "enrollment",
        });
      }
    }

    var labels = MASTER_PROJECTION_LABELS || [];
    for (var j = 0; j < labels.length; j++) {
      var col = projectedColumnForSyLabel(labels[j]);
      var pv = m[col];
      if (pv !== "" && pv != null && !isNaN(Number(pv))) {
        out.push({
          label: labels[j],
          value: Number(pv),
          segment: "projected",
        });
      }
    }
    return out;
  }

  function enrollmentLabelSortKey(label) {
    var proj = MASTER_PROJECTION_LABELS;
    if (proj && proj.indexOf(label) >= 0) {
      return 10000 + proj.indexOf(label);
    }
    var m = String(label).match(/^(\d{4})-/);
    if (m) return parseInt(m[1], 10);
    return 99999;
  }

  /** First school year shown on the scenario enrollment chart (future-focused). */
  var SCENARIO_CHART_FIRST_SY = "2025-26";

  function enrollmentSeriesLabelIsScenarioFuture(label) {
    if (label == null) return false;
    var s = String(label).trim();
    return s >= SCENARIO_CHART_FIRST_SY;
  }

  function filterEnrollmentSeriesScenarioFuture(series) {
    if (!series || !series.length) return [];
    return series.filter(function (pt) {
      return enrollmentSeriesLabelIsScenarioFuture(pt.label);
    });
  }

  var SCENARIO_STACK_MIDDLE_COLOR = "#2563eb";
  /** Dark → light; long enough that 9+ feeders do not wrap to the same hex as the first school. */
  var SCENARIO_STACK_ELEM_GREENS = [
    "#14532d",
    "#166534",
    "#15803d",
    "#16a34a",
    "#22c55e",
    "#4ade80",
    "#86efac",
    "#bbf7d0",
    "#d9f99d",
    "#ecfccb",
    "#f7fee7",
    "#ecfdf5",
    "#f0fdf4",
  ];

  /**
   * Assigns greens from dark → light in feeder-row order (checkbox list top → bottom).
   * Index uses the same palette position for every school (no modulo wrap onto the darkest).
   */
  function assignElementaryFeederGreenColors(elemMsids) {
    var order = elemMsids.slice();
    var greenByMsid = {};
    var n = SCENARIO_STACK_ELEM_GREENS.length;
    for (var gi = 0; gi < order.length; gi++) {
      var idx = gi < n ? gi : n - 1;
      greenByMsid[order[gi]] = SCENARIO_STACK_ELEM_GREENS[idx];
    }
    return greenByMsid;
  }

  /** All unique feeder elementary MSIDs for the scenario (same set used for checkbox swatches). */
  function scenarioFeederElementaryMsidsFromRows(middleMsid, feederRows) {
    var out = [];
    var seen = {};
    if (!feederRows || !feederRows.length) return out;
    for (var i = 0; i < feederRows.length; i++) {
      var m = feederRows[i].msid;
      if (m == null || isNaN(m) || m === middleMsid) continue;
      if (!seen[m]) {
        seen[m] = true;
        out.push(m);
      }
    }
    return out;
  }

  function findSeriesPointForLabel(series, label) {
    if (!series || !label) return null;
    for (var i = 0; i < series.length; i++) {
      if (series[i].label === label) return series[i];
    }
    return null;
  }

  /**
   * @param feederRows Scenario feeder rows (all elementaries for this middle); colors match checkbox swatches.
   * @returns {{ periods: { label: string, total: number, segments: { name: string, value: number, color: string, isMiddle: boolean }[] }[], maxVal: number }}
   */
  function buildScenarioStackedPeriods(
    weightedSpec,
    middleMsid,
    schoolByMsid,
    feederRows
  ) {
    var periods = [];
    var maxVal = 0;
    if (
      !weightedSpec ||
      !weightedSpec.length ||
      middleMsid == null ||
      isNaN(middleMsid) ||
      !schoolByMsid
    ) {
      return { periods: periods, maxVal: 1 };
    }

    var seriesCache = {};
    function getSeriesCached(msid) {
      var k = String(msid);
      if (!seriesCache[k]) {
        seriesCache[k] = buildEnrollmentSeries(msid);
      }
      return seriesCache[k];
    }

    var labelSet = {};
    for (var si = 0; si < weightedSpec.length; si++) {
      var ser = getSeriesCached(weightedSpec[si].msid);
      for (var sj = 0; sj < ser.length; sj++) {
        labelSet[ser[sj].label] = true;
      }
    }
    var labels = Object.keys(labelSet).sort(function (a, b) {
      return enrollmentLabelSortKey(a) - enrollmentLabelSortKey(b);
    });
    labels = labels.filter(enrollmentSeriesLabelIsScenarioFuture);

    var elemMsidsForColors =
      scenarioFeederElementaryMsidsFromRows(middleMsid, feederRows);
    if (!elemMsidsForColors.length) {
      var seenE = {};
      for (var wi = 0; wi < weightedSpec.length; wi++) {
        var wm = weightedSpec[wi].msid;
        if (wm === middleMsid || wm == null || isNaN(wm)) continue;
        if (!seenE[wm]) {
          seenE[wm] = true;
          elemMsidsForColors.push(wm);
        }
      }
    }
    var greenByMsid = assignElementaryFeederGreenColors(elemMsidsForColors);

    for (var li = 0; li < labels.length; li++) {
      var lab = labels[li];
      var segments = [];
      var total = 0;

      for (var wm = 0; wm < weightedSpec.length; wm++) {
        var ww = weightedSpec[wm];
        if (ww.msid !== middleMsid) continue;
        var pt = findSeriesPointForLabel(getSeriesCached(ww.msid), lab);
        var val = pt != null ? Math.round(Number(pt.value) * ww.weight) : 0;
        var mp = schoolByMsid[middleMsid];
        var mname = mp ? schoolNameForSelect(mp) : "Middle school";
        segments.push({
          name: mname,
          value: val,
          color: SCENARIO_STACK_MIDDLE_COLOR,
          isMiddle: true,
        });
        total += val;
      }

      var weightByElemMsid = {};
      for (var wi = 0; wi < weightedSpec.length; wi++) {
        var wx = weightedSpec[wi];
        if (wx.msid === middleMsid || wx.msid == null || isNaN(wx.msid)) continue;
        weightByElemMsid[wx.msid] = wx;
      }
      /* Lightest sits above middle (segment drawn first after middle); darkest on top — matches checkbox list top = dark, bottom = light. */
      for (var ei = elemMsidsForColors.length - 1; ei >= 0; ei--) {
        var emsid = elemMsidsForColors[ei];
        var ew = weightByElemMsid[emsid];
        if (!ew) continue;
        var ept = findSeriesPointForLabel(getSeriesCached(ew.msid), lab);
        var ev = ept != null ? Math.round(Number(ept.value) * ew.weight) : 0;
        var ep = schoolByMsid[ew.msid];
        var ename = ep ? schoolNameForSelect(ep) : String(ew.msid);
        segments.push({
          name: ename,
          value: ev,
          color: greenByMsid[ew.msid] || SCENARIO_STACK_ELEM_GREENS[0],
          isMiddle: false,
        });
        total += ev;
      }

      segments.sort(function (a, b) {
        if (a.isMiddle && !b.isMiddle) return -1;
        if (!a.isMiddle && b.isMiddle) return 1;
        return 0;
      });

      periods.push({ label: lab, segments: segments, total: total });
      if (total > maxVal) maxVal = total;
    }

    if (maxVal <= 0) maxVal = 1;
    return { periods: periods, maxVal: maxVal };
  }

  function teardownScenarioStackedChart(root) {
    if (root && typeof root._scenarioStackedCleanup === "function") {
      root._scenarioStackedCleanup();
      root._scenarioStackedCleanup = null;
    }
  }

  function renderScenarioStackedEnrollmentChartIntoRoot(root, stacked, options) {
    options = options || {};
    teardownScenarioStackedChart(root);
    if (!root) return;
    var noDataMsg =
      options.noDataMsg ||
      "No merged enrollment series from 2025-26 onward for the current selection (check data/school_master.csv).";
    if (!stacked.periods || !stacked.periods.length) {
      root.innerHTML =
        '<p class="enrollment-chart-empty">' + noDataMsg + "</p>";
      root.setAttribute(
        "aria-label",
        options.noDataAria || "Merged enrollment data is not available."
      );
      return;
    }

    var periods = stacked.periods;
    var maxVal = stacked.maxVal;
    var n = periods.length;
    var ml = 36;
    var mb = 54;
    var mt = 42;
    var mr = 10;
    var perBar = 34;
    var w = Math.min(1280, Math.max(480, ml + mr + n * perBar));
    var h = 252;
    var iw = w - ml - mr;
    var ih = h - mt - mb;
    var slot = iw / n;
    var barW = slot * 0.58;
    var gap = (slot - barW) / 2;
    var labelLift = 14;

    var parts = [];
    parts.push('<div class="scenario-enrollment-chart-wrap">');
    parts.push(
      '<div id="scenario-enrollment-tooltip" class="scenario-enrollment-tooltip" hidden></div>'
    );
    parts.push(
      '<svg xmlns="http://www.w3.org/2000/svg" class="scenario-enrollment-svg" style="min-width:' +
        w +
        'px" viewBox="0 0 ' +
        w +
        " " +
        h +
        '" aria-hidden="true">'
    );
    parts.push(
      '<line x1="' +
        ml +
        '" y1="' +
        (mt + ih) +
        '" x2="' +
        (w - mr) +
        '" y2="' +
        (mt + ih) +
        '" stroke="#e5e7eb" stroke-width="1" />'
    );

    for (var b = 0; b < n; b++) {
      var period = periods[b];
      var x = ml + b * slot + gap;
      var cum = 0;
      for (var s = 0; s < period.segments.length; s++) {
        var seg = period.segments[s];
        var sv = seg.value;
        var sh = maxVal > 0 ? (sv / maxVal) * ih : 0;
        var y = mt + ih - cum - sh;
        cum += sh;
        parts.push(
          '<rect class="scenario-stack-seg" data-bar="' +
            b +
            '" data-seg="' +
            s +
            '" x="' +
            x.toFixed(1) +
            '" y="' +
            y.toFixed(1) +
            '" width="' +
            barW.toFixed(1) +
            '" height="' +
            sh.toFixed(1) +
            '" fill="' +
            seg.color +
            '" rx="0" pointer-events="all" style="cursor:default"/>'
        );
      }
      var total = period.total;
      var topY = mt + ih - cum;
      var valY = topY - labelLift;
      parts.push(
        '<text x="' +
          (x + barW / 2) +
          '" y="' +
          valY +
          '" text-anchor="middle" dominant-baseline="alphabetic" font-size="11" font-weight="600" fill="#1f2937" font-family="Libre Franklin, sans-serif" pointer-events="none">' +
          escapeXmlText(total.toLocaleString()) +
          "</text>"
      );
      var lx = x + barW / 2;
      var ly = mt + ih + 12;
      parts.push(
        '<text x="' +
          lx +
          '" y="' +
          ly +
          '" text-anchor="end" transform="rotate(-52 ' +
          lx +
          " " +
          ly +
          ')" font-size="10" fill="#374151" font-family="Libre Franklin, sans-serif" pointer-events="none">' +
          escapeXmlText(period.label) +
          "</text>"
      );
    }
    parts.push("</svg>");
    parts.push(
      '<div class="enrollment-chart-legend" aria-hidden="true">' +
        '<span><i style="background:' +
        SCENARIO_STACK_MIDDLE_COLOR +
        '"></i> Middle school</span>' +
        '<span><i style="background:' +
        SCENARIO_STACK_ELEM_GREENS[0] +
        '"></i> Feeder elementaries (shades)</span>' +
        "</div>"
    );
    parts.push("</div>");

    root.innerHTML = parts.join("");
    root.setAttribute(
      "aria-label",
      options.ariaLabel ||
        "Stacked enrollment by school from 2025-26 forward (scenario)."
    );
    root.classList.add("enrollment-chart--stacked");

    var svg = root.querySelector(".scenario-enrollment-svg");
    var tip = document.getElementById("scenario-enrollment-tooltip");
    if (!svg || !tip) return;

    function showTooltipOne(periodLabel, seg, clientX, clientY) {
      if (!seg) {
        hideTooltip();
        return;
      }
      tip.removeAttribute("hidden");
      tip.innerHTML = "";
      var head = document.createElement("div");
      head.className = "scenario-enrollment-tooltip-title";
      head.textContent = periodLabel;
      tip.appendChild(head);
      var row = document.createElement("div");
      row.className = "scenario-enrollment-tooltip-row";
      var sw = document.createElement("span");
      sw.className = "scenario-enrollment-tooltip-swatch";
      sw.style.background = seg.color;
      row.appendChild(sw);
      row.appendChild(
        document.createTextNode(
          seg.name + ": " + Number(seg.value).toLocaleString()
        )
      );
      tip.appendChild(row);
      tip.style.left = Math.min(clientX + 14, window.innerWidth - 280) + "px";
      tip.style.top = Math.min(clientY + 14, window.innerHeight - 200) + "px";
    }

    function hideTooltip() {
      tip.setAttribute("hidden", "hidden");
    }

    function onMove(e) {
      var t = e.target;
      if (
        t &&
        t.classList &&
        t.classList.contains("scenario-stack-seg")
      ) {
        var b = parseInt(t.getAttribute("data-bar"), 10);
        var si = parseInt(t.getAttribute("data-seg"), 10);
        var period = periods[b];
        if (
          !isNaN(b) &&
          !isNaN(si) &&
          period &&
          period.segments &&
          period.segments[si]
        ) {
          showTooltipOne(
            period.label,
            period.segments[si],
            e.clientX,
            e.clientY
          );
          return;
        }
      }
      hideTooltip();
    }

    function onLeave() {
      hideTooltip();
    }

    svg.addEventListener("mousemove", onMove);
    svg.addEventListener("mouseleave", onLeave);

    root._scenarioStackedCleanup = function () {
      svg.removeEventListener("mousemove", onMove);
      svg.removeEventListener("mouseleave", onLeave);
      root.classList.remove("enrollment-chart--stacked");
    };
  }

  /** Sums calendar + projected series by label; each entry is { msid, weight }. Middle school weight is always 1. */
  function buildMergedEnrollmentSeriesWeighted(weighted) {
    var merged = {};
    for (var i = 0; i < weighted.length; i++) {
      var msid = weighted[i].msid;
      var wt = weighted[i].weight;
      if (msid == null || isNaN(msid) || wt == null || isNaN(wt)) continue;
      var series = buildEnrollmentSeries(msid);
      for (var j = 0; j < series.length; j++) {
        var s = series[j];
        if (!merged[s.label]) {
          merged[s.label] = { label: s.label, value: 0, segment: s.segment };
        }
        merged[s.label].value += s.value * wt;
      }
    }
    var labels = Object.keys(merged).sort(function (a, b) {
      return enrollmentLabelSortKey(a) - enrollmentLabelSortKey(b);
    });
    return labels.map(function (lb) {
      var pt = merged[lb];
      return {
        label: pt.label,
        value: Math.round(pt.value),
        segment: pt.segment,
      };
    });
  }

  /** Shorter labels (e.g. 66.8k) when count > 9999 to reduce overlap; full count stays on rect/title tooltip. */
  function formatEnrollmentBarAxisLabel(val) {
    if (val <= 9999) return val.toLocaleString();
    var k = val / 1000;
    var rounded = Math.round(k * 10) / 10;
    var ir = Math.round(rounded);
    if (Math.abs(rounded - ir) < 0.001) {
      return String(ir) + "k";
    }
    return rounded.toFixed(1) + "k";
  }

  function setMainEnrollmentDemographicsHeadings(isDistrict) {
    var eh = document.getElementById("main-enrollment-chart-heading");
    if (eh) {
      eh.textContent = isDistrict
        ? "Sum of Non-Charter Schools: Enrollment over Time"
        : "Enrollment over time";
    }
    var eth = document.getElementById("demographics-ethnicity-heading");
    var lunchH = document.getElementById("demographics-lunch-heading");
    if (eth) {
      eth.textContent = isDistrict
        ? "Sum of Non-Charter Schools: Race and Ethnicity"
        : "Race and Ethnicity";
    }
    if (lunchH) {
      lunchH.textContent = isDistrict
        ? "Sum of Non-Charter Schools: Free and Reduced Lunch"
        : "Free and Reduced Lunch";
    }
  }

  function renderEnrollmentChartIntoRoot(root, series, options) {
    options = options || {};
    if (!root) return;
    var noDataMsg =
      options.noDataMsg ||
      "No enrollment rows in data/school_master.csv for this school.";
    if (!series || !series.length) {
      root.innerHTML =
        '<p class="enrollment-chart-empty">' + noDataMsg + "</p>";
      root.setAttribute(
        "aria-label",
        options.noDataAria || "Enrollment data is not available."
      );
      return;
    }
    var maxVal = 0;
    for (var i = 0; i < series.length; i++) {
      if (series[i].value > maxVal) maxVal = series[i].value;
    }
    if (maxVal <= 0) maxVal = 1;
    var n = series.length;
    var hasLargeBarValues = false;
    for (var li = 0; li < series.length; li++) {
      if (series[li].value > 9999) {
        hasLargeBarValues = true;
        break;
      }
    }
    var ml = 36;
    var mb = 54;
    /** Top margin: room so value labels sit fully above bars (incl. tallest). */
    var mt = 42;
    var mr = 10;
    var perBar = hasLargeBarValues ? 38 : 34;
    var w = Math.min(1280, Math.max(480, ml + mr + n * perBar));
    var h = 252;
    var iw = w - ml - mr;
    var ih = h - mt - mb;
    var slot = iw / n;
    var barW = slot * 0.58;
    var gap = (slot - barW) / 2;
    /** Pixels from bar top to label baseline (labels render upward from baseline). */
    var labelLift = 14;

    var parts = [];
    parts.push(
      '<svg xmlns="http://www.w3.org/2000/svg" style="min-width:' +
        w +
        'px" viewBox="0 0 ' +
        w +
        " " +
        h +
        '" aria-hidden="true">'
    );
    parts.push(
      '<line x1="' +
        ml +
        '" y1="' +
        (mt + ih) +
        '" x2="' +
        (w - mr) +
        '" y2="' +
        (mt + ih) +
        '" stroke="#e5e7eb" stroke-width="1" />'
    );

    for (var b = 0; b < series.length; b++) {
      var s = series[b];
      var val = s.value;
      var bh = (val / maxVal) * ih;
      var x = ml + b * slot + gap;
      var y = mt + ih - bh;
      var fill =
        s.segment === "projected"
          ? ENCHART_COLORS.projected
          : ENCHART_COLORS.calendar;
      parts.push(
        '<rect x="' +
          x.toFixed(1) +
          '" y="' +
          y.toFixed(1) +
          '" width="' +
          barW.toFixed(1) +
          '" height="' +
          bh.toFixed(1) +
          '" fill="' +
          fill +
          '" rx="2"><title>' +
          escapeXmlText(
            s.label + ": " + val.toLocaleString() + " students"
          ) +
          "</title></rect>"
      );
      var cx = x + barW / 2;
      var valY = y - labelLift;
      var axisLabel = formatEnrollmentBarAxisLabel(val);
      parts.push(
        '<text x="' +
          cx +
          '" y="' +
          valY +
          '" text-anchor="middle" dominant-baseline="alphabetic" font-size="11" font-weight="600" fill="#1f2937" font-family="Libre Franklin, sans-serif">' +
          (val > 9999
            ? "<title>" +
              escapeXmlText(
                s.label + ": " + val.toLocaleString() + " students"
              ) +
              "</title>"
            : "") +
          escapeXmlText(axisLabel) +
          "</text>"
      );
      var lx = cx;
      var ly = mt + ih + 12;
      parts.push(
        '<text x="' +
          lx +
          '" y="' +
          ly +
          '" text-anchor="end" transform="rotate(-52 ' +
          lx +
          " " +
          ly +
          ')" font-size="10" fill="#374151" font-family="Libre Franklin, sans-serif">' +
          escapeXmlText(s.label) +
          "</text>"
      );
    }
    parts.push("</svg>");
    parts.push(
      '<div class="enrollment-chart-legend" aria-hidden="true">' +
        '<span><i style="background:' +
        ENCHART_COLORS.calendar +
        '"></i> Enrollment</span>' +
        '<span><i style="background:' +
        ENCHART_COLORS.projected +
        '"></i> Projected Enrollment</span>' +
        "</div>"
    );
    root.innerHTML = parts.join("");
    root.setAttribute(
      "aria-label",
      options.ariaLabel ||
        "Enrollment bar chart with " + n + " periods for the selected school."
    );
  }

  function renderEnrollmentChart(msid) {
    setMainEnrollmentDemographicsHeadings(msid == null || isNaN(msid));
    var root = document.getElementById("enrollment-chart");
    if (!root) return;
    if (msid == null || isNaN(msid)) {
      var distSeries = buildDistrictEnrollmentSeries();
      renderEnrollmentChartIntoRoot(root, distSeries, {
        noDataMsg:
          "No enrollment rows in data/school_master.csv for district schools (appears_in_dropdown=yes).",
        noDataAria:
          "District enrollment data is not available in data/school_master.csv.",
        ariaLabel:
          "District-wide enrollment bar chart: sum of calendar and projected membership for all schools with appears_in_dropdown=yes in data/school_master.csv.",
      });
      return;
    }
    var series = buildEnrollmentSeries(msid);
    renderEnrollmentChartIntoRoot(root, series, {
      noDataMsg:
        "No enrollment rows in data/school_master.csv for this school.",
      noDataAria:
        "Enrollment data is not available for this school in data/school_master.csv.",
      ariaLabel:
        "Enrollment bar chart with periods for the selected school.",
    });
  }

  /**
   * Fallback when an ethnicity label is not in the fixed map below (e.g. new export values).
   */
  var DEMOGRAPHICS_PIE_COLORS = [
    "#795548",
    "#e65100",
    "#fb8c00",
    "#f9a825",
    "#c0ca33",
    "#7cb342",
    "#558b2f",
    "#00897b",
    "#039be5",
    "#3949ab",
    "#7b1fa2",
    "#c2185b",
  ];

  function lunchSliceColor(label) {
    var u = String(label).toLowerCase();
    /** Must run before "reduced" — "Not free/reduced" also contains "reduced". */
    if (u.indexOf("not free") >= 0) return "#e53935";
    if (u === "free") return "#689f38";
    if (u.indexOf("reduced") >= 0) return "#fbc02d";
    return "#78909c";
  }

  /** Fixed label → color for race/ethnicity pies (not rank-based). */
  function ethnicitySliceColor(label, idx) {
    var s = String(label).trim().toLowerCase();
    if (s.indexOf("white") >= 0 && s.indexOf("non-hispanic") >= 0) {
      return "#93612c";
    }
    if (s.indexOf("black") >= 0 && s.indexOf("non-hispanic") >= 0) {
      return "#fb8c00";
    }
    if (s === "hispanic" || (s.indexOf("hispanic") >= 0 && s.indexOf("non-hispanic") < 0)) {
      return "#e65100";
    }
    if (
      s.indexOf("multi-racial") >= 0 ||
      s.indexOf("multiracial") >= 0 ||
      s.indexOf("mixed race") >= 0
    ) {
      return "#fdd835";
    }
    if (s === "asian") {
      return "#c0ca33";
    }
    if (
      s.indexOf("amer. indian") >= 0 ||
      s.indexOf("american indian") >= 0 ||
      s.indexOf("alaskan native") >= 0
    ) {
      return "#7cb342";
    }
    if (s.indexOf("hawaiian") >= 0 || s.indexOf("pacific islander") >= 0) {
      return "#00897b";
    }
    return DEMOGRAPHICS_PIE_COLORS[idx % DEMOGRAPHICS_PIE_COLORS.length];
  }

  function buildPieChartHtml(countsObj, colorForIndex) {
    var entries = Object.keys(countsObj).map(function (k) {
      return { label: k, value: Number(countsObj[k]) };
    }).filter(function (e) {
      return e.value > 0 && !isNaN(e.value);
    });
    entries.sort(function (a, b) {
      return b.value - a.value;
    });
    var total = entries.reduce(function (s, e) {
      return s + e.value;
    }, 0);
    if (total <= 0) {
      return {
        html:
          '<p class="demographics-pie-empty">No students in this category for the selected school.</p>',
        total: 0,
      };
    }
    var cx = 100;
    var cy = 100;
    var r = 88;
    var angle = -Math.PI / 2;
    var pathParts = [];
    for (var i = 0; i < entries.length; i++) {
      var slice = entries[i];
      var frac = slice.value / total;
      var a2 = angle + frac * 2 * Math.PI;
      var large = frac > 0.5 ? 1 : 0;
      var x1 = cx + r * Math.cos(angle);
      var y1 = cy + r * Math.sin(angle);
      var x2 = cx + r * Math.cos(a2);
      var y2 = cy + r * Math.sin(a2);
      var d = [
        "M",
        cx,
        cy,
        "L",
        x1.toFixed(3),
        y1.toFixed(3),
        "A",
        r,
        r,
        0,
        large,
        1,
        x2.toFixed(3),
        y2.toFixed(3),
        "Z",
      ].join(" ");
      var fill = colorForIndex(slice.label, i);
      pathParts.push(
        '<path d="' +
          d +
          '" fill="' +
          fill +
          '" stroke="#fff" stroke-width="1.5"><title>' +
          escapeXmlText(
            slice.label +
              ": " +
              slice.value +
              " (" +
              ((slice.value / total) * 100).toFixed(1) +
              "%)"
          ) +
          "</title></path>"
      );
      angle = a2;
    }
    var legendItems = [];
    for (var j = 0; j < entries.length; j++) {
      var e = entries[j];
      var pct = ((e.value / total) * 100).toFixed(1);
      var fillJ = colorForIndex(e.label, j);
      legendItems.push(
        "<li>" +
          '<span class="demographics-legend-swatch" style="background:' +
          fillJ +
          '"></span>' +
          "<span>" +
          escapeXmlText(e.label) +
          " — " +
          e.value.toLocaleString() +
          " (" +
          pct +
          "%)</span></li>"
      );
    }
    return {
      html:
        '<div class="demographics-pie-inner"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 200" aria-hidden="true">' +
        pathParts.join("") +
        '</svg><ul class="demographics-legend">' +
        legendItems.join("") +
        "</ul></div>",
      total: total,
    };
  }

  function renderDemographicsCharts(msid) {
    setMainEnrollmentDemographicsHeadings(msid == null || isNaN(msid));
    var ethEl = document.getElementById("demographics-ethnicity");
    var lunchEl = document.getElementById("demographics-lunch");
    if (!ethEl || !lunchEl) return;

    if (msid == null || isNaN(msid)) {
      var dms = getDistrictAggregateMsids();
      var weighted = dms.map(function (id) {
        return { msid: id, weight: 1 };
      });
      var agg = aggregateDemographicsMsidsWeighted(weighted);
      renderDemographicsFromAggregates(
        agg,
        ethEl,
        lunchEl,
        '<p class="demographics-pie-empty">No student demographics in data/school_master.csv for district schools (appears_in_dropdown=yes).</p>'
      );
      return;
    }
    if (!MASTER_BY_MSID) {
      ethEl.innerHTML =
        '<p class="demographics-pie-empty">School master data is not loaded.</p>';
      lunchEl.innerHTML =
        '<p class="demographics-pie-empty">School master data is not loaded.</p>';
      return;
    }
    var objs = demographicsObjectsFromMaster(masterRow(msid));
    if (!objs) {
      var msg =
        '<p class="demographics-pie-empty">No student rows for this school in the SY2025-26 export.</p>';
      ethEl.innerHTML = msg;
      lunchEl.innerHTML = msg;
      return;
    }

    var ethRes = buildPieChartHtml(objs.ethnicity || {}, ethnicitySliceColor);
    ethEl.innerHTML = ethRes.html;

    var lunchRes = buildPieChartHtml(objs.lunchStatus || {}, function (label) {
      return lunchSliceColor(label);
    });
    lunchEl.innerHTML = lunchRes.html;
  }

  function mergeCountObjScaled(dst, src, scale) {
    if (!src || scale == null || isNaN(scale)) return;
    Object.keys(src).forEach(function (k) {
      var v = Number(src[k]);
      if (isNaN(v) || v <= 0) return;
      dst[k] = (dst[k] || 0) + v * scale;
    });
  }

  function aggregateDemographicsMsidsWeighted(weighted) {
    var eth = {};
    var lunch = {};
    if (!MASTER_BY_MSID) {
      return { ethnicity: eth, lunchStatus: lunch };
    }
    for (var i = 0; i < weighted.length; i++) {
      var msid = weighted[i].msid;
      var wt = weighted[i].weight;
      if (msid == null || isNaN(msid) || wt == null || isNaN(wt)) continue;
      var objs = demographicsObjectsFromMaster(masterRow(msid));
      if (!objs) continue;
      mergeCountObjScaled(eth, objs.ethnicity || {}, wt);
      mergeCountObjScaled(lunch, objs.lunchStatus || {}, wt);
    }
    return { ethnicity: eth, lunchStatus: lunch };
  }

  function renderDemographicsFromAggregates(agg, ethEl, lunchEl, emptyMsg) {
    var emptyAgg =
      emptyMsg ||
      '<p class="demographics-pie-empty">No student rows for merged selection in the SY2025-26 export.</p>';
    if (!ethEl || !lunchEl) return;
    var ethRes = buildPieChartHtml(agg.ethnicity || {}, ethnicitySliceColor);
    var lunchRes = buildPieChartHtml(agg.lunchStatus || {}, function (label) {
      return lunchSliceColor(label);
    });
    ethEl.innerHTML = ethRes.total > 0 ? ethRes.html : emptyAgg;
    lunchEl.innerHTML = lunchRes.total > 0 ? lunchRes.html : emptyAgg;
  }

  function schoolHasEnrollmentWorkbook(msid) {
    if (msid == null || isNaN(msid)) return false;
    return buildEnrollmentSeries(msid).length > 0;
  }

  function findElementaryPropsBySankeyLabel(label, schoolsFc) {
    if (!schoolsFc || !schoolsFc.features) return null;
    for (var i = 0; i < schoolsFc.features.length; i++) {
      var p = schoolsFc.features[i].properties;
      if (!p) continue;
      var t = (p.TYPE || "").toUpperCase();
      if (t.indexOf("ELEMENTARY") < 0) continue;
      if (sankeyElementaryLabelMatchesSchool(label, p)) return p;
    }
    return null;
  }

  /** Sum of flow counts from one elementary to all middle schools (denominator for share to this middle). */
  function elementaryOutgoingTotalsMap(flows) {
    var m = {};
    if (!flows || !flows.length) return m;
    for (var i = 0; i < flows.length; i++) {
      var f = flows[i];
      if (!f || f.elementary == null) continue;
      var v = Number(f.value);
      if (isNaN(v) || v < 1) continue;
      var key = f.elementary;
      m[key] = (m[key] || 0) + v;
    }
    return m;
  }

  /**
   * Selected middle school row for the scenario feeder list (blue swatch, checkbox controls merged totals).
   */
  function buildScenarioMiddleFeederRow(middleProps, middleMsid) {
    var hasEnrollment = schoolHasEnrollmentWorkbook(middleMsid);
    return {
      sankeyLabel: "",
      msid: middleMsid,
      props: middleProps,
      hasEnrollment: hasEnrollment,
      flowValue: null,
      flowProportion: 1,
      isScenarioMiddleRow: true,
    };
  }

  function getFeederElementaryRowsForMiddle(middleProps, flows, schoolsFc) {
    var rows = [];
    if (!flows || !middleProps) return rows;
    var outgoingByEl = elementaryOutgoingTotalsMap(flows);
    var seen = {};
    for (var i = 0; i < flows.length; i++) {
      var f = flows[i];
      if (!f || f.value < 1) continue;
      if (!sankeyMiddleLabelMatchesSchool(f.middle, middleProps)) continue;
      var key = f.elementary;
      if (seen[key]) continue;
      seen[key] = true;
      var p = findElementaryPropsBySankeyLabel(f.elementary, schoolsFc);
      var msid =
        p && p.SCHOOLS_ID != null ? Number(p.SCHOOLS_ID) : null;
      var hasEnrollment = schoolHasEnrollmentWorkbook(msid);
      var totalOut = outgoingByEl[f.elementary] || 0;
      var flowProportion = totalOut > 0 ? f.value / totalOut : 1;
      rows.push({
        sankeyLabel: f.elementary,
        msid: msid,
        props: p,
        hasEnrollment: hasEnrollment,
        flowValue: f.value,
        flowProportion: flowProportion,
      });
    }
    rows.sort(function (a, b) {
      return a.sankeyLabel.localeCompare(b.sankeyLabel);
    });
    return rows;
  }

  /** Same 2025 calendar column as main dashboard ’25-26 enrollment KPI. */
  function enrollment202526CalendarForMsid(msid) {
    if (msid == null || isNaN(msid)) return null;
    var m = masterRow(msid);
    if (!m) return null;
    var v = m.enrollment_2025;
    if (v !== "" && v != null && !isNaN(Number(v))) {
      return Number(v);
    }
    return null;
  }

  function buildMiddleSchoolMsidSetFromSchoolsFc(schoolsFc) {
    var o = {};
    if (!schoolsFc || !schoolsFc.features) return o;
    for (var i = 0; i < schoolsFc.features.length; i++) {
      var p = schoolsFc.features[i].properties;
      if (!p || p.SCHOOLS_ID == null || p.SCHOOLS_ID === "") continue;
      var t = (p.TYPE || "").toUpperCase();
      if (t.indexOf("MIDDLE") >= 0 && t.indexOf("HIGH") < 0) {
        o[String(Number(p.SCHOOLS_ID))] = true;
      }
    }
    return o;
  }

  /**
   * Middle-school attendance rows only count when attendance matches the scenario middle,
   * and when that middle is included via the feeder-list checkbox (same as merged enrollment).
   * Elementary attendance respects feeder checkboxes unless ignoreFeederCheckboxes (axis extent only).
   */
  function attendancePassesScenarioTravelFilter(
    attMsid,
    selectedMiddleMsid,
    feederRows,
    ignoreFeederCheckboxes
  ) {
    var ms = Number(attMsid);
    if (isNaN(ms)) return false;
    var msStr = String(ms);
    var midSet = MIDDLE_SCHOOL_MSID_SET || {};
    if (midSet[msStr]) {
      if (ms !== selectedMiddleMsid) {
        return false;
      }
      if (ignoreFeederCheckboxes) {
        return true;
      }
      return scenarioFeederChecked[selectedMiddleMsid] !== false;
    }
    for (var i = 0; i < feederRows.length; i++) {
      var r = feederRows[i];
      if (r.msid != null && !isNaN(r.msid) && r.msid === ms) {
        if (!r.hasEnrollment) return false;
        if (ignoreFeederCheckboxes) return true;
        return scenarioFeederChecked[r.msid] !== false;
      }
    }
    return true;
  }

  /** @returns {number|null} miles, or null if invalid / omit */
  function travelMilesFromFeet(ft) {
    var n = Number(ft);
    if (!isFinite(n) || n <= 0) return null;
    return n / FEET_PER_MILE;
  }

  function medianOfNumbers(vals) {
    if (!vals || !vals.length) return null;
    var s = vals.slice().sort(function (a, b) {
      return a - b;
    });
    var n = s.length;
    var h = Math.floor(n / 2);
    if (n % 2 === 1) return s[h];
    return (s[h - 1] + s[h]) / 2;
  }

  function meanOfNumbers(vals) {
    if (!vals || !vals.length) return null;
    var s = 0;
    var i;
    for (i = 0; i < vals.length; i++) s += vals[i];
    return s / vals.length;
  }

  /** Tukey upper inner fence (Q3 + 3×IQR); null when not applicable. */
  function travelTukeyUpperFenceMiles(miles) {
    if (!miles || miles.length < 4) return null;
    var sorted = miles.slice().sort(function (a, b) {
      return a - b;
    });
    var n = sorted.length;
    var q1 = sorted[Math.floor((n - 1) * 0.25)];
    var q3 = sorted[Math.ceil((n - 1) * 0.75)];
    var iqr = q3 - q1;
    if (!(iqr > 0) || !isFinite(iqr)) return null;
    return q3 + 3 * iqr;
  }

  /** Drops values above the Tukey upper fence; keeps all when fence undefined. */
  function travelMilesExcludingUpperOutliers(miles) {
    if (!miles || !miles.length) return [];
    var fence = travelTukeyUpperFenceMiles(miles);
    if (fence == null || !isFinite(fence)) return miles.slice();
    return miles.filter(function (m) {
      return m <= fence + 1e-9;
    });
  }

  /** X-axis max (mi) from retained distances only; bucket-aligned. Outliers should already be removed. */
  function travelHistogramAxisExtentFromMiles(miles) {
    var bw = TRAVEL_BIN_MI;
    if (!miles || !miles.length) return bw;
    var maxV = Math.max.apply(null, miles);
    return Math.max(bw, Math.ceil(maxV / bw) * bw);
  }

  /** Shared x-axis for existing + scenario pair after each side drops outliers. */
  function travelHistogramPairedAxisHiMiles(milesA, milesB) {
    var bw = TRAVEL_BIN_MI;
    var ha =
      milesA && milesA.length ? travelHistogramAxisExtentFromMiles(milesA) : 0;
    var hb =
      milesB && milesB.length ? travelHistogramAxisExtentFromMiles(milesB) : 0;
    if (ha <= 0 && hb <= 0) return bw;
    return Math.max(bw, ha, hb);
  }

  function binTravelHistogramCounts(miles, axisHi) {
    var bw = TRAVEL_BIN_MI;
    var numBins = Math.max(1, Math.round(axisHi / bw));
    var counts = [];
    var b;
    for (b = 0; b < numBins; b++) counts[b] = 0;
    for (var j = 0; j < miles.length; j++) {
      var m = miles[j];
      var idx;
      if (m >= axisHi - 1e-12) idx = numBins - 1;
      else {
        idx = Math.floor(m / bw);
        if (idx < 0) idx = 0;
        if (idx >= numBins) idx = numBins - 1;
      }
      counts[idx]++;
    }
    return { counts: counts, numBins: numBins, axisHi: axisHi };
  }

  /**
   * @param {'existing'|'scenario'} mode
   * @param {boolean} [ignoreFeederCheckboxes] If true, axis-style extent: include all feeder schools regardless of checkbox.
   * @returns {number[]} distances in miles
   */
  function collectScenarioTravelMiles(
    mode,
    triples,
    selectedMiddleMsid,
    feederRows,
    ignoreFeederCheckboxes
  ) {
    var out = [];
    if (!triples || !triples.length) return out;
    var sel = selectedMiddleMsid;
    var ignFeed =
      ignoreFeederCheckboxes === true;
    for (var i = 0; i < triples.length; i++) {
      var row = triples[i];
      if (!row || row.length < 3) continue;
      var att = Number(row[0]);
      var sce = Number(row[1]);
      var mi = travelMilesFromFeet(row[2]);
      if (mi == null) continue;
      if (
        !attendancePassesScenarioTravelFilter(att, sel, feederRows, ignFeed)
      )
        continue;
      if (mode === "existing") {
        if (att !== sce) continue;
      } else {
        if (sce !== sel) continue;
      }
      out.push(mi);
    }
    return out;
  }

  function renderTravelHistogramIntoRoot(root, chartTitle, miles, options) {
    options = options || {};
    if (!root) return;
    var bw = TRAVEL_BIN_MI;
    if (!miles || !miles.length) {
      root.innerHTML =
        '<p class="travel-hist-empty">No students match the current filters.</p>';
      root.setAttribute(
        "aria-label",
        chartTitle + ": no data for the current filters."
      );
      return;
    }
    var milesUse = travelMilesExcludingUpperOutliers(miles);
    if (!milesUse.length) {
      root.innerHTML =
        '<p class="travel-hist-empty">All distances were excluded as statistical outliers (Tukey upper fence).</p>';
      root.setAttribute(
        "aria-label",
        chartTitle + ": all values excluded as outliers."
      );
      return;
    }
    var axisHi =
      options.axisHiOverride != null &&
      isFinite(options.axisHiOverride) &&
      options.axisHiOverride > 0
        ? options.axisHiOverride
        : travelHistogramAxisExtentFromMiles(milesUse);
    if (!(axisHi > 0) || !isFinite(axisHi)) axisHi = bw;
    var binRes = binTravelHistogramCounts(milesUse, axisHi);
    var counts = binRes.counts;
    var numBins = binRes.numBins;
    var maxCount = 0;
    var c;
    for (c = 0; c < counts.length; c++) {
      if (counts[c] > maxCount) maxCount = counts[c];
    }
    if (maxCount <= 0) maxCount = 1;

    var ml = 44;
    var mr = 118;
    var medTextY = 6;
    var meanTextY = 24;
    var mt = 40;
    var mb = 46;
    var cw = root.clientWidth || 0;
    if (cw < 80 && root.closest) {
      var chartCard = root.closest(".travel-impact-chart-card");
      if (chartCard) {
        cw = Math.max(0, chartCard.getBoundingClientRect().width - 24);
      }
    }
    if (cw < 80) {
      cw =
        typeof window !== "undefined"
          ? Math.min(920, Math.max(320, window.innerWidth - 120))
          : 520;
    }
    var iw = Math.max(220, cw - ml - mr);
    var ih = 156;
    var med = medianOfNumbers(milesUse);
    var avg = meanOfNumbers(milesUse);
    var medLabel =
      med != null && isFinite(med) ? "Median " + med.toFixed(2) + " mi" : "";
    var meanLabel =
      avg != null && isFinite(avg)
        ? "Average " + avg.toFixed(2) + " mi"
        : "";

    var parts = [];
    parts.push(
      '<svg xmlns="http://www.w3.org/2000/svg" class="travel-hist-svg" viewBox="0 0 ' +
        (ml + iw + mr) +
        " " +
        (mt + ih + mb) +
        '" aria-hidden="true">'
    );

    var baselineY = mt + ih;
    parts.push(
      '<line x1="' +
        ml +
        '" y1="' +
        baselineY +
        '" x2="' +
        (ml + iw) +
        '" y2="' +
        baselineY +
        '" stroke="#e5e7eb" stroke-width="1" />'
    );

    function xPix(miCoord) {
      return ml + (miCoord / axisHi) * iw;
    }

    var barGap = 1;
    for (var b = 0; b < numBins; b++) {
      var x0 = ((b * bw) / axisHi) * iw;
      var barW = Math.max(0.5, (bw / axisHi) * iw - barGap);
      var h = maxCount > 0 ? (counts[b] / maxCount) * ih : 0;
      var bx = ml + x0 + barGap / 2;
      var by = baselineY - h;
      var tip =
        counts[b] > 0
          ? "<title>" +
            escapeXmlText(
              (b * bw).toFixed(2) +
                "–" +
                Math.min(axisHi, (b + 1) * bw).toFixed(2) +
                " mi: " +
                counts[b].toLocaleString()
            ) +
            "</title>"
          : "";
      parts.push(
        '<rect class="travel-hist-bar" x="' +
          bx.toFixed(2) +
          '" y="' +
          by.toFixed(2) +
          '" width="' +
          barW.toFixed(2) +
          '" height="' +
          h.toFixed(2) +
          '" fill="#64748b" rx="1">' +
          tip +
          "</rect>"
      );
    }

    if (med != null && isFinite(med)) {
      var mx = xPix(Math.min(med, axisHi));
      parts.push(
        '<line class="travel-hist-median" x1="' +
          mx.toFixed(2) +
          '" y1="' +
          mt +
          '" x2="' +
          mx.toFixed(2) +
          '" y2="' +
          baselineY +
          '" stroke="' +
          TRAVEL_MEDIAN_COLOR +
          '" stroke-width="2" stroke-dasharray="4 3" />'
      );
      if (medLabel) {
        var flagX = mx + 6;
        parts.push(
          '<text class="travel-hist-median-flag" x="' +
            flagX.toFixed(2) +
            '" y="' +
            medTextY +
            '" font-size="13" font-weight="600" fill="' +
            TRAVEL_MEDIAN_COLOR +
            '" font-family="Libre Franklin, sans-serif" text-anchor="start" dominant-baseline="hanging" pointer-events="none">' +
            escapeXmlText(medLabel) +
            "</text>"
        );
      }
    }

    if (avg != null && isFinite(avg)) {
      var ax = xPix(Math.min(avg, axisHi));
      parts.push(
        '<line class="travel-hist-mean" x1="' +
          ax.toFixed(2) +
          '" y1="' +
          mt +
          '" x2="' +
          ax.toFixed(2) +
          '" y2="' +
          baselineY +
          '" stroke="' +
          TRAVEL_MEAN_COLOR +
          '" stroke-width="2" stroke-dasharray="6 4" />'
      );
      if (meanLabel) {
        var meanFlagX = ax + 6;
        parts.push(
          '<text class="travel-hist-mean-flag" x="' +
            meanFlagX.toFixed(2) +
            '" y="' +
            meanTextY +
            '" font-size="13" font-weight="600" fill="' +
            TRAVEL_MEAN_COLOR +
            '" font-family="Libre Franklin, sans-serif" text-anchor="start" dominant-baseline="hanging" pointer-events="none">' +
            escapeXmlText(meanLabel) +
            "</text>"
        );
      }
    }

    var maxIntTick = Math.ceil(axisHi - 1e-9);
    var ki;
    for (ki = 0; ki <= maxIntTick; ki++) {
      if (ki > axisHi + 1e-9) break;
      var tx = xPix(Math.min(ki, axisHi));
      parts.push(
        '<line x1="' +
          tx.toFixed(2) +
          '" y1="' +
          baselineY +
          '" x2="' +
          tx.toFixed(2) +
          '" y2="' +
          (baselineY + 5) +
          '" stroke="#d1d5db" stroke-width="1" />'
      );
      parts.push(
        '<text x="' +
          tx.toFixed(2) +
          '" y="' +
          (baselineY + 18) +
          '" text-anchor="middle" font-size="10" fill="#4b5563" font-family="Libre Franklin, sans-serif">' +
          escapeXmlText(String(ki)) +
          "</text>"
      );
    }

    parts.push(
      '<text x="' +
        (ml + iw / 2) +
        '" y="' +
        (baselineY + mb - 4) +
        '" text-anchor="middle" font-size="10" fill="#6b7280" font-family="Libre Franklin, sans-serif">Network Travel Distance (Miles)</text>'
    );

    parts.push("</svg>");
    root.innerHTML = parts.join("");
    root.setAttribute(
      "aria-label",
      chartTitle + ": histogram by " + bw + " mi buckets."
    );
  }

  function renderScenarioTravelImpactCharts() {
    var elEx = document.getElementById("scenario-travel-existing");
    var elSc = document.getElementById("scenario-travel-scenario");
    if (!elEx || !elSc) return;

    if (
      scenarioMiddleMsid == null ||
      isNaN(scenarioMiddleMsid) ||
      PRIORITY_SCHOOL_MSIDS.indexOf(scenarioMiddleMsid) < 0
    ) {
      elEx.innerHTML =
        '<p class="travel-hist-empty">Select a middle school to view travel distances.</p>';
      elSc.innerHTML =
        '<p class="travel-hist-empty">Select a middle school to view travel distances.</p>';
      return;
    }

    var pack =
      TRAVEL_IMPACT_ALL &&
      TRAVEL_IMPACT_ALL.byMsid &&
      TRAVEL_IMPACT_ALL.byMsid[String(scenarioMiddleMsid)];
    if (!pack || !pack.rows || !pack.rows.length) {
      var miss =
        '<p class="travel-hist-empty">Travel distance data is not loaded. Run scripts/export_travel_impact_from_xlsx.py and refresh.</p>';
      elEx.innerHTML = miss;
      elSc.innerHTML = miss;
      return;
    }

    var feederRows = scenarioLastFeederRows || [];
    var milesEx = collectScenarioTravelMiles(
      "existing",
      pack.rows,
      scenarioMiddleMsid,
      feederRows,
      false
    );
    var milesSc = collectScenarioTravelMiles(
      "scenario",
      pack.rows,
      scenarioMiddleMsid,
      feederRows,
      false
    );

    var milesExAxis = collectScenarioTravelMiles(
      "existing",
      pack.rows,
      scenarioMiddleMsid,
      feederRows,
      true
    );
    var milesScAxis = collectScenarioTravelMiles(
      "scenario",
      pack.rows,
      scenarioMiddleMsid,
      feederRows,
      true
    );

    var scenTitleEl = document.getElementById(
      "scenario-travel-scenario-chart-title"
    );
    var shortMid = scenarioMiddleShortDisplayName(scenarioMiddleMsid);
    if (scenTitleEl) {
      if (shortMid) {
        scenTitleEl.textContent =
          "Scenario Travel Distances to " + shortMid;
      } else {
        scenTitleEl.textContent =
          "Scenario Travel Distances to Selected Middle School";
      }
    }

    var scenarioChartTitle =
      scenTitleEl && scenTitleEl.textContent
        ? scenTitleEl.textContent
        : "Scenario Travel Distances to Selected Middle School";

    var useExAxis = travelMilesExcludingUpperOutliers(milesExAxis);
    var useScAxis = travelMilesExcludingUpperOutliers(milesScAxis);
    var pairedAxisHi = travelHistogramPairedAxisHiMiles(useExAxis, useScAxis);
    var histOpts = { axisHiOverride: pairedAxisHi };

    function paintTravelHistograms() {
      renderTravelHistogramIntoRoot(
        elEx,
        "Existing Travel Distances to Attendance School",
        milesEx,
        histOpts
      );
      renderTravelHistogramIntoRoot(
        elSc,
        scenarioChartTitle,
        milesSc,
        histOpts
      );
    }

    paintTravelHistograms();
    if (typeof requestAnimationFrame !== "undefined") {
      requestAnimationFrame(function travelHistLayoutReflow() {
        if (
          scenarioMiddleMsid == null ||
          isNaN(scenarioMiddleMsid) ||
          PRIORITY_SCHOOL_MSIDS.indexOf(scenarioMiddleMsid) < 0
        ) {
          return;
        }
        paintTravelHistograms();
      });
    }
  }

  function collectScenarioWeightedSpec() {
    var out = [];
    if (!scenarioLastFeederRows.length) {
      if (
        scenarioMiddleMsid != null &&
        !isNaN(scenarioMiddleMsid) &&
        scenarioFeederChecked[scenarioMiddleMsid] !== false
      ) {
        out.push({ msid: scenarioMiddleMsid, weight: 1 });
      }
      return out;
    }
    for (var i = 0; i < scenarioLastFeederRows.length; i++) {
      var r = scenarioLastFeederRows[i];
      if (!r.hasEnrollment || r.msid == null) continue;
      if (scenarioFeederChecked[r.msid] === false) continue;
      if (r.isScenarioMiddleRow) {
        out.push({ msid: r.msid, weight: 1 });
        continue;
      }
      var w =
        scenarioCompleteMerger
          ? 1
          : r.flowProportion != null && !isNaN(r.flowProportion)
            ? r.flowProportion
            : 1;
      out.push({ msid: r.msid, weight: w });
    }
    return out;
  }

  function applyScenarioMergedUpdates() {
    var weighted = collectScenarioWeightedSpec();
    var chartRoot = document.getElementById("scenario-enrollment-chart");
    teardownScenarioStackedChart(chartRoot);
    if (chartRoot) chartRoot.classList.remove("enrollment-chart--stacked");

    if (
      SCENARIO_USE_STACKED_ENROLLMENT_CHART &&
      scenarioSchoolByMsid &&
      scenarioMiddleMsid != null &&
      !isNaN(scenarioMiddleMsid)
    ) {
      var stacked = buildScenarioStackedPeriods(
        weighted,
        scenarioMiddleMsid,
        scenarioSchoolByMsid,
        scenarioLastFeederRows
      );
      renderScenarioStackedEnrollmentChartIntoRoot(chartRoot, stacked, {
        noDataMsg:
          "No merged enrollment series from 2025-26 onward for the current selection (check data/school_master.csv).",
        noDataAria: "Merged enrollment data is not available.",
        ariaLabel:
          "Stacked enrollment by middle school and feeder elementaries from 2025-26 forward.",
      });
    } else {
      var series = buildMergedEnrollmentSeriesWeighted(weighted);
      series = filterEnrollmentSeriesScenarioFuture(series);
      renderEnrollmentChartIntoRoot(chartRoot, series, {
        noDataMsg:
          "No merged enrollment series from 2025-26 onward for the current selection (check data/school_master.csv).",
        noDataAria: "Merged enrollment data is not available.",
        ariaLabel:
          "Merged K–8 enrollment bar chart from 2025-26 forward (scenario projection).",
      });
    }

    var ethEl = document.getElementById("scenario-demographics-ethnicity");
    var lunchEl = document.getElementById("scenario-demographics-lunch");
    if (!scenarioMiddleMsid || isNaN(scenarioMiddleMsid)) {
      if (ethEl) {
        ethEl.innerHTML =
          '<p class="demographics-pie-empty">Select a middle school to view merged demographics.</p>';
      }
      if (lunchEl) {
        lunchEl.innerHTML =
          '<p class="demographics-pie-empty">Select a middle school to view merged demographics.</p>';
      }
      var trEx = document.getElementById("scenario-travel-existing");
      var trSc = document.getElementById("scenario-travel-scenario");
      if (trEx) {
        trEx.innerHTML =
          '<p class="travel-hist-empty">Select a middle school to view travel distances.</p>';
      }
      if (trSc) {
        trSc.innerHTML =
          '<p class="travel-hist-empty">Select a middle school to view travel distances.</p>';
      }
      applyScenarioFeederMapHighlights();
      syncStudentHexLayer();
      return;
    }
    var agg = aggregateDemographicsMsidsWeighted(weighted);
    renderDemographicsFromAggregates(agg, ethEl, lunchEl);
    renderScenarioTravelImpactCharts();
    applyScenarioFeederMapHighlights();
    syncStudentHexLayer();
  }

  function updateScenarioSummaryText(middleProps) {
    var elP = document.getElementById("scenario-details-primary");
    var elS = document.getElementById("scenario-details-secondary");
    var elKpiPri = document.getElementById("scenario-details-kpi-primary");
    var elKpiCap = document.getElementById("scenario-details-kpi-capture");
    if (!middleProps || !elP) return;
    var pMerged = schoolPropsWithMasterType(middleProps);
    var msid =
      pMerged.SCHOOLS_ID != null && pMerged.SCHOOLS_ID !== ""
        ? Number(pMerged.SCHOOLS_ID)
        : null;
    var m = masterRow(msid);
    fillSchoolDetailsPrimarySecondary(pMerged, elP, elS);
    var hsScenario =
      msid != null && !isNaN(msid)
        ? countHomeschoolStudentsInAssignmentBoundary(msid)
        : 0;
    var kpi = getSchoolKpiDisplayParts(pMerged, m, msid, {
      includeHomeschoolInCaptureDenominator: true,
      homeschoolStudentsInBoundary: hsScenario,
    });
    if (elKpiPri) {
      elKpiPri.textContent =
        "'25-26 Enrollment: " +
        kpi.enrollmentStr +
        " | Factored Capacity: " +
        kpi.capacityStr +
        " | Utilization: " +
        kpi.utilizationStr;
      elKpiPri.classList.remove("school-details-placeholder");
      elKpiPri.title =
        "Key metrics from data/school_master.csv for the selected middle school.";
    }
    if (elKpiCap) {
      var capScenario =
        "Assignment: " +
        kpi.assignmentStr +
        " | Other district: " +
        kpi.otherDistrictStr +
        " | Choice: " +
        kpi.choiceStr +
        " | Charter: " +
        kpi.charterStr;
      if (!kpi.captureIsChoice) {
        capScenario += " | Homeschool: " + (kpi.homeschoolStr || "—");
      }
      elKpiCap.textContent = capScenario;
      elKpiCap.classList.remove("school-details-placeholder");
      if (kpi.scenarioCaptureCountsTitle) {
        elKpiCap.setAttribute("title", kpi.scenarioCaptureCountsTitle);
      } else {
        elKpiCap.removeAttribute("title");
      }
    }
  }

  /**
   * '25-26 enrollment and proportional assignment to the selected MS (rounded),
   * with proportional clamped so it never exceeds enrollment.
   * @returns {{ enr: number|null, propAmt: number|null }}
   */
  function scenarioFeederEnrollmentProportionalPair(r) {
    if (r.msid == null || isNaN(r.msid)) return { enr: null, propAmt: null };
    var enr = enrollment202526CalendarForMsid(r.msid);
    if (enr == null) return { enr: null, propAmt: null };
    var p =
      r.flowProportion != null && !isNaN(r.flowProportion)
        ? r.flowProportion
        : 1;
    var propAmt = scenarioCompleteMerger
      ? Math.round(enr)
      : Math.round(enr * p);
    if (propAmt > enr) {
      console.warn(
        "[Scenario] Proportional enrollment exceeds '25-26 enrollment for MSID " +
          r.msid +
          " (" +
          propAmt +
          " > " +
          enr +
          "). Clamping proportional to enrollment."
      );
      propAmt = enr;
    }
    return { enr: enr, propAmt: propAmt };
  }

  /**
   * Remaining ES enrollment text for the feeder row (uses checkbox + merger state).
   * @param {{ enr: number|null, propAmt: number|null }} [pairOpt] from scenarioFeederEnrollmentProportionalPair to avoid duplicate work
   */
  function scenarioRemainingEsEnrollmentText(r, pairOpt) {
    if (r && r.isScenarioMiddleRow) {
      return "--";
    }
    if (!r.props || !r.hasEnrollment || r.msid == null || isNaN(r.msid)) {
      console.warn(
        "[Scenario] Feeder list row has a disabled checkbox (unexpected): " +
          String(r && r.sankeyLabel ? r.sankeyLabel : "")
      );
      return "--";
    }
    if (scenarioFeederChecked[r.msid] === false) return "--";
    var pair = pairOpt || scenarioFeederEnrollmentProportionalPair(r);
    if (pair.enr == null || pair.propAmt == null) return "--";
    var remaining = Math.max(0, pair.enr - pair.propAmt);
    return remaining.toLocaleString();
  }

  /**
   * Feeder scenario column: whole-number percent (CSV decimal e.g. 0.84 → "84%").
   * @param {number} ratioZeroToOne
   * @returns {string|null}
   */
  function scenarioUtilizationPercentStringFromDecimalRatio(ratioZeroToOne) {
    if (ratioZeroToOne == null || isNaN(ratioZeroToOne) || !isFinite(ratioZeroToOne)) {
      return null;
    }
    var utilPctDisp = Math.round(ratioZeroToOne * 100);
    return String(utilPctDisp) + "%";
  }

  /**
   * Feeder scenario: whole-number percent from headcount ÷ factored capacity.
   * @returns {string|null}
   */
  function scenarioUtilizationPercentStringFromCountOverCapacity(count, cap) {
    if (count == null || isNaN(count) || cap == null || isNaN(cap) || cap <= 0) {
      return null;
    }
    var r = count / cap;
    if (!isFinite(r) || r < 0) return null;
    var utilPctDisp = Math.round(r * 100);
    return String(utilPctDisp) + "%";
  }

  /**
   * Current 2025-26 utilization: CSV utilization_2025_26, or enrollment_2025 ÷ factored_capacity_2025_26.
   * @returns {string|null}
   */
  function scenarioFeederEsCurrentUtilizationPercentString(m, msid) {
    if (!m) return null;
    if (m.utilization_2025_26 !== "" && m.utilization_2025_26 != null) {
      var d = Number(m.utilization_2025_26);
      if (!isNaN(d)) {
        return scenarioUtilizationPercentStringFromDecimalRatio(d);
      }
    }
    var enr = enrollment202526CalendarForMsid(msid);
    var cap = m.factored_capacity_2025_26;
    var capN = cap !== "" && cap != null ? Number(cap) : NaN;
    if (enr != null && !isNaN(enr) && !isNaN(capN) && capN > 0) {
      return scenarioUtilizationPercentStringFromCountOverCapacity(enr, capN);
    }
    return null;
  }

  /**
   * e.g. "84% → 37%" (Remaining ES headcount as numerator, factored capacity as denominator for the "new" value).
   * @param {{ enr: number|null, propAmt: number|null }} [pairOpt]
   */
  function scenarioFeederUtilizationChangeText(r, pairOpt) {
    if (r && r.isScenarioMiddleRow) {
      return "--";
    }
    if (!r.props || !r.hasEnrollment || r.msid == null || isNaN(r.msid)) {
      return "--";
    }
    if (scenarioFeederChecked[r.msid] === false) {
      return "--";
    }
    var m = masterRow(r.msid);
    if (!m) {
      return "--";
    }
    var currentStr = scenarioFeederEsCurrentUtilizationPercentString(
      m,
      r.msid
    );
    var pair = pairOpt || scenarioFeederEnrollmentProportionalPair(r);
    if (pair.enr == null || pair.propAmt == null) {
      return "--";
    }
    var remaining = Math.max(0, pair.enr - pair.propAmt);
    var cap = m.factored_capacity_2025_26;
    var capN = cap !== "" && cap != null ? Number(cap) : NaN;
    var newStr = scenarioUtilizationPercentStringFromCountOverCapacity(
      remaining,
      capN
    );
    if (!newStr) {
      return "--";
    }
    if (!currentStr) {
      return "--";
    }
    return currentStr + " → " + newStr;
  }

  function updateScenarioFeederRemainingCells() {
    var ul = document.getElementById("scenario-feeder-list");
    if (!ul || !scenarioLastFeederRows.length) return;
    var items = ul.querySelectorAll(".scenario-feeder-item");
    for (
      var i = 0;
      i < scenarioLastFeederRows.length && i < items.length;
      i++
    ) {
      var row = scenarioLastFeederRows[i];
      var pairP = scenarioFeederEnrollmentProportionalPair(row);
      var rem = items[i].querySelector(".scenario-feeder-remaining");
      if (rem) {
        rem.textContent = scenarioRemainingEsEnrollmentText(row, pairP);
      }
      var util = items[i].querySelector(".scenario-feeder-util-change");
      if (util) {
        util.textContent = scenarioFeederUtilizationChangeText(row, pairP);
      }
    }
  }

  function renderScenarioFeederList(middleMsid, rows) {
    var ul = document.getElementById("scenario-feeder-list");
    var alerts = document.getElementById("scenario-data-alerts");
    if (!ul) return;
    ul.innerHTML = "";
    var feederElemMsids = scenarioFeederElementaryMsidsFromRows(
      middleMsid,
      rows
    );
    var greenMap = assignElementaryFeederGreenColors(feederElemMsids);
    var warnings = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!r.props) {
        console.warn(
          '[Scenario] Feeder elementary "' +
            r.sankeyLabel +
            '" was not matched to a GeoJSON elementary school.'
        );
        warnings.push(
          "No map/school match for feeder label \"" +
            escapeHtml(r.sankeyLabel) +
            "\"."
        );
      }
      if (r.msid != null && !r.hasEnrollment) {
        console.warn(
          '[Scenario] Feeder elementary "' +
            r.sankeyLabel +
            '" (MSID ' +
            r.msid +
            ") has no enrollment row in data/school_master.csv."
        );
        warnings.push(
          "No enrollment row in data/school_master.csv for \"" +
            escapeHtml(r.sankeyLabel) +
            "\" (MSID " +
            r.msid +
            ")."
        );
      }
      var li = document.createElement("li");
      li.className = "scenario-feeder-item";
      var id = "scenario-feeder-" + middleMsid + "-" + i;
      var label = document.createElement("label");
      label.className = "scenario-feeder-label";
      var swatch = document.createElement("span");
      swatch.className = "scenario-feeder-swatch";
      swatch.setAttribute("aria-hidden", "true");
      if (r.isScenarioMiddleRow) {
        swatch.style.background = SCENARIO_STACK_MIDDLE_COLOR;
      } else if (
        r.msid != null &&
        !isNaN(r.msid) &&
        greenMap[r.msid]
      ) {
        swatch.style.background = greenMap[r.msid];
      } else {
        swatch.style.background = "#e5e7eb";
      }
      var cb = document.createElement("input");
      cb.type = "checkbox";
      cb.id = id;
      if (r.msid != null && !isNaN(r.msid)) {
        cb.dataset.msid = String(r.msid);
      }
      var displayName = r.props
        ? schoolNameForSelect(r.props)
        : r.sankeyLabel;
      if (!r.props || !r.hasEnrollment || r.msid == null) {
        cb.disabled = true;
        cb.checked = false;
      } else {
        cb.checked = scenarioFeederChecked[r.msid] !== false;
        cb.addEventListener("change", function (e) {
          var tgt = e.target;
          var ms = Number(tgt && tgt.dataset ? tgt.dataset.msid : NaN);
          if (isNaN(ms)) return;
          scenarioFeederChecked[ms] = tgt.checked;
          applyScenarioMergedUpdates();
          updateScenarioFeederRemainingCells();
        });
      }
      label.appendChild(swatch);
      label.appendChild(cb);
      var span = document.createElement("span");
      span.className = "scenario-feeder-name";
      var pairPP = scenarioFeederEnrollmentProportionalPair(r);
      var enr = pairPP.enr;
      var propAmt = pairPP.propAmt;
      var enrStr = enr != null ? enr.toLocaleString() : "—";
      var propStr = propAmt != null ? propAmt.toLocaleString() : "—";
      span.textContent =
        displayName +
        " ('25-26 enrollment: " +
        enrStr +
        "; proportional: " +
        propStr +
        ")";
      label.appendChild(span);
      li.appendChild(label);
      var remSpan = document.createElement("span");
      remSpan.className = "scenario-feeder-name scenario-feeder-remaining";
      remSpan.setAttribute(
        "aria-labelledby",
        "scenario-feeder-remaining-heading"
      );
      if (r.isScenarioMiddleRow) {
        remSpan.setAttribute(
          "title",
          "Not applicable — this column is for remaining enrollment at feeder elementary schools."
        );
      }
      remSpan.textContent = scenarioRemainingEsEnrollmentText(r, pairPP);
      var metricsWrap = document.createElement("div");
      metricsWrap.className = "scenario-feeder-item-metrics";
      metricsWrap.appendChild(remSpan);
      var utilSpan = document.createElement("span");
      utilSpan.className = "scenario-feeder-util-change";
      utilSpan.setAttribute("aria-labelledby", "scenario-feeder-util-heading");
      utilSpan.setAttribute(
        "title",
        r.isScenarioMiddleRow
          ? "Not applicable — remaining elementary enrollment and utilization change apply to feeder elementary schools only."
          : "2025-26 utilization (school_master.csv). New value: Remaining ES Enrollment ÷ factored capacity (2025-26). Percentages are rounded to whole numbers."
      );
      utilSpan.textContent = scenarioFeederUtilizationChangeText(r, pairPP);
      metricsWrap.appendChild(utilSpan);
      li.appendChild(metricsWrap);
      if (!r.props) {
        var un = document.createElement("span");
        un.className = "scenario-feeder-flag";
        un.textContent = "No school match";
        li.appendChild(un);
      } else if (!r.hasEnrollment || r.msid == null) {
        var fl = document.createElement("span");
        fl.className = "scenario-feeder-flag";
        fl.textContent = "No enrollment row";
        li.appendChild(fl);
      }
      ul.appendChild(li);
    }
    if (alerts) {
      if (warnings.length) {
        alerts.hidden = false;
        alerts.innerHTML =
          '<strong class="scenario-data-alerts-title">Data checks</strong><ul class="scenario-data-alerts-list"><li>' +
          warnings.join("</li><li>") +
          "</li></ul>";
      } else {
        alerts.hidden = true;
        alerts.innerHTML = "";
      }
    }
  }

  function resetScenarioPanel() {
    scenarioMiddleMsid = null;
    scenarioLastFeederRows = [];
    scenarioFeederChecked = {};
    scenarioCompleteMerger = false;
    var mergerCb = document.getElementById("scenario-complete-merger");
    if (mergerCb) mergerCb.checked = false;
    var p1 = document.getElementById("scenario-details-primary");
    if (p1) {
      p1.textContent = "Name of School | Grades Served | Address";
      p1.classList.add("school-details-placeholder");
    }
    var p2 = document.getElementById("scenario-details-secondary");
    if (p2) {
      p2.textContent =
        "Year Opened | Age of Site | Year of Last Major Renovation | Size of Site (Acres) | Count of On-Site BPS Employees";
      p2.classList.add("school-details-placeholder");
      p2.removeAttribute("title");
    }
    var p3a = document.getElementById("scenario-details-kpi-primary");
    if (p3a) {
      p3a.textContent = "'25-26 Enrollment: — | Factored Capacity: — | Utilization: —";
      p3a.classList.add("school-details-placeholder");
      p3a.removeAttribute("title");
    }
    var p3b = document.getElementById("scenario-details-kpi-capture");
    if (p3b) {
      p3b.textContent =
        "Assignment: — | Other district: — | Choice: — | Charter: — | Homeschool: —";
      p3b.classList.add("school-details-placeholder");
      p3b.removeAttribute("title");
    }
    var alerts = document.getElementById("scenario-data-alerts");
    if (alerts) {
      alerts.hidden = true;
      alerts.innerHTML = "";
    }
    var ul = document.getElementById("scenario-feeder-list");
    if (ul) ul.innerHTML = "";
    var chartRoot = document.getElementById("scenario-enrollment-chart");
    if (chartRoot) {
      teardownScenarioStackedChart(chartRoot);
      chartRoot.classList.remove("enrollment-chart--stacked");
      chartRoot.innerHTML =
        '<p class="enrollment-chart-empty">Select a middle school to view merged enrollment trends.</p>';
      chartRoot.removeAttribute("aria-label");
    }
    var ethEl = document.getElementById("scenario-demographics-ethnicity");
    var lunchEl = document.getElementById("scenario-demographics-lunch");
    if (ethEl) {
      ethEl.innerHTML =
        '<p class="demographics-pie-empty">Select a middle school to view merged demographics.</p>';
    }
    if (lunchEl) {
      lunchEl.innerHTML =
        '<p class="demographics-pie-empty">Select a middle school to view merged demographics.</p>';
    }
    var trExReset = document.getElementById("scenario-travel-existing");
    var trScReset = document.getElementById("scenario-travel-scenario");
    if (trExReset) {
      trExReset.innerHTML =
        '<p class="travel-hist-empty">Select a middle school to view travel distances.</p>';
    }
    if (trScReset) {
      trScReset.innerHTML =
        '<p class="travel-hist-empty">Select a middle school to view travel distances.</p>';
    }
    var scTrTitle = document.getElementById(
      "scenario-travel-scenario-chart-title"
    );
    if (scTrTitle) {
      scTrTitle.textContent =
        "Scenario Travel Distances to Selected Middle School";
    }
    applyScenarioFeederMapHighlights();
    syncStudentHexLayer();
    syncTravelShedLayerFilter();
  }

  function runScenarioForMiddleMsid(msid, schoolByMsid, schoolsFc) {
    scenarioSchoolByMsid = schoolByMsid;
    scenarioMiddleMsid = msid;
    scenarioFeederChecked = {};
    var p = schoolByMsid[msid];
    if (!p) return;
    var flows = SANKEY_CACHE && SANKEY_CACHE.flows ? SANKEY_CACHE.flows : [];
    scenarioLastFeederRows = getFeederElementaryRowsForMiddle(
      p,
      flows,
      schoolsFc
    ).concat([buildScenarioMiddleFeederRow(p, msid)]);
    for (var i = 0; i < scenarioLastFeederRows.length; i++) {
      var r = scenarioLastFeederRows[i];
      if (r.hasEnrollment && r.msid != null) {
        scenarioFeederChecked[r.msid] = true;
      }
    }
    updateScenarioSummaryText(p);
    renderScenarioFeederList(msid, scenarioLastFeederRows);
    applyScenarioMergedUpdates();
    applySelectedSchoolHighlight(msid);
    zoomToSchoolAssignment(msid, schoolByMsid);
    syncStudentHexLayer();
    syncTravelShedLayerFilter();
  }

  function populateScenarioSchoolSelect(schoolsFc) {
    var sel = document.getElementById("scenario-school-select");
    if (!sel || !schoolsFc || !schoolsFc.features) return;
    sel.innerHTML = "";
    var placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Select a middle school";
    sel.appendChild(placeholder);
    var byId = {};
    schoolsFc.features.forEach(function (ft) {
      var pr = ft.properties;
      if (pr && pr.SCHOOLS_ID != null) byId[pr.SCHOOLS_ID] = pr;
    });
    PRIORITY_SCHOOL_MSIDS.forEach(function (msid) {
      var pr = byId[msid];
      if (!pr) return;
      var opt = document.createElement("option");
      opt.value = String(msid);
      opt.textContent = schoolNameForSelect(pr);
      sel.appendChild(opt);
    });
    sel.value = "";
    sel.disabled = false;
  }

  function setupScenarioSchoolSelection(schoolByMsid, schoolsFc) {
    scenarioSchoolByMsid = schoolByMsid;
    var sel = document.getElementById("scenario-school-select");
    if (!sel) return;
    sel.addEventListener("change", function () {
      var v = sel.value;
      if (!v) {
        clearSelectedSchoolHighlight();
        resetScenarioPanel();
        return;
      }
      var msid = Number(v);
      if (isNaN(msid)) return;
      if (PRIORITY_SCHOOL_MSIDS.indexOf(msid) < 0) return;
      runScenarioForMiddleMsid(msid, schoolByMsid, schoolsFc);
    });
  }

  function refreshScenarioPanelIfVisible() {
    var panel = document.getElementById("page-scenario");
    if (!panel || panel.hidden) return;
    if (scenarioMiddleMsid != null && !isNaN(scenarioMiddleMsid)) {
      applyScenarioMergedUpdates();
    }
  }

  /**
   * Fills the two school detail lines (name | grades | address; opened | age | renovation | acres) from GeoJSON + master CSV.
   * @param {Object} p school feature properties
   * @param {HTMLElement|null} elP primary line
   * @param {HTMLElement|null} elS secondary line
   */
  function fillSchoolDetailsPrimarySecondary(p, elP, elS) {
    if (!p) return;
    var msid =
      p.SCHOOLS_ID != null && p.SCHOOLS_ID !== ""
        ? Number(p.SCHOOLS_ID)
        : null;
    var m = masterRow(msid);

    if (elP) {
      var name = schoolDisplayNameFromProps(p);
      var grades = m
        ? m.grades_served
          ? standardCapitalization(normalizeGradesServedForUi(m.grades_served))
          : "—"
        : p.Grades
          ? standardCapitalization(normalizeGradesServedForUi(p.Grades))
          : "—";
      var addrLine = "—";
      if (m) {
        var sa = m.address ? standardCapitalization(m.address) : "";
        var sc = m.city_state_zip ? formatCityStateZip(m.city_state_zip) : "";
        if (sa && sc) addrLine = sa + ", " + sc;
        else if (sa) addrLine = sa;
        else if (sc) addrLine = sc;
      } else if (p.ADDRESS) {
        addrLine = standardCapitalization(p.ADDRESS);
      }
      elP.textContent = [name, grades, addrLine].join(" | ");
      elP.classList.remove("school-details-placeholder");
    }
    if (elS) {
      var acres =
        m && m.site_acres !== "" && m.site_acres != null
          ? String(m.site_acres)
          : "—";
      var openedYearRaw = "";
      if (m) {
        if (m.opened_year !== "" && m.opened_year != null) {
          openedYearRaw = m.opened_year;
        } else if (m.constructed_year !== "" && m.constructed_year != null) {
          openedYearRaw = m.constructed_year;
        }
      }
      var opened =
        openedYearRaw !== "" && !isNaN(Number(openedYearRaw))
          ? String(openedYearRaw)
          : "—";
      var age =
        m &&
        m.age_of_site_2026 !== "" &&
        m.age_of_site_2026 != null &&
        !isNaN(Number(m.age_of_site_2026))
          ? String(m.age_of_site_2026)
          : "—";
      var renovation = "—";
      if (
        m &&
        m.last_major_renovation_year !== "" &&
        m.last_major_renovation_year != null
      ) {
        renovation = String(m.last_major_renovation_year).trim();
        if (renovation === "No Major Renovation") renovation = "N/A";
      }
      var bpsEmployees = bpsOnSiteEmployeeCountDisplay(msid);
      elS.textContent =
        "Opened: " +
        opened +
        " | Age of Site: " +
        age +
        " | Year of Last Major Renovation: " +
        renovation +
        " | Size of Site (Acres): " +
        acres +
        " | Count of On-Site BPS Employees: " +
        bpsEmployees;
      elS.classList.remove("school-details-placeholder");
      elS.removeAttribute("title");
    }
  }

  function fromToResidentDenominatorForMaster(m) {
    if (!m || m.fromto_resident_denominator === "" || m.fromto_resident_denominator == null) {
      return NaN;
    }
    var d = parseInt(String(m.fromto_resident_denominator).trim(), 10);
    if (isNaN(d) || d <= 0) return NaN;
    return d;
  }

  function fromToStudentCount(m, key) {
    if (!m || m[key] === "" || m[key] == null) return NaN;
    var n = parseInt(String(m[key]).trim(), 10);
    return isNaN(n) ? NaN : n;
  }

  /** Tooltip text for From-To capture: "N of D boundary public-school students", or null if counts missing. */
  function boundaryPublicSchoolStudentsPhrase(m, countKey) {
    var den = fromToResidentDenominatorForMaster(m);
    var num = fromToStudentCount(m, countKey);
    if (!isNaN(den) && !isNaN(num)) {
      return (
        num.toLocaleString() +
        " of " +
        den.toLocaleString() +
        " boundary public-school students"
      );
    }
    return null;
  }

  /**
   * Tooltip line when capture denominators use an adjusted total (e.g. including homeschool residents).
   */
  function boundaryStudentsPhraseAdjusted(m, countKey, denominatorAdjusted, homeschoolStudentCount) {
    var num = fromToStudentCount(m, countKey);
    var den = denominatorAdjusted;
    if (isNaN(den) || den <= 0 || isNaN(num)) {
      return null;
    }
    var s =
      num.toLocaleString() +
      " of " +
      den.toLocaleString() +
      " students residing in the attendance boundary";
    if (homeschoolStudentCount != null && homeschoolStudentCount > 0) {
      s +=
        " (denominator includes " +
        Number(homeschoolStudentCount).toLocaleString() +
        " grade-eligible homeschool students)";
    }
    return s;
  }

  /**
   * Shared display strings for KPI cards and scenario summary line (same rules).
   * From-To capture decimals: assignment_capture_rate, other_district_capture_rate, choice_capture_rate, charter_capture_rate.
   * @param {Object|null} captureOpts optional; when `includeHomeschoolInCaptureDenominator` and `homeschoolStudentsInBoundary` are set, recomputes % from CSV numerators over expanded denominator.
   * @returns {Object}
   */
  function getSchoolKpiDisplayParts(p, m, msid, captureOpts) {
    var enrollmentStr = "—";
    if (m && m.enrollment_2025 !== "" && m.enrollment_2025 != null) {
      var ev = Number(m.enrollment_2025);
      if (!isNaN(ev)) enrollmentStr = ev.toLocaleString();
    }

    var capacityStr = "—";
    if (
      m &&
      m.factored_capacity_2025_26 !== "" &&
      m.factored_capacity_2025_26 != null
    ) {
      var cn = Number(m.factored_capacity_2025_26);
      if (!isNaN(cn)) capacityStr = cn.toLocaleString();
    }

    var utilizationStr = "—";
    if (m && m.utilization_2025_26 !== "" && m.utilization_2025_26 != null) {
      var utilDec = Number(m.utilization_2025_26);
      if (!isNaN(utilDec)) {
        var utilPctDisp = utilDec * 100;
        utilizationStr =
          (utilPctDisp % 1 === 0 ? String(utilPctDisp) : utilPctDisp.toFixed(1)) +
          "%";
      }
    }

    var captureIsChoice = schoolIsChoiceFromProps(p);
    var choiceNaStr = "N/A (Choice School)";
    var choiceNaTitle =
      "Choice schools have no assignment-area residence row in the From-To analysis; these capture rates do not apply.";

    function pctFromCsvDecimal(raw, title) {
      if (raw === "" || raw == null) {
        return { str: "—", title: null };
      }
      var d = Number(raw);
      if (isNaN(d)) {
        return { str: "—", title: null };
      }
      var pctDisp = d * 100;
      var str =
        (pctDisp % 1 === 0 ? String(pctDisp) : pctDisp.toFixed(1)) + "%";
      return { str: str, title: title };
    }

    var assignmentStr = "—";
    var otherDistrictStr = "—";
    var choiceStr = "—";
    var charterStr = "—";
    var assignmentTitle = null;
    var otherDistrictTitle = null;
    var choiceTitle = null;
    var charterTitle = null;
    var captureHoverAssignment = null;
    var captureHoverOtherDistrict = null;
    var captureHoverChoice = null;
    var captureHoverCharter = null;
    var homeschoolStr = "—";
    var homeschoolTitle = null;
    var captureHoverHomeschool = null;
    var scenarioCaptureCountsTitle = null;

    if (captureIsChoice) {
      assignmentStr = choiceNaStr;
      otherDistrictStr = choiceNaStr;
      choiceStr = choiceNaStr;
      charterStr = choiceNaStr;
      assignmentTitle = choiceNaTitle;
      otherDistrictTitle = choiceNaTitle;
      choiceTitle = choiceNaTitle;
      charterTitle = choiceNaTitle;
      captureHoverAssignment = null;
      captureHoverOtherDistrict = null;
      captureHoverChoice = null;
      captureHoverCharter = null;
      scenarioCaptureCountsTitle = null;
    } else if (m) {
      var useHsDen =
        captureOpts &&
        captureOpts.includeHomeschoolInCaptureDenominator === true;
      var Hhs =
        useHsDen && typeof captureOpts.homeschoolStudentsInBoundary === "number"
          ? Math.max(0, Math.floor(Number(captureOpts.homeschoolStudentsInBoundary)))
          : 0;
      var Dbase = fromToResidentDenominatorForMaster(m);

      function pctStrFromCounts(num, den) {
        if (isNaN(num) || isNaN(den) || den <= 0) {
          return "—";
        }
        var pctDisp = (num / den) * 100;
        return (
          (pctDisp % 1 === 0 ? String(pctDisp) : pctDisp.toFixed(1)) + "%"
        );
      }

      if (useHsDen && !isNaN(Dbase) && Dbase > 0) {
        var Dadj = Dbase + Hhs;
        var na = fromToStudentCount(m, "assignment_capture_students");
        var no = fromToStudentCount(m, "other_district_capture_students");
        var nc = fromToStudentCount(m, "choice_capture_students");
        var nv = fromToStudentCount(m, "charter_capture_students");

        assignmentStr = pctStrFromCounts(na, Dadj);
        otherDistrictStr = pctStrFromCounts(no, Dadj);
        choiceStr = pctStrFromCounts(nc, Dadj);
        charterStr = pctStrFromCounts(nv, Dadj);
        homeschoolStr = pctStrFromCounts(Hhs, Dadj);

        captureHoverAssignment = boundaryStudentsPhraseAdjusted(
          m,
          "assignment_capture_students",
          Dadj,
          Hhs
        );
        captureHoverOtherDistrict = boundaryStudentsPhraseAdjusted(
          m,
          "other_district_capture_students",
          Dadj,
          Hhs
        );
        captureHoverChoice = boundaryStudentsPhraseAdjusted(
          m,
          "choice_capture_students",
          Dadj,
          Hhs
        );
        captureHoverCharter = boundaryStudentsPhraseAdjusted(
          m,
          "charter_capture_students",
          Dadj,
          Hhs
        );
        if (!isNaN(Hhs) && !isNaN(Dadj)) {
          captureHoverHomeschool =
            Hhs.toLocaleString() +
            " of " +
            Dadj.toLocaleString() +
            " students residing in the attendance boundary (grade-eligible homeschool students)";
        }

        var countKeysAdj = [
          "assignment_capture_students",
          "other_district_capture_students",
          "choice_capture_students",
          "charter_capture_students",
        ];
        var partsCtAdj = [];
        for (var cj = 0; cj < countKeysAdj.length; cj++) {
          var phA = boundaryStudentsPhraseAdjusted(
            m,
            countKeysAdj[cj],
            Dadj,
            Hhs
          );
          if (phA) {
            partsCtAdj.push(phA);
          }
        }
        if (captureHoverHomeschool) {
          partsCtAdj.push(captureHoverHomeschool);
        }
        if (partsCtAdj.length) {
          scenarioCaptureCountsTitle = partsCtAdj.join(" · ");
        }
      } else {
        var a = pctFromCsvDecimal(m.assignment_capture_rate, null);
        assignmentStr = a.str;
        assignmentTitle = a.title;
        var o = pctFromCsvDecimal(m.other_district_capture_rate, null);
        otherDistrictStr = o.str;
        otherDistrictTitle = o.title;
        var ch = pctFromCsvDecimal(m.choice_capture_rate, null);
        choiceStr = ch.str;
        choiceTitle = ch.title;
        var chrt = pctFromCsvDecimal(m.charter_capture_rate, null);
        charterStr = chrt.str;
        charterTitle = chrt.title;

        captureHoverAssignment = boundaryPublicSchoolStudentsPhrase(
          m,
          "assignment_capture_students"
        );
        captureHoverOtherDistrict = boundaryPublicSchoolStudentsPhrase(
          m,
          "other_district_capture_students"
        );
        captureHoverChoice = boundaryPublicSchoolStudentsPhrase(m, "choice_capture_students");
        captureHoverCharter = boundaryPublicSchoolStudentsPhrase(
          m,
          "charter_capture_students"
        );

        var countKeys = [
          "assignment_capture_students",
          "other_district_capture_students",
          "choice_capture_students",
          "charter_capture_students",
        ];
        var partsCt = [];
        for (var ci = 0; ci < countKeys.length; ci++) {
          var phrase = boundaryPublicSchoolStudentsPhrase(m, countKeys[ci]);
          if (phrase) {
            partsCt.push(phrase);
          }
        }
        if (partsCt.length) {
          scenarioCaptureCountsTitle = partsCt.join(" · ");
        }
      }
    }

    return {
      enrollmentStr: enrollmentStr,
      capacityStr: capacityStr,
      utilizationStr: utilizationStr,
      assignmentStr: assignmentStr,
      otherDistrictStr: otherDistrictStr,
      choiceStr: choiceStr,
      charterStr: charterStr,
      captureIsChoice: captureIsChoice,
      assignmentTitle: assignmentTitle,
      otherDistrictTitle: otherDistrictTitle,
      choiceTitle: choiceTitle,
      charterTitle: charterTitle,
      captureHoverAssignment: captureHoverAssignment,
      captureHoverOtherDistrict: captureHoverOtherDistrict,
      captureHoverChoice: captureHoverChoice,
      captureHoverCharter: captureHoverCharter,
      homeschoolStr: homeschoolStr,
      homeschoolTitle: homeschoolTitle,
      captureHoverHomeschool: captureHoverHomeschool,
      scenarioCaptureCountsTitle: scenarioCaptureCountsTitle,
      /** @deprecated scenario line — use assignmentStr */
      captureStr: assignmentStr,
      /** @deprecated scenario line — use charterStr */
      charterStr: charterStr,
      captureTitle: assignmentTitle,
      charterTitle: charterTitle,
    };
  }

  /**
   * Fills #ese-feeder-tbody from ESE_FEEDER_MATRIX for the selected school MSID.
   * Row labels use short titles; title attribute carries full Excel column header text.
   */
  function renderEseFeederFlowsTable(msid) {
    var tbody = document.getElementById("ese-feeder-tbody");
    if (!tbody) return;

    function rowPlaceholder(msg) {
      tbody.innerHTML = "";
      var tr = document.createElement("tr");
      var td = document.createElement("td");
      td.colSpan = 3;
      td.className = "ese-feeder-placeholder";
      td.textContent = msg;
      tr.appendChild(td);
      tbody.appendChild(tr);
    }

    if (!ESE_FEEDER_MATRIX || !ESE_FEEDER_MATRIX.programs || !ESE_FEEDER_MATRIX.programs.length) {
      rowPlaceholder(
        "ESE feeder matrix could not be loaded (data/processed/ese_feeder_matrix.json)."
      );
      return;
    }

    if (msid == null || isNaN(Number(msid))) {
      rowPlaceholder("Select a school to view ESE feeder flows.");
      return;
    }

    var sidStr = String(Number(msid));
    var rowMap = ESE_FEEDER_MATRIX.rows ? ESE_FEEDER_MATRIX.rows[sidStr] : null;
    var accAll =
      ESE_FEEDER_MATRIX.acceptsFrom && ESE_FEEDER_MATRIX.acceptsFrom[sidStr]
        ? ESE_FEEDER_MATRIX.acceptsFrom[sidStr]
        : {};

    tbody.innerHTML = "";
    ESE_FEEDER_MATRIX.programs.forEach(function (prog) {
      var tr = document.createElement("tr");
      var tdLabel = document.createElement("th");
      tdLabel.scope = "row";
      tdLabel.className = "ese-feeder-program-label";
      tdLabel.textContent = prog.shortLabel || prog.key;
      if (prog.headerFull) {
        tdLabel.title = String(prog.headerFull).replace(/\s*\n\s*/g, " ").trim();
      }
      var tdAccept = document.createElement("td");
      var tdSend = document.createElement("td");
      var acc = accAll[prog.key] || [];
      var sends = rowMap && rowMap[prog.key] ? rowMap[prog.key] : [];
      var acceptNames = eseFilteredSortedSchoolNames(acc, msid);
      var sendNames = eseFilteredSortedSchoolNames(sends, msid);
      tdAccept.textContent = acceptNames.length ? acceptNames.join(", ") : "—";
      tdSend.textContent = sendNames.length ? sendNames.join(", ") : "—";
      tr.appendChild(tdLabel);
      tr.appendChild(tdAccept);
      tr.appendChild(tdSend);
      tbody.appendChild(tr);
    });
  }

  function updateLeftPanelFromSchool(p) {
    var elP = document.getElementById("school-details-primary");
    var elS = document.getElementById("school-details-secondary");
    var msid =
      p.SCHOOLS_ID != null && p.SCHOOLS_ID !== ""
        ? Number(p.SCHOOLS_ID)
        : null;
    var m = masterRow(msid);

    fillSchoolDetailsPrimarySecondary(p, elP, elS);

    var hsCapCb = document.getElementById("toggle-include-homeschool-capture");
    var includeHsCapture = !!(hsCapCb && hsCapCb.checked);
    var hsInBoundary =
      includeHsCapture && msid != null && !isNaN(msid)
        ? countHomeschoolStudentsInAssignmentBoundary(msid)
        : 0;
    var parts = getSchoolKpiDisplayParts(p, m, msid, {
      includeHomeschoolInCaptureDenominator: includeHsCapture,
      homeschoolStudentsInBoundary: hsInBoundary,
    });

    var capEl = document.getElementById("kpi-capacity");
    if (capEl) {
      if (parts.capacityStr !== "—") {
        capEl.textContent = parts.capacityStr;
        capEl.classList.remove("kpi-value--placeholder");
        capEl.title = "Includes capacity from portables.";
      } else {
        capEl.textContent = "—";
        capEl.classList.add("kpi-value--placeholder");
        capEl.removeAttribute("title");
      }
    }

    var enrollEl = document.getElementById("kpi-enrollment");
    if (enrollEl) {
      if (parts.enrollmentStr !== "—") {
        enrollEl.textContent = parts.enrollmentStr;
        enrollEl.classList.remove("kpi-value--placeholder");
        enrollEl.title = "2025 calendar-year membership from data/school_master.csv.";
      } else {
        enrollEl.textContent = "—";
        enrollEl.classList.add("kpi-value--placeholder");
        enrollEl.removeAttribute("title");
      }
    }

    var utilEl = document.getElementById("kpi-utilization");
    if (utilEl) {
      if (parts.utilizationStr !== "—") {
        utilEl.textContent = parts.utilizationStr;
        utilEl.classList.remove("kpi-value--placeholder");
        utilEl.title = "'25-26 Enrollment by Factored Capacity";
      } else {
        utilEl.textContent = "—";
        utilEl.classList.add("kpi-value--placeholder");
        utilEl.removeAttribute("title");
      }
    }

    renderEnrollmentChart(msid);
    renderDemographicsCharts(msid);
    renderSankeyPanel(schoolPropsWithMasterType(p));
    renderEseFeederFlowsTable(msid);

    var capAssignedLbl = document.getElementById("kpi-capture-assigned-label");
    if (capAssignedLbl) {
      capAssignedLbl.textContent = captureRateAssignedSchoolLabelUpper(p);
    }

    function applyCaptureKpi(el, displayStr, cardHoverTitle, captureIsChoice) {
      if (!el) return;
      var card = el.closest ? el.closest(".kpi-card") : null;
      el.removeAttribute("title");
      el.classList.remove("kpi-value--choice-na");
      if (captureIsChoice) {
        el.textContent = displayStr;
        el.classList.remove("kpi-value--placeholder");
        el.classList.add("kpi-value--choice-na");
        if (card) {
          if (cardHoverTitle) {
            card.setAttribute("title", cardHoverTitle);
          } else {
            card.removeAttribute("title");
          }
        }
        return;
      }
      if (displayStr !== "—") {
        el.textContent = displayStr;
        el.classList.remove("kpi-value--placeholder");
      } else {
        el.textContent = "—";
        el.classList.add("kpi-value--placeholder");
      }
      if (card) {
        if (cardHoverTitle) {
          card.setAttribute("title", cardHoverTitle);
        } else {
          card.removeAttribute("title");
        }
      }
    }

    applyCaptureKpi(
      document.getElementById("kpi-assignment-capture"),
      parts.assignmentStr,
      parts.captureHoverAssignment,
      parts.captureIsChoice
    );
    applyCaptureKpi(
      document.getElementById("kpi-other-district-capture"),
      parts.otherDistrictStr,
      parts.captureHoverOtherDistrict,
      parts.captureIsChoice
    );
    applyCaptureKpi(
      document.getElementById("kpi-choice-capture"),
      parts.choiceStr,
      parts.captureHoverChoice,
      parts.captureIsChoice
    );
    applyCaptureKpi(
      document.getElementById("kpi-charter-capture"),
      parts.charterStr,
      parts.captureHoverCharter,
      parts.captureIsChoice
    );

    var gridCap = document.getElementById("kpi-grid-capture");
    var cardHs = document.getElementById("kpi-card-homeschool-capture");
    var showHsCard = includeHsCapture && !parts.captureIsChoice;
    if (gridCap) {
      gridCap.classList.toggle("kpi-grid--capture--five", !!showHsCard);
    }
    if (cardHs) {
      cardHs.hidden = !showHsCard;
    }
    var elHsCap = document.getElementById("kpi-homeschool-capture");
    if (elHsCap) {
      if (showHsCard && parts.homeschoolStr != null && parts.homeschoolStr !== "—") {
        elHsCap.textContent = parts.homeschoolStr;
        elHsCap.classList.remove("kpi-value--placeholder");
        if (parts.captureHoverHomeschool) {
          elHsCap.removeAttribute("title");
          cardHs.setAttribute("title", parts.captureHoverHomeschool);
        } else {
          cardHs.removeAttribute("title");
        }
      } else if (showHsCard) {
        elHsCap.textContent = parts.homeschoolStr || "—";
        elHsCap.classList.toggle(
          "kpi-value--placeholder",
          !parts.homeschoolStr || parts.homeschoolStr === "—"
        );
        cardHs.removeAttribute("title");
      } else {
        elHsCap.textContent = "—";
        elHsCap.classList.add("kpi-value--placeholder");
        elHsCap.removeAttribute("title");
        if (cardHs) {
          cardHs.removeAttribute("title");
        }
      }
    }
  }

  function resetLeftPanelPlaceholders() {
    var elP = document.getElementById("school-details-primary");
    var elS = document.getElementById("school-details-secondary");
    if (elP) {
      elP.textContent = "Name of School | Grades Served | Address";
      elP.classList.add("school-details-placeholder");
    }
    if (elS) {
      elS.textContent =
        "Year Opened | Age of Site | Year of Last Major Renovation | Size of Site (Acres) | Count of On-Site BPS Employees";
      elS.classList.add("school-details-placeholder");
      elS.removeAttribute("title");
    }
    var capAsgLbl = document.getElementById("kpi-capture-assigned-label");
    if (capAsgLbl) {
      capAsgLbl.textContent = "Selected School";
    }
    [
      "kpi-enrollment",
      "kpi-capacity",
      "kpi-utilization",
      "kpi-assignment-capture",
      "kpi-other-district-capture",
      "kpi-choice-capture",
      "kpi-charter-capture",
      "kpi-homeschool-capture",
    ].forEach(function (id) {
      var k = document.getElementById(id);
      if (k) {
        k.textContent = "—";
        k.classList.add("kpi-value--placeholder");
        k.classList.remove("kpi-value--choice-na");
        k.removeAttribute("title");
        var card = k.closest && k.closest(".kpi-card");
        if (card) {
          card.removeAttribute("title");
        }
      }
    });
    var gridCapReset = document.getElementById("kpi-grid-capture");
    if (gridCapReset) {
      gridCapReset.classList.remove("kpi-grid--capture--five");
    }
    var cardHsReset = document.getElementById("kpi-card-homeschool-capture");
    if (cardHsReset) {
      cardHsReset.hidden = true;
    }
    renderEnrollmentChart(null);
    renderDemographicsCharts(null);
    renderSankeyPanel(null);
    renderEseFeederFlowsTable(null);
  }

  /**
   * Applies the current #school-select value: highlight, map frame, left panel, student hex.
   * If `pendingMapSelectFrame` is "centerOnSchool" (set just before a map point/parcel pick), pans to the
   * school with no zoom change. Otherwise (dropdown or map boundary) uses `zoomToSchoolAssignment`.
   */
  function applyExistingSchoolFromSelectValue(schoolByMsid) {
    var sel = document.getElementById("school-select");
    if (!sel) return;
    var v = sel.value;
    if (!v) {
      pendingMapSelectFrame = null;
      clearSelectedSchoolHighlight();
      resetLeftPanelPlaceholders();
      syncStudentHexLayer();
      syncTravelShedLayerFilter();
      return;
    }
    var msid = Number(v);
    if (isNaN(msid)) return;
    var p = schoolByMsid[msid];
    if (!p) return;
    var mapFrame = pendingMapSelectFrame;
    pendingMapSelectFrame = null;
    if (mapFrame !== "centerOnSchool" && mapFrame !== "assignment") {
      mapFrame = "assignment";
    }
    applySelectedSchoolHighlight(msid);
    if (mapFrame === "centerOnSchool") {
      centerMapOnSchoolPoint(msid, schoolByMsid);
    } else {
      zoomToSchoolAssignment(msid, schoolByMsid);
    }
    updateLeftPanelFromSchool(p);
    syncStudentHexLayer();
    syncTravelShedLayerFilter();
  }

  function setupSchoolSelection(schoolByMsid) {
    var sel = document.getElementById("school-select");
    if (!sel) return;

    sel.addEventListener("change", function () {
      applyExistingSchoolFromSelectValue(schoolByMsid);
    });

    var hsCapStorageKey = "brevardK8IncludeHomeschoolCapture";
    var hsCapCb = document.getElementById("toggle-include-homeschool-capture");
    if (hsCapCb) {
      try {
        if (sessionStorage.getItem(hsCapStorageKey) === "1") {
          hsCapCb.checked = true;
        }
      } catch (eSs) {
        /* ignore */
      }
      hsCapCb.addEventListener("change", function () {
        try {
          sessionStorage.setItem(hsCapStorageKey, hsCapCb.checked ? "1" : "0");
        } catch (eS2) {
          /* ignore */
        }
        var v = sel.value;
        if (!v) {
          return;
        }
        var mid = Number(v);
        if (isNaN(mid)) {
          return;
        }
        var sp = schoolByMsid[mid];
        if (sp) {
          updateLeftPanelFromSchool(sp);
        }
      });
    }
  }

  function setupMapInteractions(schoolByMsid) {
    var boundaryHoverPopup = new mapboxgl.Popup({
      closeButton: false,
      closeOnClick: false,
      maxWidth: "260px",
      className: "boundary-hover-popup",
      offset: 12,
    });

    var schoolHoverPopup = new mapboxgl.Popup({
      closeButton: false,
      closeOnClick: false,
      maxWidth: "300px",
      className: "school-hover-popup",
      offset: 10,
    });

    var schoolBoardHoverPopup = new mapboxgl.Popup({
      closeButton: false,
      closeOnClick: false,
      maxWidth: "260px",
      className: "school-board-hover-popup",
      offset: 12,
    });

    var studentHexHoverPopup = new mapboxgl.Popup({
      closeButton: false,
      closeOnClick: false,
      maxWidth: "320px",
      className: "student-hex-hover-popup",
      offset: 10,
    });

    var travelShedHoverPopup = new mapboxgl.Popup({
      closeButton: false,
      closeOnClick: false,
      maxWidth: "430px",
      className: "travel-shed-hover-popup",
      offset: 8,
    });

    var ttDensityCb = document.getElementById("toggle-student-hex-density-tooltip");
    if (ttDensityCb) {
      ttDensityCb.addEventListener("change", function () {
        if (!ttDensityCb.checked) {
          studentHexHoverPopup.remove();
        }
      });
    }

    dismissStudentHexDensityTooltip = function () {
      try {
        studentHexHoverPopup.remove();
      } catch (eRm) {
        /* ignore */
      }
    };
    syncStudentHexTooltipCheckboxVisibility();

    var lastRingMsid = null;
    var lastOutline = { source: null, id: null };

    function clearOutlineHighlight() {
      if (lastOutline.source != null && lastOutline.id != null) {
        try {
          map.setFeatureState(
            { source: lastOutline.source, id: lastOutline.id },
            { highlight: false }
          );
        } catch (e) {
          /* ignore */
        }
      }
      lastOutline.source = null;
      lastOutline.id = null;
    }

    /** Clears hover ring only; dropdown selection uses feature-state "selected". */
    function clearHoverRing() {
      if (lastRingMsid != null) {
        try {
          map.setFeatureState({ source: "schools", id: lastRingMsid }, { ring: false });
        } catch (e) {
          /* ignore */
        }
        lastRingMsid = null;
      }
    }

    function clearMunicipalHoverStroke() {
      try {
        if (map.getLayer("municipal-boundaries-hover")) {
          map.setFilter("municipal-boundaries-hover", [
            "==",
            ["to-string", ["get", "OBJECTID"]],
            MUN_HOVER_FILTER_OFF,
          ]);
        }
      } catch (errM) {
        try {
          if (map.getLayer("municipal-boundaries-hover")) {
            map.setFilter("municipal-boundaries-hover", [
              "==",
              ["to-string", ["get", "OBJECTID"]],
              MUN_HOVER_FILTER_OFF,
            ]);
          }
        } catch (errM2) {
          /* ignore */
        }
      }
    }

    function applyMunicipalHoverStroke(feature) {
      if (!feature) return;
      var props = feature.properties || {};
      var oidRaw =
        props.OBJECTID != null
          ? props.OBJECTID
          : props.objectid != null
            ? props.objectid
            : feature.id;
      if (oidRaw == null) return;
      /** Compare as strings so source + query features match regardless of number vs string. */
      var oidKey = String(oidRaw).trim();
      if (!oidKey) return;
      try {
        map.setFilter("municipal-boundaries-hover", [
          "==",
          ["to-string", ["get", "OBJECTID"]],
          oidKey,
        ]);
      } catch (errA) {
        try {
          map.setFilter("municipal-boundaries-hover", [
            "==",
            ["get", "OBJECTID"],
            oidRaw,
          ]);
        } catch (errB) {
          /* ignore */
        }
      }
    }

    function clearBoundaryHoverUi() {
      clearTravelShedResidenceDebounce();
      clearOutlineHighlight();
      clearHoverRing();
      clearMunicipalHoverStroke();
      boundaryHoverPopup.remove();
      schoolBoardHoverPopup.remove();
      studentHexHoverPopup.remove();
      travelShedHoverPopup.remove();
      map.getCanvas().style.cursor = "";
    }

    function clearSchoolHoverUi() {
      clearTravelShedResidenceDebounce();
      schoolHoverPopup.remove();
      studentHexHoverPopup.remove();
      travelShedHoverPopup.remove();
    }

    /**
     * When hovering a school location, show only that school's assignment outline (hover highlight)
     * if the matching es/ms/hs fill is on. Temporarily clears the dropdown-selected assignment outline
     * so only the hovered assignment is visible; call refreshAssignmentBoundaryHighlight when leaving
     * (invalid msid, no zoned area, or layer off) to restore the selection.
     */
    function setAssignmentHoverHighlightForSchoolMsid(msid) {
      if (msid == null || isNaN(msid)) {
        clearOutlineHighlight();
        clearHoverRing();
        refreshAssignmentBoundaryHighlight();
        return;
      }
      var src = findBoundarySourceForMsid(msid);
      if (!src || !boundaryFillVisibleForSource(src)) {
        clearOutlineHighlight();
        clearHoverRing();
        refreshAssignmentBoundaryHighlight();
        return;
      }
      clearSelectedAssignmentBoundary();
      if (lastOutline.source !== src || lastOutline.id !== msid) {
        clearOutlineHighlight();
        lastOutline.source = src;
        lastOutline.id = msid;
        try {
          map.setFeatureState({ source: src, id: msid }, { highlight: true });
        } catch (eH) {
          /* ignore */
        }
      }
      if (schoolByMsid[msid]) {
        if (lastRingMsid !== msid) {
          clearHoverRing();
          lastRingMsid = msid;
          try {
            map.setFeatureState({ source: "schools", id: msid }, { ring: true });
          } catch (eR) {
            /* ignore */
          }
        }
      } else {
        clearHoverRing();
      }
    }

    function boundaryTitleText(props) {
      var msid = props.MSID != null ? Number(props.MSID) : null;
      var raw;
      if (msid != null && !isNaN(msid) && schoolByMsid[msid]) {
        var sp = schoolByMsid[msid];
        var fromMaster = schoolDisplayNamePreferMaster(sp);
        if (fromMaster) return fromMaster;
        raw = sp.NAME || sp.CommonName || String(msid);
      } else {
        raw =
          props.Elem_Commo ||
          props.Middle_Com ||
          props.High_Commo ||
          "Assignment area";
      }
      return formatSchoolDisplayName(
        standardCapitalization(expandElemSchoolName(raw))
      );
    }

    function schoolBoardDistrictHtml(props) {
      var rawName = props && props.NAME != null ? String(props.NAME) : "";
      var rawMember = props && props.SchBoardMe != null ? String(props.SchBoardMe) : "";
      var name = escapeHtml(standardCapitalization(rawName || "District"));
      var member = rawMember ? escapeHtml(standardCapitalization(rawMember)) : "";
      return (
        '<div class="school-board-hover-inner">' +
        '<div class="school-board-hover-title">' +
        name +
        "</div>" +
        (member ? '<div class="school-board-hover-member">' + member + "</div>" : "") +
        "</div>"
      );
    }

    function municipalBoundaryHtml(props) {
      var rawName = props && props.CITY_NAME != null ? String(props.CITY_NAME) : "";
      var name = escapeHtml(standardCapitalization(rawName || "Municipality"));
      return (
        '<div class="school-board-hover-inner">' +
        '<div class="school-board-hover-title">' +
        name +
        "</div></div>"
      );
    }

    function isStudentHexDensityTooltipEnabled() {
      if (residenceDensityHeatmapHiddenAtCurrentZoom()) {
        return false;
      }
      var el = document.getElementById("toggle-student-hex-density-tooltip");
      return !el || el.checked;
    }

    function visibleOverlayHitLayers() {
      var out = [];
      for (var i = 0; i < MAP_OVERLAY_HIT_LAYER_ORDER_TOP_FIRST.length; i++) {
        var lid = MAP_OVERLAY_HIT_LAYER_ORDER_TOP_FIRST[i];
        if (
          (lid === "student-hex-hit-fill" ||
            lid === "charter-student-hex-hit-fill" ||
            lid === "homeschool-student-hex-hit-fill") &&
          !isStudentHexDensityTooltipEnabled()
        ) {
          continue;
        }
        try {
          if (map.getLayer(lid) && map.getLayoutProperty(lid, "visibility") === "visible") {
            out.push(lid);
          }
        } catch (err) {
          /* ignore */
        }
      }
      return out;
    }

    map.on("mousemove", function (e) {
      var hitLayers = visibleOverlayHitLayers();
      if (!hitLayers.length) {
        clearBoundaryHoverUi();
        clearSchoolHoverUi();
        refreshAssignmentBoundaryHighlight();
        return;
      }

      var feats = map.queryRenderedFeatures(e.point, { layers: hitLayers });
      if (!feats.length) {
        clearBoundaryHoverUi();
        clearSchoolHoverUi();
        refreshAssignmentBoundaryHighlight();
        return;
      }

      var top = feats[0];
      var layerId = top.layer.id;

      if (layerId === "schools-private") {
        clearBoundaryHoverUi();
        clearOutlineHighlight();
        clearHoverRing();
        refreshAssignmentBoundaryHighlight();
        map.getCanvas().style.cursor = "pointer";
        schoolHoverPopup
          .setLngLat(e.lngLat)
          .setHTML(privateSchoolDetailHtml(top.properties))
          .addTo(map);
        return;
      }

      if (
        layerId === "schools-elementary" ||
        layerId === "schools-middle" ||
        layerId === "schools-high" ||
        layerId === "schools-charter"
      ) {
        clearBoundaryHoverUi();
        var p = top.properties;
        var hMsid = p.SCHOOLS_ID != null ? Number(p.SCHOOLS_ID) : null;
        if (hMsid != null && !isNaN(hMsid)) {
          setAssignmentHoverHighlightForSchoolMsid(hMsid);
        } else {
          clearOutlineHighlight();
          clearHoverRing();
          refreshAssignmentBoundaryHighlight();
        }
        map.getCanvas().style.cursor = "pointer";
        schoolHoverPopup.setLngLat(e.lngLat).setHTML(schoolDetailHtml(p)).addTo(map);
        return;
      }

      if (
        layerId === "student-hex-hit-fill" ||
        layerId === "charter-student-hex-hit-fill" ||
        layerId === "homeschool-student-hex-hit-fill"
      ) {
        clearBoundaryHoverUi();
        clearSchoolHoverUi();
        refreshAssignmentBoundaryHighlight();
        map.getCanvas().style.cursor = "default";
        var wantB =
          isStudentResidenceLayerEnabled() &&
          map.getLayer("student-hex-hit-fill") &&
          map.getLayoutProperty("student-hex-hit-fill", "visibility") === "visible";
        var wantC =
          isCharterStudentResidenceLayerEnabled() &&
          map.getLayer("charter-student-hex-hit-fill") &&
          map.getLayoutProperty("charter-student-hex-hit-fill", "visibility") === "visible";
        var wantH =
          isHomeschoolStudentResidenceLayerEnabled() &&
          map.getLayer("homeschool-student-hex-hit-fill") &&
          map.getLayoutProperty("homeschool-student-hex-hit-fill", "visibility") === "visible";
        var qLayers = [];
        if (wantB) {
          qLayers.push("student-hex-hit-fill");
        }
        if (wantC) {
          qLayers.push("charter-student-hex-hit-fill");
        }
        if (wantH) {
          qLayers.push("homeschool-student-hex-hit-fill");
        }
        if (!qLayers.length) {
          studentHexHoverPopup.remove();
        } else {
          var pair = map.queryRenderedFeatures(e.point, { layers: qLayers });
          var propsB = null;
          var propsC = null;
          var propsH = null;
          for (var ip = 0; ip < pair.length; ip++) {
            var lId = pair[ip].layer && pair[ip].layer.id;
            if (lId === "student-hex-hit-fill" && !propsB) {
              propsB = pair[ip].properties;
            } else if (lId === "charter-student-hex-hit-fill" && !propsC) {
              propsC = pair[ip].properties;
            } else if (lId === "homeschool-student-hex-hit-fill" && !propsH) {
              propsH = pair[ip].properties;
            }
          }
          var cohortPhrase = studentResidenceCohortTooltipPhrase();
          studentHexHoverPopup
            .setLngLat(e.lngLat)
            .setHTML(
              combinedResidenceHexHoverHtml(
                propsB,
                propsC,
                propsH,
                wantB,
                wantC,
                wantH,
                cohortPhrase
              )
            )
            .addTo(map);
        }
        return;
      }

      if (layerId === "school-isochrones-fill" || layerId === "school-isochrones-outline") {
        clearTravelShedResidenceDebounce();
        clearOutlineHighlight();
        clearHoverRing();
        clearMunicipalHoverStroke();
        schoolHoverPopup.remove();
        studentHexHoverPopup.remove();
        boundaryHoverPopup.remove();
        schoolBoardHoverPopup.remove();
        refreshAssignmentBoundaryHighlight();
        map.getCanvas().style.cursor = "default";
        var miTravel =
          top.properties && top.properties.iso_miles != null
            ? Math.round(Number(top.properties.iso_miles))
            : NaN;
        var gIso = top.geometry;
        var ptLng = e.lngLat.lng;
        var ptLat = e.lngLat.lat;
        var seq = ++travelShedResidenceHoverGen;
        travelShedResidenceDebounceId = setTimeout(function () {
          travelShedResidenceDebounceId = null;
          if (seq !== travelShedResidenceHoverGen) {
            return;
          }
          var mShed = masterRow(getActiveTravelShedMsid());
          var byCanon =
            gIso && (gIso.type === "Polygon" || gIso.type === "MultiPolygon")
              ? travelShedResidenceCountsInIsochrone(gIso)
              : null;
          if (byCanon == null) {
            byCanon = {};
          }
          var htmlR = formatTravelShedResidenceHtml(byCanon, mShed, miTravel);
          try {
            travelShedHoverPopup
              .setLngLat([ptLng, ptLat])
              .setHTML(htmlR)
              .addTo(map);
          } catch (eTs) {
            /* ignore */
          }
        }, 150);
        return;
      }

      if (
        layerId === "school-parcels-high" ||
        layerId === "school-parcels-jr-sr" ||
        layerId === "school-parcels-middle" ||
        layerId === "school-parcels-elementary"
      ) {
        clearBoundaryHoverUi();
        clearSchoolHoverUi();
        map.getCanvas().style.cursor = "";
        refreshAssignmentBoundaryHighlight();
        return;
      }

      if (layerId === "school-board-districts-fill" || layerId === "school-board-districts-outline") {
        clearBoundaryHoverUi();
        clearSchoolHoverUi();
        map.getCanvas().style.cursor = "pointer";
        schoolBoardHoverPopup
          .setLngLat(e.lngLat)
          .setHTML(schoolBoardDistrictHtml(top.properties))
          .addTo(map);
        refreshAssignmentBoundaryHighlight();
        return;
      }

      if (layerId === "municipal-boundaries-fill" || layerId === "municipal-boundaries-outline") {
        clearBoundaryHoverUi();
        clearSchoolHoverUi();
        applyMunicipalHoverStroke(top);
        map.getCanvas().style.cursor = "pointer";
        schoolBoardHoverPopup
          .setLngLat(e.lngLat)
          .setHTML(municipalBoundaryHtml(top.properties))
          .addTo(map);
        refreshAssignmentBoundaryHighlight();
        return;
      }

      if (BOUNDARY_FILL_LAYERS.indexOf(layerId) === -1 && boundaryLayerIdToSource(layerId) == null) {
        clearBoundaryHoverUi();
        clearSchoolHoverUi();
        refreshAssignmentBoundaryHighlight();
        return;
      }

      clearSchoolHoverUi();
      schoolBoardHoverPopup.remove();
      clearMunicipalHoverStroke();

      var f = top;
      var props = f.properties;
      var msid = props.MSID != null ? Number(props.MSID) : null;
      if (msid != null && isNaN(msid)) msid = null;
      var src = boundaryLayerIdToSource(layerId);

      var hoveringDifferentAssignment =
        msid != null &&
        selectedSchoolMsid != null &&
        msid !== selectedSchoolMsid;

      if (msid != null && selectedSchoolMsid != null) {
        if (msid !== selectedSchoolMsid) {
          clearSelectedAssignmentBoundary();
        } else {
          applySelectedAssignmentBoundary(msid);
        }
      }

      if (!hoveringDifferentAssignment) {
        refreshAssignmentBoundaryHighlight();
      }

      map.getCanvas().style.cursor = "pointer";

      boundaryHoverPopup
        .setLngLat(e.lngLat)
        .setHTML(escapeHtml(boundaryTitleText(props)))
        .addTo(map);

      if (src && msid != null) {
        if (lastOutline.source !== src || lastOutline.id !== msid) {
          clearOutlineHighlight();
          lastOutline.source = src;
          lastOutline.id = msid;
          try {
            map.setFeatureState({ source: src, id: msid }, { highlight: true });
          } catch (e2) {
            /* ignore */
          }
        }
      } else {
        clearOutlineHighlight();
      }

      if (msid != null && schoolByMsid[msid]) {
        if (lastRingMsid !== msid) {
          clearHoverRing();
          lastRingMsid = msid;
          try {
            map.setFeatureState({ source: "schools", id: msid }, { ring: true });
          } catch (e3) {
            /* ignore */
          }
        }
      } else {
        clearHoverRing();
      }
    });

    map.on("mouseout", function () {
      clearBoundaryHoverUi();
      clearSchoolHoverUi();
      refreshAssignmentBoundaryHighlight();
    });

    function visibleClickLayers(orderedIds) {
      var out = [];
      for (var i = 0; i < orderedIds.length; i++) {
        var lid = orderedIds[i];
        try {
          if (!map.getLayer(lid)) continue;
          var v = map.getLayoutProperty(lid, "visibility");
          if (v === "none") continue;
          if (v === "visible" || v === undefined) out.push(lid);
        } catch (errC) {
          /* layer missing */
        }
      }
      return out;
    }

    function firstTopFeatureInLayers(e, orderedLayerIds) {
      var vis = visibleClickLayers(orderedLayerIds);
      if (!vis.length) return null;
      var feats = map.queryRenderedFeatures(e.point, { layers: vis });
      return feats && feats.length ? feats[0] : null;
    }

    function msidFromMapPickFeature(f) {
      if (!f || !f.properties) return null;
      var lid = f.layer && f.layer.id ? f.layer.id : "";
      if (SCHOOL_LAYER_IDS.indexOf(lid) >= 0) {
        var s = f.properties.SCHOOLS_ID;
        if (s == null || s === "") return null;
        var m = Number(s);
        return isNaN(m) ? null : m;
      }
      if (SCHOOL_PARCEL_LAYERS_CLICK_TOP_FIRST.indexOf(lid) >= 0) {
        var s2 = f.properties.SCHOOLS_ID;
        if (s2 == null || s2 === "") return null;
        var m2 = Number(s2);
        return isNaN(m2) ? null : m2;
      }
      if (ASSIGNMENT_BOUNDARY_LAYERS_CLICK_TOP_FIRST.indexOf(lid) >= 0) {
        var s3 = f.properties.MSID;
        if (s3 == null || s3 === "") return null;
        var m3 = Number(s3);
        return isNaN(m3) ? null : m3;
      }
      return null;
    }

    map.on("click", function (e) {
      if (!isExistingConditionsViewActive()) return;
      var fSch = firstTopFeatureInLayers(e, SCHOOL_LAYERS_CLICK_TOP_FIRST);
      var fParc = fSch
        ? null
        : firstTopFeatureInLayers(e, SCHOOL_PARCEL_LAYERS_CLICK_TOP_FIRST);
      var fBnd =
        fSch || fParc
          ? null
          : firstTopFeatureInLayers(e, ASSIGNMENT_BOUNDARY_LAYERS_CLICK_TOP_FIRST);
      var f = fSch || fParc || fBnd;
      if (!f) return;
      var msid = msidFromMapPickFeature(f);
      if (msid == null) return;
      if (!isMsidInSchoolSelectDropdown(msid)) return;
      var sel = document.getElementById("school-select");
      if (!sel) return;
      if (fSch || fParc) {
        pendingMapSelectFrame = "centerOnSchool";
      } else {
        pendingMapSelectFrame = "assignment";
      }
      if (String(sel.value) === String(msid)) {
        applyExistingSchoolFromSelectValue(schoolByMsid);
        return;
      }
      sel.value = String(msid);
      sel.dispatchEvent(new Event("change", { bubbles: true }));
    });

    function boundarySandboxMapMouseMoveForPaintLasso(e) {
      if (!BOUNDARY_SANDBOX_PAINT.active && !BOUNDARY_SANDBOX_LASSO.active) {
        return;
      }
      if (BOUNDARY_SANDBOX_PAINT.active) {
        var dx = e.point.x - BOUNDARY_SANDBOX_PAINT.startX;
        var dy = e.point.y - BOUNDARY_SANDBOX_PAINT.startY;
        if (dx * dx + dy * dy > BOUNDARY_SANDBOX_BRUSH_DRAG_THRESH2) {
          BOUNDARY_SANDBOX_PAINT.isDrag = true;
        }
        if (BOUNDARY_SANDBOX_PAINT.isDrag) {
          tryBrushDragAtPoint(e.point);
        }
      } else if (BOUNDARY_SANDBOX_LASSO.active && BOUNDARY_SANDBOX_LASSO.points) {
        BOUNDARY_SANDBOX_LASSO.points.push([e.lngLat.lng, e.lngLat.lat]);
        setBoundarySandboxLassoSource({
          type: "FeatureCollection",
          features: [
            {
              type: "Feature",
              properties: {},
              geometry: { type: "LineString", coordinates: BOUNDARY_SANDBOX_LASSO.points },
            },
          ],
        });
      }
    }

    map.on("mousedown", function (e) {
      if (!isBoundarySandboxViewActive()) {
        return;
      }
      if (e.originalEvent && e.originalEvent.button !== 0) {
        return;
      }
      try {
        if (map.getLayoutProperty("boundary-sandbox-hex-fill", "visibility") !== "visible") {
          return;
        }
      } catch (eV0) {
        return;
      }
      var toolM = getBoundarySandboxSelectTool();
      e.preventDefault();
      if (toolM === "brush") {
        clearBoundarySandboxLassoRegionFill();
        BOUNDARY_SANDBOX_PAINT.active = true;
        BOUNDARY_SANDBOX_PAINT.lastKey = null;
        BOUNDARY_SANDBOX_PAINT.isDrag = false;
        BOUNDARY_SANDBOX_PAINT.startX = e.point.x;
        BOUNDARY_SANDBOX_PAINT.startY = e.point.y;
        BOUNDARY_SANDBOX_PAINT.clickKey = querySandboxHexKeyAtPoint(e.point);
        try {
          map.dragPan.disable();
          map.getCanvas().style.cursor = "crosshair";
        } catch (eB0) {
          /* ignore */
        }
        return;
      }
      if (toolM === "lasso") {
        BOUNDARY_SANDBOX_LASSO.active = true;
        BOUNDARY_SANDBOX_LASSO.points = [[e.lngLat.lng, e.lngLat.lat]];
        try {
          map.dragPan.disable();
          map.getCanvas().style.cursor = "crosshair";
        } catch (eL1) {
          /* ignore */
        }
        setBoundarySandboxLassoSource({
          type: "FeatureCollection",
          features: [
            {
              type: "Feature",
              properties: {},
              geometry: {
                type: "LineString",
                coordinates: BOUNDARY_SANDBOX_LASSO.points,
              },
            },
          ],
        });
      }
    });
    map.on("mousemove", boundarySandboxMapMouseMoveForPaintLasso);
    map.on("mouseup", function () {
      endBoundarySandboxPaintOrLassoFromWindow();
    });
    if (typeof window !== "undefined") {
      window.addEventListener("mouseup", endBoundarySandboxPaintOrLassoFromWindow);
    }
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function ringCentroid(ring) {
    if (!ring || ring.length < 2) return null;
    var n = ring.length;
    var last = ring[n - 1];
    var first = ring[0];
    if (last[0] === first[0] && last[1] === first[1]) {
      n -= 1;
    }
    var sx = 0;
    var sy = 0;
    for (var i = 0; i < n; i++) {
      sx += ring[i][0];
      sy += ring[i][1];
    }
    return [sx / n, sy / n];
  }

  /** Approximate interior point for hex polygons (ArcGIS-style centroid). */
  function polygonCentroid(geometry) {
    if (!geometry || !geometry.type) return null;
    if (geometry.type === "Polygon") {
      var ring = geometry.coordinates[0];
      return ringCentroid(ring);
    }
    if (geometry.type === "MultiPolygon") {
      var best = null;
      var bestLen = -1;
      for (var p = 0; p < geometry.coordinates.length; p++) {
        var ring = geometry.coordinates[p][0];
        if (!ring || ring.length < 2) continue;
        var c = ringCentroid(ring);
        if (!c) continue;
        if (ring.length > bestLen) {
          bestLen = ring.length;
          best = c;
        }
      }
      return best;
    }
    return null;
  }

  /**
   * Undirected adjacency: two hexes touch on an edge. Built once from all hex geometries.
   * O(candidates) with coarse centroid grid; pair tests use turf.booleanTouches when available.
   * @param {Object<string, *>} geometryByHexKey
   * @returns {Object<string, string[]>|null} hexKey -> adjacent hex keys; null = skip adjacency
   */
  function buildHexNeighborMap(geometryByHexKey) {
    if (!geometryByHexKey) {
      return null;
    }
    var boolTouches = null;
    if (typeof turf !== "undefined" && turf) {
      if (typeof turf.booleanTouches === "function") {
        boolTouches = turf.booleanTouches;
      } else if (typeof turf.booleanTouch === "function") {
        boolTouches = turf.booleanTouch;
      }
    }
    if (boolTouches == null || typeof turf.feature !== "function") {
      return null;
    }
    var keys = Object.keys(geometryByHexKey);
    if (!keys.length) {
      return {};
    }
    var n = keys.length;
    if (n === 1) {
      var o1 = Object.create(null);
      o1[keys[0]] = [];
      return o1;
    }
    var CELL = 0.12;
    var bucket = Object.create(null);
    for (var bi = 0; bi < n; bi++) {
      var kB = keys[bi];
      var cB = polygonCentroid(geometryByHexKey[kB]);
      if (!cB || cB.length < 2) {
        continue;
      }
      var cxb = Math.floor(cB[0] / CELL);
      var cyb = Math.floor(cB[1] / CELL);
      var bid = cxb + "," + cyb;
      if (!bucket[bid]) {
        bucket[bid] = [];
      }
      bucket[bid].push(kB);
    }
    var neighbors = Object.create(null);
    for (var ni = 0; ni < n; ni++) {
      neighbors[keys[ni]] = [];
    }
    for (var i = 0; i < n; i++) {
      var k1 = keys[i];
      var c1 = polygonCentroid(geometryByHexKey[k1]);
      if (!c1 || c1.length < 2) {
        continue;
      }
      var cx1 = Math.floor(c1[0] / CELL);
      var cy1 = Math.floor(c1[1] / CELL);
      for (var ddx = -1; ddx <= 1; ddx++) {
        for (var ddy = -1; ddy <= 1; ddy++) {
          var bList = bucket[cx1 + ddx + "," + (cy1 + ddy)];
          if (!bList) {
            continue;
          }
          for (var t = 0; t < bList.length; t++) {
            var k2 = bList[t];
            if (k2 === k1) {
              continue;
            }
            if (k2 <= k1) {
              continue;
            }
            var g1 = geometryByHexKey[k1];
            var g2 = geometryByHexKey[k2];
            if (!g1 || !g2) {
              continue;
            }
            var touches = false;
            try {
              touches = boolTouches(turf.feature(g1), turf.feature(g2));
            } catch (eAdj) {
              /* ignore */
            }
            if (touches) {
              neighbors[k1].push(k2);
              neighbors[k2].push(k1);
            }
          }
        }
      }
    }
    return neighbors;
  }

  /**
   * @param {Object|null|undefined} p
   * @returns {string|null} e.g. "id:123" when a stable hex id is present, else null
   */
  function studentHexIdKeyFromProperties(p) {
    if (!p) {
      return null;
    }
    var id =
      p.GRID_ID != null
        ? p.GRID_ID
        : p.HEX_ID != null
          ? p.HEX_ID
          : p.HexID != null
            ? p.HexID
            : p.hex_id != null
              ? p.hex_id
              : p.OBJECTID != null
                ? p.OBJECTID
                : p.FID != null
                  ? p.FID
                  : null;
    if (id != null && id !== "") {
      return "id:" + String(id);
    }
    return null;
  }

  function studentHexKey(feature) {
    var p = feature.properties || {};
    var fromId = studentHexIdKeyFromProperties(p);
    if (fromId) {
      return fromId;
    }
    return "geom:" + JSON.stringify(feature.geometry);
  }

  /**
   * One increment per homeschool student row; hex id from GRID_ID matches main student hex keys.
   * @param {Object|null} homeschoolFc
   * @returns {Object<string, number>}
   */
  function buildHomeschoolHexCounts(homeschoolFc) {
    var o = Object.create(null);
    if (!homeschoolFc || !homeschoolFc.features) {
      return o;
    }
    for (var i = 0; i < homeschoolFc.features.length; i++) {
      var k = studentHexKey(homeschoolFc.features[i]);
      o[k] = (o[k] || 0) + 1;
    }
    return o;
  }

  /**
   * First polygon/MultiPolygon per hex key from homeschool features — used when that hex is absent from the student bundle.
   * @param {Object|null} homeschoolFc
   * @returns {Object<string, GeoJSON.Geometry>}
   */
  function buildHomeschoolHexGeometryFallback(homeschoolFc) {
    var out = Object.create(null);
    if (!homeschoolFc || !homeschoolFc.features) {
      return out;
    }
    for (var i = 0; i < homeschoolFc.features.length; i++) {
      var f = homeschoolFc.features[i];
      if (!f || !f.geometry) {
        continue;
      }
      var t = f.geometry.type;
      if (t !== "Polygon" && t !== "MultiPolygon") {
        continue;
      }
      var k = studentHexKey(f);
      if (!out[k]) {
        out[k] = f.geometry;
      }
    }
    return out;
  }

  /**
   * Detail row for boundary sandbox: mirrors student hex fields (`ELEM_` / `MID_` / `HIGH_`) from homeschool export zoned columns.
   */
  /**
   * @param {Object|null} zoningTriplet from `attendanceZoningTripletAtLngLat` / per-hex cache (`elem`/`mid`/`high`).
   */
  function homeschoolSandboxDetailFromProperties(props, zoningTriplet) {
    props = props || {};
    var zt = zoningTriplet || {};
    function merged(propVal, inferredNum) {
      if (msidNormForZoning(propVal) != null) {
        return propVal;
      }
      if (inferredNum != null && !isNaN(Number(inferredNum)) && Number(inferredNum) > 0) {
        return Math.round(Number(inferredNum));
      }
      return propVal;
    }
    return {
      Grade: props.Grade,
      MSID: HOMESCHOOL_ATTENDANCE_MSID,
      ELEM_: merged(props.Zoned_Elem, zt.elem),
      MID_: merged(props.Zoned_Midd, zt.mid),
      HIGH_: merged(props.Zoned_High, zt.high),
      INT_: null,
      lunch_stat: null,
      ethnicity: null,
      __homeschool: true,
    };
  }

  /**
   * Homeschool students grouped by `studentHexKey` for sandbox aggregation (same keys as `HOMESCHOOL_HEX_COUNTS`).
   */
  function buildHomeschoolDetailsByHexKey(homeschoolFc) {
    var byHex = Object.create(null);
    if (!homeschoolFc || !homeschoolFc.features) {
      return byHex;
    }
    var zoningByHex = Object.create(null);
    for (var i = 0; i < homeschoolFc.features.length; i++) {
      var f = homeschoolFc.features[i];
      var k = studentHexKey(f);
      if (!Object.prototype.hasOwnProperty.call(zoningByHex, k)) {
        zoningByHex[k] = homeschoolAttendanceZoningTripletForHex(k, f);
      }
      var det = homeschoolSandboxDetailFromProperties(f.properties, zoningByHex[k]);
      if (!byHex[k]) {
        byHex[k] = [];
      }
      byHex[k].push(det);
    }
    return byHex;
  }

  /**
   * Resolver for homeschool map layers and density tooltips: main student hex geometry when present,
   * else homeschool-source hex polygon for GRIDs not in the bundle.
   */
  function homeschoolHexGeometry(hexKey) {
    var k = String(hexKey);
    if (
      STUDENT_HEX_INDEX &&
      STUDENT_HEX_INDEX.geometryByHexKey &&
      STUDENT_HEX_INDEX.geometryByHexKey[k]
    ) {
      return STUDENT_HEX_INDEX.geometryByHexKey[k];
    }
    if (HOMESCHOOL_HEX_GEOMETRY_FALLBACK && HOMESCHOOL_HEX_GEOMETRY_FALLBACK[k]) {
      return HOMESCHOOL_HEX_GEOMETRY_FALLBACK[k];
    }
    return null;
  }

  /**
   * Hex keys where homeschool students live and the hex centroid lies inside the school’s assignment polygon.
   * Same geographic rule as capture KPIs / density alignment (not “zoned from student index only”).
   * @returns {Object<string, true>}
   */
  function homeschoolHexKeysWithCentroidInAssignmentBoundary(msid) {
    var out = Object.create(null);
    if (msid == null || isNaN(Number(msid))) {
      return out;
    }
    if (
      typeof turf === "undefined" ||
      !turf ||
      typeof turf.point !== "function" ||
      typeof turf.feature !== "function" ||
      typeof turf.booleanPointInPolygon !== "function"
    ) {
      return out;
    }
    if (!HOMESCHOOL_HEX_COUNTS) {
      return out;
    }
    var bf = findBoundaryFeatureForMsid(Number(msid));
    if (!bf || !bf.geometry) {
      return out;
    }
    var polyFeat;
    try {
      polyFeat = turf.feature(bf.geometry);
    } catch (ePoly) {
      return out;
    }
    for (var hexKey in HOMESCHOOL_HEX_COUNTS) {
      if (!Object.prototype.hasOwnProperty.call(HOMESCHOOL_HEX_COUNTS, hexKey)) {
        continue;
      }
      var cnt = Number(HOMESCHOOL_HEX_COUNTS[hexKey]) || 0;
      if (cnt <= 0) {
        continue;
      }
      var gHex = homeschoolHexGeometry(hexKey);
      if (!gHex) {
        continue;
      }
      var ctr = polygonCentroid(gHex);
      if (!ctr || ctr.length < 2) {
        continue;
      }
      var inside = false;
      try {
        inside = turf.booleanPointInPolygon(turf.point(ctr), polyFeat);
      } catch (eIn) {
        inside = false;
      }
      if (inside) {
        out[hexKey] = true;
      }
    }
    return out;
  }

  /**
   * Grade-eligible homeschool students where the hex centroid lies inside the school’s assignment polygon.
   * When per-student homeschool rows exist, counts only grades that match the school’s level band (same as From-To “resident” notion).
   */
  function countHomeschoolStudentsInAssignmentBoundary(msid) {
    if (msid == null || isNaN(Number(msid))) {
      return 0;
    }
    if (
      typeof turf === "undefined" ||
      !turf ||
      typeof turf.point !== "function" ||
      typeof turf.feature !== "function" ||
      typeof turf.booleanPointInPolygon !== "function"
    ) {
      return 0;
    }
    if (!HOMESCHOOL_HEX_COUNTS) {
      return 0;
    }
    var keyCache = String(Number(msid));
    if (Object.prototype.hasOwnProperty.call(homeschoolInBoundaryByMsidCache, keyCache)) {
      return homeschoolInBoundaryByMsidCache[keyCache];
    }
    var keyBag = homeschoolHexKeysWithCentroidInAssignmentBoundary(Number(msid));
    var m = masterRow(Number(msid));
    var total = 0;
    for (var hk in keyBag) {
      if (!keyBag[hk]) {
        continue;
      }
      var cnt = Number(HOMESCHOOL_HEX_COUNTS[hk]) || 0;
      var rows = HOMESCHOOL_DETAILS_BY_HEX_KEY && HOMESCHOOL_DETAILS_BY_HEX_KEY[hk];
      if (m && rows && rows.length) {
        for (var ir = 0; ir < rows.length; ir++) {
          var rd = rows[ir];
          if (rd && studentGradeInSelectedSchoolBand(rd.Grade, m, false)) {
            total += 1;
          }
        }
      } else {
        total += cnt;
      }
    }
    homeschoolInBoundaryByMsidCache[keyCache] = total;
    return total;
  }

  /**
   * Unpacks `v:2` { g: hexId -> geometry, r: property[] } to a standard FeatureCollection.
   * Falls back to a plain FeatureCollection. Used to avoid repeating hex geometry in JSON.
   * @param {*} raw
   * @returns {Object|null}
   */
  function expandStudentHexBundleToFeatureCollection(raw) {
    if (!raw) {
      return null;
    }
    if (raw.v === 2 && raw.g && Array.isArray(raw.r)) {
      var geoms = raw.g;
      var rows = raw.r;
      var out = [];
      for (var i = 0; i < rows.length; i++) {
        var pr = rows[i] || {};
        var hk = studentHexIdKeyFromProperties(pr);
        if (!hk) {
          continue;
        }
        var geom = geoms[hk];
        if (!geom) {
          continue;
        }
        out.push({ type: "Feature", properties: pr, geometry: geom });
      }
      return { type: "FeatureCollection", features: out };
    }
    if (raw.type === "FeatureCollection") {
      return raw;
    }
    return null;
  }

  /**
   * One object per student for filters (grade / zoned MSIDs). ArcGIS exports use `Grade`
   * inconsistently; fall back to `grade` or `StudGRD` when `Grade` is empty.
   * @param {Object} p feature.properties
   * @returns {{ Grade: string, MSID: string, ELEM_: *, MID_: *, INT_: *, HIGH_: * }}
   */
  function studentHexDetailFromProps(p) {
    var g = "";
    if (p.Grade != null && String(p.Grade).trim() !== "") {
      g = String(p.Grade).trim();
    } else if (p.grade != null && String(p.grade).trim() !== "") {
      g = String(p.grade).trim();
    } else if (p.StudGRD != null && String(p.StudGRD).trim() !== "") {
      g = String(p.StudGRD).trim();
    }
    var oid = "";
    if (p.OBJECTID != null && String(p.OBJECTID).trim() !== "") {
      oid = "o:" + String(p.OBJECTID).trim();
    } else if (p.JOIN_FID != null && String(p.JOIN_FID).trim() !== "") {
      oid = "j:" + String(p.JOIN_FID).trim();
    } else if (p.TARGET_FID != null && String(p.TARGET_FID).trim() !== "") {
      oid = "t:" + String(p.TARGET_FID).trim();
    }
    return {
      Grade: g,
      MSID: p.MSID != null ? String(p.MSID).trim() : "",
      ELEM_: p.ELEM_,
      MID_: p.MID_,
      INT_: p.INT_,
      HIGH_: p.HIGH_,
      _oid: oid,
      ethnicity: p.ethnicity != null ? String(p.ethnicity).trim() : "",
      lunch_stat: p.lunch_stat != null ? String(p.lunch_stat).trim() : "",
    };
  }

  /**
   * Attendance MSID in district charter 65xx / 66xx range (residential density layer).
   */
  function attendanceMsidIsCharterDistrictResidentialRange(msid) {
    var n = Number(msid);
    if (!isFinite(n) || isNaN(n)) return false;
    return n >= 6500 && n <= 6699;
  }

  function buildStudentHexIndex(fc) {
    var countsByMsid = {};
    var geometryByHexKey = {};
    var detailsByMsid = {};
    var charterDistrictHexCounts = {};
    if (!fc || !fc.features) {
      return {
        countsByMsid: countsByMsid,
        geometryByHexKey: geometryByHexKey,
        detailsByMsid: detailsByMsid,
        charterDistrictHexCounts: charterDistrictHexCounts,
        neighborsByHexKey: Object.create(null),
      };
    }
    for (var i = 0; i < fc.features.length; i++) {
      var f = fc.features[i];
      var p = f.properties || {};
      var msid = Number(
        p.MSID != null ? p.MSID : p.SCHOOLS_ID != null ? p.SCHOOLS_ID : NaN
      );
      if (isNaN(msid)) continue;
      var key = studentHexKey(f);
      if (!geometryByHexKey[key]) {
        geometryByHexKey[key] = f.geometry;
      }
      var sk = String(msid);
      if (!countsByMsid[sk]) countsByMsid[sk] = {};
      var inc = 1;
      if (p.count != null && isFinite(Number(p.count))) {
        inc = Number(p.count);
      }
      countsByMsid[sk][key] = (countsByMsid[sk][key] || 0) + inc;

      if (attendanceMsidIsCharterDistrictResidentialRange(msid)) {
        charterDistrictHexCounts[key] =
          (charterDistrictHexCounts[key] || 0) + inc;
      }

      if (!detailsByMsid[sk]) detailsByMsid[sk] = {};
      if (!detailsByMsid[sk][key]) detailsByMsid[sk][key] = [];
      var det = studentHexDetailFromProps(p);
      if (inc === 1) {
        detailsByMsid[sk][key].push(det);
      } else {
        for (var jd = 0; jd < inc; jd++) {
          detailsByMsid[sk][key].push(Object.assign({}, det));
        }
      }
    }
    var neighborsByHexKey = buildHexNeighborMap(geometryByHexKey);
    if (!neighborsByHexKey) {
      neighborsByHexKey = Object.create(null);
    }
    return {
      countsByMsid: countsByMsid,
      geometryByHexKey: geometryByHexKey,
      detailsByMsid: detailsByMsid,
      charterDistrictHexCounts: charterDistrictHexCounts,
      neighborsByHexKey: neighborsByHexKey,
    };
  }

  /**
   * Every feature in the student hex file (not filtered by attendance MSID);
   * per-hex grade counts and centroids for travel-shed PIP aggregation.
   */
  function buildTravelShedResidenceIndex(fc) {
    var gradeCountsByHex = {};
    var geometryByHexKey = {};
    if (!fc || !fc.features) {
      return { gradeCountsByHex: {}, centroidsByHex: {}, hexKeyList: [] };
    }
    for (var i0 = 0; i0 < fc.features.length; i0++) {
      var f0 = fc.features[i0];
      if (!f0 || !f0.geometry) continue;
      var key0 = studentHexKey(f0);
      if (!geometryByHexKey[key0]) {
        geometryByHexKey[key0] = f0.geometry;
      }
      var p0 = f0.properties || {};
      var inc0 = 1;
      if (p0.count != null && isFinite(Number(p0.count))) {
        inc0 = Number(p0.count);
      }
      var det0 = studentHexDetailFromProps(p0);
      var gCanon = canonicalStudentGradeCode(det0.Grade);
      if (gCanon == null || gCanon === "") {
        gCanon = "__UNK__";
      }
      if (!gradeCountsByHex[key0]) gradeCountsByHex[key0] = {};
      gradeCountsByHex[key0][gCanon] = (gradeCountsByHex[key0][gCanon] || 0) + inc0;
    }
    var centroidsByHex = {};
    var hexKeyList = [];
    for (var k0 in geometryByHexKey) {
      if (!Object.prototype.hasOwnProperty.call(geometryByHexKey, k0)) continue;
      var c0 = polygonCentroid(geometryByHexKey[k0]);
      if (c0 && c0.length === 2) {
        centroidsByHex[k0] = c0;
        hexKeyList.push(k0);
      }
    }
    return {
      gradeCountsByHex: gradeCountsByHex,
      centroidsByHex: centroidsByHex,
      hexKeyList: hexKeyList,
    };
  }

  function travelShedGradeSortKey(canon) {
    if (canon === "PK") return 0;
    if (canon === "K") return 1;
    if (canon === "__NOGRADE__") return 9998;
    if (canon === "__UNK__") return 9999;
    if (/^0[1-9]$/.test(canon)) return 2 + parseInt(canon, 10);
    if (/^1[0-2]$/.test(canon)) return 2 + parseInt(canon, 10);
    return 5000;
  }

  function travelShedGradeDisplayLabel(canon) {
    if (canon === "__NOGRADE__") return "No Grade";
    if (canon === "__UNK__") return "—";
    if (canon === "PK" || canon === "K") return canon;
    if (/^0[1-9]$/.test(canon)) return String(parseInt(canon, 10));
    return String(canon);
  }

  /**
   * Representative string so `studentGradeInSelectedSchoolBand` re-canonicalizes like source Grade.
   */
  function travelShedRawGradeStringForBand(canon) {
    if (canon === "__UNK__") return "";
    if (canon === "PK" || canon === "K") return canon;
    if (/^0[1-9]$/.test(canon)) return String(parseInt(canon, 10));
    if (/^1[0-2]$/.test(canon)) return canon;
    return String(canon);
  }

  function clearTravelShedResidenceDebounce() {
    if (travelShedResidenceDebounceId != null) {
      try {
        clearTimeout(travelShedResidenceDebounceId);
      } catch (e) {
        /* ignore */
      }
      travelShedResidenceDebounceId = null;
    }
  }

  /**
   * Sums all hex-residence grade buckets whose hex **centroid** lies inside the isochrone polygon
   * (or MultiPolygon). Returns map canonical grade key -> count.
   */
  function travelShedResidenceCountsInIsochrone(isoGeometry) {
    if (!TRAVEL_SHED_RESIDENCE_INDEX || !isoGeometry) return null;
    if (
      typeof turf === "undefined" ||
      !turf ||
      typeof turf.point !== "function" ||
      typeof turf.feature !== "function" ||
      typeof turf.booleanPointInPolygon !== "function"
    ) {
      return null;
    }
    var idx = TRAVEL_SHED_RESIDENCE_INDEX;
    if (!idx.hexKeyList || !idx.hexKeyList.length) {
      return {};
    }
    var polyFeat;
    try {
      polyFeat = turf.feature(isoGeometry);
    } catch (ePoly) {
      return null;
    }
    var bbox;
    try {
      bbox = turf.bbox(polyFeat);
    } catch (eB) {
      bbox = null;
    }
    var totalByCanon = {};
    var hlist = idx.hexKeyList;
    for (var i1 = 0; i1 < hlist.length; i1++) {
      var hkx = hlist[i1];
      var c1 = idx.centroidsByHex[hkx];
      if (!c1 || c1.length < 2) continue;
      if (bbox && bbox.length === 4) {
        if (
          c1[0] < bbox[0] ||
          c1[0] > bbox[2] ||
          c1[1] < bbox[1] ||
          c1[1] > bbox[3]
        ) {
          continue;
        }
      }
      var ptf;
      try {
        ptf = turf.point(c1);
      } catch (eP) {
        continue;
      }
      var ins;
      try {
        ins = turf.booleanPointInPolygon(ptf, polyFeat);
      } catch (eI) {
        continue;
      }
      if (!ins) continue;
      var gch = idx.gradeCountsByHex[hkx];
      if (!gch) continue;
      for (var gkx in gch) {
        if (!Object.prototype.hasOwnProperty.call(gch, gkx)) continue;
        totalByCanon[gkx] = (totalByCanon[gkx] || 0) + gch[gkx];
      }
    }
    return totalByCanon;
  }

  function formatTravelShedResidenceHtml(totalByCanon, m, milesRounded) {
    var titleSchool = travelShedTitleSchoolNameForMsid(m);
    var miN =
      milesRounded != null && isFinite(milesRounded)
        ? Math.round(milesRounded)
        : 0;
    var titleLine =
      escapeHtml(titleSchool) + ": " + (miN > 0 ? miN : "—") + " Mi Travel Shed";
    if (!totalByCanon || !Object.keys(totalByCanon).length) {
      return (
        '<div class="travel-shed-hover-inner travel-shed-hover-inner--residence">' +
        '<div class="travel-shed-hover-title">' +
        titleLine +
        "</div>" +
        '<p class="travel-shed-residence-empty">No student residence hexes with centroids inside this area (or residence data not loaded yet).</p></div>'
      );
    }
    var keys = Object.keys(totalByCanon);
    keys.sort(function (a, b) {
      return travelShedGradeSortKey(a) - travelShedGradeSortKey(b);
    });
    var headerHtml =
      '<div class="travel-shed-residence-header">' +
      '<span class="travel-shed-residence-h-grade">Grade</span>' +
      '<span class="travel-shed-residence-h-n">Student Residences</span></div>';
    function rowHtml(ckey) {
      var nct = totalByCanon[ckey];
      var gRaw = travelShedRawGradeStringForBand(ckey);
      var isServed =
        m && gRaw !== "" && studentGradeInSelectedSchoolBand(gRaw, m, false);
      var numStr = (nct != null ? Number(nct) : 0).toLocaleString();
      var lab = travelShedGradeDisplayLabel(ckey);
      var h =
        '<div class="travel-shed-residence-row' +
        (isServed ? " travel-shed-residence-row--served" : "") +
        '"><span class="travel-shed-residence-grade">' +
        escapeHtml(lab) +
        '</span><span class="travel-shed-residence-n">';
      if (isServed) {
        h += "<strong>" + escapeHtml(numStr) + "</strong>";
      } else {
        h += escapeHtml(numStr);
      }
      return h + "</span></div>";
    }
    var nK = keys.length;
    var useTwoCols = nK > 5;
    var mid = Math.ceil(nK / 2);
    var kLeft = useTwoCols ? keys.slice(0, mid) : keys;
    var kRight = useTwoCols ? keys.slice(mid) : [];
    var i3;
    var leftRows = [];
    for (i3 = 0; i3 < kLeft.length; i3++) {
      leftRows.push(rowHtml(kLeft[i3]));
    }
    var rightRows = [];
    for (i3 = 0; i3 < kRight.length; i3++) {
      rightRows.push(rowHtml(kRight[i3]));
    }
    var tableBody;
    if (!useTwoCols) {
      tableBody =
        '<div class="travel-shed-residence-grades travel-shed-residence-grades--1col">' +
        headerHtml +
        '<div class="travel-shed-residence-rows">' +
        leftRows.join("") +
        "</div></div>";
    } else {
      tableBody =
        '<div class="travel-shed-residence-grades travel-shed-residence-grades--2col">' +
        '<div class="travel-shed-residence-pane">' +
        headerHtml +
        '<div class="travel-shed-residence-rows">' +
        leftRows.join("") +
        "</div></div>" +
        '<div class="travel-shed-residence-pane">' +
        headerHtml +
        '<div class="travel-shed-residence-rows">' +
        rightRows.join("") +
        "</div></div></div>";
    }
    return (
      '<div class="travel-shed-hover-inner travel-shed-hover-inner--residence">' +
      '<div class="travel-shed-hover-title">' +
      titleLine +
      "</div>" +
      tableBody +
      "</div>"
    );
  }

  function travelShedTitleSchoolNameForMsid(m) {
    if (m && m.school_name) {
      return eseTableAbbreviatedSchoolName(m);
    }
    return "School";
  }

  function getActiveDashboardSchoolMsid() {
    if (isBoundarySandboxViewActive()) {
      return getSandboxBaseSchoolMsid();
    }
    var panelScenario = document.getElementById("page-scenario");
    if (panelScenario && !panelScenario.hidden) {
      if (scenarioMiddleMsid != null && !isNaN(scenarioMiddleMsid)) {
        return scenarioMiddleMsid;
      }
      return null;
    }
    var sel = document.getElementById("school-select");
    if (!sel || !sel.value) return null;
    var v = Number(sel.value);
    return isNaN(v) ? null : v;
  }

  function isStudentResidenceLayerEnabled() {
    var inp = document.getElementById("toggle-student-hex");
    return !inp || inp.checked;
  }

  function isCharterStudentResidenceLayerEnabled() {
    var inp = document.getElementById("toggle-charter-student-hex");
    return !inp || inp.checked;
  }

  function isHomeschoolStudentResidenceLayerEnabled() {
    var inp = document.getElementById("toggle-homeschool-student-hex");
    return !inp || inp.checked;
  }

  /**
   * Scenario: hex rows are keyed by each student's school MSID.
   * Always include students enrolled at the selected middle school, then add
   * checked feeder elementaries (same feeder rules as collectScenarioWeightedSpec).
   */
  function buildMergedScenarioStudentHexCounts() {
    var combined = {};
    if (!STUDENT_HEX_INDEX || !STUDENT_HEX_INDEX.countsByMsid) {
      return combined;
    }
    var byMs = STUDENT_HEX_INDEX.countsByMsid;

    function addPart(msid) {
      if (msid == null || isNaN(msid)) return;
      var part = byMs[String(msid)];
      if (!part) return;
      for (var hexKey in part) {
        if (!Object.prototype.hasOwnProperty.call(part, hexKey)) continue;
        combined[hexKey] = (combined[hexKey] || 0) + part[hexKey];
      }
    }

    for (var i = 0; i < scenarioLastFeederRows.length; i++) {
      var r = scenarioLastFeederRows[i];
      if (!r.hasEnrollment || r.msid == null || isNaN(r.msid)) continue;
      if (scenarioFeederChecked[r.msid] === false) continue;
      addPart(r.msid);
    }
    return combined;
  }

  /**
   * Same MSIDs as buildMergedScenarioStudentHexCounts: concat per-student rows per hex
   * (for grade / zoned-school filters in scenario mode).
   */
  function buildMergedScenarioStudentHexDetailsByHex() {
    var combined = {};
    if (!STUDENT_HEX_INDEX || !STUDENT_HEX_INDEX.detailsByMsid) {
      return combined;
    }
    var byDet = STUDENT_HEX_INDEX.detailsByMsid;

    function appendPart(msid) {
      if (msid == null || isNaN(msid)) return;
      var part = byDet[String(msid)];
      if (!part) return;
      for (var hexKey in part) {
        if (!Object.prototype.hasOwnProperty.call(part, hexKey)) continue;
        var arr = part[hexKey];
        if (!arr || !arr.length) continue;
        if (!combined[hexKey]) combined[hexKey] = [];
        for (var t = 0; t < arr.length; t++) {
          combined[hexKey].push(arr[t]);
        }
      }
    }

    for (var j = 0; j < scenarioLastFeederRows.length; j++) {
      var r2 = scenarioLastFeederRows[j];
      if (!r2.hasEnrollment || r2.msid == null || isNaN(r2.msid)) continue;
      if (scenarioFeederChecked[r2.msid] === false) continue;
      appendPart(r2.msid);
    }
    return combined;
  }

  /**
   * Per-hex student rows for the active dashboard cohort (selected school or scenario merge).
   * @returns {Object<string, Array<{Grade: string, MSID: string, ELEM_: *, MID_: *, INT_: *, HIGH_: *}>>}
   */
  function getStudentHexCohortDetailsByHex() {
    if (!STUDENT_HEX_INDEX || !STUDENT_HEX_INDEX.detailsByMsid) {
      return {};
    }
    var panelScenario = document.getElementById("page-scenario");
    var onScenario = panelScenario && !panelScenario.hidden;
    if (
      onScenario &&
      scenarioMiddleMsid != null &&
      !isNaN(scenarioMiddleMsid)
    ) {
      return buildMergedScenarioStudentHexDetailsByHex();
    }
    var msid = getActiveDashboardSchoolMsid();
    if (msid == null || isNaN(msid)) return {};
    return STUDENT_HEX_INDEX.detailsByMsid[String(msid)] || {};
  }

  /** Meadowlane Primary (2041) K–2; Intermediate (2031) 3–6; other elementaries PK–6. */
  var MEADOWLANE_PRIMARY_MSID = 2041;
  var MEADOWLANE_INTERMEDIATE_MSID = 2031;

  function studentHexDedupeKey(d) {
    if (d && d._oid != null && String(d._oid) !== "") {
      return String(d._oid);
    }
    return (
      "c:" +
      String((d && d.MSID) || "") +
      "|g:" +
      String((d && d.Grade) || "") +
      "|e:" +
      String((d && d.ELEM_) != null ? d.ELEM_ : "") +
      "|m:" +
      String((d && d.MID_) != null ? d.MID_ : "") +
      "|i:" +
      String((d && d.INT_) != null ? d.INT_ : "") +
      "|h:" +
      String((d && d.HIGH_) != null ? d.HIGH_ : "")
    );
  }

  function msidNormForZoning(v) {
    var n = Number(v);
    if (!isFinite(n) || n <= 0) return null;
    return Math.round(n);
  }

  /**
   * @param {string} raw from Grade / grade / StudGRD
   * @returns {string|null} PK | K | 01..12
   */
  function canonicalStudentGradeCode(raw) {
    if (raw == null) return null;
    var t = String(raw).trim();
    if (!t) return null;
    var u = t.toUpperCase();
    if (/^(PK|PRE-?K|PREK|VPK)$/.test(u)) return "PK";
    if (/^(K|KG|KIN|KINDERGARTEN)$/.test(u)) return "K";
    var n = parseInt(t.replace(/^0+/, "") || t, 10);
    if (isNaN(n)) return null;
    if (n === 0) return "K";
    if (n >= 1 && n <= 9) return "0" + n;
    if (n >= 10 && n <= 12) return String(n);
    return null;
  }

  /** @returns {Object<string, true>} */
  function elementaryGradeAllowedSet(msidNum) {
    var o = {};
    if (msidNum === MEADOWLANE_PRIMARY_MSID) {
      o.K = true;
      o["01"] = true;
      o["02"] = true;
      return o;
    }
    if (msidNum === MEADOWLANE_INTERMEDIATE_MSID) {
      o["03"] = true;
      o["04"] = true;
      o["05"] = true;
      o["06"] = true;
      return o;
    }
    o.PK = true;
    o.K = true;
    for (var g = 1; g <= 6; g++) {
      o[g < 10 ? "0" + g : String(g)] = true;
    }
    return o;
  }

  /**
   * @param {string} gradeRaw
   * @param {Object|null} m master row for selected / scenario middle school
   * @param {boolean} scenarioMiddleZoned grade 07–08 only (MID_ zoning to scenario middle)
   */
  function studentGradeInSelectedSchoolBand(gradeRaw, m, scenarioMiddleZoned) {
    if (!m) return false;
    if (scenarioMiddleZoned) {
      var cm = canonicalStudentGradeCode(gradeRaw);
      return cm === "07" || cm === "08";
    }
    var g = canonicalStudentGradeCode(gradeRaw);
    if (!g) return false;
    var msidNum = parseInt(String(m.msid || "").trim(), 10);
    var lv = String(m.school_level || "").toLowerCase().trim();
    if (lv === "elementary") {
      var setE = elementaryGradeAllowedSet(msidNum);
      return !!setE[g];
    }
    if (lv === "middle") {
      return g === "07" || g === "08";
    }
    if (lv === "high") {
      return g === "09" || g === "10" || g === "11" || g === "12";
    }
    if (lv === "jr_sr_high") {
      return (
        g === "07" ||
        g === "08" ||
        g === "09" ||
        g === "10" ||
        g === "11" ||
        g === "12"
      );
    }
    return false;
  }

  function detailMatchesZonedTargetMsid(d, targetNum, schoolLevel) {
    var lv = String(schoolLevel || "").toLowerCase().trim();
    if (lv === "elementary") {
      return msidNormForZoning(d.ELEM_) === targetNum;
    }
    if (lv === "middle") {
      return msidNormForZoning(d.MID_) === targetNum;
    }
    if (lv === "high") {
      return msidNormForZoning(d.HIGH_) === targetNum;
    }
    if (lv === "jr_sr_high") {
      return (
        msidNormForZoning(d.MID_) === targetNum ||
        msidNormForZoning(d.INT_) === targetNum ||
        msidNormForZoning(d.HIGH_) === targetNum
      );
    }
    return false;
  }

  /**
   * Zoned assignment MSID for a student for aggregate charts (ELEM_ / MID_ / HIGH_ by grade band).
   * Does not use a “target” school — mirrors typical PK–6, 7–8, 9–12 column use in the layer.
   */
  function zonedMsidForDetailForAggregate(d) {
    if (!d) {
      return null;
    }
    var g = canonicalStudentGradeCode(d.Grade);
    if (!g) {
      return null;
    }
    if (g === "PK" || g === "K" || g === "01" || g === "02" || g === "03" || g === "04" || g === "05" || g === "06") {
      return msidNormForZoning(d.ELEM_);
    }
    if (g === "07" || g === "08") {
      return msidNormForZoning(d.MID_) || msidNormForZoning(d.INT_);
    }
    if (g === "09" || g === "10" || g === "11" || g === "12") {
      return msidNormForZoning(d.HIGH_) || msidNormForZoning(d.INT_);
    }
    return null;
  }

  /** Aligned with school_master lunch columns: blank / missing → Not free/reduced (same as existing & scenario views). */
  function normalizeSandboxLunchStatForPie(raw) {
    if (raw == null) {
      return "Not free/reduced";
    }
    var t = String(raw).trim();
    if (!t) {
      return "Not free/reduced";
    }
    var u = t.toLowerCase();
    if (u === "f" || u === "free") {
      return "Free";
    }
    if (u === "r" || u === "reduced" || u.indexOf("reduced") >= 0) {
      return "Reduced";
    }
    if (u === "n" || u.indexOf("not free") >= 0) {
      return "Not free/reduced";
    }
    if (u === "unspecified" || u === "unknown" || u === "—" || u === "-") {
      return "Not free/reduced";
    }
    return t;
  }

  /**
   * All students (any attendance MSID) zoned to target school in `m`'s level band,
   * or scenario middle (MID_ === target, grades 07–08).
   */
  function collectZonedDetailsByHex(targetMsid, m, scenarioMiddleZoned) {
    var out = {};
    if (
      !STUDENT_HEX_INDEX ||
      !STUDENT_HEX_INDEX.detailsByMsid ||
      targetMsid == null ||
      isNaN(targetMsid) ||
      !m
    ) {
      return out;
    }
    var tgt = Number(targetMsid);
    var lvl = String(m.school_level || "").toLowerCase().trim();
    if (scenarioMiddleZoned) {
      lvl = "middle";
    }
    var byDet = STUDENT_HEX_INDEX.detailsByMsid;
    for (var attMs in byDet) {
      if (!Object.prototype.hasOwnProperty.call(byDet, attMs)) continue;
      var hexMap = byDet[attMs];
      for (var hk in hexMap) {
        if (!Object.prototype.hasOwnProperty.call(hexMap, hk)) continue;
        var arr = hexMap[hk];
        if (!arr || !arr.length) continue;
        for (var i = 0; i < arr.length; i++) {
          var d = arr[i];
          if (!studentGradeInSelectedSchoolBand(d.Grade, m, scenarioMiddleZoned)) {
            continue;
          }
          if (!detailMatchesZonedTargetMsid(d, tgt, lvl)) continue;
          if (!out[hk]) out[hk] = [];
          out[hk].push(d);
        }
      }
    }
    return out;
  }

  /**
   * Sums per-hex student counts over every school MSID in the index (no selection filter).
   * @returns {Object<string, number>|null}
   */
  function buildAllSchoolsHexDisplayCountsByHex() {
    if (!STUDENT_HEX_INDEX || !STUDENT_HEX_INDEX.countsByMsid) {
      return null;
    }
    var byM = STUDENT_HEX_INDEX.countsByMsid;
    var out = Object.create(null);
    for (var msk in byM) {
      if (!Object.prototype.hasOwnProperty.call(byM, msk)) continue;
      var hmap = byM[msk];
      if (!hmap || typeof hmap !== "object") continue;
      for (var hk in hmap) {
        if (!Object.prototype.hasOwnProperty.call(hmap, hk)) continue;
        var c = Number(hmap[hk]) || 0;
        if (c <= 0) continue;
        out[hk] = (out[hk] || 0) + c;
      }
    }
    return Object.keys(out).length ? out : null;
  }

  /**
   * Per-hex counts for map overlay from attending / zoned toggles (union deduped per hex).
   * @returns {Object<string, number>|null} null = no overlay; object may be empty
   */
  function buildStudentHexDisplayCountsByHex() {
    var attEl = document.getElementById("toggle-student-hex-attending");
    var zonEl = document.getElementById("toggle-student-hex-zoned");
    var attOn = !attEl || attEl.checked;
    var zonedOn = !!(zonEl && zonEl.checked && !zonEl.disabled);

    if (!attOn && !zonedOn) {
      return null;
    }
    if (!STUDENT_HEX_INDEX || !STUDENT_HEX_INDEX.countsByMsid) {
      return null;
    }

    var panelScenario = document.getElementById("page-scenario");
    var onScenario = panelScenario && !panelScenario.hidden;
    var targetMsid = getActiveDashboardSchoolMsid();
    if (targetMsid == null || isNaN(targetMsid)) {
      if (zonedOn && !attOn) {
        return null;
      }
      if (attOn) {
        return buildAllSchoolsHexDisplayCountsByHex();
      }
      return null;
    }
    var m = masterRow(targetMsid);

    if (attOn && !zonedOn) {
      if (
        onScenario &&
        scenarioMiddleMsid != null &&
        !isNaN(scenarioMiddleMsid)
      ) {
        return buildMergedScenarioStudentHexCounts();
      }
      var simple = STUDENT_HEX_INDEX.countsByMsid[String(targetMsid)];
      return simple && typeof simple === "object" ? simple : null;
    }

    var perHexKeys = {};
    function touch(hk, dedupeKey) {
      if (!dedupeKey) return;
      if (!perHexKeys[hk]) perHexKeys[hk] = {};
      if (perHexKeys[hk][dedupeKey]) return;
      perHexKeys[hk][dedupeKey] = true;
    }

    if (attOn) {
      var detByHex =
        onScenario &&
        scenarioMiddleMsid != null &&
        !isNaN(scenarioMiddleMsid)
          ? buildMergedScenarioStudentHexDetailsByHex()
          : STUDENT_HEX_INDEX.detailsByMsid[String(targetMsid)] || {};
      for (var hka in detByHex) {
        if (!Object.prototype.hasOwnProperty.call(detByHex, hka)) continue;
        var arrA = detByHex[hka];
        if (!arrA) continue;
        for (var ia = 0; ia < arrA.length; ia++) {
          touch(hka, studentHexDedupeKey(arrA[ia]));
        }
      }
    }

    if (zonedOn && m) {
      var isScenarioZoned =
        !!(
          onScenario &&
          scenarioMiddleMsid != null &&
          !isNaN(scenarioMiddleMsid)
        );
      var zMap = collectZonedDetailsByHex(targetMsid, m, isScenarioZoned);
      for (var hkz in zMap) {
        if (!Object.prototype.hasOwnProperty.call(zMap, hkz)) continue;
        var arrZ = zMap[hkz];
        if (!arrZ) continue;
        for (var iz = 0; iz < arrZ.length; iz++) {
          touch(hkz, studentHexDedupeKey(arrZ[iz]));
        }
      }
    }

    var idxOut = {};
    for (var hkf in perHexKeys) {
      if (!Object.prototype.hasOwnProperty.call(perHexKeys, hkf)) continue;
      var bucket = perHexKeys[hkf];
      var n = 0;
      for (var kk in bucket) {
        if (Object.prototype.hasOwnProperty.call(bucket, kk)) n++;
      }
      if (n > 0) idxOut[hkf] = n;
    }
    return idxOut;
  }

  function syncStudentHexResidenceSubToggleAvailability() {
    var panelScenario = document.getElementById("page-scenario");
    var onScenario = panelScenario && !panelScenario.hidden;
    var msid = onScenario ? scenarioMiddleMsid : null;
    if (!onScenario) {
      var sel = document.getElementById("school-select");
      msid =
        sel && sel.value !== ""
          ? Number(sel.value)
          : null;
    }
    var dis =
      msid == null ||
      isNaN(msid) ||
      selectedSchoolDisallowsZonedStudentHex(msid);
    var zcb = document.getElementById("toggle-student-hex-zoned");
    if (zcb) {
      zcb.disabled = !!dis;
      if (dis) {
        zcb.checked = false;
      }
    }
  }

  function hexPolygonAreaSqMeters(geom) {
    if (
      typeof turf === "undefined" ||
      !turf ||
      typeof turf.area !== "function" ||
      !geom ||
      (geom.type !== "Polygon" && geom.type !== "MultiPolygon")
    ) {
      return null;
    }
    try {
      var sq = turf.area({ type: "Feature", geometry: geom });
      return sq != null && isFinite(sq) && sq > 0 ? sq : null;
    } catch (errA) {
      return null;
    }
  }

  /** Students per square mile for the student count placed across the hex polygon area. */
  function studentsPerSqMiFromCountAndGeom(count, geom) {
    if (count == null || count <= 0 || !geom) return null;
    var sqM = hexPolygonAreaSqMeters(geom);
    if (sqM == null || sqM <= 0) return null;
    var sqMi = sqM / SQ_METERS_PER_SQ_MI;
    if (!(sqMi > 0)) return null;
    var v = count / sqMi;
    if (!isFinite(v)) return null;
    return Math.round(v);
  }

  function formatStudentsPerSqMiForUi(v) {
    if (v == null || !isFinite(v)) return "—";
    return Math.round(Number(v)).toLocaleString();
  }

  /**
   * Mean of students per sq mi (including zeros) over the hovered hex and its geometric
   * neighbors, using the current school residence cohort counts and hex geometries.
   */
  function neighborhoodAverageSchoolResidenceStudentsPerSqMi(centerHexKey, prebuiltIdx) {
    if (!STUDENT_HEX_INDEX || !STUDENT_HEX_INDEX.geometryByHexKey) {
      return null;
    }
    var geomBy = STUDENT_HEX_INDEX.geometryByHexKey;
    var hk0 = String(centerHexKey);
    if (!Object.prototype.hasOwnProperty.call(geomBy, hk0)) {
      return null;
    }
    var nbrs =
      (STUDENT_HEX_INDEX.neighborsByHexKey && STUDENT_HEX_INDEX.neighborsByHexKey[hk0]) || [];
    var idx;
    if (prebuiltIdx !== undefined) {
      idx = prebuiltIdx;
    } else {
      idx = buildStudentHexDisplayCountsByHex();
    }
    if (idx == null) {
      idx = Object.create(null);
    }
    var totalD = 0;
    var nH = 0;
    var keysH = [hk0].concat(nbrs);
    var seenH = Object.create(null);
    for (var i = 0; i < keysH.length; i++) {
      var hk = keysH[i];
      if (seenH[hk]) {
        continue;
      }
      seenH[hk] = true;
      if (!Object.prototype.hasOwnProperty.call(geomBy, hk)) {
        continue;
      }
      nH += 1;
      var g = geomBy[hk];
      var cnt = 0;
      if (Object.prototype.hasOwnProperty.call(idx, hk)) {
        cnt = Number(idx[hk]) || 0;
      }
      if (cnt <= 0) {
        continue;
      }
      var dens = studentsPerSqMiFromCountAndGeom(cnt, g);
      totalD += dens != null && isFinite(dens) ? dens : 0;
    }
    if (nH === 0) {
      return null;
    }
    return Math.round(totalD / nH);
  }

  /**
   * Mean of charter students per sq mi (including zeros) over the hovered hex and
   * geometric neighbors, using `charterDistrictHexCounts`.
   */
  function neighborhoodAverageCharterResidenceStudentsPerSqMi(centerHexKey, prebuiltCh) {
    if (!STUDENT_HEX_INDEX || !STUDENT_HEX_INDEX.geometryByHexKey) {
      return null;
    }
    var geomBy = STUDENT_HEX_INDEX.geometryByHexKey;
    var ch;
    if (prebuiltCh !== undefined) {
      ch = prebuiltCh;
    } else {
      ch =
        (STUDENT_HEX_INDEX && STUDENT_HEX_INDEX.charterDistrictHexCounts) || Object.create(null);
    }
    var hk0 = String(centerHexKey);
    if (!Object.prototype.hasOwnProperty.call(geomBy, hk0)) {
      return null;
    }
    var nbrs =
      (STUDENT_HEX_INDEX.neighborsByHexKey && STUDENT_HEX_INDEX.neighborsByHexKey[hk0]) || [];
    var totalC = 0;
    var nC = 0;
    var keysC = [hk0].concat(nbrs);
    var seenC = Object.create(null);
    for (var j = 0; j < keysC.length; j++) {
      var hkc = keysC[j];
      if (seenC[hkc]) {
        continue;
      }
      seenC[hkc] = true;
      if (!Object.prototype.hasOwnProperty.call(geomBy, hkc)) {
        continue;
      }
      nC += 1;
      var gc = geomBy[hkc];
      var cnt2 = 0;
      if (Object.prototype.hasOwnProperty.call(ch, hkc)) {
        cnt2 = Number(ch[hkc]) || 0;
      }
      if (cnt2 <= 0) {
        continue;
      }
      var dens2 = studentsPerSqMiFromCountAndGeom(cnt2, gc);
      totalC += dens2 != null && isFinite(dens2) ? dens2 : 0;
    }
    if (nC === 0) {
      return null;
    }
    return Math.round(totalC / nC);
  }

  /**
   * Mean of homeschool students per sq mi (including zeros) over the hovered hex and
   * geometric neighbors, using aggregated `HOMESCHOOL_HEX_COUNTS`.
   */
  function neighborhoodAverageHomeschoolResidenceStudentsPerSqMi(centerHexKey, prebuiltHm) {
    var hm;
    if (prebuiltHm !== undefined) {
      hm = prebuiltHm;
    } else {
      hm = HOMESCHOOL_HEX_COUNTS || Object.create(null);
    }
    var hk0 = String(centerHexKey);
    if (!homeschoolHexGeometry(hk0)) {
      return null;
    }
    var nbrs =
      (STUDENT_HEX_INDEX &&
        STUDENT_HEX_INDEX.neighborsByHexKey &&
        STUDENT_HEX_INDEX.neighborsByHexKey[hk0]) ||
      [];
    var totalH = 0;
    var nH = 0;
    var keysH = [hk0].concat(nbrs);
    var seenH = Object.create(null);
    for (var j = 0; j < keysH.length; j++) {
      var hkh = keysH[j];
      if (seenH[hkh]) {
        continue;
      }
      seenH[hkh] = true;
      var gh = homeschoolHexGeometry(hkh);
      if (!gh) {
        continue;
      }
      nH += 1;
      var cnt2 = 0;
      if (Object.prototype.hasOwnProperty.call(hm, hkh)) {
        cnt2 = Number(hm[hkh]) || 0;
      }
      if (cnt2 <= 0) {
        continue;
      }
      var dens2 = studentsPerSqMiFromCountAndGeom(cnt2, gh);
      totalH += dens2 != null && isFinite(dens2) ? dens2 : 0;
    }
    if (nH === 0) {
      return null;
    }
    return Math.round(totalH / nH);
  }

  /** Short label (e.g. "McNair MS") for student-hex map tooltips; matches eseTableAbbreviatedSchoolName. */
  function studentResidenceTooltipSchoolLabel() {
    var msid = getActiveDashboardSchoolMsid();
    if (msid == null || isNaN(msid)) {
      return "selected school";
    }
    var m = masterRow(msid);
    if (!m || !m.school_name) {
      return "selected school";
    }
    var s = eseTableAbbreviatedSchoolName(m);
    s = s && String(s).trim() ? String(s).trim() : "";
    return s || "selected school";
  }

  function studentResidenceCohortTooltipPhrase() {
    var name = studentResidenceTooltipSchoolLabel();
    var attEl = document.getElementById("toggle-student-hex-attending");
    var zonEl = document.getElementById("toggle-student-hex-zoned");
    var attOn = !attEl || attEl.checked;
    var zonedOn = !!(zonEl && zonEl.checked && !zonEl.disabled);
    if (attOn && !zonedOn) {
      return "attending " + name;
    }
    if (zonedOn && !attOn) {
      return "zoned to " + name;
    }
    if (attOn && zonedOn) {
      return "attending or zoned to " + name;
    }
    return "selected cohort";
  }

  function studentHexResidenceHoverLinesHtml(props, cohortPhrase) {
    var showD;
    var hk = props && props._hexKey != null ? String(props._hexKey) : null;
    if (hk) {
      var aggB = neighborhoodAverageSchoolResidenceStudentsPerSqMi(hk);
      if (aggB != null && isFinite(aggB)) {
        showD = aggB;
      }
    }
    if (showD == null || !isFinite(showD)) {
      var rawD =
        props && props.students_per_sq_mi != null
          ? Number(props.students_per_sq_mi)
          : NaN;
      showD = rawD;
    }
    var rawC = props && props.count != null ? Number(props.count) : NaN;
    var phrase =
      cohortPhrase != null && String(cohortPhrase).trim() !== ""
        ? String(cohortPhrase).trim()
        : "selected cohort";
    var main =
      '<div class="student-hex-hover-line">' +
      '<span class="student-hex-hover-value">' +
      escapeHtml(formatStudentsPerSqMiForUi(showD)) +
      "</span>" +
      '<span class="student-hex-hover-unit"> grade-eligible students per square mile (' +
      escapeHtml(phrase) +
      ")</span>" +
      "</div>";
    var sub = "";
    if (!isNaN(rawC) && rawC > 3) {
      sub =
        '<div class="student-hex-hover-sub">' +
        escapeHtml(rawC.toLocaleString()) +
        " student residence" +
        (rawC === 1 ? "" : "s") +
        " in this hex</div>";
    }
    return main + sub;
  }

  function charterStudentHexResidenceLinesHtml(props) {
    var showC;
    var hkc1 = props && props._hexKey != null ? String(props._hexKey) : null;
    if (hkc1) {
      var aggC = neighborhoodAverageCharterResidenceStudentsPerSqMi(hkc1);
      if (aggC != null && isFinite(aggC)) {
        showC = aggC;
      }
    }
    if (showC == null || !isFinite(showC)) {
      var rawD0 =
        props && props.students_per_sq_mi != null
          ? Number(props.students_per_sq_mi)
          : NaN;
      showC = rawD0;
    }
    var rawC = props && props.count != null ? Number(props.count) : NaN;
    var mainLine =
      '<div class="student-hex-hover-line">' +
      '<span class="student-hex-hover-value">' +
      escapeHtml(formatStudentsPerSqMiForUi(showC)) +
      "</span>" +
      '<span class="student-hex-hover-unit"> grade-eligible charter students per square mile (districtwide)</span></div>';
    var sub = "";
    if (!isNaN(rawC) && rawC > 3) {
      var resWord;
      if (rawC === 1) {
        resWord = "1 charter student residence in this hex";
      } else {
        resWord =
          escapeHtml(rawC.toLocaleString()) + " charter student residences in this hex";
      }
      sub = '<div class="student-hex-hover-sub">' + resWord + "</div>";
    }
    return mainLine + sub;
  }

  function homeschoolStudentHexResidenceLinesHtml(props) {
    var showH;
    var hkh = props && props._hexKey != null ? String(props._hexKey) : null;
    if (hkh) {
      var aggH = neighborhoodAverageHomeschoolResidenceStudentsPerSqMi(hkh);
      if (aggH != null && isFinite(aggH)) {
        showH = aggH;
      }
    }
    if (showH == null || !isFinite(showH)) {
      var rawHd =
        props && props.students_per_sq_mi != null
          ? Number(props.students_per_sq_mi)
          : NaN;
      showH = rawHd;
    }
    var rawCnt = props && props.count != null ? Number(props.count) : NaN;
    var mainLine =
      '<div class="student-hex-hover-line">' +
      '<span class="student-hex-hover-value">' +
      escapeHtml(formatStudentsPerSqMiForUi(showH)) +
      "</span>" +
      '<span class="student-hex-hover-unit"> grade-eligible homeschool students per square mile (districtwide)</span></div>';
    var sub = "";
    if (!isNaN(rawCnt) && rawCnt > 3) {
      var resWord;
      if (rawCnt === 1) {
        resWord = "1 homeschool student residence in this hex";
      } else {
        resWord =
          escapeHtml(rawCnt.toLocaleString()) + " homeschool student residences in this hex";
      }
      sub = '<div class="student-hex-hover-sub">' + resWord + "</div>";
    }
    return mainLine + sub;
  }

  function studentHexResidenceHoverHtml(props, cohortPhrase) {
    return (
      '<div class="student-hex-hover-inner">' +
      studentHexResidenceHoverLinesHtml(props, cohortPhrase) +
      "</div>"
    );
  }

  function charterStudentHexResidenceHoverHtml(props) {
    return (
      '<div class="student-hex-hover-inner">' +
      charterStudentHexResidenceLinesHtml(props) +
      "</div>"
    );
  }

  function homeschoolStudentHexResidenceHoverHtml(props) {
    return (
      '<div class="student-hex-hover-inner">' +
      homeschoolStudentHexResidenceLinesHtml(props) +
      "</div>"
    );
  }

  function combinedResidenceHexHoverHtml(bProps, cProps, hProps, wantB, wantC, wantH, cohortPhrase) {
    var parts = [];
    var schoolShort = studentResidenceTooltipSchoolLabel();
    if (wantB && bProps) {
      parts.push(
        '<div class="student-hex-hover-section">' +
        '<div class="student-hex-hover-section-title">Selected School: ' +
        escapeHtml(schoolShort) +
        "</div>" +
        '<div class="student-hex-hover-inner">' +
        studentHexResidenceHoverLinesHtml(bProps, cohortPhrase) +
        "</div></div>"
      );
    }
    if (wantC && cProps) {
      parts.push(
        '<div class="student-hex-hover-section">' +
        '<div class="student-hex-hover-section-title">Charter (districtwide)</div>' +
        '<div class="student-hex-hover-inner">' +
        charterStudentHexResidenceLinesHtml(cProps) +
        "</div></div>"
      );
    }
    if (wantH && hProps) {
      parts.push(
        '<div class="student-hex-hover-section">' +
        '<div class="student-hex-hover-section-title">Homeschool (districtwide)</div>' +
        '<div class="student-hex-hover-inner">' +
        homeschoolStudentHexResidenceLinesHtml(hProps) +
        "</div></div>"
      );
    }
    if (parts.length) {
      return '<div class="student-hex-hover-dual">' + parts.join("") + "</div>";
    }
    return (
      '<div class="student-hex-hover-inner">No student residence data for the enabled layers at this location.</div>'
    );
  }

  /** Districtwide charter attendance (MSID 65xx–66xx) residential density; not tied to dropdown selection. */
  function syncCharterDistrictStudentHexLayer() {
    if (!map || !map.getSource || !map.getSource("charter-student-hex")) {
      return;
    }

    function emptyCharterHexAndHide() {
      map.getSource("charter-student-hex").setData({
        type: "FeatureCollection",
        features: [],
      });
      if (map.getSource("charter-student-hex-hit")) {
        map.getSource("charter-student-hex-hit").setData({
          type: "FeatureCollection",
          features: [],
        });
      }
      if (map.getLayer("charter-student-hex-hit-fill")) {
        map.setLayoutProperty("charter-student-hex-hit-fill", "visibility", "none");
      }
      syncResidenceDensityHeatmapZoomVisibility();
    }

    if (
      !STUDENT_HEX_INDEX ||
      !STUDENT_HEX_INDEX.charterDistrictHexCounts ||
      !STUDENT_HEX_INDEX.geometryByHexKey
    ) {
      emptyCharterHexAndHide();
      return;
    }

    var idx = STUDENT_HEX_INDEX.charterDistrictHexCounts;
    var features = [];
    var hitFeatures = [];
    for (var key in idx) {
      if (!Object.prototype.hasOwnProperty.call(idx, key)) continue;
      var cnt = idx[key];
      if (cnt <= 0) continue;
      var geom = STUDENT_HEX_INDEX.geometryByHexKey[key];
      if (!geom) continue;
      var pt = polygonCentroid(geom);
      if (!pt) continue;
      var dens = studentsPerSqMiFromCountAndGeom(cnt, geom);
      features.push({
        type: "Feature",
        properties: { _hexKey: key, count: cnt, students_per_sq_mi: dens },
        geometry: { type: "Point", coordinates: pt },
      });
      hitFeatures.push({
        type: "Feature",
        properties: {
          _hexKey: key,
          count: cnt,
          students_per_sq_mi: dens,
        },
        geometry: geom,
      });
    }
    if (features.length === 0) {
      emptyCharterHexAndHide();
      return;
    }
    map.getSource("charter-student-hex").setData({
      type: "FeatureCollection",
      features: features,
    });
    if (map.getSource("charter-student-hex-hit")) {
      map.getSource("charter-student-hex-hit").setData({
        type: "FeatureCollection",
        features: hitFeatures,
      });
    }
    var inp = document.getElementById("toggle-charter-student-hex");
    var vis = inp && inp.checked ? "visible" : "none";
    if (map.getLayer("charter-student-hex-hit-fill")) {
      map.setLayoutProperty("charter-student-hex-hit-fill", "visibility", vis);
    }
    syncResidenceDensityHeatmapZoomVisibility();
  }

  /** Districtwide homeschool residential density; hex geometries from main student hex index. */
  function syncHomeschoolStudentHexLayer() {
    if (!map || !map.getSource || !map.getSource("homeschool-student-hex")) {
      return;
    }

    function emptyHomeschoolHexAndHide() {
      map.getSource("homeschool-student-hex").setData({
        type: "FeatureCollection",
        features: [],
      });
      if (map.getSource("homeschool-student-hex-hit")) {
        map.getSource("homeschool-student-hex-hit").setData({
          type: "FeatureCollection",
          features: [],
        });
      }
      if (map.getLayer("homeschool-student-hex-hit-fill")) {
        map.setLayoutProperty("homeschool-student-hex-hit-fill", "visibility", "none");
      }
      syncResidenceDensityHeatmapZoomVisibility();
    }

    if (!HOMESCHOOL_HEX_COUNTS) {
      emptyHomeschoolHexAndHide();
      return;
    }

    var idx = HOMESCHOOL_HEX_COUNTS;
    var features = [];
    var hitFeatures = [];
    for (var key in idx) {
      if (!Object.prototype.hasOwnProperty.call(idx, key)) continue;
      var cnt = idx[key];
      if (cnt <= 0) continue;
      var geom = homeschoolHexGeometry(key);
      if (!geom) continue;
      var pt = polygonCentroid(geom);
      if (!pt) continue;
      var dens = studentsPerSqMiFromCountAndGeom(cnt, geom);
      features.push({
        type: "Feature",
        properties: { _hexKey: key, count: cnt, students_per_sq_mi: dens },
        geometry: { type: "Point", coordinates: pt },
      });
      hitFeatures.push({
        type: "Feature",
        properties: {
          _hexKey: key,
          count: cnt,
          students_per_sq_mi: dens,
        },
        geometry: geom,
      });
    }
    if (features.length === 0) {
      emptyHomeschoolHexAndHide();
      return;
    }
    map.getSource("homeschool-student-hex").setData({
      type: "FeatureCollection",
      features: features,
    });
    if (map.getSource("homeschool-student-hex-hit")) {
      map.getSource("homeschool-student-hex-hit").setData({
        type: "FeatureCollection",
        features: hitFeatures,
      });
    }
    var inp = document.getElementById("toggle-homeschool-student-hex");
    var vis = inp && inp.checked ? "visible" : "none";
    if (map.getLayer("homeschool-student-hex-hit-fill")) {
      map.setLayoutProperty("homeschool-student-hex-hit-fill", "visibility", vis);
    }
    syncResidenceDensityHeatmapZoomVisibility();
  }

  function syncStudentHexLayer() {
    if (!map || !map.getSource || !map.getSource("student-hex")) return;
    syncStudentHexResidenceSubToggleAvailability();

    try {
      function emptyStudentHexSourcesAndHide() {
        map.getSource("student-hex").setData({
          type: "FeatureCollection",
          features: [],
        });
        if (map.getSource("student-hex-hit")) {
          map.getSource("student-hex-hit").setData({
            type: "FeatureCollection",
            features: [],
          });
        }
        if (map.getLayer("student-hex-hit-fill")) {
          map.setLayoutProperty("student-hex-hit-fill", "visibility", "none");
        }
      }

      if (!STUDENT_HEX_INDEX || !STUDENT_HEX_INDEX.countsByMsid) {
        emptyStudentHexSourcesAndHide();
        return;
      }
      var idx = buildStudentHexDisplayCountsByHex();
      if (idx == null || typeof idx !== "object") {
        emptyStudentHexSourcesAndHide();
        return;
      }
      var features = [];
      var hitFeatures = [];
      for (var key in idx) {
        if (!Object.prototype.hasOwnProperty.call(idx, key)) continue;
        var cnt = idx[key];
        if (cnt <= 0) continue;
        var geom = STUDENT_HEX_INDEX.geometryByHexKey[key];
        if (!geom) continue;
        var pt = polygonCentroid(geom);
        if (!pt) continue;
        var dens = studentsPerSqMiFromCountAndGeom(cnt, geom);
        features.push({
          type: "Feature",
          properties: { _hexKey: key, count: cnt, students_per_sq_mi: dens },
          geometry: { type: "Point", coordinates: pt },
        });
        hitFeatures.push({
          type: "Feature",
          properties: {
            _hexKey: key,
            count: cnt,
            students_per_sq_mi: dens,
          },
          geometry: geom,
        });
      }
      if (features.length === 0) {
        emptyStudentHexSourcesAndHide();
        return;
      }
      map.getSource("student-hex").setData({
        type: "FeatureCollection",
        features: features,
      });
      if (map.getSource("student-hex-hit")) {
        map.getSource("student-hex-hit").setData({
          type: "FeatureCollection",
          features: hitFeatures,
        });
      }
      var showHex = isStudentResidenceLayerEnabled();
      var vis = showHex ? "visible" : "none";
      if (map.getLayer("student-hex-hit-fill")) {
        map.setLayoutProperty("student-hex-hit-fill", "visibility", vis);
      }
    } finally {
      syncCharterDistrictStudentHexLayer();
      syncHomeschoolStudentHexLayer();
    }
    scheduleRefreshMapDensityLegendValueRanges();
    applyResidenceHeatmapSymbology();
  }

  /** Draggable vertical splitter between data panel and map. */
  function initDashboardResizer(map) {
    var dashboard = document.getElementById("dashboard");
    var sidebar = document.getElementById("dashboard-sidebar");
    var resizer = document.getElementById("dashboard-resizer");
    if (!dashboard || !sidebar || !resizer) return;

    var dragging = false;

    function clampSidebarWidth(px) {
      var rect = dashboard.getBoundingClientRect();
      var resizerW = resizer.offsetWidth || 8;
      var minSide = 240;
      var minMap = 280;
      var max = rect.width - resizerW - minMap;
      return Math.max(minSide, Math.min(max, px));
    }

    function setSidebarWidth(px) {
      px = clampSidebarWidth(px);
      sidebar.style.flex = "0 0 " + px + "px";
      sidebar.style.width = px + "px";
      map.resize();
    }

    resizer.addEventListener("mousedown", function (e) {
      dragging = true;
      e.preventDefault();
      document.body.style.cursor = "col-resize";
      document.body.style.userSelect = "none";
    });

    document.addEventListener("mousemove", function (e) {
      if (!dragging) return;
      var rect = dashboard.getBoundingClientRect();
      setSidebarWidth(e.clientX - rect.left);
    });

    document.addEventListener("mouseup", function () {
      if (!dragging) return;
      dragging = false;
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
      map.resize();
    });

    resizer.addEventListener("keydown", function (e) {
      var step = 24;
      var current = sidebar.getBoundingClientRect().width;
      if (e.key === "ArrowLeft") {
        e.preventDefault();
        setSidebarWidth(current - step);
      } else if (e.key === "ArrowRight") {
        e.preventDefault();
        setSidebarWidth(current + step);
      }
    });

    window.addEventListener("resize", function () {
      if (window.innerWidth <= 960) {
        map.resize();
        return;
      }
      var rect = dashboard.getBoundingClientRect();
      var sw = sidebar.getBoundingClientRect().width;
      var resizerW = resizer.offsetWidth || 8;
      if (sw + resizerW > rect.width - 200) {
        setSidebarWidth((rect.width - resizerW) * 0.5);
      } else {
        map.resize();
      }
    });
  }

  (function initToolbar() {
    var btn = document.getElementById("toolbar-toggle");
    var toolbar = document.getElementById("toolbar");
    if (!btn || !toolbar) return;
    btn.addEventListener("click", function () {
      var collapsed = toolbar.classList.toggle("toolbar--collapsed");
      btn.setAttribute("aria-expanded", collapsed ? "false" : "true");
    });
  })();

  (function setupScenarioMergerControl() {
    var el = document.getElementById("scenario-complete-merger");
    if (!el) return;
    el.checked = false;
    scenarioCompleteMerger = false;
    el.addEventListener("change", function () {
      scenarioCompleteMerger = el.checked;
      applyScenarioMergedUpdates();
      if (
        scenarioMiddleMsid != null &&
        !isNaN(scenarioMiddleMsid) &&
        scenarioLastFeederRows.length
      ) {
        renderScenarioFeederList(
          scenarioMiddleMsid,
          scenarioLastFeederRows
        );
      }
    });
  })();

  (function setupPageSwitcher() {
    var titleEl = document.getElementById("sidebar-view-title");
    var tabExisting = document.getElementById("page-tab-existing");
    var tabScenario = document.getElementById("page-tab-scenario");
    var tabSandbox = document.getElementById("page-tab-sandbox");
    var panelExisting = document.getElementById("page-existing");
    var panelScenario = document.getElementById("page-scenario");
    var panelSandbox = document.getElementById("page-sandbox");
    if (
      !tabExisting ||
      !tabScenario ||
      !tabSandbox ||
      !panelExisting ||
      !panelScenario ||
      !panelSandbox
    ) {
      return;
    }

    var labels = {
      existing: "Existing Conditions",
      scenario: "Scenario Planning",
      sandbox: "Boundary Sandbox",
    };

    function setPage(page) {
      var p =
        page === "existing" ? "existing" : page === "scenario" ? "scenario" : "sandbox";
      if (titleEl) {
        titleEl.textContent = labels[p] || labels.existing;
      }
      var isExisting = p === "existing";
      var isScenario = p === "scenario";
      var isSandbox = p === "sandbox";
      tabExisting.setAttribute("aria-selected", isExisting ? "true" : "false");
      tabScenario.setAttribute("aria-selected", isScenario ? "true" : "false");
      tabSandbox.setAttribute("aria-selected", isSandbox ? "true" : "false");
      tabExisting.classList.toggle("is-active", isExisting);
      tabScenario.classList.toggle("is-active", isScenario);
      tabSandbox.classList.toggle("is-active", isSandbox);
      panelExisting.hidden = !isExisting;
      panelScenario.hidden = !isScenario;
      panelSandbox.hidden = !isSandbox;
      if (isScenario) {
        refreshScenarioPanelIfVisible();
      }
      applyScenarioFeederMapHighlights();
      syncStudentHexLayer();
      syncTravelShedLayerFilter();
      syncBoundarySandboxMapLayers();
      if (p === "sandbox") {
        updateSandboxSelectedHexCountUi();
      }
      requestAnimationFrame(function () {
        if (map && typeof map.resize === "function") {
          map.resize();
        }
      });
    }

    tabExisting.addEventListener("click", function () {
      setPage("existing");
    });
    tabScenario.addEventListener("click", function () {
      setPage("scenario");
    });
    tabSandbox.addEventListener("click", function () {
      setPage("sandbox");
    });
  })();

  (function setupBoundarySandboxBaseSchoolChange() {
    var sb = document.getElementById("sandbox-base-school");
    if (!sb) {
      return;
    }
    sb.addEventListener("change", function () {
      try {
        var raw = sb.value;
        var ms = raw !== "" && raw != null ? Number(raw) : null;
        if (raw !== "" && raw != null && !isNaN(ms)) {
          clearBoundarySandboxLassoRegionFill();
          prefillBoundarySandboxZonedHexesForBaseMsid(ms);
          requestApplyBoundarySandboxSelectionOnIdle();
        }
        syncTravelShedLayerFilter();
        syncStudentHexLayer();
        updateSandboxSelectedHexCountUi();
      } catch (eSb) {
        /* ignore */
      }
    });
  })();

  (function setupBoundarySandboxConfirm() {
    var cbtn = document.getElementById("sandbox-confirm-btn");
    if (!cbtn) {
      return;
    }
    cbtn.addEventListener("click", function () {
      if (cbtn.getAttribute("aria-disabled") === "true") {
        return;
      }
      BOUNDARY_SANDBOX.confirmedHexKeysSnapshot = shallowCopyHexKeyBag(BOUNDARY_SANDBOX.selectedHexKeys);
      BOUNDARY_SANDBOX.selectionConfirmed = true;
      updateSandboxSelectedHexCountUi();
      updateBoundarySandboxSelectionOutline();
    });
  })();

  (function setupBoundarySandboxClearButton() {
    var cl = document.getElementById("sandbox-clear-btn");
    if (!cl) {
      return;
    }
    cl.addEventListener("click", function () {
      clearBoundarySandboxGeographicSelection();
    });
  })();

  (function setupBoundarySandboxGradeAndSchoolListUi() {
    var gB = document.getElementById("sandbox-card-body-grade");
    if (gB) {
      gB.addEventListener("change", function (e) {
        var t = e.target;
        if (!t || !t.classList) {
          return;
        }
        if (t.classList.contains("sandbox-grade-select-all")) {
          var wantAll = t.checked;
          BOUNDARY_SANDBOX.gradeToggles = BOUNDARY_SANDBOX.gradeToggles || Object.create(null);
          var rowInputs = gB.querySelectorAll("input.sandbox-grade-toggle[data-grade-canon]");
          for (var si = 0; si < rowInputs.length; si++) {
            var bx = rowInputs[si];
            var gcx = bx.getAttribute("data-grade-canon");
            if (gcx == null) {
              continue;
            }
            BOUNDARY_SANDBOX.gradeToggles[gcx] = wantAll;
          }
          updateSandboxStatsPanelSummary();
          return;
        }
        if (!t.classList.contains("sandbox-grade-toggle")) {
          return;
        }
        var gc = t.getAttribute("data-grade-canon");
        if (gc == null) {
          return;
        }
        BOUNDARY_SANDBOX.gradeToggles = BOUNDARY_SANDBOX.gradeToggles || Object.create(null);
        BOUNDARY_SANDBOX.gradeToggles[gc] = t.checked;
        updateSandboxStatsPanelSummary();
      });
    }
    var aT = document.getElementById("sandbox-card-body-attendance-type");
    if (aT) {
      aT.addEventListener("change", function (e) {
        var t2 = e.target;
        if (!t2 || !t2.classList || !t2.classList.contains("sandbox-attendance-type-toggle")) {
          return;
        }
        var atp = t2.getAttribute("data-atype");
        if (atp == null) {
          return;
        }
        BOUNDARY_SANDBOX.attendanceTypeToggles =
          BOUNDARY_SANDBOX.attendanceTypeToggles || Object.create(null);
        BOUNDARY_SANDBOX.attendanceTypeToggles[atp] = t2.checked;
        updateSandboxStatsPanelSummary();
      });
    }
    var pSand = document.getElementById("page-sandbox");
    if (pSand) {
      pSand.addEventListener("click", function (e) {
        var t = e.target;
        if (!t || !t.classList || !t.classList.contains("sandbox-school-expand")) {
          return;
        }
        e.preventDefault();
        var pan = t.getAttribute("data-panel");
        if (!pan) {
          return;
        }
        BOUNDARY_SANDBOX.schoolListExpanded = BOUNDARY_SANDBOX.schoolListExpanded || {
          attendance: false,
          zoned: false,
        };
        BOUNDARY_SANDBOX.schoolListExpanded[pan] = !BOUNDARY_SANDBOX.schoolListExpanded[pan];
        updateSandboxStatsPanelSummary();
      });
    }
  })();

  (function setupSchoolMasterCsvDownload() {
    var btn = document.getElementById("download-school-master-btn");
    if (!btn) return;
    btn.addEventListener("click", function () {
      fetch(DATA.masterCsv)
        .then(function (r) {
          if (!r.ok) throw new Error("HTTP " + r.status);
          return r.text();
        })
        .then(function (text) {
          var grid = parseCsvRows(text);
          grid = filterCsvGridToDropdownSchools(grid);
          grid = applyChoiceSchoolCaptureToCsvGrid(grid);
          grid = applyBpsEmployeeCountToCsvGrid(grid);
          var out = grid.map(joinCsvQuotedRow).join("\r\n");
          var blob = new Blob([out], { type: "text/csv;charset=utf-8" });
          var url = URL.createObjectURL(blob);
          var a = document.createElement("a");
          a.href = url;
          a.download = "school_master.csv";
          a.setAttribute("aria-hidden", "true");
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(url);
        })
        .catch(function () {
          alert(
            "Could not download school_master.csv. Serve this folder over HTTP (e.g. Live Server), not as a file:// URL."
          );
        });
    });
  })();

  syncSandboxConfirmEditButtonStates();
})();
