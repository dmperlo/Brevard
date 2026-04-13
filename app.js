(function () {
  "use strict";

  var DATA = {
    es: "geo/ESBoundaries.json",
    ms: "geo/MSBoundaries.json",
    hs: "geo/HSBoundaries.json",
    schools: "geo/SchoolLocations.json",
    enrollment: "data/processed/enrollment.json",
    facilityAge: "data/processed/facility_age.json",
    demographics: "data/processed/demographics_by_msid.json",
    capture: "data/processed/capture_by_msid.json",
    sankeyEsMs: "data/processed/sankey_es_ms.json",
    studentHexagons: "geo/StudentHexagons.geojson",
    schoolParcels: "geo/SchoolParcels.geojson",
  };

  /** Set after GeoJSON loads; used to zoom to assignment boundaries. */
  var GEO_CACHE = { es: null, ms: null, hs: null, schools: null };
  /** From export of projected enrollment workbook; null if missing or failed to load. */
  var ENROLLMENT_CACHE = null;
  /** From Age of all Facilities xls (by school name); null if missing or failed to load. */
  var FACILITY_CACHE = null;
  /** Student SY2025-26 aggregates by MSID (ethnicity, lunch); null if missing. */
  var DEMOGRAPHICS_CACHE = null;
  /** Capture rate by MSID and level (elementary/middle/high); null if missing. */
  var CAPTURE_CACHE = null;
  /** ES→MS flows from SankeyFlowHelper export; null if missing. */
  var SANKEY_CACHE = null;
  /** Pre-aggregated student hex counts by school MSID (from one polygon per student). */
  var STUDENT_HEX_INDEX = null;
  /** Dropdown-driven selection; map clicks do not change this. */
  var selectedSchoolMsid = null;
  /** { source, id } for assignment outline emphasis when a school is chosen from the dropdown. */
  var selectedAssignmentBoundary = null;

  /** Scenario Testing: merged K–8 tool state (middle MS + feeder checkboxes). */
  var scenarioSchoolByMsid = null;
  var scenarioMiddleMsid = null;
  var scenarioLastFeederRows = [];
  var scenarioFeederChecked = {};
  /** MSIDs last given map feature-state `scenarioFeeder`; cleared before each update. */
  var lastScenarioFeederHighlightMsids = [];
  /** When true, each selected elementary counts at 100%; when false (default), use flow proportion × enrollment. */
  var scenarioCompleteMerger = false;
  /** Set to false to restore the single aggregated bar chart on the Scenario page. */
  var SCENARIO_USE_STACKED_ENROLLMENT_CHART = true;

  var ENCHART_COLORS = { calendar: "#94a3b8", projected: "#93c5fd" };

  /** Matches school location dot colors (elementary / middle / high). */
  var PALETTE = {
    elementary: { fill: "#16a34a", line: "#15803d" },
    middle: { fill: "#2563eb", line: "#1d4ed8" },
    high: { fill: "#9333ea", line: "#7e22ce" },
  };

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

  /** Raster basemaps: Esri Light Gray Canvas (base + reference), same family as “Light Gray Reference” in ArcGIS. */
  var BASEMAP_STYLE = {
    version: 8,
    sources: {
      "basemap-gray-base": {
        type: "raster",
        tiles: [
          "https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{z}/{y}/{x}",
        ],
        tileSize: 256,
        attribution:
          '<a href="https://www.esri.com/">Esri</a> — Light Gray Canvas',
      },
      "basemap-gray-reference": {
        type: "raster",
        tiles: [
          "https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Reference/MapServer/tile/{z}/{y}/{x}",
        ],
        tileSize: 256,
        attribution: "",
      },
      "basemap-satellite": {
        type: "raster",
        tiles: [
          "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
        ],
        tileSize: 256,
        attribution:
          '<a href="https://www.esri.com/">Esri</a> — World Imagery',
      },
    },
    layers: [
      {
        id: "basemap-gray-base",
        type: "raster",
        source: "basemap-gray-base",
        minzoom: 0,
        maxzoom: 22,
      },
      {
        id: "basemap-gray-reference",
        type: "raster",
        source: "basemap-gray-reference",
        minzoom: 0,
        maxzoom: 22,
      },
      {
        id: "basemap-satellite",
        type: "raster",
        source: "basemap-satellite",
        minzoom: 0,
        maxzoom: 22,
        layout: { visibility: "none" },
      },
    ],
  };

  var map = new maplibregl.Map({
    container: "map",
    style: BASEMAP_STYLE,
    center: [-80.7, 28.2],
    zoom: 8,
    maxZoom: 19,
  });

  map.addControl(new maplibregl.NavigationControl(), "top-left");
  map.addControl(new maplibregl.AttributionControl({ compact: true }), "bottom-right");

  function setBasemap(mode) {
    var streets = mode === "streets";
    ["basemap-gray-base", "basemap-gray-reference"].forEach(function (id) {
      if (map.getLayer(id)) {
        map.setLayoutProperty(id, "visibility", streets ? "visible" : "none");
      }
    });
    if (map.getLayer("basemap-satellite")) {
      map.setLayoutProperty("basemap-satellite", "visibility", streets ? "none" : "visible");
    }
    var root = document.getElementById("basemap-toggle");
    if (root) {
      root.querySelectorAll("[data-basemap]").forEach(function (btn) {
        var active = btn.getAttribute("data-basemap") === mode;
        btn.classList.toggle("is-active", active);
        btn.setAttribute("aria-pressed", active ? "true" : "false");
      });
    }
  }

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

  map.on("load", function () {
    var basemapRoot = document.getElementById("basemap-toggle");
    if (basemapRoot) {
      basemapRoot.querySelectorAll("[data-basemap]").forEach(function (btn) {
        btn.addEventListener("click", function () {
          var mode = btn.getAttribute("data-basemap");
          if (mode === "streets" || mode === "satellite") setBasemap(mode);
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
      fetch(DATA.enrollment)
        .then(function (r) {
          return r.ok ? r.json() : null;
        })
        .catch(function () {
          return null;
        }),
      fetch(DATA.facilityAge)
        .then(function (r) {
          return r.ok ? r.json() : null;
        })
        .catch(function () {
          return null;
        }),
      fetch(DATA.demographics)
        .then(function (r) {
          return r.ok ? r.json() : null;
        })
        .catch(function () {
          return null;
        }),
      fetch(DATA.capture)
        .then(function (r) {
          return r.ok ? r.json() : null;
        })
        .catch(function () {
          return null;
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
    ])
      .then(function (results) {
        var es = results[0];
        var ms = results[1];
        var hs = results[2];
        var schools = results[3];
        ENROLLMENT_CACHE = results[4];
        FACILITY_CACHE = results[5];
        DEMOGRAPHICS_CACHE = results[6];
        CAPTURE_CACHE = results[7];
        SANKEY_CACHE = results[8];
        var studentHexFc = results[9];
        var schoolParcelsRaw = results[10];

        var boundarySourceOpts = { type: "geojson", promoteId: "MSID" };

        map.addSource("es-boundaries", Object.assign({ data: es }, boundarySourceOpts));
        map.addSource("ms-boundaries", Object.assign({ data: ms }, boundarySourceOpts));
        map.addSource("hs-boundaries", Object.assign({ data: hs }, boundarySourceOpts));
        map.addSource("schools", {
          type: "geojson",
          data: schools,
          promoteId: "SCHOOLS_ID",
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
          id: "school-parcels-elementary",
          type: "line",
          source: "school-parcels",
          filter: ["==", ["get", "_parcelLevel"], "elementary"],
          layout: schoolParcelLineLayout,
          paint: Object.assign({}, schoolParcelLinePaintBase, {
            "line-color": PALETTE.elementary.line,
          }),
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
            "heatmap-weight": ["get", "count"],
            "heatmap-intensity": [
              "interpolate",
              ["linear"],
              ["zoom"],
              8,
              0.28,
              11,
              0.38,
              14,
              0.5,
            ],
            "heatmap-radius": [
              "interpolate",
              ["linear"],
              ["zoom"],
              8,
              18,
              11,
              32,
              14,
              48,
              17,
              62,
            ],
            "heatmap-opacity": 0.88,
            "heatmap-color": [
              "interpolate",
              ["linear"],
              ["heatmap-density"],
              0,
              "rgba(0, 0, 0, 0)",
              0.04,
              "rgba(127, 205, 187, 0.32)",
              0.1,
              "rgba(90, 198, 194, 0.45)",
              0.16,
              "rgba(65, 182, 196, 0.55)",
              0.22,
              "rgba(56, 178, 172, 0.62)",
              0.28,
              "rgba(99, 102, 241, 0.58)",
              0.34,
              "rgba(139, 92, 246, 0.7)",
              0.4,
              "rgba(168, 85, 247, 0.76)",
              0.46,
              "rgba(192, 38, 211, 0.8)",
              0.52,
              "rgba(197, 27, 125, 0.84)",
              0.58,
              "rgba(219, 39, 119, 0.87)",
              0.64,
              "rgba(225, 29, 72, 0.89)",
              0.7,
              "rgba(220, 38, 38, 0.91)",
              0.76,
              "rgba(220, 60, 30, 0.92)",
              0.81,
              "rgba(234, 88, 12, 0.93)",
              0.85,
              "rgba(234, 95, 20, 0.94)",
              0.88,
              "rgba(245, 120, 20, 0.94)",
              0.91,
              "rgba(249, 130, 25, 0.95)",
              0.93,
              "rgba(251, 146, 40, 0.95)",
              0.95,
              "rgba(253, 170, 55, 0.96)",
              0.965,
              "rgba(255, 200, 85, 0.97)",
              0.98,
              "rgba(255, 228, 120, 0.98)",
              0.992,
              "rgba(255, 248, 170, 0.99)",
              1,
              "rgba(255, 255, 210, 1)",
            ],
          },
          layout: { visibility: "none" },
        });

        var schoolCirclePaint = {
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
            [
              "any",
              ["==", ["feature-state", "ring"], true],
              ["==", ["feature-state", "selected"], true],
              ["==", ["feature-state", "scenarioFeeder"], true],
            ],
            5,
            1,
          ],
          "circle-stroke-color": "#ffffff",
          "circle-opacity": 0.92,
        };

        map.addLayer({
          id: "schools-elementary",
          type: "circle",
          source: "schools",
          filter: ["==", ["get", "TYPE"], "ELEMENTARY"],
          paint: Object.assign({}, schoolCirclePaint, {
            "circle-color": PALETTE.elementary.fill,
          }),
        });
        map.addLayer({
          id: "schools-middle",
          type: "circle",
          source: "schools",
          filter: ["==", ["get", "TYPE"], "MIDDLE"],
          paint: Object.assign({}, schoolCirclePaint, {
            "circle-color": PALETTE.middle.fill,
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
          paint: Object.assign({}, schoolCirclePaint, {
            "circle-color": [
              "match",
              ["get", "TYPE"],
              "HIGH",
              PALETTE.high.fill,
              "JR SR HIGH",
              "#ea580c",
              PALETTE.high.fill,
            ],
          }),
        });

        ["schools-elementary", "schools-middle", "schools-high"].forEach(function (lid) {
          if (map.getLayer(lid)) map.moveLayer(lid);
        });

        var combined = null;
        combined = mergeBbox(combined, computeBbox(es));
        combined = mergeBbox(combined, computeBbox(ms));
        combined = mergeBbox(combined, computeBbox(hs));
        combined = mergeBbox(combined, computeBbox(schools));
        combined = mergeBbox(combined, computeBbox(schoolParcelsFc));

        map.resize();
        if (combined) {
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
        } else {
          STUDENT_HEX_INDEX = null;
        }

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
    var ab = String(sp.SchAB_Type || "").toUpperCase();
    if (ab === "CHOICE") return true;
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
      var lvl = schoolParcelLevelFromType(sp);
      if (!lvl) continue;
      var geom = ft.geometry;
      if (!geom || (geom.type !== "Polygon" && geom.type !== "MultiPolygon")) {
        continue;
      }
      out.features.push({
        type: "Feature",
        geometry: geom,
        properties: { _parcelLevel: lvl },
      });
    }
    return out;
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
      gs.className = "toggle-gradient-strip";
      gs.setAttribute("aria-hidden", "true");
      label.appendChild(gs);
    } else if (def.swatchColor) {
      var sw = document.createElement("span");
      sw.className = "swatch";
      sw.style.background = def.swatchColor;
      sw.setAttribute("aria-hidden", "true");
      label.appendChild(sw);
    }
    label.appendChild(document.createTextNode(def.label));
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
        layerIds: ["schools-high"],
        swatchColor: PALETTE.high.fill,
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
        layerIds: ["school-parcels-high"],
        swatchColor: PALETTE.high.fill,
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
          layerIds: ["student-hex-heatmap"],
          gradientStrip: true,
          defaultChecked: false,
        },
        function () {
          var inp = document.getElementById("toggle-student-hex");
          if (inp && inp.checked) {
            syncStudentHexLayer();
          }
        }
      );
    }
  }

  var BOUNDARY_FILL_LAYERS = ["es-fill", "ms-fill", "hs-fill"];
  var SCHOOL_LAYER_IDS = ["schools-elementary", "schools-middle", "schools-high"];

  function fillLayerIdToSource(layerId) {
    if (layerId === "es-fill") return "es-boundaries";
    if (layerId === "ms-fill") return "ms-boundaries";
    if (layerId === "hs-fill") return "hs-boundaries";
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

  /** "CITY, ST 12345" → "City, ST 12345" */
  function formatCityStateZip(str) {
    if (!str) return "";
    var t = String(str).trim();
    var m = t.match(/^(.+),\s*([A-Za-z]{2})\s+(.+)$/);
    if (m) {
      return (
        standardCapitalization(m[1].trim()) +
        ", " +
        m[2].toUpperCase() +
        " " +
        m[3].trim()
      );
    }
    return standardCapitalization(t);
  }

  function schoolDetailHtml(p) {
    var rawName = p.NAME || p.CommonName || "School";
    var name = standardCapitalization(expandElemSchoolName(rawName));
    var grades = p.Grades || "";
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

  /** Shown first in the school dropdown; order is Johnson, McNair, Stone. MSIDs match SCHOOLS_ID in GeoJSON. */
  var PRIORITY_SCHOOL_MSIDS = [3031, 1081, 2071];

  function schoolNameForSelect(p) {
    return standardCapitalization(
      expandElemSchoolName(p.NAME || p.CommonName || "School")
    );
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
        var na = expandElemSchoolName(a.NAME || a.CommonName || "").toLowerCase();
        var nb = expandElemSchoolName(b.NAME || b.CommonName || "").toLowerCase();
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

  function zoomToSchoolAssignment(msid, schoolByMsid) {
    var boundaryFt = findBoundaryFeatureForMsid(msid);
    var bbox;
    if (boundaryFt) {
      bbox = computeBbox({
        type: "FeatureCollection",
        features: [boundaryFt],
      });
    } else {
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
      if (lon == null || lat == null || isNaN(lon) || isNaN(lat)) return;
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
    if (!panelScenario || panelScenario.hidden) return;
    if (
      scenarioMiddleMsid == null ||
      isNaN(scenarioMiddleMsid) ||
      !scenarioSchoolByMsid
    ) {
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

  /** @returns {null | { yearSchoolOpened: number|null, yearPropertyPurchased: number|null, ageAsOf2026: number|null }} */
  function lookupFacilityForSchool(p) {
    if (!FACILITY_CACHE || !FACILITY_CACHE.byNameKey || !p) return null;
    var n = normalizeSchoolNameKey(p.NAME);
    var keys = [
      n,
      normalizeSchoolNameKey(p.CommonName),
    ];
    if (n && n.indexOf("JR SR") >= 0 && n.indexOf("HIGH") < 0) {
      keys.push(n + " HIGH");
    }
    for (var i = 0; i < keys.length; i++) {
      if (!keys[i]) continue;
      var row = FACILITY_CACHE.byNameKey[keys[i]];
      if (row) return row;
    }
    return null;
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
    var L = normalizeSchoolNameKey(label);
    var cn = normalizeSchoolNameKey(p.CommonName || "");
    var nm = normalizeSchoolNameKey(p.NAME || "");
    if (!L) return false;
    if (cn && (L === cn || nm.indexOf(L) !== -1 || cn.indexOf(L) !== -1))
      return true;
    var parts = nm.split(" ").filter(Boolean);
    if (parts.length && L === parts[0]) return true;
    return false;
  }

  /**
   * @returns {{ elementary: string, middle: string, value: number, emphasis: boolean }[]}
   *   emphasis = flow is a primary focus for the selection (all ES→MS links when ES selected;
   *   ES→selected-MS when middle selected; other MS destinations from same feeders are emphasis:false).
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
    return [];
  }

  function sankeyHighLabelMatchesSchool(label, p) {
    var L = normalizeSchoolNameKey(label);
    var cn = normalizeSchoolNameKey(p.CommonName || "");
    var nm = normalizeSchoolNameKey(p.NAME || "");
    if (!L) return false;
    if (cn && (L === cn || nm.indexOf(L) !== -1 || cn.indexOf(L) !== -1))
      return true;
    var parts = nm.split(" ").filter(Boolean);
    if (parts.length && L === parts[0]) return true;
    return false;
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

  /**
   * @param {{ from: string, to: string, value: number, emphasis: boolean }[]} normFlows
   * @param {{ leftFill: string, rightFill: string, emphStroke: string, ariaLabel: string, secondaryTooltip: string }} cfg
   */
  function renderBipartiteSankey(root, normFlows, cfg) {
    if (!normFlows.length) {
      root.innerHTML =
        '<p class="sankey-empty">No matching flows in SankeyFlowHelper for this selection.</p>';
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
      var rect = document.createElementNS(svgNs, "rect");
      rect.setAttribute("x", d.x0);
      rect.setAttribute("y", d.y0);
      rect.setAttribute("width", Math.max(1, d.x1 - d.x0));
      rect.setAttribute("height", Math.max(1, d.y1 - d.y0));
      var isLeft = i < leftList.length;
      rect.setAttribute(
        "fill",
        isLeft ? cfg.leftFill : cfg.rightFill
      );
      rect.setAttribute("rx", "2");
      rect.setAttribute("class", "sankey-node");
      gNodes.appendChild(rect);
      var nm = String(d.name);
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
    if (!schoolTypeIsElemOrMiddle(p.TYPE)) {
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
    var typeU = (p.TYPE || "").toUpperCase();
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
      emphStroke: PALETTE.middle.fill,
      ariaLabel:
        "Sankey diagram of student flows from middle schools to high schools",
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

    row.className = "sankey-row";
    if (isMid) {
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

  function buildEnrollmentSeries(msid) {
    if (!ENROLLMENT_CACHE || msid == null || isNaN(msid)) return [];
    var key = String(msid);
    var sy = ENROLLMENT_CACHE.byMsid && ENROLLMENT_CACHE.byMsid[key];
    var cal =
      ENROLLMENT_CACHE.calendarByMsid && ENROLLMENT_CACHE.calendarByMsid[key];
    var out = [];

    if (cal) {
      var years = Object.keys(cal)
        .map(function (k) {
          return parseInt(k, 10);
        })
        .filter(function (y) {
          return !isNaN(y) && y >= 2010 && y <= 2025;
        })
        .sort(function (a, b) {
          return a - b;
        });
      for (var i = 0; i < years.length; i++) {
        var y = years[i];
        var yk = String(y);
        var v = cal[yk];
        if (v != null && !isNaN(Number(v))) {
          out.push({
            label: schoolYearLabelFromExcelYear(y),
            value: Number(v),
            segment: "enrollment",
          });
        }
      }
    }

    var labels = ENROLLMENT_CACHE.schoolYearLabels || [];
    if (sy && sy.projected && labels.length) {
      for (var j = 0; j < labels.length; j++) {
        var pv = sy.projected[j];
        if (pv != null && !isNaN(Number(pv))) {
          out.push({
            label: labels[j],
            value: Number(pv),
            segment: "projected",
          });
        }
      }
    }
    return out;
  }

  function enrollmentLabelSortKey(label) {
    var proj = ENROLLMENT_CACHE && ENROLLMENT_CACHE.schoolYearLabels;
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
      "No merged enrollment series from 2025-26 onward for the current selection (check workbook data).";
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

  function renderEnrollmentChartIntoRoot(root, series, options) {
    options = options || {};
    if (!root) return;
    var noDataMsg =
      options.noDataMsg ||
      "No enrollment rows in the published workbook for this school.";
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
    var ml = 36;
    var mb = 54;
    /** Top margin: room so value labels sit fully above bars (incl. tallest). */
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
      parts.push(
        '<text x="' +
          cx +
          '" y="' +
          valY +
          '" text-anchor="middle" dominant-baseline="alphabetic" font-size="11" font-weight="600" fill="#1f2937" font-family="Libre Franklin, sans-serif">' +
          escapeXmlText(val.toLocaleString()) +
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
    var root = document.getElementById("enrollment-chart");
    if (!root) return;
    if (msid == null || isNaN(msid)) {
      root.innerHTML =
        '<p class="enrollment-chart-empty">Select a school to view enrollment trends.</p>';
      root.removeAttribute("aria-label");
      return;
    }
    var series = buildEnrollmentSeries(msid);
    renderEnrollmentChartIntoRoot(root, series, {
      noDataMsg:
        "No enrollment rows in the published workbook for this school.",
      noDataAria:
        "Enrollment data is not available for this school in the workbook.",
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
    var ethEl = document.getElementById("demographics-ethnicity");
    var lunchEl = document.getElementById("demographics-lunch");
    if (!ethEl || !lunchEl) return;

    var emptySelect =
      '<p class="demographics-pie-empty">Select a school to view student demographics.</p>';
    if (msid == null || isNaN(msid)) {
      ethEl.innerHTML = emptySelect;
      lunchEl.innerHTML = emptySelect;
      return;
    }
    if (!DEMOGRAPHICS_CACHE || !DEMOGRAPHICS_CACHE.byMsid) {
      ethEl.innerHTML =
        '<p class="demographics-pie-empty">Demographics data is not loaded.</p>';
      lunchEl.innerHTML =
        '<p class="demographics-pie-empty">Demographics data is not loaded.</p>';
      return;
    }
    var row = DEMOGRAPHICS_CACHE.byMsid[String(msid)];
    if (!row) {
      var msg =
        '<p class="demographics-pie-empty">No student rows for this school in the SY2025-26 export.</p>';
      ethEl.innerHTML = msg;
      lunchEl.innerHTML = msg;
      return;
    }

    var ethRes = buildPieChartHtml(row.ethnicity || {}, ethnicitySliceColor);
    ethEl.innerHTML = ethRes.html;

    var lunchRes = buildPieChartHtml(row.lunchStatus || {}, function (label) {
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
    if (!DEMOGRAPHICS_CACHE || !DEMOGRAPHICS_CACHE.byMsid) {
      return { ethnicity: eth, lunchStatus: lunch };
    }
    for (var i = 0; i < weighted.length; i++) {
      var msid = weighted[i].msid;
      var wt = weighted[i].weight;
      if (msid == null || isNaN(msid) || wt == null || isNaN(wt)) continue;
      var row = DEMOGRAPHICS_CACHE.byMsid[String(msid)];
      if (!row) continue;
      mergeCountObjScaled(eth, row.ethnicity || {}, wt);
      mergeCountObjScaled(lunch, row.lunchStatus || {}, wt);
    }
    return { ethnicity: eth, lunchStatus: lunch };
  }

  function renderDemographicsFromAggregates(agg, ethEl, lunchEl) {
    var emptyAgg =
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
    var cal =
      ENROLLMENT_CACHE &&
      ENROLLMENT_CACHE.calendarByMsid &&
      ENROLLMENT_CACHE.calendarByMsid[String(msid)];
    if (cal && cal["2025"] != null && !isNaN(Number(cal["2025"]))) {
      return Number(cal["2025"]);
    }
    return null;
  }

  function collectScenarioWeightedSpec() {
    var out = [];
    if (scenarioMiddleMsid != null && !isNaN(scenarioMiddleMsid)) {
      out.push({ msid: scenarioMiddleMsid, weight: 1 });
    }
    for (var i = 0; i < scenarioLastFeederRows.length; i++) {
      var r = scenarioLastFeederRows[i];
      if (!r.hasEnrollment || r.msid == null) continue;
      if (scenarioFeederChecked[r.msid] === false) continue;
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
          "No merged enrollment series from 2025-26 onward for the current selection (check workbook data).",
        noDataAria: "Merged enrollment data is not available.",
        ariaLabel:
          "Stacked enrollment by middle school and feeder elementaries from 2025-26 forward.",
      });
    } else {
      var series = buildMergedEnrollmentSeriesWeighted(weighted);
      series = filterEnrollmentSeriesScenarioFuture(series);
      renderEnrollmentChartIntoRoot(chartRoot, series, {
        noDataMsg:
          "No merged enrollment series from 2025-26 onward for the current selection (check workbook data).",
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
      applyScenarioFeederMapHighlights();
      syncStudentHexLayer();
      return;
    }
    var agg = aggregateDemographicsMsidsWeighted(weighted);
    renderDemographicsFromAggregates(agg, ethEl, lunchEl);
    applyScenarioFeederMapHighlights();
    syncStudentHexLayer();
  }

  function updateScenarioSummaryText(middleProps) {
    var p1 = document.getElementById("scenario-details-primary");
    if (!middleProps || !p1) return;
    var name = schoolNameForSelect(middleProps);
    p1.textContent =
      "Scenario: " +
      name +
      " — merged K–8 enrollment and demographics for this middle school and selected elementary feeders.";
    p1.classList.remove("school-details-placeholder");
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
            ") has no enrollment workbook row."
        );
        warnings.push(
          "No enrollment workbook row for \"" +
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
      if (
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
        });
      }
      label.appendChild(swatch);
      label.appendChild(cb);
      var span = document.createElement("span");
      span.className = "scenario-feeder-name";
      var enr =
        r.msid != null && !isNaN(r.msid)
          ? enrollment202526CalendarForMsid(r.msid)
          : null;
      var enrStr =
        enr != null ? enr.toLocaleString() : "—";
      var p =
        r.flowProportion != null && !isNaN(r.flowProportion)
          ? r.flowProportion
          : 1;
      var propAmt = null;
      if (enr != null) {
        propAmt = scenarioCompleteMerger
          ? Math.round(enr)
          : Math.round(enr * p);
      }
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
      p1.textContent =
        "Select a middle school for a merged K–8 scenario summary.";
      p1.classList.add("school-details-placeholder");
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
    applyScenarioFeederMapHighlights();
    syncStudentHexLayer();
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
    );
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

  function updateLeftPanelFromSchool(p) {
    var elP = document.getElementById("school-details-primary");
    var elS = document.getElementById("school-details-secondary");
    if (elP) {
      var name = standardCapitalization(
        expandElemSchoolName(p.NAME || p.CommonName || "")
      );
      var grades = p.Grades ? standardCapitalization(p.Grades) : "—";
      var addr = p.ADDRESS ? standardCapitalization(p.ADDRESS) : "—";
      elP.textContent = [name, grades, addr].join(" | ");
      elP.classList.remove("school-details-placeholder");
    }
    if (elS) {
      var acres =
        p.ACREAGE != null && p.ACREAGE !== "" ? String(p.ACREAGE) : "—";
      var fac = lookupFacilityForSchool(p);
      var opened =
        fac && fac.yearSchoolOpened != null && !isNaN(Number(fac.yearSchoolOpened))
          ? String(fac.yearSchoolOpened)
          : "—";
      var age = "—";
      if (fac && fac.ageAsOf2026 != null && !isNaN(Number(fac.ageAsOf2026))) {
        age = String(fac.ageAsOf2026);
      } else if (p.FacilityAg != null && p.FacilityAg !== "") {
        age = String(p.FacilityAg);
      }
      elS.textContent =
        "Constructed: " +
        opened +
        " | Age of Site: " +
        age +
        " | Size of Site (Acres): " +
        acres;
      elS.classList.remove("school-details-placeholder");
      if (fac) {
        elS.title =
          "Year opened and age as of 2026 from Age of all Facilities 2026; acres from map layer.";
      } else {
        elS.removeAttribute("title");
      }
    }
    var msid =
      p.SCHOOLS_ID != null && p.SCHOOLS_ID !== ""
        ? Number(p.SCHOOLS_ID)
        : null;
    var enrollRow = null;
    if (
      ENROLLMENT_CACHE &&
      ENROLLMENT_CACHE.byMsid &&
      msid != null &&
      !isNaN(msid)
    ) {
      enrollRow = ENROLLMENT_CACHE.byMsid[String(msid)];
    }

    var capEl = document.getElementById("kpi-capacity");
    if (capEl) {
      var capNum =
        enrollRow &&
        enrollRow.factoredCapacity202526 != null &&
        !isNaN(Number(enrollRow.factoredCapacity202526))
          ? Number(enrollRow.factoredCapacity202526)
          : null;
      if (capNum != null) {
        capEl.textContent = capNum.toLocaleString();
        capEl.classList.remove("kpi-value--placeholder");
        capEl.title =
          "2025-26 factored capacity (column J) from the enrollment workbook, not map GeoJSON.";
      } else {
        capEl.textContent = "—";
        capEl.classList.add("kpi-value--placeholder");
        capEl.removeAttribute("title");
      }
    }

    var enrollEl = document.getElementById("kpi-enrollment");
    if (enrollEl) {
      var cur = null;
      if (
        ENROLLMENT_CACHE &&
        ENROLLMENT_CACHE.calendarByMsid &&
        msid != null &&
        !isNaN(msid)
      ) {
        var calRow = ENROLLMENT_CACHE.calendarByMsid[String(msid)];
        if (calRow && calRow["2025"] != null && !isNaN(Number(calRow["2025"]))) {
          cur = Number(calRow["2025"]);
        }
      }
      if (cur != null) {
        enrollEl.textContent = cur.toLocaleString();
        enrollEl.classList.remove("kpi-value--placeholder");
        enrollEl.title =
          "2025 calendar-year membership column from the enrollment workbook (not from map attributes).";
      } else {
        enrollEl.textContent = "—";
        enrollEl.classList.add("kpi-value--placeholder");
        enrollEl.removeAttribute("title");
      }
    }

    var utilEl = document.getElementById("kpi-utilization");
    if (utilEl) {
      var utilPct =
        enrollRow &&
        enrollRow.utilization202526Pct != null &&
        !isNaN(Number(enrollRow.utilization202526Pct))
          ? Number(enrollRow.utilization202526Pct)
          : null;
      if (utilPct != null) {
        var utilStr =
          utilPct % 1 === 0 ? String(utilPct) : utilPct.toFixed(1);
        utilEl.textContent = utilStr + "%";
        utilEl.classList.remove("kpi-value--placeholder");
        utilEl.title =
          "2025-26 actual utilization % (column Q) from the enrollment workbook.";
      } else {
        utilEl.textContent = "—";
        utilEl.classList.add("kpi-value--placeholder");
        utilEl.removeAttribute("title");
      }
    }

    renderEnrollmentChart(msid);
    renderDemographicsCharts(msid);
    renderSankeyPanel(p);
    var captureEl = document.getElementById("kpi-capture");
    if (captureEl) {
      var capRow =
        CAPTURE_CACHE &&
        CAPTURE_CACHE.byMsid &&
        CAPTURE_CACHE.byMsid[String(msid)];
      var capKey = schoolPaletteKeyFromType(p.TYPE);
      var bucket = capRow && capRow[capKey];
      var pct =
        bucket &&
        bucket.captureRatePct != null &&
        !isNaN(Number(bucket.captureRatePct))
          ? Number(bucket.captureRatePct)
          : null;
      if (pct != null) {
        var capStr = pct % 1 === 0 ? String(pct) : pct.toFixed(1);
        captureEl.textContent = capStr + "%";
        captureEl.classList.remove("kpi-value--placeholder");
        captureEl.title =
          "Students attending this school (col A) vs. students zoned here (V/W/Y) within the same grade band; from SY2025-26 StuData export.";
      } else {
        captureEl.textContent = "—";
        captureEl.classList.add("kpi-value--placeholder");
        captureEl.removeAttribute("title");
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
        "Constructed: 19XX | Age of Site: XX | Size of Site (Acres): XX";
      elS.classList.add("school-details-placeholder");
    }
    ["kpi-enrollment", "kpi-capacity", "kpi-utilization", "kpi-capture"].forEach(
      function (id) {
        var k = document.getElementById(id);
        if (k) {
          k.textContent = "—";
          k.classList.add("kpi-value--placeholder");
          if (
            id === "kpi-enrollment" ||
            id === "kpi-capacity" ||
            id === "kpi-utilization" ||
            id === "kpi-capture"
          ) {
            k.removeAttribute("title");
          }
        }
      }
    );
    renderEnrollmentChart(null);
    renderDemographicsCharts(null);
    renderSankeyPanel(null);
  }

  /** Dropdown drives map framing, selection highlight, and left panel; map clicks do not call this. */
  function setupSchoolSelection(schoolByMsid) {
    var sel = document.getElementById("school-select");
    if (!sel) return;

    sel.addEventListener("change", function () {
      var v = sel.value;
      if (!v) {
        clearSelectedSchoolHighlight();
        resetLeftPanelPlaceholders();
        syncStudentHexLayer();
        return;
      }
      var msid = Number(v);
      if (isNaN(msid)) return;
      var p = schoolByMsid[msid];
      if (!p) return;

      applySelectedSchoolHighlight(msid);
      zoomToSchoolAssignment(msid, schoolByMsid);
      updateLeftPanelFromSchool(p);
      syncStudentHexLayer();
    });
  }

  function setupMapInteractions(schoolByMsid) {
    var boundaryHoverPopup = new maplibregl.Popup({
      closeButton: false,
      closeOnClick: false,
      maxWidth: "260px",
      className: "boundary-hover-popup",
      offset: 12,
    });

    var schoolHoverPopup = new maplibregl.Popup({
      closeButton: false,
      closeOnClick: false,
      maxWidth: "300px",
      className: "school-hover-popup",
      offset: 10,
    });

    var schoolClickPopup = new maplibregl.Popup({
      closeButton: true,
      closeOnClick: true,
      maxWidth: "300px",
    });

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

    function clearBoundaryHoverUi() {
      clearOutlineHighlight();
      clearHoverRing();
      boundaryHoverPopup.remove();
      map.getCanvas().style.cursor = "";
    }

    function clearSchoolHoverUi() {
      schoolHoverPopup.remove();
    }

    function boundaryTitleText(props) {
      var msid = props.MSID != null ? Number(props.MSID) : null;
      var raw;
      if (msid != null && !isNaN(msid) && schoolByMsid[msid]) {
        var sp = schoolByMsid[msid];
        raw = sp.NAME || sp.CommonName || String(msid);
      } else {
        raw =
          props.Elem_Commo ||
          props.Middle_Com ||
          props.High_Commo ||
          "Assignment area";
      }
      return standardCapitalization(expandElemSchoolName(raw));
    }

    map.on("mousemove", function (e) {
      var schoolFeats = map.queryRenderedFeatures(e.point, {
        layers: SCHOOL_LAYER_IDS,
      });
      if (schoolFeats.length) {
        clearBoundaryHoverUi();
        var p = schoolFeats[0].properties;
        map.getCanvas().style.cursor = "pointer";
        schoolHoverPopup.setLngLat(e.lngLat).setHTML(schoolDetailHtml(p)).addTo(map);
        refreshAssignmentBoundaryHighlight();
        return;
      }

      clearSchoolHoverUi();

      var feats = map.queryRenderedFeatures(e.point, {
        layers: BOUNDARY_FILL_LAYERS,
      });
      if (!feats.length) {
        clearBoundaryHoverUi();
        refreshAssignmentBoundaryHighlight();
        return;
      }

      var f = feats[0];
      var props = f.properties;
      var layerId = f.layer.id;
      var msid = props.MSID != null ? Number(props.MSID) : null;
      if (msid != null && isNaN(msid)) msid = null;
      var src = fillLayerIdToSource(layerId);

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

    function onSchoolClick(e) {
      var f = e.features && e.features[0];
      if (!f || !f.properties) return;
      schoolClickPopup.setLngLat(e.lngLat).setHTML(schoolDetailHtml(f.properties)).addTo(map);
    }

    SCHOOL_LAYER_IDS.forEach(function (layerId) {
      map.on("click", layerId, onSchoolClick);
    });
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

  function studentHexKey(feature) {
    var p = feature.properties || {};
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
    return "geom:" + JSON.stringify(feature.geometry);
  }

  function buildStudentHexIndex(fc) {
    var countsByMsid = {};
    var geometryByHexKey = {};
    if (!fc || !fc.features) {
      return { countsByMsid: countsByMsid, geometryByHexKey: geometryByHexKey };
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
    }
    return { countsByMsid: countsByMsid, geometryByHexKey: geometryByHexKey };
  }

  function getActiveDashboardSchoolMsid() {
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

    addPart(scenarioMiddleMsid);

    var midStr =
      scenarioMiddleMsid != null && !isNaN(scenarioMiddleMsid)
        ? String(scenarioMiddleMsid)
        : null;

    for (var i = 0; i < scenarioLastFeederRows.length; i++) {
      var r = scenarioLastFeederRows[i];
      if (!r.hasEnrollment || r.msid == null || isNaN(r.msid)) continue;
      if (scenarioFeederChecked[r.msid] === false) continue;
      if (midStr != null && String(r.msid) === midStr) continue;
      addPart(r.msid);
    }
    return combined;
  }

  function syncStudentHexLayer() {
    if (!map || !map.getSource || !map.getSource("student-hex")) return;
    var msid = getActiveDashboardSchoolMsid();
    if (
      !STUDENT_HEX_INDEX ||
      msid == null ||
      isNaN(msid) ||
      !STUDENT_HEX_INDEX.countsByMsid
    ) {
      map.getSource("student-hex").setData({
        type: "FeatureCollection",
        features: [],
      });
      if (map.getLayer("student-hex-heatmap")) {
        map.setLayoutProperty("student-hex-heatmap", "visibility", "none");
      }
      return;
    }
    var panelScenario = document.getElementById("page-scenario");
    var onScenario = panelScenario && !panelScenario.hidden;
    var idx;
    if (
      onScenario &&
      scenarioMiddleMsid != null &&
      !isNaN(scenarioMiddleMsid)
    ) {
      idx = buildMergedScenarioStudentHexCounts();
    } else {
      idx = STUDENT_HEX_INDEX.countsByMsid[String(msid)];
    }
    if (!idx) {
      map.getSource("student-hex").setData({
        type: "FeatureCollection",
        features: [],
      });
      if (map.getLayer("student-hex-heatmap")) {
        map.setLayoutProperty("student-hex-heatmap", "visibility", "none");
      }
      return;
    }
    var features = [];
    for (var key in idx) {
      if (!Object.prototype.hasOwnProperty.call(idx, key)) continue;
      var cnt = idx[key];
      if (cnt <= 0) continue;
      var geom = STUDENT_HEX_INDEX.geometryByHexKey[key];
      if (!geom) continue;
      var pt = polygonCentroid(geom);
      if (!pt) continue;
      features.push({
        type: "Feature",
        properties: { count: cnt },
        geometry: { type: "Point", coordinates: pt },
      });
    }
    if (features.length === 0) {
      map.getSource("student-hex").setData({
        type: "FeatureCollection",
        features: [],
      });
      if (map.getLayer("student-hex-heatmap")) {
        map.setLayoutProperty("student-hex-heatmap", "visibility", "none");
      }
      return;
    }
    map.getSource("student-hex").setData({
      type: "FeatureCollection",
      features: features,
    });
    var showHex = isStudentResidenceLayerEnabled();
    map.setLayoutProperty(
      "student-hex-heatmap",
      "visibility",
      showHex ? "visible" : "none"
    );
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
    var panelExisting = document.getElementById("page-existing");
    var panelScenario = document.getElementById("page-scenario");
    if (!tabExisting || !tabScenario || !panelExisting || !panelScenario) return;

    var labels = {
      existing: "Existing Conditions",
      scenario: "Scenario Planning",
    };

    function setPage(page) {
      var isExisting = page === "existing";
      if (titleEl) {
        titleEl.textContent = isExisting ? labels.existing : labels.scenario;
      }
      tabExisting.setAttribute("aria-selected", isExisting ? "true" : "false");
      tabScenario.setAttribute("aria-selected", isExisting ? "false" : "true");
      tabExisting.classList.toggle("is-active", isExisting);
      tabScenario.classList.toggle("is-active", !isExisting);
      panelExisting.hidden = !isExisting;
      panelScenario.hidden = isExisting;
      if (isExisting) {
        applyScenarioFeederMapHighlights();
      } else {
        refreshScenarioPanelIfVisible();
      }
      syncStudentHexLayer();
      requestAnimationFrame(function () {
        if (map && typeof map.resize === "function") map.resize();
      });
    }

    tabExisting.addEventListener("click", function () {
      setPage("existing");
    });
    tabScenario.addEventListener("click", function () {
      setPage("scenario");
    });
  })();
})();
