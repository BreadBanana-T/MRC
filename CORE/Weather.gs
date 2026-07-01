/**
 * WEATHER MODULE (Open-Meteo + ECCC Alerts API - National Radar Edition)
 * - Uses Open-Meteo for hyper-accurate, unblockable weather/wind/forecasts.
 * - Uses the official Environment Canada GeoMet OGC API for active Weather Alerts.
 * - Actively scans over 70 major/minor Canadian markets to provide coast-to-coast alert coverage.
 */

var WeatherService = {
  fetch: function() {
    var cities = [
      { name: "Toronto",       lat: 43.7001, lon: -79.4163, province: "ON", id: "on-143" },
      { name: "Ottawa",        lat: 45.4112, lon: -75.6981, province: "ON", id: "on-118" },
      { name: "Calgary",       lat: 51.0501, lon: -114.085, province: "AB", id: "ab-52"  },
      { name: "Edmonton",      lat: 53.5501, lon: -113.468, province: "AB", id: "ab-50"  },
      { name: "Vancouver",     lat: 49.2500, lon: -123.120, province: "BC", id: "bc-74"  },
      { name: "Prince George", lat: 53.9169, lon: -122.749, province: "BC", id: "bc-79"  },
      { name: "Montreal",      lat: 45.5088, lon: -73.5878, province: "QC", id: "qc-147" },
      { name: "Quebec City",   lat: 46.8123, lon: -71.2145, province: "QC", id: "qc-133" }
    ];

    var weatherData = {};
    for(var i = 0; i < cities.length; i++) { weatherData[cities[i].province] = []; }
    var activeAlerts = [];

    // --- 1. FETCH WEATHER FROM OPEN-METEO ---
    var wmoToText = function(code) {
        var map = {
            0: "Clear", 1: "Mainly Clear", 2: "Partly Cloudy", 3: "Overcast",
            45: "Fog", 48: "Rime Fog", 
            51: "Light Drizzle", 53: "Drizzle", 55: "Heavy Drizzle",
            56: "Freezing Drizzle", 57: "Heavy Freezing Drizzle", 
            61: "Light Rain", 63: "Rain", 65: "Heavy Rain",
            66: "Freezing Rain", 67: "Heavy Freezing Rain", 
            71: "Light Snow", 73: "Snow", 75: "Heavy Snow", 77: "Snow Grains", 
            80: "Light Showers", 81: "Showers", 82: "Heavy Showers",
            85: "Snow Showers", 86: "Heavy Snow Showers", 
            95: "Thunderstorm", 96: "Thunderstorm + Hail", 99: "Heavy Thunderstorm"
        };
        return map[code] || "Unknown";
    };

    var getDayName = function(dateStr) {
        var d = new Date(dateStr + "T12:00:00Z");
        var days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
        return days[d.getUTCDay()];
    };

    // --- 1. FETCH WEATHER (cached + 3-provider fallback chain) ---
    // No single weather host is reliable from Google's shared egress IPs (they
    // get rate-limited). So we try three independent providers IN ORDER until
    // one returns usable data:  open-meteo -> MET Norway -> wttr.in.
    // Successful pulls are cached (~20 min) so we rarely hit the network, and
    // mirrored to Script Properties as a durable last-good backup. If all three
    // fail AND we have no backup, only then do we show "API Offline".
    var WX_KEY = 'WX_DATA_V1';
    var wxCache = null; try { wxCache = CacheService.getScriptCache(); } catch (eC) {}
    var wxProps = null; try { wxProps = PropertiesService.getScriptProperties(); } catch (eP) {}

    // Map MET Norway symbol_code -> our condition text.
    var symbolToText = function(s) {
        s = String(s || "").replace(/_(day|night|polartwilight)$/, "");
        var m = {
            clearsky: "Clear", fair: "Mainly Clear", partlycloudy: "Partly Cloudy", cloudy: "Overcast",
            fog: "Fog", lightrain: "Light Rain", rain: "Rain", heavyrain: "Heavy Rain",
            lightrainshowers: "Light Showers", rainshowers: "Showers", heavyrainshowers: "Heavy Showers",
            lightsnow: "Light Snow", snow: "Snow", heavysnow: "Heavy Snow",
            lightsnowshowers: "Light Snow", snowshowers: "Snow Showers", heavysnowshowers: "Heavy Snow Showers",
            sleet: "Sleet", lightsleet: "Sleet", heavysleet: "Heavy Sleet",
            rainandthunder: "Thunderstorm", heavyrainandthunder: "Heavy Thunderstorm", thunderstorm: "Thunderstorm"
        };
        return m[s] || (s ? s.charAt(0).toUpperCase() + s.slice(1) : "Unknown");
    };
    var blankData = function() { var w = {}; for (var i = 0; i < cities.length; i++) { w[cities[i].province] = []; } return w; };
    var enough = function(w) { if (!w) return false; var n = 0; for (var k in w) { if (w.hasOwnProperty(k)) n += w[k].length; } return n >= Math.ceil(cities.length / 2); };

    // Provider A — open-meteo (one bulk call, WMO codes).
    var pOpenMeteo = function() {
        var lats = cities.map(function(c){return c.lat;}).join(",");
        var lons = cities.map(function(c){return c.lon;}).join(",");
        var url = "https://api.open-meteo.com/v1/forecast?latitude=" + lats + "&longitude=" + lons + "&current=temperature_2m,weather_code,wind_speed_10m&daily=weather_code,temperature_2m_max,temperature_2m_min,wind_speed_10m_max&timezone=auto";
        var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (res.getResponseCode() !== 200) { console.error("open-meteo non-200: " + res.getResponseCode()); return null; }
        var data = JSON.parse(res.getContentText());
        var w = blankData();
        for (var i = 0; i < cities.length; i++) {
            try {
                var ld = data[i];
                var fc = [];
                for (var d = 1; d <= 3; d++) {
                    fc.push({ day: getDayName(ld.daily.time[d]), temp: Math.round(ld.daily.temperature_2m_max[d]) + "°", wind: Math.round(ld.daily.wind_speed_10m_max[d]).toString() });
                }
                w[cities[i].province].push({ id: cities[i].id, name: cities[i].name, temp: Math.round(ld.current.temperature_2m), wind: Math.round(ld.current.wind_speed_10m).toString(), condition: wmoToText(ld.current.weather_code), forecast: fc });
            } catch (e) {}
        }
        return w;
    };

    // Provider B — MET Norway (per-city; needs a User-Agent; wind m/s -> km/h).
    var pMetNo = function() {
        var reqs = cities.map(function(c){ return { url: "https://api.met.no/weatherapi/locationforecast/2.0/compact?lat=" + c.lat + "&lon=" + c.lon, headers: { "User-Agent": "MRC-Ops-Dashboard/1.0 (github.com/BreadBanana-T/MRC)" }, muteHttpExceptions: true }; });
        var resps = UrlFetchApp.fetchAll(reqs);
        var w = blankData();
        for (var i = 0; i < cities.length; i++) {
            try {
                if (resps[i].getResponseCode() !== 200) continue;
                var ts = JSON.parse(resps[i].getContentText()).properties.timeseries;
                var inst = ts[0].data.instant.details;
                var sym = ((ts[0].data.next_1_hours || ts[0].data.next_6_hours || { summary: {} }).summary || {}).symbol_code || "";
                var byDay = {};
                for (var k = 0; k < ts.length; k++) {
                    var day = ts[k].time.substring(0, 10); var det = ts[k].data.instant.details;
                    if (!byDay[day]) byDay[day] = { t: -999, wn: 0 };
                    if (det.air_temperature != null) byDay[day].t = Math.max(byDay[day].t, det.air_temperature);
                    if (det.wind_speed != null) byDay[day].wn = Math.max(byDay[day].wn, det.wind_speed * 3.6);
                }
                var days = Object.keys(byDay).sort(); var fc = [];
                for (var dn = 1; dn < days.length && fc.length < 3; dn++) {
                    fc.push({ day: getDayName(days[dn]), temp: Math.round(byDay[days[dn]].t) + "°", wind: Math.round(byDay[days[dn]].wn).toString() });
                }
                w[cities[i].province].push({ id: cities[i].id, name: cities[i].name, temp: Math.round(inst.air_temperature), wind: Math.round((inst.wind_speed || 0) * 3.6).toString(), condition: symbolToText(sym), forecast: fc });
            } catch (e) {}
        }
        return w;
    };

    // Provider C — wttr.in (per-city; already km/h, plain-text conditions).
    var pWttr = function() {
        var reqs = cities.map(function(c){ return { url: "https://wttr.in/" + c.lat + "," + c.lon + "?format=j1", muteHttpExceptions: true }; });
        var resps = UrlFetchApp.fetchAll(reqs);
        var w = blankData();
        for (var i = 0; i < cities.length; i++) {
            try {
                if (resps[i].getResponseCode() !== 200) continue;
                var j = JSON.parse(resps[i].getContentText());
                var cur = j.current_condition[0];
                var fc = [];
                for (var d = 1; d < (j.weather || []).length && fc.length < 3; d++) {
                    var wd = j.weather[d];
                    var wWind = (wd.hourly && wd.hourly[4]) ? wd.hourly[4].windspeedKmph : ((wd.hourly && wd.hourly[0]) ? wd.hourly[0].windspeedKmph : "0");
                    fc.push({ day: getDayName(wd.date), temp: Math.round(parseFloat(wd.maxtempC)) + "°", wind: Math.round(parseFloat(wWind)).toString() });
                }
                var desc = (cur.weatherDesc && cur.weatherDesc[0]) ? cur.weatherDesc[0].value : "";
                w[cities[i].province].push({ id: cities[i].id, name: cities[i].name, temp: Math.round(parseFloat(cur.temp_C)), wind: Math.round(parseFloat(cur.windspeedKmph)).toString(), condition: desc, forecast: fc });
            } catch (e) {}
        }
        return w;
    };

    var freshStr = null;
    if (wxCache) { try { freshStr = wxCache.get(WX_KEY); } catch (e) {} }
    if (freshStr) { try { weatherData = JSON.parse(freshStr); } catch (e) { freshStr = null; } }

    if (!freshStr) {
        var providers = [pOpenMeteo, pMetNo, pWttr];
        var got = null;
        for (var pi = 0; pi < providers.length && !enough(got); pi++) {
            try { var r = providers[pi](); if (enough(r)) got = r; } catch (e) { console.error("WX provider " + pi + " failed: " + e); }
        }

        if (got) {
            weatherData = got;
            var okStr = JSON.stringify(weatherData);
            if (wxCache) { try { wxCache.put(WX_KEY, okStr, 1200); } catch (e) {} }   // 20-min cache
            if (wxProps) { try { wxProps.setProperty(WX_KEY, okStr); } catch (e) {} } // durable backup
        } else {
            // All three providers failed — serve the last good result; only show
            // "API Offline" if we've genuinely never succeeded.
            var backup = null;
            if (wxProps) { try { backup = wxProps.getProperty(WX_KEY); } catch (e) {} }
            if (backup) { try { weatherData = JSON.parse(backup); } catch (e) { backup = null; } }
            if (!backup) {
                for (var c = 0; c < cities.length; c++) {
                    weatherData[cities[c].province].push({ id: cities[c].id, name: cities[c].name, temp: "--", wind: "ERR", condition: "API Offline", forecast: [] });
                }
            }
        }
    }

    // --- 2. FETCH OFFICIAL ALERTS FROM ECCC GEOMET API ---
    // The GeoMet weather-alerts feature schema (confirmed live) gives us everything
    // directly, no guessing needed:
    //   province        -> "QC" / "ON" / ...        (province attribution)
    //   alert_name_en   -> "wind warning"           (exact hazard type)
    //   feature_name_en -> "Alma - Desbiens area"   (real affected region name)
    //   risk_colour_en  -> "yellow" / "red"         (severity — RED = most severe)
    //   status_en       -> "continued" / "ended"    (lifecycle)
    //
    // "RED-level" filters on risk_colour_en === "red". This is important: many
    // items are alert_type="warning" yet only YELLOW (e.g. heat warnings), so
    // filtering by type alone flooded the banner. Watches/advisories/statements
    // and ended alerts are dropped. Alerts are aggregated per province + hazard,
    // counting DISTINCT affected regions. Geometry / city-text remain only as a
    // fallback for the rare feature missing a province code.
    var cityWeights = {
        // Major
        "TORONTO": 10, "MONTREAL": 10, "VANCOUVER": 10, "CALGARY": 10, "OTTAWA": 10, "EDMONTON": 10,
        // Mid
        "WINNIPEG": 6, "QUEBEC": 6, "HAMILTON": 6, "KITCHENER": 6, "LONDON": 5,
        "MARKHAM": 5, "MISSISSAUGA": 5, "BRAMPTON": 5, "VAUGHAN": 5,
        "HALIFAX": 6, "VICTORIA": 6, "SURREY": 5, "BURNABY": 5, "REGINA": 5,
        "SASKATOON": 5, "GATINEAU": 5, "LAVAL": 5, "LONGUEUIL": 5, "SHERBROOKE": 4
    };
    var alertRadar = {
        "ON": ["TORONTO", "OTTAWA", "MISSISSAUGA", "BRAMPTON", "HAMILTON", "LONDON", "MARKHAM", "VAUGHAN", "KITCHENER", "WINDSOR", "BURLINGTON", "SUDBURY", "OSHAWA", "BARRIE", "KINGSTON", "NIAGARA", "THUNDER BAY", "PETERBOROUGH", "GUELPH", "WATERLOO", "BELLEVILLE"],
        "BC": ["VANCOUVER", "PRINCE GEORGE", "VICTORIA", "SURREY", "BURNABY", "RICHMOND", "ABBOTSFORD", "KELOWNA", "KAMLOOPS", "NANAIMO", "CHILLIWACK", "COQUITLAM", "LANGLEY", "DELTA", "WHISTLER", "SQUAMISH"],
        "AB": ["CALGARY", "EDMONTON", "RED DEER", "LETHBRIDGE", "FORT MCMURRAY", "MEDICINE HAT", "GRANDE PRAIRIE", "AIRDRIE", "BANFF", "JASPER", "LLOYDMINSTER"],
        "QC": ["MONTREAL", "QUEBEC", "LAVAL", "GATINEAU", "LONGUEUIL", "SHERBROOKE", "TROIS-RIVIERES", "CHICOUTIMI", "SAINT-JEAN", "BROSSARD", "LEVIS", "DRUMMONDVILLE", "SAGUENAY", "GRANBY"],
        "MB": ["WINNIPEG", "BRANDON", "THOMPSON", "PORTAGE LA PRAIRIE", "STEINBACH"],
        "SK": ["SASKATOON", "REGINA", "PRINCE ALBERT", "MOOSE JAW", "SWIFT CURRENT", "YORKTON"],
        "NS": ["HALIFAX", "DARTMOUTH", "SYDNEY", "TRURO", "NEW GLASGOW", "CAPE BRETON"],
        "NB": ["MONCTON", "SAINT JOHN", "FREDERICTON", "DIEPPE", "MIRAMICHI", "EDMUNDSTON"],
        "NL": ["ST. JOHN'S", "CORNER BROOK", "GANDER", "MOUNT PEARL", "CONCEPTION BAY"],
        "PE": ["CHARLOTTETOWN", "SUMMERSIDE", "STRATFORD"]
    };
    var toTitleCase = function(str) {
        return str.replace(/\w\S*/g, function(t){ return t.charAt(0).toUpperCase() + t.substr(1).toLowerCase(); });
    };

    // Full province/territory names (+ common variants) for text attribution.
    var provNames = {
        "ON": ["ONTARIO"], "QC": ["QUEBEC", "QUÉBEC"], "BC": ["BRITISH COLUMBIA"],
        "AB": ["ALBERTA"], "SK": ["SASKATCHEWAN"], "MB": ["MANITOBA"],
        "NS": ["NOVA SCOTIA"], "NB": ["NEW BRUNSWICK"], "PE": ["PRINCE EDWARD ISLAND"],
        "NL": ["NEWFOUNDLAND", "LABRADOR"], "YT": ["YUKON"],
        "NT": ["NORTHWEST TERRITORIES"], "NU": ["NUNAVUT"]
    };

    // Rough province bounding boxes: [minLon, minLat, maxLon, maxLat]. Boxes over
    // non-rectangular provinces overlap near borders; when a point falls in more
    // than one we keep the province whose box centre is closest.
    var provBBox = {
        "BC": [-139.1, 48.2, -114.0, 60.0], "AB": [-120.0, 48.9, -110.0, 60.0],
        "SK": [-110.0, 48.9, -101.3, 60.0], "MB": [-102.1, 48.9, -88.9, 60.0],
        "ON": [-95.2, 41.6, -74.3, 56.9],   "QC": [-79.8, 44.9, -57.1, 62.6],
        "NB": [-69.1, 44.5, -63.7, 48.1],   "NS": [-66.4, 43.3, -59.7, 47.1],
        "PE": [-64.5, 45.9, -61.9, 47.1],   "NL": [-67.9, 46.5, -52.5, 60.4],
        "YT": [-141.1, 60.0, -123.8, 69.7], "NT": [-136.5, 60.0, -101.9, 78.8],
        "NU": [-120.0, 60.0, -61.0, 83.2]
    };

    // Flatten any GeoJSON geometry down to [lon,lat] pairs and average -> centroid.
    var geomCentroid = function(g) {
        if (!g || !g.coordinates) return null;
        var pts = [];
        var walk = function(a) {
            if (!a || !a.length) return;
            if (typeof a[0] === "number") { pts.push(a); return; }
            for (var i = 0; i < a.length; i++) walk(a[i]);
        };
        walk(g.coordinates);
        if (!pts.length) return null;
        var sx = 0, sy = 0;
        for (var i = 0; i < pts.length; i++) { sx += pts[i][0]; sy += pts[i][1]; }
        return { lon: sx / pts.length, lat: sy / pts.length };
    };
    var provinceForPoint = function(lon, lat) {
        var best = null, bestD = Infinity;
        for (var pc in provBBox) {
            var b = provBBox[pc];
            if (lon >= b[0] && lon <= b[2] && lat >= b[1] && lat <= b[3]) {
                var cx = (b[0] + b[2]) / 2, cy = (b[1] + b[3]) / 2;
                var d = (lon - cx) * (lon - cx) + (lat - cy) * (lat - cy);
                if (d < bestD) { bestD = d; best = pc; }
            }
        }
        return best;
    };

    // Read a field from a properties object trying several possible names.
    var pick = function(obj, names) {
        for (var i = 0; i < names.length; i++) {
            var v = obj[names[i]];
            if (v !== undefined && v !== null && v !== "") return v;
        }
        return "";
    };

    // ECCC names every warning as "<hazard> warning" (e.g. "wind warning",
    // "snowfall warning"). We MUST anchor on that exact phrase — matching a bare
    // keyword like "wind" anywhere in the description body mis-classifies almost
    // everything as a wind warning, because nearly every alert description mentions
    // wind ("blowing snow due to strong winds", "gusty winds", etc.).
    // Order matters: more specific phrases first.
    var HAZARD_MAP = [
        ["tornado warning", "Tornado warning"], ["hurricane warning", "Hurricane warning"],
        ["tropical storm warning", "Tropical storm warning"], ["tsunami warning", "Tsunami warning"],
        ["storm surge warning", "Storm surge warning"], ["blizzard warning", "Blizzard warning"],
        ["winter storm warning", "Winter storm warning"], ["snow squall warning", "Snow squall warning"],
        ["snowfall warning", "Snowfall warning"], ["flash freeze warning", "Flash freeze warning"],
        ["freezing rain warning", "Freezing rain warning"], ["freezing drizzle warning", "Freezing drizzle warning"],
        ["severe thunderstorm warning", "Severe thunderstorm warning"], ["thunderstorm warning", "Thunderstorm warning"],
        ["rainfall warning", "Rainfall warning"], ["arctic outflow warning", "Arctic outflow warning"],
        ["les suetes wind warning", "Wind warning"], ["wind warning", "Wind warning"],
        ["extreme cold warning", "Extreme cold warning"], ["cold warning", "Extreme cold warning"],
        ["heat warning", "Heat warning"], ["frost warning", "Frost warning"],
        ["coastal flooding warning", "Coastal flooding warning"], ["fog warning", "Fog warning"],
        ["dust storm warning", "Dust storm warning"], ["squall warning", "Squall warning"]
    ];
    var classifyHazard = function(text) {
        var t = String(text || "").toLowerCase();
        for (var i = 0; i < HAZARD_MAP.length; i++) { if (t.indexOf(HAZARD_MAP[i][0]) !== -1) return HAZARD_MAP[i][1]; }
        return "Warning"; // couldn't identify a specific hazard phrase
    };
    // ECCC risk-based colours: yellow / orange / red. "RED-level" = the most severe.
    // Set to { red:true, orange:true } if you ever want to include amber alerts too.
    var RED_COLOURS = { "red": true };
    // Severity rank drives sort order and colour intensity (higher = more dangerous).
    var hazardRank = function(name) {
        var t = String(name || "").toLowerCase();
        if (t.indexOf("tornado") !== -1) return 100;
        if (t.indexOf("hurricane") !== -1 || t.indexOf("tsunami") !== -1) return 90;
        if (t.indexOf("blizzard") !== -1 || t.indexOf("winter storm") !== -1 || t.indexOf("storm surge") !== -1) return 80;
        if (t.indexOf("severe thunderstorm") !== -1 || t.indexOf("snow squall") !== -1) return 70;
        if (t.indexOf("wind") !== -1 || t.indexOf("freezing") !== -1 || t.indexOf("flash freeze") !== -1) return 60;
        if (t.indexOf("snowfall") !== -1 || t.indexOf("rainfall") !== -1 || t.indexOf("thunderstorm") !== -1) return 50;
        if (t.indexOf("extreme cold") !== -1 || t.indexOf("heat") !== -1 || t.indexOf("arctic") !== -1) return 40;
        return 30;
    };
    // Harvest every text value on a feature's properties (incl. description arrays).
    var harvestText = function(props) {
        var out = "";
        for (var key in props) {
            if (!props.hasOwnProperty(key)) continue;
            var val = props[key];
            if (typeof val === "string") { out += " " + val; }
            else if (val && typeof val === "object" && typeof val.length === "number") {
                for (var d = 0; d < val.length; d++) {
                    var it = val[d];
                    if (typeof it === "string") out += " " + it;
                    else if (it && typeof it === "object") { if (it.text) out += " " + it.text; if (it.event) out += " " + it.event; if (it.headline) out += " " + it.headline; }
                }
            }
        }
        return out;
    };

    try {
        // limit=1000 pulls all currently active records; default limit is 10.
        var alertUrl = "https://api.weather.gc.ca/collections/weather-alerts/items?f=json&lang=en&limit=1000";
        var alertRes = UrlFetchApp.fetch(alertUrl, { muteHttpExceptions: true, headers: { "User-Agent": "MRC-Ops-Dashboard/1.0" } });
        var alertData = (alertRes.getResponseCode() === 200) ? JSON.parse(alertRes.getContentText()) : null;
        var features = (alertData && alertData.features) ? alertData.features : [];

        // Aggregate per province -> { hazards: {name:{count,areas:Set}}, ids:Set, zoneCount }
        // Grouping by province (not per-zone) keeps the banner readable: one row per
        // province listing the distinct hazard TYPES, instead of hundreds of "cities".
        var provAgg = {};

        for (var j = 0; j < features.length; j++) {
            var feat = features[j] || {};
            var props = feat.properties || {};

            // --- Real ECCC GeoMet schema (confirmed via debugWeatherAlerts) ---
            var alertType = String(pick(props, ["alert_type", "type"])).toLowerCase();
            var status    = String(pick(props, ["status_en", "status"])).toLowerCase();
            var colour    = String(pick(props, ["risk_colour_en", "risk_colour", "risk_color_en"])).toLowerCase();
            var name      = String(pick(props, ["alert_name_en", "alert_short_name_en", "headline", "event"]));
            var region    = String(pick(props, ["feature_name_en", "area", "areas", "zone", "location"]));
            var prov      = String(pick(props, ["province"])).toUpperCase();

            // Skip ended / expired alerts.
            if (status === "ended" || status === "expired" || alertType === "ended") continue;

            // RED-LEVEL ONLY — filter on ECCC's official risk colour. This is what
            // drops the flood of yellow/orange alerts (e.g. heat warnings), which
            // are "warnings" by type but not red. Fall back to alert_type only when
            // the feed omits the colour.
            var isRed = RED_COLOURS[colour] === true;
            if (!colour && alertType.indexOf("warning") !== -1) isRed = true;
            if (!isRed) continue;
            // Never surface watches / advisories / statements.
            if (alertType.indexOf("watch") !== -1 || alertType.indexOf("advisory") !== -1 || alertType.indexOf("statement") !== -1) continue;

            // Hazard name straight from the feed ("wind warning" -> "Wind warning");
            // fall back to text sniffing only if the field is missing.
            var hazard = name ? name.replace(/\s+/g, " ").trim() : classifyHazard(harvestText(props));
            if (hazard) hazard = hazard.charAt(0).toUpperCase() + hazard.slice(1);

            // Province comes straight from the feed; geometry is only a fallback.
            var provCode = (prov && prov.length === 2) ? prov : null;
            if (!provCode) {
                var ctr = geomCentroid(feat.geometry);
                if (ctr) provCode = provinceForPoint(ctr.lon, ctr.lat);
            }
            if (!provCode) continue; // unattributable — skip rather than mislabel

            if (!provAgg[provCode]) provAgg[provCode] = { hazards: {}, ids: {}, zoneCount: 0 };
            var pa = provAgg[provCode];
            if (!pa.hazards[hazard]) pa.hazards[hazard] = { count: 0, areas: {} };

            // Count DISTINCT affected regions per hazard (feature_name_en is the
            // region, e.g. "Alma - Desbiens area"), so the same region under the
            // same hazard is never double-counted.
            var areaLabel = region ? region.replace(/\s+/g, " ").trim() : "";
            var dkey = hazard + "|" + (areaLabel || ("_" + j));
            if (!pa.ids[dkey]) { pa.ids[dkey] = true; pa.hazards[hazard].count += 1; pa.zoneCount += 1; }
            if (areaLabel) pa.hazards[hazard].areas[areaLabel] = true;
        }

        // Materialize: one entry per province, hazards sorted most-dangerous first.
        Object.keys(provAgg).forEach(function(prov) {
            var pa = provAgg[prov];
            var hazList = Object.keys(pa.hazards).map(function(name) {
                var h = pa.hazards[name];
                return { type: name, count: h.count, areas: Object.keys(h.areas), rank: hazardRank(name) };
            });
            hazList.sort(function(a, b) { return (b.rank - a.rank) || (b.count - a.count); });
            activeAlerts.push({
                province: prov,
                hazards: hazList,
                zoneCount: pa.zoneCount,
                topRank: hazList.length ? hazList[0].rank : 0,
                summary: hazList.map(function(h) { return h.type; }).join(" · ")
            });
        });
        // Most-dangerous province first, then by number of affected zones.
        activeAlerts.sort(function(a, b) { return (b.topRank - a.topRank) || (b.zoneCount - a.zoneCount); });
    } catch (e) {
        console.error("ECCC Alerts Error: " + e.toString());
    }

    return { weather: weatherData, alerts: activeAlerts };
  }
};

/**
 * DIAGNOSTIC — run this from the Apps Script editor (Run > debugWeatherAlerts)
 * then open View > Logs. It prints the ACTUAL property field names ECCC returns
 * for a warning feature, so we can wire the exact fields for region names /
 * hazard type. Paste the log output back and we can lock the parser to the real
 * schema instead of guessing. Safe / read-only.
 */
function debugWeatherAlerts() {
  var url = "https://api.weather.gc.ca/collections/weather-alerts/items?f=json&lang=en&limit=1000";
  var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { "User-Agent": "MRC-Ops-Dashboard/1.0" } });
  if (res.getResponseCode() !== 200) { Logger.log("HTTP " + res.getResponseCode()); return "HTTP " + res.getResponseCode(); }
  var data = JSON.parse(res.getContentText());
  var feats = data.features || [];
  // Prefer a feature that looks like a warning so we see a populated example.
  var sample = null;
  for (var i = 0; i < feats.length; i++) {
    var p = feats[i].properties || {};
    var blob = JSON.stringify(p).toLowerCase();
    if (blob.indexOf("warning") !== -1) { sample = feats[i]; break; }
  }
  if (!sample) sample = feats[0];
  var out = {
    totalFeatures: feats.length,
    propertyKeys: sample ? Object.keys(sample.properties || {}) : [],
    geometryType: sample && sample.geometry ? sample.geometry.type : null,
    sampleProperties: sample ? sample.properties : null
  };
  var s = JSON.stringify(out, null, 2);
  Logger.log(s);
  return s;
}
