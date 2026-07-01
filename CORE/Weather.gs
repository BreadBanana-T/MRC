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
    // RED-level = anything ECCC classifies as a WARNING (that's what shows red on
    // weather.gc.ca). WATCHES / ADVISORIES / STATEMENTS / ENDED alerts are dropped.
    // We do NOT additionally require CAP severity Severe/Extreme — plenty of real
    // red warnings (e.g. snowfall) are tagged "Moderate", and dropping them was
    // causing red provinces to show "no warnings".
    //
    // Each alert is attributed to a province primarily by GEOMETRY (polygon centroid
    // -> province bounding box) so region warnings whose zone names don't contain a
    // major-city string ("R.M. of Portage la Prairie", "southern Manitoba", a rural
    // tornado warning) are no longer silently dropped. Province-name and known-city
    // text are used as extra/faster signals. Alerts are aggregated per province +
    // hazard, and the ACTUAL affected region names from the API are shown.
    //
    // City weights (impact heat only): Major metros (>=1M) = 10, Mid (>=250k) = 5.
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

    try {
        // limit=1000 pulls all currently active records; default limit is 10.
        var alertUrl = "https://api.weather.gc.ca/collections/weather-alerts/items?f=json&lang=en&limit=1000";
        var alertRes = UrlFetchApp.fetch(alertUrl, { muteHttpExceptions: true, headers: { "User-Agent": "MRC-Ops-Dashboard/1.0" } });
        var alertData = (alertRes.getResponseCode() === 200) ? JSON.parse(alertRes.getContentText()) : null;
        var features = (alertData && alertData.features) ? alertData.features : [];

        // Bucket = province + hazard type -> { hazard, areas:Set, count, weight }
        var buckets = {};

        for (var j = 0; j < features.length; j++) {
            var feat = features[j] || {};
            var props = feat.properties || {};

            // --- Robust field reads (ECCC field names vary across the API) ---
            var headline = String(pick(props, ["headline", "event", "summary", "title"]));
            if (!headline) {
                var descs = props.descriptions || props.description;
                if (descs && descs.length) headline = String(descs[0].text || descs[0].headline || descs[0] || "");
            }
            var alertType = String(pick(props, ["alert_type", "type", "msg_type", "category"])).toLowerCase();
            var areaText  = String(pick(props, ["area", "areas", "zone", "location", "region"]));
            var blob = (headline + " " + alertType + " " + String(props.status || "")).toLowerCase();

            // Skip ended / cancelled alerts.
            if (alertType === "ended" || alertType === "cancel" || /\b(ended|cancel(l)?ed|expired)\b/.test(blob)) continue;

            // RED-only: keep WARNINGS, drop watches / advisories / statements.
            var isWarning = alertType ? (alertType.indexOf("warning") !== -1) : /\bwarning\b/.test(blob);
            if (!isWarning) continue;

            // Hazard label: strip trailing status words ("in effect", "issued"...).
            var hazard = headline || (alertType ? alertType : "Weather Warning");
            var cut = hazard.toLowerCase().search(/\s+(in effect|issued|ended|continued|updated|has ended)/);
            if (cut !== -1) hazard = hazard.substring(0, cut).trim();
            hazard = hazard.replace(/\s+/g, " ").trim();
            if (hazard) hazard = hazard.charAt(0).toUpperCase() + hazard.slice(1);

            // --- Attribute to a province ---
            var upperBlob = (areaText + " " + headline).toUpperCase();
            var provCode = null;

            // 1. Explicit province/territory name in the area or headline.
            for (var pc in provNames) {
                var variants = provNames[pc];
                for (var v = 0; v < variants.length; v++) {
                    if (upperBlob.indexOf(variants[v]) !== -1) { provCode = pc; break; }
                }
                if (provCode) break;
            }
            // 2. A known major city in the area or headline. Always scanned (even
            //    if the province is already known) so we can weight impact by city.
            var matchedCityWeight = 1;
            for (var pc2 in alertRadar) {
                var cityList = alertRadar[pc2];
                for (var ci = 0; ci < cityList.length; ci++) {
                    if (upperBlob.indexOf(cityList[ci]) !== -1) {
                        if (!provCode) provCode = pc2;
                        matchedCityWeight = Math.max(matchedCityWeight, cityWeights[cityList[ci]] || 1);
                    }
                }
            }
            // 3. Geometry centroid -> province bounding box (the reliable fallback).
            if (!provCode) {
                var ctr = geomCentroid(feat.geometry);
                if (ctr) provCode = provinceForPoint(ctr.lon, ctr.lat);
            }
            if (!provCode) continue; // unattributable — skip rather than mislabel

            var bucketKey = provCode + "|" + hazard;
            if (!buckets[bucketKey]) buckets[bucketKey] = { province: provCode, hazard: hazard, areas: {}, count: 0, weight: 0 };
            var bk = buckets[bucketKey];
            var areaLabel = areaText ? areaText.replace(/\s+/g, " ").trim() : "";
            if (areaLabel && !bk.areas[areaLabel]) bk.areas[areaLabel] = true;
            bk.count += 1;
            bk.weight += matchedCityWeight;
        }

        // Materialize and sort (tornado first, then by impact weight).
        Object.keys(buckets).forEach(function(k) {
            var b = buckets[k];
            var areaNames = Object.keys(b.areas);
            var shown = areaNames.slice(0, 3).join(", ");
            if (areaNames.length > 3) shown += " + " + (areaNames.length - 3) + " more";
            var cnt = areaNames.length || b.count;
            activeAlerts.push({
                province: b.province,
                type: b.hazard,
                cities: areaNames,        // affected region names (key kept for UI compat)
                cityCount: cnt,
                weight: b.weight,
                displayCities: shown
            });
        });
        activeAlerts.sort(function(a, b) {
            var pa = /tornado/i.test(a.type) ? 1 : 0;
            var pb = /tornado/i.test(b.type) ? 1 : 0;
            if (pa !== pb) return pb - pa;
            return b.weight - a.weight;
        });
        // Cap so a flood of alerts can't take over the banner.
        if (activeAlerts.length > 12) activeAlerts = activeAlerts.slice(0, 12);
    } catch (e) {
        console.error("ECCC Alerts Error: " + e.toString());
    }

    return { weather: weatherData, alerts: activeAlerts };
  }
};
