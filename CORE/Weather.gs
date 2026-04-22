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

    try {
        var lats = cities.map(function(c){return c.lat;}).join(",");
        var lons = cities.map(function(c){return c.lon;}).join(",");
        var url = "https://api.open-meteo.com/v1/forecast?latitude=" + lats + "&longitude=" + lons + "&current=temperature_2m,weather_code,wind_speed_10m&daily=weather_code,temperature_2m_max,temperature_2m_min,wind_speed_10m_max&timezone=auto";
        
        var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        
        if (res.getResponseCode() === 200) {
            var data = JSON.parse(res.getContentText());
            
            for (var i = 0; i < cities.length; i++) {
                var city = cities[i];
                var locData = data[i]; 
                
                var curTemp = Math.round(locData.current.temperature_2m);
                var curCond = wmoToText(locData.current.weather_code);
                
                var wSpeed = Math.round(locData.current.wind_speed_10m);
                var windStr = wSpeed > 0 ? wSpeed.toString() : "0";
                
                var forecastData = [];
                for (var d = 1; d <= 3; d++) {
                    var dayName = getDayName(locData.daily.time[d]);
                    var fTemp = Math.round(locData.daily.temperature_2m_max[d]) + "°";
                    var fWind = Math.round(locData.daily.wind_speed_10m_max[d]).toString();
                    forecastData.push({ day: dayName, temp: fTemp, wind: fWind });
                }
                
                weatherData[city.province].push({
                    id: city.id, name: city.name, temp: curTemp, wind: windStr, condition: curCond, forecast: forecastData
                });
            }
        }
    } catch (e) {
        console.error("Open-Meteo Error: " + e.toString());
        for(var c=0; c<cities.length; c++) {
            weatherData[cities[c].province].push({ id: cities[c].id, name: cities[c].name, temp: "--", wind: "ERR", condition: "API Offline", forecast: [] });
        }
    }

    // --- 2. FETCH OFFICIAL ALERTS FROM ECCC GEOMET API ---
    // Only RED-level alerts (WARNING + severity Severe/Extreme). WATCHES, ADVISORIES,
    // STATEMENTS are dropped. Alerts are aggregated per-province by hazard type and
    // weighted by the population of the cities they touch, so Toronto + Mississauga +
    // Brampton under one freezing-rain warning produces a single heavy row instead of
    // three duplicate banner entries.
    //
    // City weights: Major metros (>=1M) = 10, Mid (>=250k) = 5, others = 1.
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

    try {
        // limit=1000 pulls all currently active records; default limit is 10.
        var alertUrl = "https://api.weather.gc.ca/collections/weather-alerts/items?f=json&lang=en&limit=1000";
        var alertRes = UrlFetchApp.fetch(alertUrl, { muteHttpExceptions: true, headers: { "User-Agent": "MRC-Ops-Dashboard/1.0" } });
        var alertData = (alertRes.getResponseCode() === 200) ? JSON.parse(alertRes.getContentText()) : null;
        var features = (alertData && alertData.features) ? alertData.features : [];

        // Bucket = province + hazard type -> { weight, cities:Set }
        var buckets = {};

        for (var j = 0; j < features.length; j++) {
            var props = features[j].properties || {};
            var headline = props.headline || props.event || "";
            if (!headline) continue;

            // RED-only filter: must be a WARNING (not watch/advisory/statement) AND
            // severity must be Severe or Extreme when present.
            var headUp = headline.toUpperCase();
            if (headUp.indexOf("WARNING") === -1) continue;
            var severity = String(props.severity || "").toLowerCase();
            if (severity && severity !== "severe" && severity !== "extreme") continue;

            var areas = String(props.areas || "").toUpperCase();

            // Extract the hazard name (strip trailing " in effect", " issued", etc.)
            var hazard = headline;
            var cut = hazard.toLowerCase().search(/\s+(in effect|issued|ended)/);
            if (cut !== -1) hazard = hazard.substring(0, cut).trim();

            for (var provCode in alertRadar) {
                var cityList = alertRadar[provCode];
                for (var c = 0; c < cityList.length; c++) {
                    var city = cityList[c];
                    if (areas.indexOf(city) === -1 && headUp.indexOf(city) === -1) continue;

                    var bucketKey = provCode + "|" + hazard;
                    if (!buckets[bucketKey]) buckets[bucketKey] = { province: provCode, hazard: hazard, cities: {}, weight: 0 };
                    if (!buckets[bucketKey].cities[city]) {
                        buckets[bucketKey].cities[city] = true;
                        buckets[bucketKey].weight += (cityWeights[city] || 1);
                    }
                }
            }
        }

        // Materialize and sort by weight (most-impactful first)
        Object.keys(buckets).forEach(function(k) {
            var b = buckets[k];
            var cityNames = Object.keys(b.cities).map(toTitleCase);
            var shown = cityNames.slice(0, 3).join(", ");
            if (cityNames.length > 3) shown += " + " + (cityNames.length - 3) + " more";
            activeAlerts.push({
                province: b.province,
                type: b.hazard,
                cities: cityNames,
                cityCount: cityNames.length,
                weight: b.weight,
                displayCities: shown
            });
        });
        activeAlerts.sort(function(a, b) { return b.weight - a.weight; });
        // Cap so a flood of alerts can't take over the banner.
        if (activeAlerts.length > 8) activeAlerts = activeAlerts.slice(0, 8);
    } catch (e) {
        console.error("ECCC Alerts Error: " + e.toString());
    }

    return { weather: weatherData, alerts: activeAlerts };
  }
};
