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

    // --- 2. FETCH OFFICIAL ALERTS FROM ECCC GEOMET API (NATIONAL RADAR) ---
    try {
        var alertUrl = "https://api.weather.gc.ca/collections/weather-alerts/items?f=json&lang=en";
        var alertRes = UrlFetchApp.fetch(alertUrl, { muteHttpExceptions: true });
        
        if (alertRes.getResponseCode() === 200) {
            var alertData = JSON.parse(alertRes.getContentText());
            if (alertData && alertData.features) {
                var seenAlerts = new Set();
                
                // MASSIVE DICTIONARY: Scans all 10 provinces for major and minor regional centers
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

                // Helper to pretty-print uppercase cities (e.g. "MISSISSAUGA" -> "Mississauga")
                var toTitleCase = function(str) {
                    return str.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
                };
                
                var provKeys = Object.keys(alertRadar);
                
                // Scan all active alerts in Canada
                for (var j = 0; j < alertData.features.length; j++) {
                    var props = alertData.features[j].properties || {};
                    var headline = props.headline || props.event || "";
                    if (!headline) continue;
                    
                    var areas = (props.areas || "").toUpperCase();
                    var headUp = headline.toUpperCase();

                    // Check our dictionary
                    for (var p = 0; p < provKeys.length; p++) {
                        var provCode = provKeys[p];
                        var cityList = alertRadar[provCode];
                        var matchedCities = [];

                        // Find all minor/major cities caught in this specific alert
                        for (var c = 0; c < cityList.length; c++) {
                            var checkCity = cityList[c];
                            if (areas.indexOf(checkCity) !== -1 || headUp.indexOf(checkCity) !== -1) {
                                matchedCities.push(toTitleCase(checkCity));
                            }
                        }

                        if (matchedCities.length > 0) {
                            var alertText = headline;
                            if (alertText.toLowerCase().indexOf(" in effect") !== -1) {
                                alertText = alertText.substring(0, alertText.toLowerCase().indexOf(" in effect")).trim();
                            }
                            
                            // Combine cities neatly (e.g. "Toronto, Mississauga, Brampton")
                            // Limits to 3 to prevent the dashboard banner from taking up half the screen
                            var displayCities = matchedCities.slice(0, 3).join(", ");
                            if (matchedCities.length > 3) displayCities += " + " + (matchedCities.length - 3) + " more";

                            var alertKey = provCode + "-" + alertText + "-" + displayCities;
                            
                            if (!seenAlerts.has(alertKey)) {
                                seenAlerts.add(alertKey);
                                activeAlerts.push({ province: provCode, type: alertText + " (" + displayCities + ")", count: 1 });
                            }
                        }
                    }
                }
            }
        }
    } catch (e) {
        console.error("ECCC Alerts Error: " + e.toString());
    }

    return { weather: weatherData, alerts: activeAlerts };
  }
};
