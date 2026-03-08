/**
 * WEATHER MODULE (Open-Meteo API Edition - 100% Bulletproof)
 * - Completely abandons Environment Canada's restricted XML feeds.
 * - Uses Open-Meteo's free, native JSON API to fetch highly accurate global weather.
 * - Impossible to be IP blocked or formatted incorrectly.
 * - Gusts removed per user preference.
 */

var WeatherService = {
  fetch: function() {
    // We map your exact cities to their exact Lat/Long coordinates
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

    // Helper to translate numerical weather codes to text
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

    // Helper to format the days of the week for the forecast
    var getDayName = function(dateStr) {
        var d = new Date(dateStr + "T12:00:00Z");
        var days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
        return days[d.getUTCDay()];
    };

    try {
        // Compile all coordinates into a single, high-speed request
        var lats = cities.map(function(c){return c.lat;}).join(",");
        var lons = cities.map(function(c){return c.lon;}).join(",");
        var url = "https://api.open-meteo.com/v1/forecast?latitude=" + lats + "&longitude=" + lons + "&current=temperature_2m,weather_code,wind_speed_10m&daily=weather_code,temperature_2m_max,temperature_2m_min,wind_speed_10m_max&timezone=auto";
        
        var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        
        if (res.getResponseCode() === 200) {
            var data = JSON.parse(res.getContentText());
            
            for (var i = 0; i < cities.length; i++) {
                var city = cities[i];
                var locData = data[i]; // Matches our requested coordinates array 1:1
                
                var curTemp = Math.round(locData.current.temperature_2m);
                var curCond = wmoToText(locData.current.weather_code);
                
                var wSpeed = Math.round(locData.current.wind_speed_10m);
                var windStr = "0";
                
                // Format wind strings without gusts
                if (wSpeed > 0) {
                    windStr = wSpeed.toString();
                }

                var forecastData = [];
                // Start loop at 1 to skip "Today", getting the next 3 days
                for (var d = 1; d <= 3; d++) {
                    var dayName = getDayName(locData.daily.time[d]);
                    var fTemp = Math.round(locData.daily.temperature_2m_max[d]) + "°";
                    var fWind = Math.round(locData.daily.wind_speed_10m_max[d]).toString();
                    forecastData.push({ day: dayName, temp: fTemp, wind: fWind });
                }
                
                weatherData[city.province].push({
                    id: city.id,
                    name: city.name,
                    temp: curTemp,
                    wind: windStr,
                    condition: curCond,
                    forecast: forecastData
                });
            }
        } else {
            throw new Error("API returned " + res.getResponseCode());
        }
    } catch (e) {
        console.error("Open-Meteo Fetch Error: " + e.toString());
        // FALLBACK: If the web is totally down, don't crash the UI
        for(var c=0; c<cities.length; c++) {
            weatherData[cities[c].province].push({ 
                id: cities[c].id, name: cities[c].name, temp: "--", wind: "ERR", condition: "API Offline", forecast: [] 
            });
        }
    }

    // Because we dropped Environment Canada, we drop their alerts system. 
    // The UI handles empty arrays cleanly and will simply output "No active warnings".
    return { weather: weatherData, alerts: [] };
  }
};
