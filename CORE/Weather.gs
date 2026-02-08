/**
 * WEATHER MODULE (Powered by Open-Meteo)
 * Replaces EC RSS to bypass Google IP blocking.
 * V3: Removes Gusts (User Request) & Adds Null Safety.
 */

const WeatherService = {
  fetch: function() { return fetchWeatherOpenMeteo(); }
};

function fetchWeatherOpenMeteo() {
  // Mapping Cities to Lat/Lon for Open-Meteo
  const cities = [
    { name: "Toronto",       lat: 43.70, lon: -79.42, province: "ON", code: "on-143" },
    { name: "Ottawa",        lat: 45.42, lon: -75.70, province: "ON", code: "on-118" },
    { name: "Calgary",       lat: 51.05, lon: -114.07, province: "AB", code: "ab-52" },
    { name: "Edmonton",      lat: 53.54, lon: -113.49, province: "AB", code: "ab-50" },
    { name: "Vancouver",     lat: 49.25, lon: -123.12, province: "BC", code: "bc-74" },
    { name: "Prince George", lat: 53.91, lon: -122.74, province: "BC", code: "bc-79" },
    { name: "Montreal",      lat: 45.50, lon: -73.57, province: "QC", code: "qc-147" },
    { name: "Quebec City",   lat: 46.81, lon: -71.21, province: "QC", code: "qc-133" }
  ];

  const weatherData = {};
  
  cities.forEach(city => {
    if (!weatherData[city.province]) weatherData[city.province] = [];

    try {
      // Fetch Current (Speed Only) + Daily Forecast
      // Removed 'wind_gusts_10m' to ensure clean data for spreadsheet copy
      const url = `https://api.open-meteo.com/v1/forecast?latitude=${city.lat}&longitude=${city.lon}&current=temperature_2m,wind_speed_10m,weather_code&daily=weather_code,temperature_2m_max,temperature_2m_min,wind_speed_10m_max&timezone=auto`;
      
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) throw new Error("API Error " + res.getResponseCode());
      
      const data = JSON.parse(res.getContentText());
      
      // 1. Current Conditions
      const current = data.current || {};
      const curCode = current.weather_code !== undefined ? current.weather_code : 0;
      const curTemp = Math.round(current.temperature_2m || 0);
      
      // Wind Logic: Speed ONLY (No Gusts), defaulted to 0 to prevent 'NaN' or errors
      const speed = Math.round(current.wind_speed_10m || 0);
      const windStr = speed.toString();
      
      // 2. Forecast (Next 3 days)
      const daily = data.daily || {};
      const forecastData = [];
      
      if (daily.time) {
        for (let i = 1; i <= 3; i++) {
           if (!daily.time[i]) break;
           
           const dDate = new Date(daily.time[i] + "T00:00:00");
           const dayName = Utilities.formatDate(dDate, "America/Toronto", "EEEE");
           const high = Math.round(daily.temperature_2m_max[i]);
           const fSpeed = Math.round(daily.wind_speed_10m_max[i] || 0).toString();
           
           forecastData.push({
             day: dayName,
             temp: `${high}°`,
             wind: fSpeed
           });
        }
      }

      weatherData[city.province].push({ 
          id: city.code, 
          name: city.name, 
          temp: curTemp, 
          wind: windStr, // Pure speed string (e.g. "20")
          condition: wmoToCondition(curCode),
          forecast: forecastData 
      });

    } catch (e) {
      console.warn(`Weather fail for ${city.name}: ${e.message}`);
      weatherData[city.province].push({ 
        id: city.code, name: city.name, 
        temp: "--", wind: "ERR", condition: "Offline", forecast: [] 
      });
    }
  });

  return { weather: weatherData, alerts: [] };
}

function wmoToCondition(code) {
  if (code === 0) return "Clear";
  if (code <= 3) return "Cloudy";
  if (code === 45 || code === 48) return "Fog";
  if (code >= 51 && code <= 57) return "Drizzle";
  if (code >= 61 && code <= 67) return "Rain";
  if (code >= 71 && code <= 77) return "Snow";
  if (code >= 80 && code <= 82) return "Showers";
  if (code >= 85 && code <= 86) return "Snow Showers";
  if (code >= 95) return "Thunderstorm";
  return "Unknown";
}
