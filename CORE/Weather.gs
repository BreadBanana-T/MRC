/**
 * WEATHER MODULE (Severity Priority: Warning > Watch)
 */

const WeatherService = {
  fetch: function() { return fetchWeatherRSS(); }
};

function fetchWeatherRSS() {
  const cities = [
    { name: "Toronto",       code: "on-143", province: "ON" },
    { name: "Ottawa",        code: "on-118", province: "ON" },
    { name: "Calgary",       code: "ab-52",  province: "AB" },
    { name: "Edmonton",      code: "ab-50",  province: "AB" },
    { name: "Vancouver",     code: "bc-74",  province: "BC" },
    { name: "Prince George", code: "bc-79",  province: "BC" }, // [FIXED] Was split across lines
    { name: "Montreal",      code: "qc-147", province: "QC" },
    { name: "Quebec City",   code: "qc-133", province: "QC" }
  ];

  const weatherData = {};
  const alertStats = {}; 
  
  cities.forEach(function(city) {
    if (!weatherData[city.province]) weatherData[city.province] = [];
    if (!alertStats[city.province]) alertStats[city.province] = {};

    try {
      const url = `https://weather.gc.ca/rss/city/${city.code}_e.xml`;
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      
      if (response.getResponseCode() !== 200) {
         weatherData[city.province].push({ id: city.code, name: city.name, temp: "--", wind: "OFF", condition: "Offline", forecast: [] });
         return;
      }

      const xml = response.getContentText();
      const document = XmlService.parse(xml);
      const root = document.getRootElement();
      // [FIX] Use getNamespace() to correctly parse Atom feeds
      const ns = root.getNamespace(); 
      const entries = root.getChildren("entry", ns);

      let temp = "--", windSpeed = "0", condition = "Unknown";
      const forecastData = [];

      for (let i = 0; i < entries.length; i++) {
          const entry = entries[i];
          const title = entry.getChild("title", ns).getText();
          const summary = entry.getChild("summary", ns).getText();
          
          let category = "Unknown";
          const catEl = entry.getChild("category", ns);
          if (catEl) category = catEl.getAttribute("term").getValue();

          // A. CURRENT CONDITIONS
          if (category === "Current Conditions") {
              const parts = title.split(":");
              if (parts.length > 1) {
                  const info = parts[1].trim();
                  const infoParts = info.split(",");
                  condition = infoParts[0].trim();
                  const tempMatch = info.match(/(-?\d+(\.\d+)?)°C/);
                  if (tempMatch) temp = Math.round(parseFloat(tempMatch[1]));
              }
              
              const windMatch = summary.match(/Wind.*?\s+(\d{1,3})\s*km\/h/i);
              const gustMatch = summary.match(/gusting to\s*(\d+)/i);
              
              if (windMatch) {
                  windSpeed = windMatch[1];
                  if(gustMatch) windSpeed += `G${gustMatch[1]}`;
              } else if(summary.match(/calm/i)) {
                  windSpeed = "0";
              }
          }

          // B. WARNINGS
          else if (category === "Warnings and Watches") {
              if (!title.includes("No watches or warnings")) {
                  const cleanDesc = title.replace(/ warning| watch| statement/i, "").trim();
                  
                  let severity = 1;
                  if (title.toLowerCase().includes("warning")) severity = 3;
                  else if (title.toLowerCase().includes("watch")) severity = 2;
                  
                  if (!alertStats[city.province][cleanDesc]) {
                      alertStats[city.province][cleanDesc] = { count: 0, severity: severity };
                  }
                  alertStats[city.province][cleanDesc].count++;
              }
          }

          // C. FORECAST
          else if (category === "Weather Forecasts") {
              // [FIX] REMOVED the "length > 1" check. 
              // This ensures we get the day even if title is just "Wednesday" (no colon).
              let fDay = title.trim();
              if (title.includes(":")) {
                  fDay = title.split(":")[0].trim();
              }

              // Ensure we only grab the next 5-7 days
              if (forecastData.length < 7) {
                  let fWind = "Light";
                  const fWindMatch = summary.match(/wind.*?\s+(\d{1,3})\s*km\/h/i);
                  const fGustMatch = summary.match(/gusting to\s*(\d+)/i);
                  
                  if (fWindMatch) {
                      fWind = fWindMatch[1];
                      if(fGustMatch) fWind += `G${fGustMatch[1]}`;
                  }
                  let fTemp = "";
                  const tMatch = summary.match(/(High|Low|Temperature|Steady).*?(-?\d+)/i);
                  if(tMatch) fTemp = tMatch[2] + "°";

                  forecastData.push({ day: fDay, temp: fTemp, wind: fWind });
              }
          }
      }

      weatherData[city.province].push({ 
          id: city.code, name: city.name, 
          temp: temp, wind: windSpeed, condition: condition,
          forecast: forecastData 
      });
    } catch (e) {
      if (weatherData[city.province]) {
         weatherData[city.province].push({ id: city.code, name: city.name, temp: "--", wind: "ERR", condition: "Error", forecast: [] });
      }
    }
  });

  // --- FINAL ALERT AGGREGATION ---
  const finalAlerts = [];
  Object.keys(alertStats).forEach(prov => {
      const alerts = alertStats[prov];
      let warnings = [];
      let watches = [];

      for (const [type, data] of Object.entries(alerts)) {
          if (data.severity === 3) warnings.push({ type, count: data.count, severity: 3 });
          else if (data.severity === 2) watches.push({ type, count: data.count, severity: 2 });
      }

      let winner = null;
      if (warnings.length > 0) {
          warnings.sort((a,b) => b.count - a.count);
          winner = warnings[0];
      } else if (watches.length > 0) {
          watches.sort((a,b) => b.count - a.count);
          winner = watches[0];
      }

      if (winner) {
          finalAlerts.push({ 
              province: prov, 
              type: winner.type, 
              count: winner.count,
              severity: winner.severity,
              isOverride: false
          });
      }
  });
  return { weather: weatherData, alerts: finalAlerts };
}
