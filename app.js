const thingSpeakChannelId = '3531601';
const apiUrl = `https://api.thingspeak.com/channels/${thingSpeakChannelId}/feeds.json?results=48`;
const temperatureElement = document.querySelector('#temperature');
const humidityElement = document.querySelector('#humidity');
const temperatureLowElement = document.querySelector('#temperatureLow');
const temperatureHighElement = document.querySelector('#temperatureHigh');
const humidityLowElement = document.querySelector('#humidityLow');
const humidityHighElement = document.querySelector('#humidityHigh');
const heatIndexElement = document.querySelector('#heatIndex');
const dewPointElement = document.querySelector('#dewPoint');
const comfortElement = document.querySelector('#comfortStatus');
const connectionElement = document.querySelector('#connectionText');
const chart = document.querySelector('#historyChart');
const context = chart.getContext('2d');

let temperatures = [];
let humidities = [];

function number(value, digits = 1) {
  return value === null || value === undefined ? '--' : Number(value).toFixed(digits);
}

function updateExtremes(data) {
  const temperatureValues = temperatures.length ? temperatures : [data.temperature];
  const humidityValues = humidities.length ? humidities : [data.humidity];
  temperatureLowElement.textContent = `${number(data.temperatureMin ?? Math.min(...temperatureValues))}°C`;
  temperatureHighElement.textContent = `${number(data.temperatureMax ?? Math.max(...temperatureValues))}°C`;
  humidityLowElement.textContent = `${number(data.humidityMin ?? Math.min(...humidityValues), 0)}%`;
  humidityHighElement.textContent = `${number(data.humidityMax ?? Math.max(...humidityValues), 0)}%`;
}

function drawChart() {
  const width = chart.width;
  const height = chart.height;
  context.clearRect(0, 0, width, height);
  context.strokeStyle = '#dce1d9';
  context.lineWidth = 1;
  for (let row = 1; row < 5; row += 1) {
    const y = row * height / 5;
    context.beginPath(); context.moveTo(0, y); context.lineTo(width, y); context.stroke();
  }
  const plot = (values, color, min, max) => {
    if (values.length < 2) return;
    context.strokeStyle = color; context.lineWidth = 4; context.lineJoin = 'round'; context.lineCap = 'round'; context.beginPath();
    values.forEach((value, index) => {
      const x = index * width / (values.length - 1);
      const y = height - Math.max(0, Math.min(1, (value - min) / (max - min))) * height;
      if (index === 0) context.moveTo(x, y); else context.lineTo(x, y);
    });
    context.stroke();
  };
  plot(temperatures, '#e87652', 18, 32);
  plot(humidities, '#168b83', 0, 100);
}

function showReading(data) {
  temperatures = (data.temperatureHistory || []).map(Number);
  humidities = (data.humidityHistory || []).map(Number);
  temperatureElement.textContent = number(data.temperature);
  humidityElement.textContent = number(data.humidity, 0);
  heatIndexElement.textContent = `${number(data.heatIndex)}°C`;
  dewPointElement.textContent = `${number(data.dewPoint)}°C`;
  comfortElement.textContent = data.comfort || 'Waiting...';
  updateExtremes(data);
  drawChart();
  document.querySelector('#lastUpdate').textContent = new Date().toLocaleTimeString([], { hour12: false });
}

async function readLiveSensor() {
  try {
    const response = await fetch(apiUrl, { cache: 'no-store' });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const feed = await response.json();
    const readings = feed.feeds.map((entry) => ({
      temperature: Number(entry.field1),
      humidity: Number(entry.field2),
      heatIndex: Number(entry.field3),
      dewPoint: Number(entry.field4),
    })).filter((entry) => Number.isFinite(entry.temperature) && Number.isFinite(entry.humidity));
    if (!readings.length) throw new Error('No sensor readings');
    const latest = readings[readings.length - 1];
    showReading({
      ...latest,
      temperatureMin: Math.min(...readings.map((entry) => entry.temperature)),
      temperatureMax: Math.max(...readings.map((entry) => entry.temperature)),
      humidityMin: Math.min(...readings.map((entry) => entry.humidity)),
      humidityMax: Math.max(...readings.map((entry) => entry.humidity)),
      temperatureHistory: readings.map((entry) => entry.temperature),
      humidityHistory: readings.map((entry) => entry.humidity),
      comfort: latest.temperature >= 35 || latest.humidity >= 80 ? 'Caution: hot or humid' : 'Live sensor data',
    });
    connectionElement.textContent = 'Live ESP32 stream';
  } catch (error) {
    connectionElement.textContent = 'ESP32 offline - connect to its Wi-Fi';
  }
}

function clock() { document.querySelector('#clock').textContent = new Date().toLocaleTimeString([], { hour12: false }); }

drawChart(); clock(); readLiveSensor();
setInterval(readLiveSensor, 2000);
setInterval(clock, 1000);
window.addEventListener('resize', drawChart);
