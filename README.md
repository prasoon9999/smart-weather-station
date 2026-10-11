# Smart Weather Station

> A live, cloud-connected weather dashboard powered by an **ESP32 + DHT22** sensor.  
> Readings are published to ThingSpeak every 20 seconds and displayed in real-time at the link below.

🌐 **Live site:** https://smart-weather-station.vercel.app/  
📦 **GitHub:** https://github.com/prasoon9999/smart-weather-station

---

## Features

- 🌡️ Live **temperature** and **humidity** from a DHT22 sensor via ESP32
- 🔢 Calculated **heat index** and **dew point**
- 📈 **24-hour history chart** (canvas, no external chart library)
- 🔴 Live **min / max** extremes for the current session
- ✅ **Comfort status** indicator
- 📡 Auto-refreshes every **2 seconds** via ThingSpeak public API
- 📱 Fully **responsive** (mobile + desktop)
- ⚡ Zero npm dependencies — pure HTML / CSS / JavaScript

---

## Hardware

| Component | Role |
|-----------|------|
| ESP32 | Wi-Fi microcontroller |
| DHT22 | Temperature & humidity sensor |
| SSD1306 OLED | Local display |

### Wiring

| Pin | Connects to |
|-----|-------------|
| OLED SDA | GPIO 21 |
| OLED SCK | GPIO 22 |
| DHT22 OUT | GPIO 4 |

---

## How it works

```
DHT22 → ESP32 → ThingSpeak (cloud) → This dashboard (Vercel)
```

1. The ESP32 reads temperature and humidity from the DHT22 every 20 s.
2. It calculates heat index and dew point on-device.
3. It publishes all four values to a **public ThingSpeak channel** (fields 1–4).
4. This static dashboard fetches the last 48 entries from ThingSpeak every 2 s and renders them live.

---

## Run locally

Just open `index.html` in any browser — no build step, no server needed.

```bash
# Optional: serve with Node
npx serve .
```

---

## Deploy to Vercel

1. Push this folder to a GitHub repository.
2. Import the repository at [vercel.com/new](https://vercel.com/new).
3. Framework preset → **Other**
4. Build command → *(leave empty)*
5. Output directory → `.`
6. Click **Deploy**.

---

## Live ThingSpeak setup (connect your own ESP32)

1. Create a [ThingSpeak](https://thingspeak.com/) channel with four fields:
   - Field 1 → Temperature (°C)
   - Field 2 → Humidity (%)
   - Field 3 → Heat Index (°C)
   - Field 4 → Dew Point (°C)
2. Copy the **Channel ID** from the ThingSpeak dashboard.
3. Open `app.js` and set `thingSpeakChannelId` to your Channel ID.
4. In the ESP32 firmware (`firmware/main.cpp`), fill in:
   - Your phone hotspot SSID and password
   - Your ThingSpeak **write API key**
5. Flash the firmware while your hotspot is active.
6. Deploy this folder to Vercel.

> ⚠️ **Never commit your ThingSpeak write key or hotspot password to a public repo.**

---

## Project structure

```
smart-weather-station/
├── index.html      — Dashboard markup
├── style.css       — All styles (Manrope + DM Mono)
├── app.js          — Data fetching, chart, live updates
├── package.json    — Metadata (no runtime dependencies)
├── README.md       — This file
└── firmware/       — ESP32 PlatformIO project (main.cpp)
```

---

## License

MIT
