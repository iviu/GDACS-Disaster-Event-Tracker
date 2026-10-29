<p align="center"> GDACS-Disaster-Event-Tracker is
a desktop app that pulls live disaster alerts from the UN's GDACS API, deduplicates them across weeks, and lets you slice the data with filters, charts, and a broken map 😭 </p><p align="center"> <img alt="Python" src="https://img.shields.io/badge/python-3.13%2B-blueviolet?style=for-the-badge&logo=python"> <img alt="License" src="https://img.shields.io/badge/license-MIT-lightgrey?style=for-the-badge"> <img alt="Data" src="https://img.shields.io/badge/data-GDACS%20API-ff69b4?style=for-the-badge"></p>

---

### hi, i was bored
#### it was a tuesday. i had three hours before my next class. i thought "what if i could see every earthquake, cyclone, flood, volcano and wildfire that the UN is currently tracking in real time on my laptop?"

#### three hours later this thing existed. so. here we are.

#### GDACS is the Global Disaster Alert and Coordination System — a joint thing by the UN and the EU that tracks major disasters worldwide. their API is public, it's JSON, and it's beautifully undocumented in the fun way where you just curl it and see what comes back.

---

## features:
- pulls everything in the last N days that's rated red or orange (the serious ones)
- deduplicates across runs so you don't see the same earthquake 47 times
- saves everything into a growing JSON log
- shows you a live table with search, filters, sorting, and column chooser
- draws duration timelines, a world map of events, and a trend line of new vs skipped per run
- exports to CSV, Excel, JSON, Markdown or clipboard
- has dark mode (because why not)

## what it actually does
###  live data, properly deduplicated
- fetches from the official GDACS geteventlist/SEARCH endpoint
- paginates through all pages (up to 200, page size 10–500)
- filters by event type (EQ/TC/FL/VO/WF/DR) and alert level (red/orange)
- dedupes by event ID across a configurable window (default 28 days) so a week-old earthquake that's still listed doesn't get re-added
- appends only new events to a weekly JSON log

### filtering that doesn't suck
- text search across ID/type/alert/country/headline
- dropdown filters for type, alert level, country
- minimum duration filter (only events lasting ≥ N days)
- debounced (300 ms) so typing doesn't re-render every keystroke
- saved filter presets you can name, reuse, and delete
- drill-down — double-click a country in the summary tab and it filters the main view

### 5 analytics tabs
- Countries — event count per country, with flag emoji where available
- Event Types — count per type (EQ, TC, FL, …)
- Alert Levels — red vs orange counts
- Duration — average event duration by country and by type
- Timeline — horizontal bar chart of the last 80 events, colored by alert level
- Map — scatter plot of every event's latitude/longitude, colored by alert level
- Trends — new vs skipped events per run over time

### save anything, anywhere
- CSV (Ctrl+E)
- Excel .xlsx (needs openpyxl)
- JSON (round-trippable — you can re-import it)
- Markdown report with auto-generated summary tables
- raw analysis text (auto-generated every run)
- copy current view to clipboard (or just the selected rows)

### settings that persist
- lookback window, dedupe window, page size
- auto-refresh interval (0 = off, up to 1440 minutes)
- HTTP cache TTL (so re-running within N minutes doesn't hammer the API)
- alert level & event type toggles
- auto-export CSV to a chosen directory after each fetch
- dark/light theme
- window geometry, column order, and hidden columns all remembered

### it's also a nice engineering story
- atomic file writes (temp file + os.replace) so a crash mid-write doesn't corrupt your log
- corrupt log recovery — a corrupt log.json is backed up and you start fresh, no crash
- rate-limit & 5xx retries with jitter (via urllib3.Retry)
- background worker threads — the UI never freezes
- cancellable fetches — hit Cancel and it stops mid-pagination
- undo last fetch — accidentally imported a week you didn't want? Ctrl+Z (well, Ctrl+Z conceptually, the button says "Undo last fetch")
- run history with duration, new, skipped counts

---

## installation
#### you'll need Python 3.10+ and these:
```
pip install requests pandas
```
#### optional but nice:
```
pip install matplotlib openpyxl
```
- `matplotlib` → the Timeline, Map and Trends tabs (without it, those tabs just say "install matplotlib")
- `openpyxl` → the "Export → Excel" button (without it, it politely refuses)
## run it
```
python weather.py
```
##### (i know the file is named weather.py. i named it that at 2 AM because it looked "weather-like events" 😭)
## quick start
### 1. launch
```
python weather.py
```
#### it auto-fetches on startup. after ~3 seconds you have a table of every red/orange disaster from the last 7 days.
### 2. look at a specific event
#### click a row. the right panel shows the event ID, type, alert level, country, dates, duration, lat/lon, and the headline.
### 3. open it on a map (need a fix)
#### select a row with coordinates, hit Open in OSM (or double-click the row, or press Enter). OpenStreetMap opens in your browser, centered on the event.
### 4. filter
- type in the search box (it debounces 300 ms)
- pick a Type from the dropdown (EQ = earthquake, TC = tropical cyclone, FL = flood, VO = volcano, WF = wildfire, DR = drought)
- pick an Alert (red or orange)
- pick a Country (only countries actually in the dataset show up)
- enter a Min days value to hide short-lived events

### 5. save a preset
#### Presets → Save current… → name it → next time you're one click away.

### 6. export
#### Export ▾ menu — CSV, Excel, JSON, Markdown, or copy to clipboard.

### 7. change the fetch window
#### Settings tab → "Lookback days" → for example, 30 to see a month of data. Save. Next refresh uses the new window.

### 8. rate-limit yourself
#### by default there's no HTTP cache. if you want to re-run repeatedly without hammering the API, set "HTTP cache TTL" to 5 — repeated runs within 5 minutes reuse the cached response.

---

## what's in each tab
### Events
#### the main table. sortable columns, right-click menu (open in OSM, copy ID, copy rows, filter by this country/type/alert, clear filters), multi-select, and columns you can toggle and reorder.

### Countries / Event Types / Alert Levels
#### one summary table per dimension. double-click a row to filter the main table to that value. yes, this means you can drill from "countries with the most earthquakes" to "those events" in two clicks.

### Duration
#### plain-text stats: average event duration by country and by event type. useful for "which type of disaster lasts longest?" (spoiler: droughts, by a lot)

### Timeline
#### a horizontal bar chart of the last 80 events. each bar is one event's fromdate to todate, colored red or orange by alert level.

### Broken Map
#### a scatter plot of every event that has usable lat/lon. red/orange by alert. it's not a real map (no coastlines) — it's a quick "are these events clustered somewhere?" view. works fine.

### Trends
#### new vs skipped counts per run, over your last ~200 fetches. useful for "am I actually finding new events or just re-skipping the same 200?"

### History
#### table of every past run with timestamp, new count, skipped count, rows in log, duration, and whether it was cancelled. Clear history button if you want a fresh slate.

### Settings
#### lookback, dedupe window, page size, auto-refresh, HTTP cache TTL, alert/type toggles, auto-export settings, flag toggle. Save / Reset / Open app log.

### Log
#### every log line the app has ever produced, with colors for ok/warn/error. Save log to file… button. clears automatically when it hits 5000 lines.

----

## known limitations
- `country` deduplication is imperfect — an event affecting "USA, Mexico" is one row per country in the final table. good for statistics, but if you export CSV you'll see the event ID twice.
- map is not a real geographic map — it's a scatter plot. if you want an actual map with coastlines, install cartopy or folium — I didn't because matplotlib's basemap is dead and cartopy is a heavy dependency for an exam project.
- no scheduled background fetches when the app is closed — auto-refresh only fires while it's running.
- rate-limit handling is real but not perfect — the GDACS API is generally very accommodating but if you hammer it with page_size=500 and lookback=365 on a tight loop you will get 429s. the HTTP cache TTL exists for this exact reason.
- dedupe window is time-based on fromdate — an event that happened 30 days ago but is still ongoing will be re-fetched and skipped if the dedupe window is ≥ 30 days. that's usually what you want, but it's a trade-off.
- no authentication on the API — GDACS is public, so there's nothing to configure, but also no way to get higher rate limits.
- weather.py filename — yes, I know. it's historical. renaming it would break everyone's shortcuts.

---

## license
MIT. do what you want with it. if you use it to actually track disasters, please also check the official GDACS site — this is a toy client and they're the ones actually doing the important work.

---

<p align="center"> made with love and unreasonable amount of atomic file writes</p>
