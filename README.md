# Canine Companions Geographic Dashboard

I built this to help me (and eventually other departments) get a better visual on where our clients and constituents are across the country. It started as a client heat map and grew into a full geographic tool — you can drill down by region, state, and county, pull up individual client info, and even plan out field visit routes.

It runs entirely in the browser. No server, no login, nothing to install. Just open it and upload a CSV.

---

## How it works

Upload your Salesforce report and the dashboard maps everyone out by county. Counties get shaded based on how many clients are there — light blue for low, orange/amber for high. You can adjust the scale with sliders if your numbers are smaller or larger than the defaults.

From there:
- Click a region to zoom in and see which states have the most clients
- Click a state to go deeper and see the county breakdown
- Click a county to pull up every client there with their contact info and a link straight to their Salesforce record
- Use the CC Locations button to overlay all training centers, the field office, and all volunteer chapters on the map
- Use Optimize Routes to generate a visit schedule for any region — it clusters clients into stops within your chosen travel radius and gives you an ordered route starting from the regional training center

---

## Files

```
index.html   — the page structure
styles.css   — all the styling (works on mobile too)
app.js       — all the logic
README.md    — you're reading it
```

All three need to stay in the same folder together.

---

## Using it

Open it in Chrome, Edge, or Firefox. You'll need internet the first time so it can load the US county boundary data. After that it caches everything and works offline.

1. Hit **Upload Data CSV** and pick your Salesforce export
2. A column mapper pops up — just confirm which column is County, State, First Name, Last Name, and Client ID. Everything else gets pulled in automatically.
3. Map fills in. Filter away.

The CSV needs at minimum a county column and a state column. Everything else (names, contact info, Client ID) is optional but unlocks the full client detail view and the Salesforce links.

Columns it looks for automatically:
- County
- Mailing State/Province
- First Name / Last Name
- Client ID

---

## Route planning

Hit **Optimize Routes**, set your travel radius with the slider (25–250 miles), then pick a region. It finds all the counties with clients, groups the ones that are close together into stops, and orders the whole trip starting from the training center. For Northwest it also locks in the Puget Sound Field Office as a required stop.

Each stop shows how many clients are within range and which counties are covered. Click a stop to zoom the map to it. The dashed circles show your selected radius.

---

## Regions

Follows Canine Companions' official six regions. California, Nevada, and Pennsylvania are split by county.

| Region | States |
|---|---|
| Northeast | ME, NH, VT, MA, RI, CT, NY, NJ, Eastern PA, DE, MD, DC, VA, WV |
| North Central | OH, KY, MI, IN, IL, WI, MO, IA, MN, KS, NE, ND, SD, Western PA |
| Southeast | FL, GA, TN, NC, SC, MS, AL |
| South Central | TX, AR, LA, OK |
| Southwest | AZ, UT, CO, NM, HI, Southern CA, Southern NV |
| Northwest | WA, OR, ID, MT, WY, AK, Northern CA, Northern NV |

**Pennsylvania split** — 28 western/central counties go to North Central (Erie, Crawford, Allegheny, Pittsburgh area, etc.). The remaining 39 eastern counties go to Northeast (Philadelphia, Lehigh Valley, Scranton area, etc.).

**California** — Southern CA counties (LA, San Diego, Orange, Riverside, San Bernardino, Ventura, Santa Barbara, San Luis Obispo, Imperial, Kern) go to Southwest. Everything north goes to Northwest.

**Nevada** — Clark County (Las Vegas) is Southwest. Everything else is Northwest.

---

## Updating

New CSV? Just upload it — no need to refresh the page.

Code change? Replace the file in the GitHub repo and the live site updates in a couple minutes.

---

## A few things worth knowing

- Nothing gets sent anywhere. The CSV stays in your browser the whole time.
- First load needs internet to pull the map data. After that it's cached and works offline.
- Salesforce record links go to: `canine.lightning.force.com/lightning/r/Account/[Client ID]/view`
- Works on mobile but honestly it's a lot of information — desktop is better.
