# Canine Companions Geographic Dashboard

I built this to help visualize where our clients and constituents live across the country. Started as a simple client heat map and turned into something a lot more useful. You can filter by region, state, and county, pull up individual client info, and plan field visit routes. Other departments can use it too since the column mapping is flexible enough to work with different Salesforce reports.

Runs completely in the browser. Nothing to install, no server, just open it and upload a CSV.

---

## What it does

Upload a Salesforce report and it maps everything out by county. Counties shade from light blue to amber depending on how many clients are there. The scale is adjustable so it works whether you have 20 clients in a region or 600.

A few things you can do once it loads:

- Filter by region or state and the map zooms in automatically
- Click a county to see every client there with their contact info and a direct link to their Salesforce record
- Toggle on all CC locations including training centers, the field office, and all 55+ volunteer chapters
- Run the route optimizer for any region to get a suggested visit schedule based on where clients are clustered

---

## Files

```
index.html   — page layout
styles.css   — all styling, mobile friendly
app.js       — all the logic
README.md    — this
```

All three files need to be in the same folder or it won't work.

---

## How to use it

Open in Chrome, Edge, or Firefox. First time needs internet to load the US county map data, after that it caches and works offline.

1. Click Upload Data CSV and select your Salesforce export
2. A column mapper shows up, confirm which column is County, State, First Name, Last Name, and Client ID. Everything else gets captured automatically.
3. Map populates, start filtering

At minimum the CSV needs a county and state column. Names and Client ID are optional but you need them to see the full client detail and the Salesforce links.

Columns it auto-detects:
- County
- Mailing State/Province
- First Name / Last Name
- Client ID

---

## Route planning

Click Optimize Routes, set a travel radius (anywhere from 25 to 250 miles depending on the region), then pick which region. It clusters client counties into stops and orders the route starting from the training center. Northwest also locks in the Puget Sound Field Office as a stop.

The radius slider makes a big difference. Something like New York you might want 50 miles. Montana or Wyoming you could go up to 200+.

---

## Regions

Based on Canine Companions official territory. California, Nevada, and Pennsylvania are all split at the county level since they each fall into two regions.

| Region | States |
|---|---|
| Northeast | ME, NH, VT, MA, RI, CT, NY, NJ, Eastern PA, DE, MD, DC, VA, WV |
| North Central | OH, KY, MI, IN, IL, WI, MO, IA, MN, KS, NE, ND, SD, Western PA |
| Southeast | FL, GA, TN, NC, SC, MS, AL |
| South Central | TX, AR, LA, OK |
| Southwest | AZ, UT, CO, NM, HI, Southern CA, Southern NV |
| Northwest | WA, OR, ID, MT, WY, AK, Northern CA, Northern NV |

Pennsylvania: 28 western and central counties go to North Central (Allegheny, Erie, Pittsburgh area, etc.) and the other 39 go to Northeast (Philadelphia, Lehigh Valley, Scranton area, etc.)

California: Southern counties like LA, San Diego, Orange, Riverside, and San Bernardino go to Southwest. Everything north is Northwest.

Nevada: Just Clark County (Las Vegas) is Southwest, everything else is Northwest.

---

## Updating

New data just upload a new CSV, no need to refresh. For code updates replace the file on GitHub and the site rebuilds within a few minutes.

---

## Notes

- CSV data never leaves your browser, nothing is sent anywhere
- Map boundary data loads from a CDN on first open then caches locally
- Salesforce links open to canine.lightning.force.com/lightning/r/Account/[Client ID]/view
- Works on mobile but there's a lot going on, desktop is way better
