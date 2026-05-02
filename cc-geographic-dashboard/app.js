// ─── REGION CONFIG ────────────────────────────────────────────────
// Canine Companions brand colors: #0073FF (blue), #FECB00 (yellow), #425563 (grey)
// 6 region colors derived from CC brand palette
const REGION_COLORS = [
  '#0073FF',  // Northeast     — CC primary blue
  '#00b4d8',  // North Central — cyan/teal (clearly different from blue)
  '#FECB00',  // Southeast     — CC yellow
  '#e05c00',  // South Central — burnt orange (clearly different from yellow)
  '#425563',  // Southwest     — CC slate grey
  '#6a0dad',  // Northwest     — purple (clearly different from all others)
];
// 0=Northeast, 1=North Central, 2=Southeast, 3=South Central, 4=Southwest, 5=Northwest
let regionNames = ['Northeast','North Central','Southeast','South Central','Southwest','Northwest'];

// Canine Companions official 6-region mapping (canine.org)
// Note: CA and NV are split between Southwest (south) and Northwest (north) — mapped to Northwest
// as the majority of those states' area falls in that territory; user can override via CSV region column
const STATE_REGION_DEFAULT = {
  // Northeast: NY, NJ, CT, DE, PA, MD, DC, VA, WV, MA, RI, VT, NH, ME
  'ME':0,'NH':0,'VT':0,'MA':0,'RI':0,'CT':0,'NY':0,'NJ':0,'PA':0,
  'DE':0,'MD':0,'DC':0,'VA':0,'WV':0,
  // North Central: OH, KY, MI, IN, IL, WI, MO, IA, MN, KS, NE, ND, SD
  'OH':1,'KY':1,'MI':1,'IN':1,'IL':1,'WI':1,'MO':1,'IA':1,'MN':1,'KS':1,'NE':1,'ND':1,'SD':1,
  // Southeast: FL, GA, TN, NC, SC, MS, AL
  'FL':2,'GA':2,'TN':2,'NC':2,'SC':2,'MS':2,'AL':2,
  // South Central: TX, AR, LA, OK
  'TX':3,'AR':3,'LA':3,'OK':3,
  // Southwest: AZ, UT, CO, NM, HI (+ Southern CA/NV — statewide defaulted here)
  'AZ':4,'UT':4,'CO':4,'NM':4,'HI':4,
  // Northwest: WA, OR, CA, ID, MT, WY, AK, NV
  'WA':5,'OR':5,'CA':5,'ID':5,'MT':5,'WY':5,'AK':5,'NV':5
};

// ─── STATE ────────────────────────────────────────────────────────
let clientData = {}; // { "county_state": { count, region, clients[] } }
let activeRegion = 'all';  // 'all' or index string '0'-'5'
let activeState  = 'all';  // 'all' or state abbr e.g. 'CA'
let csvHeaders = [];
let csvRows = [];
let countyPaths = {}; // fips → SVGElement
let mapLoaded = false;
let allCountyFeatures = [];
let allStateFeatures  = [];
let geoPathFn = null;
let activeCounty = null; // { fips, name, state } when a county is selected
let sfBaseUrl    = '';   // Salesforce org base URL, set in column mapping modal

// ─── MAP LOADING ──────────────────────────────────────────────────
async function loadMap() {
  const loadingEl = document.getElementById('loading');
  try {
    const sources = [
      'https://cdn.jsdelivr.net/npm/us-atlas@3/counties-albers-10m.json',
      'https://unpkg.com/us-atlas@3/counties-albers-10m.json',
    ];
    let topo = null, lastErr = null;
    for (const url of sources) {
      try {
        const r = await fetch(url);
        if (!r.ok) throw new Error(`HTTP ${r.status}`);
        topo = await r.json();
        break;
      } catch(e) { lastErr = e; }
    }
    if (!topo) throw lastErr;
    renderMap(topo);
    loadingEl.style.display = 'none';
    mapLoaded = true;
  } catch(e) {
    loadingEl.innerHTML = `
      <div style="text-align:center;padding:20px;max-width:360px">
        <div style="font-size:2rem;margin-bottom:12px">🗺️</div>
        <div style="color:#1a2a3a;font-weight:600;margin-bottom:8px">Map data could not be loaded</div>
        <div style="color:#6b7e8f;font-size:0.78rem;line-height:1.6;margin-bottom:16px">
          This map requires an internet connection to load US county boundaries.<br>
          Open this file directly in Chrome, Edge, or Firefox while online.
        </div>
        <button onclick="loadMap()" style="background:#0073FF;color:#fff;border:none;border-radius:7px;padding:9px 20px;font-family:Poppins,sans-serif;font-size:0.78rem;font-weight:600;cursor:pointer">↻ Retry</button>
        <div style="color:#aab8c5;font-size:0.65rem;margin-top:10px">${e?.message || 'Network error'}</div>
      </div>`;
  }
}

function renderMap(topo) {
  const svg = document.getElementById('map-svg');
  const counties = topojson.feature(topo, topo.objects.counties);
  const states   = topojson.feature(topo, topo.objects.states);

  geoPathFn = d3.geoPath();
  allCountyFeatures = counties.features;
  allStateFeatures  = states.features;

  const countyGroup = document.createElementNS('http://www.w3.org/2000/svg','g');
  counties.features.forEach(f => {
    const el = document.createElementNS('http://www.w3.org/2000/svg','path');
    el.setAttribute('d', geoPathFn(f));
    el.setAttribute('class','county-path');
    el.dataset.fips = f.id;
    el.dataset.name = f.properties.name;
    countyPaths[f.id] = el;

    const regionIdx = getRegionByFips(f.id);
    el.setAttribute('fill', '#ffffff');
    el.setAttribute('stroke', REGION_COLORS[regionIdx]);
    el.setAttribute('stroke-width', '0.6');

    el.addEventListener('mousemove', e => showTooltip(e, f));
    el.addEventListener('mouseleave', hideTooltip);
    el.addEventListener('click', () => selectCounty(f));
    countyGroup.appendChild(el);
  });

  const stateGroup = document.createElementNS('http://www.w3.org/2000/svg','g');
  states.features.forEach(f => {
    const el = document.createElementNS('http://www.w3.org/2000/svg','path');
    el.setAttribute('d', geoPathFn(f));
    el.setAttribute('class','state-path');
    stateGroup.appendChild(el);
  });

  // Region boundary layer — one thick colored path per region edge
  // Uses topojson.mesh with a filter: only segments where the two
  // neighboring counties belong to DIFFERENT regions (outer region border)
  const regionBoundaryGroup = document.createElementNS('http://www.w3.org/2000/svg','g');
  REGION_COLORS.forEach((color, rIdx) => {
    // mesh filter: keep arcs that are either on the exterior (b===a, i.e. border of US)
    // OR separate two counties where exactly one side is this region
    const mesh = topojson.mesh(topo, topo.objects.counties, (a, b) => {
      const ra = getRegionByFips(a.id);
      const rb = getRegionByFips(b.id);
      // Keep this arc if it borders region rIdx from either side, and the other side is a different region
      return (ra === rIdx || rb === rIdx) && ra !== rb;
    });
    const d = geoPathFn(mesh);
    if (!d || d === 'M' || d.length < 4) return; // skip empty
    const el = document.createElementNS('http://www.w3.org/2000/svg','path');
    el.setAttribute('d', d);
    el.setAttribute('class','region-boundary');
    el.setAttribute('stroke', '#0073FF');
    el.setAttribute('stroke-width', '2');
    regionBoundaryGroup.appendChild(el);
  });

  svg.appendChild(countyGroup);
  svg.appendChild(stateGroup);
  svg.appendChild(regionBoundaryGroup);

  // Build the county name → FIPS lookup now that all paths are registered
  buildReverseLookup();
  // Initialize the Albers projection for org location markers
  initOrgProjection();
}

// ─── ZOOM ─────────────────────────────────────────────────────────
const BASE_VIEWBOX = { x: 0, y: 0, w: 960, h: 600 };
let currentViewBox = { ...BASE_VIEWBOX };
let zoomAnimFrame = null;

function zoomToSelection() {
  if (!geoPathFn || !allCountyFeatures.length) return;

  // Collect the features that match the current filter
  let features = [];
  if (activeState !== 'all') {
    // Zoom to counties within that state
    features = allCountyFeatures.filter(f => fipsToState(f.id) === activeState);
  } else if (activeRegion !== 'all') {
    // Zoom to all counties in the region
    features = allCountyFeatures.filter(f => getRegionByFips(f.id) === +activeRegion);
  } else {
    zoomToViewBox(BASE_VIEWBOX);
    return;
  }

  if (!features.length) return;

  // Compute union bounding box of all selected features
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  features.forEach(f => {
    const b = geoPathFn.bounds(f);
    if (b[0][0] < minX) minX = b[0][0];
    if (b[0][1] < minY) minY = b[0][1];
    if (b[1][0] > maxX) maxX = b[1][0];
    if (b[1][1] > maxY) maxY = b[1][1];
  });

  // Add padding
  const pad = activeState !== 'all' ? 30 : 20;
  minX -= pad; minY -= pad; maxX += pad; maxY += pad;

  // Clamp to map bounds
  minX = Math.max(0, minX);
  minY = Math.max(0, minY);
  maxX = Math.min(960, maxX);
  maxY = Math.min(600, maxY);

  zoomToViewBox({ x: minX, y: minY, w: maxX - minX, h: maxY - minY });
  document.getElementById('zoom-out-btn').classList.add('visible');
}

function zoomToViewBox(target) {
  if (zoomAnimFrame) cancelAnimationFrame(zoomAnimFrame);
  const svg = document.getElementById('map-svg');
  const start = { ...currentViewBox };
  const duration = 520; // ms
  const startTime = performance.now();

  function easeInOut(t) {
    return t < 0.5 ? 2 * t * t : -1 + (4 - 2 * t) * t;
  }

  function step(now) {
    const elapsed = now - startTime;
    const t = Math.min(elapsed / duration, 1);
    const e = easeInOut(t);

    currentViewBox = {
      x: start.x + (target.x - start.x) * e,
      y: start.y + (target.y - start.y) * e,
      w: start.w + (target.w - start.w) * e,
      h: start.h + (target.h - start.h) * e,
    };

    svg.setAttribute('viewBox', `${currentViewBox.x} ${currentViewBox.y} ${currentViewBox.w} ${currentViewBox.h}`);
    syncRouteViewBox();

    if (t < 1) {
      zoomAnimFrame = requestAnimationFrame(step);
    } else {
      zoomAnimFrame = null;
    }
  }

  zoomAnimFrame = requestAnimationFrame(step);
}

// ─── COUNTY SELECTION ─────────────────────────────────────────────
function selectCounty(feature) {
  const fips  = feature.id;
  const name  = feature.properties.name;
  const state = fipsToState(fips);

  // Toggle off if clicking same county
  if (activeCounty && activeCounty.fips === fips) {
    closeCountyAndZoom();
    return;
  }

  // Clear previous highlight
  if (activeCounty) {
    const prev = countyPaths[activeCounty.fips];
    if (prev) prev.classList.remove('county-highlighted');
  }

  activeCounty = { fips, name, state };

  const el = countyPaths[fips];
  if (el) el.classList.add('county-highlighted');

  zoomToCounty(feature);
  openCountyPanel(name, state, fips);
}

function zoomToCounty(feature) {
  if (!geoPathFn) return;
  const b   = geoPathFn.bounds(feature);
  const pad = 40;
  let minX = Math.max(0,   b[0][0] - pad);
  let minY = Math.max(0,   b[0][1] - pad);
  let maxX = Math.min(960, b[1][0] + pad);
  let maxY = Math.min(600, b[1][1] + pad);

  // Minimum visible area for tiny counties
  const minSize = 80;
  if (maxX - minX < minSize) { const cx=(minX+maxX)/2; minX=cx-minSize/2; maxX=cx+minSize/2; }
  if (maxY - minY < minSize) { const cy=(minY+maxY)/2; minY=cy-minSize/2; maxY=cy+minSize/2; }

  zoomToViewBox({ x: minX, y: minY, w: maxX - minX, h: maxY - minY });
  document.getElementById('zoom-out-btn').classList.add('visible');
}

function openCountyPanel(countyName, stateAbbr, fips) {
  const panel     = document.getElementById('county-panel');
  const rIdx      = getRegionByFips(fips);
  const color     = REGION_COLORS[rIdx];
  const stateName = (ALL_STATES.find(s => s.abbr === stateAbbr) || {}).name || stateAbbr;

  document.getElementById('cp-name').textContent = countyName + ' County';
  document.getElementById('cp-sub').innerHTML =
    `<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${color};margin-right:4px"></span>${stateName} · ${regionNames[rIdx]}`;

  const key   = makeKey(countyName, stateAbbr);
  const info  = clientData[key];
  // Respect active region filter — if this county belongs to a different region, treat as no data
  const regionMismatch = activeRegion !== 'all' && info && info.regionIdx !== +activeRegion;
  const listEl  = document.getElementById('cp-client-list');
  const countEl = document.getElementById('cp-count');
  listEl.innerHTML = '';

  if (!info || info.count === 0 || regionMismatch) {
    countEl.textContent = 'No clients on record';
    listEl.innerHTML = `
      <div class="county-no-data">
        <div style="font-size:1.8rem;margin-bottom:8px">📭</div>
        No clients mapped to this county yet.<br>
        Upload your Salesforce CSV to populate client data.
      </div>`;
  } else {
    countEl.textContent = `${info.count} client${info.count !== 1 ? 's' : ''} in this county`;

    info.clients.forEach(client => {
      const f       = client.fields || {};
      const firstName = f['First Name'] || '';
      const lastName  = f['Last Name']  || '';
      const fullName  = client.name || [firstName, lastName].filter(Boolean).join(' ') || '(No name)';
      const tracker   = client.tracker || '';
      const initials  = fullName.trim().split(/\s+/).map(w => w[0]).slice(0,2).join('').toUpperCase() || '?';
      const sfLink    = tracker && sfBaseUrl ? `${sfBaseUrl}/lightning/r/Account/${tracker}/view` : '';

      let fieldsHtml = '';
      DISPLAY_FIELDS.forEach(({ header, label, icon }) => {
        const val = f[header];
        if (!val) return;
        const displayLabel = label || header;
        fieldsHtml += `
          <div class="client-field client-field-row">
            <span class="client-field-icon">${icon}</span>
            <span class="client-field-label">${displayLabel}:</span>
            <span class="client-field-value">${header === 'Email'
              ? `<a href="mailto:${val}" style="color:var(--cc-blue);text-decoration:none">${val}</a>`
              : header === 'Phone'
              ? `<a href="tel:${val}" style="color:var(--cc-blue);text-decoration:none">${val}</a>`
              : val}</span>
          </div>`;
      });

      const row = document.createElement('div');
      row.className = 'county-client-row';
      row.innerHTML = `
        <div class="client-avatar" style="background:${color}">${initials}</div>
        <div class="client-details">
          <div class="client-name">${fullName}</div>
          ${tracker ? `<div class="client-field">${sfLink
            ? `<a href="${sfLink}" target="_blank" class="client-tracker-link">🔗 ID: ${tracker} ↗</a>`
            : `<span class="client-tracker">ID: ${tracker}</span>`
          }</div>` : ''}
          ${fieldsHtml}
        </div>`;
      listEl.appendChild(row);
    });
  }

  panel.classList.add('open');
}

function closeCountyPanel() {
  document.getElementById('county-panel').classList.remove('open');
  if (activeCounty) {
    const el = countyPaths[activeCounty.fips];
    if (el) el.classList.remove('county-highlighted');
  }
  activeCounty = null;
}

function closeCountyAndZoom() {
  closeCountyPanel();
  // Zoom back to state or region level, or full map
  if (activeState !== 'all' || activeRegion !== 'all') {
    zoomToSelection();
  } else {
    zoomToViewBox(BASE_VIEWBOX);
    document.getElementById('zoom-out-btn').classList.remove('visible');
  }
}

function zoomOut() {
  closeCountyPanel();
  // Reset all filters and zoom to full US
  activeRegion = 'all';
  activeState  = 'all';
  updateRegionFilters();
  updateStateList();
  updateActiveChips();
  updateStats();
  updateContextLegend();
  colorMap();
  document.getElementById('stat-region').textContent = 'All';
  zoomToViewBox(BASE_VIEWBOX);
  document.getElementById('zoom-out-btn').classList.remove('visible');
}

function showTooltip(e, feature) {
  const tt = document.getElementById('tooltip');
  const fips = feature.id;
  const name = feature.properties.name;
  const stateAbbr = fipsToState(fips);

  const key = makeKey(name, stateAbbr);
  const info = clientData[key];

  // Always show region based on FIPS (works even before CSV upload)
  const fipsRegion = getRegionByFips(fips);
  const regionLabel = regionNames[fipsRegion];

  document.getElementById('tt-county').textContent = name + ' County';
  document.getElementById('tt-state').textContent = stateAbbr || '';

  if (info && info.count > 0) {
    document.getElementById('tt-clients').textContent = `${info.count} client${info.count !== 1 ? 's' : ''}`;
    document.getElementById('tt-region').innerHTML = `<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${REGION_COLORS[fipsRegion]};margin-right:4px;vertical-align:middle"></span>${regionLabel}`;
  } else {
    document.getElementById('tt-clients').textContent = 'No clients';
    document.getElementById('tt-region').innerHTML = `<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${REGION_COLORS[fipsRegion]};margin-right:4px;vertical-align:middle"></span>${regionLabel}`;
  }

  const rect = document.querySelector('.map-container').getBoundingClientRect();
  let x = e.clientX - rect.left + 14;
  let y = e.clientY - rect.top - 10;
  if (x + 200 > rect.width) x -= 210;
  tt.style.left = x + 'px';
  tt.style.top  = y + 'px';
  tt.style.display = 'block';
}

function hideTooltip() {
  document.getElementById('tooltip').style.display = 'none';
}

// ─── CSV HANDLING ─────────────────────────────────────────────────
document.getElementById('csv-upload').addEventListener('change', function(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => parseCSV(ev.target.result);
  reader.readAsText(file);
});

function parseCSV(text) {
  const lines = text.trim().split('\n');
  csvHeaders = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g,''));
  csvRows = lines.slice(1).map(line => {
    // Handle quoted commas
    const cols = [];
    let inQ = false, cur = '';
    for (let c of line) {
      if (c === '"') { inQ = !inQ; }
      else if (c === ',' && !inQ) { cols.push(cur.trim()); cur = ''; }
      else cur += c;
    }
    cols.push(cur.trim());
    return cols;
  });
  openModal();
}

// ─── MODAL ────────────────────────────────────────────────────────
function openModal() {
  const modal = document.getElementById('col-modal');
  const selectors = ['map-county','map-state','map-first','map-last','map-tracker','map-region'];
  selectors.forEach(id => {
    const sel = document.getElementById(id);
    const isRequired = id === 'map-county' || id === 'map-state';
    sel.innerHTML = isRequired ? '<option value="">— select —</option>' : '<option value="">— skip —</option>';
    csvHeaders.forEach((h, i) => {
      const opt = document.createElement('option');
      opt.value = i;
      opt.textContent = h;
      const hl = h.toLowerCase().trim();
      if (id === 'map-county'  && hl.includes('county')) opt.selected = true;
      if (id === 'map-state'   && (hl === 'state' || hl === 'billing state' || hl === 'billing state/province' || hl === 'mailing state' || hl === 'mailing state/province')) opt.selected = true;
      if (id === 'map-first'   && (hl === 'first name' || hl === 'firstname' || hl === 'first')) opt.selected = true;
      if (id === 'map-last'    && (hl === 'last name'  || hl === 'lastname'  || hl === 'last' || hl === 'surname')) opt.selected = true;
      if (id === 'map-tracker' && (hl === 'client id' || hl === 'id' || hl === 'cci number' || hl === 'account number' || hl === 'tracker' || hl.includes('client id') || hl.includes('tracker #'))) opt.selected = true;
      if (id === 'map-region'  && hl.includes('region')) opt.selected = true;
      sel.appendChild(opt);
    });
  });

  // Restore saved SF URL
  document.getElementById('sf-base-url').value = sfBaseUrl;

  // Region name inputs
  const container = document.getElementById('region-name-inputs');
  container.innerHTML = '';
  for (let i = 0; i < 6; i++) {
    const row = document.createElement('div');
    row.className = 'region-name-row';
    row.innerHTML = `<div class="dot" style="background:${REGION_COLORS[i]}"></div>
      <input type="text" placeholder="Region ${i+1}" value="${regionNames[i]}" data-idx="${i}">`;
    container.appendChild(row);
  }

  modal.classList.add('open');
}

function closeModal() {
  document.getElementById('col-modal').classList.remove('open');
  document.getElementById('modal-error').style.display = 'none';
}

function confirmMapping() {
  const countyIdx  = document.getElementById('map-county').value;
  const stateIdx   = document.getElementById('map-state').value;

  if (countyIdx === '' || stateIdx === '') {
    document.getElementById('modal-error').style.display = 'block';
    return;
  }

  const firstIdx   = document.getElementById('map-first').value;
  const lastIdx    = document.getElementById('map-last').value;
  const trackerIdx = document.getElementById('map-tracker').value;
  const regionIdx  = document.getElementById('map-region').value;

  sfBaseUrl = (document.getElementById('sf-base-url').value || '').trim().replace(/\/$/, '');

  document.querySelectorAll('#region-name-inputs input').forEach(inp => {
    regionNames[+inp.dataset.idx] = inp.value || `Region ${+inp.dataset.idx + 1}`;
  });

  closeModal();
  processData(
    +countyIdx,
    +stateIdx,
    firstIdx   !== '' ? +firstIdx   : null,
    lastIdx    !== '' ? +lastIdx    : null,
    trackerIdx !== '' ? +trackerIdx : null,
    regionIdx  !== '' ? +regionIdx  : null,
  );
}

// ─── DATA PROCESSING ──────────────────────────────────────────────
// Special column indices tracked globally so openCountyPanel knows which are name/tracker
let colFirstIdx   = null;
let colLastIdx    = null;
let colTrackerIdx = null;
let colCountyIdx  = null;
let colStateIdx   = null;
let intlClients   = []; // clients with non-US state/location

// The 12 ordered display fields for client cards (exact Salesforce header names)
const DISPLAY_FIELDS = [
  { header: 'Account Name',                     icon: '🏢' },
  { header: 'First Name',                       icon: '👤' },
  { header: 'Last Name',                        icon: '👤' },
  { header: 'Current Placement: Dog Name',      label: 'Dog Name',             icon: '🐾' },
  { header: 'Current Placement: Outcome Category', label: 'Outcome Category',    icon: '🏷️' },
  { header: 'Mailing Street',                   icon: '📍' },
  { header: 'Mailing City',                     icon: '📍' },
  { header: 'Mailing State/Province',           icon: '📍' },
  { header: 'Mailing Zip/Postal Code',          icon: '📍' },
  { header: 'County',                           icon: '📍' },
  { header: 'Phone',                            icon: '📞' },
  { header: 'Email',                            icon: '✉️' },
];

function processData(countyCol, stateCol, firstCol, lastCol, trackerCol, regionCol) {
  clientData    = {};
  intlClients   = [];
  colFirstIdx   = firstCol;
  colLastIdx    = lastCol;
  colTrackerIdx = trackerCol;
  colCountyIdx  = countyCol;
  colStateIdx   = stateCol;

  csvRows.forEach(row => {
    const rawState  = (row[stateCol] || '').trim();
    const state     = normalizeState(rawState);
    let county      = (row[countyCol] || '').replace(/\s*county\s*$/i,'').trim();

    const allFields = {};
    csvHeaders.forEach((h, i) => { allFields[h] = (row[i] || '').trim(); });

    const firstName = firstCol   !== null ? (row[firstCol]   || '').trim() : (allFields['First Name'] || '');
    const lastName  = lastCol    !== null ? (row[lastCol]    || '').trim() : (allFields['Last Name']  || '');
    const fullName  = [firstName, lastName].filter(Boolean).join(' ');

    const clientObj = {
      name:    fullName,
      tracker: trackerCol !== null ? (row[trackerCol] || '').trim() : '',
      fields:  allFields,
    };

    const isUS = state && (STATE_NAME_MAP[rawState.toLowerCase()] !== undefined
      || Object.values(FIPS_STATE).includes(state));

    if (!isUS && rawState) {
      intlClients.push(clientObj);
      return;
    }

    if (!county || !state) return;

    let regionIdx;
    if (regionCol !== null && row[regionCol]) {
      const rv = row[regionCol].trim();
      const found = regionNames.findIndex(n => n.toLowerCase() === rv.toLowerCase());
      regionIdx = found >= 0 ? found : getDefaultRegion(state, county);
    } else {
      regionIdx = getDefaultRegion(state, county);
    }

    const key = makeKey(county, state);
    if (!clientData[key]) {
      clientData[key] = { count: 0, regionIdx, county, state, clients: [] };
    }
    clientData[key].count++;
    clientData[key].clients.push(clientObj);
  });

  updateStats();
  updateRegionFilters();
  updateStateList();
  updateActiveChips();
  updateContextLegend();
  colorMap();
  updateFieldDisplay(countyCol, stateCol, firstCol, lastCol, trackerCol);
  const intlCountEl = document.getElementById('intl-count');
  if (intlCountEl) intlCountEl.textContent = intlClients.length;
}


// ─── MAP COLORING ─────────────────────────────────────────────────
function colorMap() {
  Object.entries(countyPaths).forEach(([fips, el]) => {
    const regionIdx  = getRegionByFips(fips);
    const stateAbbr  = fipsToState(fips);
    const regionMatch = activeRegion === 'all' || regionIdx === +activeRegion;
    const stateMatch  = activeState  === 'all' || stateAbbr === activeState;

    if (!regionMatch || !stateMatch) {
      // Dimmed / out of filter
      el.setAttribute('fill', '#f5f7fa');
      el.setAttribute('stroke', '#e0e8f0');
      el.setAttribute('stroke-width', '0.3');
      el.style.opacity = '0.3';
    } else {
      // In-filter: white fill, color-coded region outline
      el.setAttribute('fill', '#ffffff');
      el.setAttribute('stroke', REGION_COLORS[regionIdx]);
      el.setAttribute('stroke-width', '0.6');
      el.style.opacity = '1';
    }
  });

  // Overlay client heat fill on matching counties
  Object.entries(clientData).forEach(([key, info]) => {
    const regionMatch = activeRegion === 'all' || info.regionIdx === +activeRegion;
    const stateMatch  = activeState  === 'all' || info.state === activeState;
    if (!regionMatch || !stateMatch) return;

    const fips = findFips(info.county, info.state);
    if (!fips) return;
    const el = countyPaths[fips];
    if (!el) return;

    const { fill, stroke } = getShade(REGION_COLORS[info.regionIdx], info.count);
    el.setAttribute('fill', fill);
    el.setAttribute('stroke', stroke);
    el.setAttribute('stroke-width', '0.6');
    el.style.opacity = '1';
  });
}

// ─── DENSITY SCALE (dynamic) ─────────────────────────────────────
let densityMax   = 50;
let densityBands = 10;

const DENSITY_PALETTE = [
  { fill: '#e8f4ff', stroke: '#b8d9f5' },
  { fill: '#c0dff7', stroke: '#7ab8f5' },
  { fill: '#7ab8f5', stroke: '#4a9ae0' },
  { fill: '#3d9be0', stroke: '#1f7dc4' },
  { fill: '#0073FF', stroke: '#0055cc' },
  { fill: '#0058cc', stroke: '#003d99' },
  { fill: '#003d99', stroke: '#002a6b' },
  { fill: '#FECB00', stroke: '#c9a000' },
  { fill: '#f59e0b', stroke: '#c47d08' },
  { fill: '#e07b00', stroke: '#b55f00' },
];

const NO_CLIENT_STOP = { fill: '#ffffff', stroke: '#d0dcea' };

function getShade(hexColor, count) {
  if (count === 0) return NO_CLIENT_STOP;
  const clamped   = Math.min(count, densityMax);
  const bandSize  = densityMax / densityBands;
  const bandIdx   = Math.min(Math.floor((clamped - 1) / bandSize), densityBands - 1);
  const paletteIdx = Math.round((bandIdx / Math.max(densityBands - 1, 1)) * (DENSITY_PALETTE.length - 1));
  return DENSITY_PALETTE[paletteIdx];
}

function onDensityChange() {
  densityMax   = +document.getElementById('density-max-slider').value;
  densityBands = +document.getElementById('density-bands-slider').value;
  document.getElementById('density-max-val').textContent   = densityMax;
  document.getElementById('density-bands-val').textContent = densityBands;
  buildDensityLegend();
  colorMap();
}

function buildDensityLegend() {
  const container = document.getElementById('context-legend');
  if (!container) return;
  if (activeRegion !== 'all' || activeState !== 'all') { updateContextLegend(); return; }
  const bandSize = densityMax / densityBands;
  let html = '<div class="legend-scale">';
  html += '<div class="legend-row"><div class="legend-swatch" style="background:#ffffff;border:1px solid #d0dcea"></div> No records</div>';
  for (let i = 0; i < densityBands; i++) {
    const lo  = Math.round(i * bandSize) + 1;
    const hi  = Math.round((i + 1) * bandSize);
    const pIdx = Math.round((i / Math.max(densityBands - 1, 1)) * (DENSITY_PALETTE.length - 1));
    const { fill } = DENSITY_PALETTE[pIdx];
    const label = i === densityBands - 1 ? lo + '+' : lo + '\u2013' + hi;
    html += '<div class="legend-row"><div class="legend-swatch" style="background:' + fill + '"></div> ' + label + '</div>';
  }
  html += '</div>';
  container.innerHTML = html;
}

// ─── TAB SWITCHING ────────────────────────────────────────────────
function switchTab(tab) {
  document.querySelectorAll('.filter-tab').forEach(t => t.classList.toggle('active', t.dataset.tab === tab));
  document.querySelectorAll('.filter-panel').forEach(p => p.classList.toggle('active', p.id === 'panel-' + tab));
}

// ─── REGION FILTERS ───────────────────────────────────────────────
function updateRegionFilters() {
  const container = document.getElementById('region-filters');
  container.innerHTML = '';

  const totalClients = Object.values(clientData).reduce((s,v) => s + v.count, 0);
  const allBtn = document.createElement('button');
  allBtn.className = 'region-btn' + (activeRegion === 'all' ? ' active' : '');
  allBtn.dataset.region = 'all';
  allBtn.innerHTML = `<span class="region-dot" style="background:#aab8c5"></span>All Regions<span class="filter-count">${totalClients}</span>`;
  allBtn.onclick = () => setRegion('all');
  container.appendChild(allBtn);

  for (let i = 0; i < 6; i++) {
    const count = Object.values(clientData).filter(d => d.regionIdx === i).reduce((s,v) => s+v.count, 0);
    const btn = document.createElement('button');
    const color = REGION_COLORS[i];
    const isActive = activeRegion === String(i);
    btn.className = 'region-btn' + (isActive ? ' active' : '');
    if (isActive) {
      btn.style.setProperty('--active-color', color);
      btn.style.setProperty('--active-bg', hexToRgba(color, 0.1));
      btn.style.color = color;
      btn.style.borderColor = color;
      btn.style.background = hexToRgba(color, 0.1);
    }
    btn.dataset.region = i;
    btn.innerHTML = `<span class="region-dot" style="background:${color}"></span>${regionNames[i]}<span class="filter-count">${count}</span>`;
    btn.onclick = () => setRegion(String(i));
    container.appendChild(btn);
  }
}

function setRegion(r) {
  activeRegion = r;
  activeState = 'all';
  document.getElementById('stat-region').textContent = r === 'all' ? 'All' : regionNames[+r];
  updateRegionFilters();
  updateStateList();
  updateActiveChips();
  updateStats();
  updateContextLegend();
  colorMap();
  if (r === 'all') {
    zoomOut();
  } else {
    zoomToSelection();
  }
}

// ─── STATE FILTERS ────────────────────────────────────────────────
// Full US state list with abbreviations
const ALL_STATES = [
  {abbr:'AL',name:'Alabama'},{abbr:'AK',name:'Alaska'},{abbr:'AZ',name:'Arizona'},
  {abbr:'AR',name:'Arkansas'},{abbr:'CA',name:'California'},{abbr:'CO',name:'Colorado'},
  {abbr:'CT',name:'Connecticut'},{abbr:'DE',name:'Delaware'},{abbr:'FL',name:'Florida'},
  {abbr:'GA',name:'Georgia'},{abbr:'HI',name:'Hawaii'},{abbr:'ID',name:'Idaho'},
  {abbr:'IL',name:'Illinois'},{abbr:'IN',name:'Indiana'},{abbr:'IA',name:'Iowa'},
  {abbr:'KS',name:'Kansas'},{abbr:'KY',name:'Kentucky'},{abbr:'LA',name:'Louisiana'},
  {abbr:'ME',name:'Maine'},{abbr:'MD',name:'Maryland'},{abbr:'MA',name:'Massachusetts'},
  {abbr:'MI',name:'Michigan'},{abbr:'MN',name:'Minnesota'},{abbr:'MS',name:'Mississippi'},
  {abbr:'MO',name:'Missouri'},{abbr:'MT',name:'Montana'},{abbr:'NE',name:'Nebraska'},
  {abbr:'NV',name:'Nevada'},{abbr:'NH',name:'New Hampshire'},{abbr:'NJ',name:'New Jersey'},
  {abbr:'NM',name:'New Mexico'},{abbr:'NY',name:'New York'},{abbr:'NC',name:'North Carolina'},
  {abbr:'ND',name:'North Dakota'},{abbr:'OH',name:'Ohio'},{abbr:'OK',name:'Oklahoma'},
  {abbr:'OR',name:'Oregon'},{abbr:'PA',name:'Pennsylvania'},{abbr:'RI',name:'Rhode Island'},
  {abbr:'SC',name:'South Carolina'},{abbr:'SD',name:'South Dakota'},{abbr:'TN',name:'Tennessee'},
  {abbr:'TX',name:'Texas'},{abbr:'UT',name:'Utah'},{abbr:'VT',name:'Vermont'},
  {abbr:'VA',name:'Virginia'},{abbr:'WA',name:'Washington'},{abbr:'WV',name:'West Virginia'},
  {abbr:'WI',name:'Wisconsin'},{abbr:'WY',name:'Wyoming'},{abbr:'DC',name:'Washington DC'}
];

function getStateRegionIdx(abbr) {
  return STATE_REGION_DEFAULT[abbr] ?? 0;
}

function updateStateList(query) {
  const container = document.getElementById('state-list');
  container.innerHTML = '';
  const q = (query || '').toLowerCase();

  // Group states by region
  const groups = Array.from({length: 6}, () => []);
  ALL_STATES.forEach(s => {
    if (q && !s.name.toLowerCase().includes(q) && !s.abbr.toLowerCase().includes(q)) return;
    const rIdx = getStateRegionIdx(s.abbr);
    groups[rIdx].push(s);
  });

  // Only show regions that are visible given activeRegion filter
  groups.forEach((states, rIdx) => {
    if (!states.length) return;
    if (activeRegion !== 'all' && +activeRegion !== rIdx) return;

    const group = document.createElement('div');
    group.className = 'state-region-group';

    const label = document.createElement('div');
    label.className = 'state-region-label';
    label.style.color = REGION_COLORS[rIdx];
    label.innerHTML = `<span class="dot" style="background:${REGION_COLORS[rIdx]}"></span>${regionNames[rIdx]}`;
    group.appendChild(label);

    states.forEach(s => {
      // Count clients in this state
      const count = Object.values(clientData)
        .filter(d => d.state === s.abbr)
        .reduce((sum, d) => sum + d.count, 0);

      const btn = document.createElement('button');
      const isActive = activeState === s.abbr;
      btn.className = 'state-btn' + (isActive ? ' active' : '');
      btn.innerHTML = `<span class="state-abbr">${s.abbr}</span>${s.name}<span class="filter-count">${count > 0 ? count : ''}</span>`;
      btn.onclick = () => setState(s.abbr, s.name);
      group.appendChild(btn);
    });

    container.appendChild(group);
  });
}

function filterStateList(q) {
  updateStateList(q);
}

function setState(abbr, name) {
  if (activeState === abbr) {
    activeState = 'all';
    // Zoom back to region level if one is still active
    if (activeRegion !== 'all') {
      zoomToSelection();
    } else {
      zoomOut();
    }
  } else {
    activeState = abbr;
    const rIdx = getStateRegionIdx(abbr);
    activeRegion = String(rIdx);
    zoomToSelection();
  }
  document.getElementById('stat-region').textContent =
    activeState !== 'all' ? name : (activeRegion === 'all' ? 'All' : regionNames[+activeRegion]);
  updateRegionFilters();
  updateStateList(document.getElementById('state-search').value);
  updateActiveChips();
  updateStats();
  updateContextLegend();
  colorMap();
}

// ─── ACTIVE CHIPS ─────────────────────────────────────────────────
function updateActiveChips() {
  const container = document.getElementById('active-filters');
  container.innerHTML = '';
  if (activeRegion !== 'all') {
    const chip = document.createElement('div');
    chip.className = 'filter-chip';
    chip.innerHTML = `<span class="chip-dot" style="background:${REGION_COLORS[+activeRegion]}"></span>${regionNames[+activeRegion]}<span class="chip-x">✕</span>`;
    chip.onclick = () => setRegion('all');
    container.appendChild(chip);
  }
  if (activeState !== 'all') {
    const s = ALL_STATES.find(s => s.abbr === activeState);
    const chip = document.createElement('div');
    chip.className = 'filter-chip';
    chip.innerHTML = `${s ? s.name : activeState}<span class="chip-x">✕</span>`;
    chip.onclick = () => setState(activeState, '');
    container.appendChild(chip);
  }
}

// ─── INTERNATIONAL CLIENTS PANEL ─────────────────────────────────
function openIntlPanel() {
  document.getElementById('intl-search').value = '';
  renderIntlClients('');
  document.getElementById('intl-modal').classList.add('open');
}

function closeIntlPanel() {
  document.getElementById('intl-modal').classList.remove('open');
}

function filterIntlList(q) {
  renderIntlClients(q);
}

function renderIntlClients(query) {
  const container = document.getElementById('intl-client-list');
  const sub       = document.getElementById('intl-modal-sub');
  const q         = (query || '').toLowerCase();

  const filtered = intlClients.filter(c => {
    if (!q) return true;
    const f = c.fields || {};
    return Object.values(f).some(v => v.toLowerCase().includes(q))
      || (c.name || '').toLowerCase().includes(q);
  });

  sub.textContent = `${filtered.length} client${filtered.length !== 1 ? 's' : ''} outside the United States`;

  if (!filtered.length) {
    container.innerHTML = `<div class="intl-empty">
      ${intlClients.length === 0
        ? '📭 No international clients found in your CSV.'
        : '🔍 No clients match your search.'}
    </div>`;
    return;
  }

  container.innerHTML = '';

  // Group by country (Mailing State/Province used as country proxy for intl)
  const byCountry = {};
  filtered.forEach(c => {
    const country = c.fields?.['Mailing State/Province'] || c.fields?.['Mailing Country'] || 'Unknown Country';
    if (!byCountry[country]) byCountry[country] = [];
    byCountry[country].push(c);
  });

  Object.entries(byCountry).sort(([a],[b]) => a.localeCompare(b)).forEach(([country, clients]) => {
    // Country group header
    const groupHeader = document.createElement('div');
    groupHeader.style.cssText = 'padding:8px 20px 4px;font-size:0.62rem;font-weight:700;text-transform:uppercase;letter-spacing:0.1em;color:var(--muted);background:#f8fafc;border-bottom:1px solid var(--border)';
    groupHeader.textContent = country;
    container.appendChild(groupHeader);

    clients.forEach(client => {
      const f         = client.fields || {};
      const firstName = f['First Name'] || '';
      const lastName  = f['Last Name']  || '';
      const fullName  = client.name || [firstName, lastName].filter(Boolean).join(' ') || '(No name)';
      const tracker   = client.tracker || '';
      const sfLink    = tracker && sfBaseUrl ? `${sfBaseUrl}/lightning/r/Account/${tracker}/view` : '';

      let fieldsHtml = '';
      DISPLAY_FIELDS.forEach(({ header, label, icon }) => {
        const val = f[header];
        if (!val) return;
        const displayLabel = label || header;
        fieldsHtml += `<div class="intl-field">
          <span class="intl-field-label">${icon} ${displayLabel}:</span>
          <span class="intl-field-value">${
            header === 'Email'
              ? `<a href="mailto:${val}" style="color:var(--cc-blue);text-decoration:none">${val}</a>`
              : header === 'Phone'
              ? `<a href="tel:${val}" style="color:var(--cc-blue);text-decoration:none">${val}</a>`
              : val
          }</span>
        </div>`;
      });

      const row = document.createElement('div');
      row.className = 'intl-client-row';
      row.innerHTML = `
        <div class="intl-client-name">
          ${fullName}
          ${tracker ? `<span class="intl-country-badge">${sfLink
            ? `<a href="${sfLink}" target="_blank" style="color:inherit;text-decoration:none">ID: ${tracker} ↗</a>`
            : `ID: ${tracker}`}</span>` : ''}
        </div>
        <div class="intl-fields">${fieldsHtml}</div>`;
      container.appendChild(row);
    });
  });
}


function updateContextLegend() {
  const title = document.getElementById('context-legend-title');
  const container = document.getElementById('context-legend');

  // Case 1: A state is selected → show counties breakdown
  if (activeState !== 'all') {
    const stateName = (ALL_STATES.find(s => s.abbr === activeState) || {}).name || activeState;
    const rIdx = getStateRegionIdx(activeState);
    const color = REGION_COLORS[rIdx];
    title.textContent = `${stateName} — Counties`;

    // Gather counties in this state, restricted to the active region
    const counties = Object.values(clientData)
      .filter(d => d.state === activeState && (activeRegion === 'all' || d.regionIdx === +activeRegion))
      .sort((a,b) => b.count - a.count);

    if (!counties.length) {
      container.innerHTML = '<div class="ctx-empty">No clients in this state yet.</div>';
      return;
    }

    const maxCount = counties[0].count;
    const list = document.createElement('div');
    list.className = 'context-list';

    counties.forEach(d => {
      const row = document.createElement('div');
      row.className = 'context-row';
      const pct = maxCount > 0 ? Math.round((d.count / maxCount) * 100) : 0;
      row.innerHTML = `
        <div class="ctx-label">${d.county}</div>
        <div class="ctx-bar-wrap"><div class="ctx-bar" style="width:${pct}%;background:${color}"></div></div>
        <div class="ctx-count ${d.count === 0 ? 'zero' : ''}">${d.count}</div>`;
      list.appendChild(row);
    });

    container.innerHTML = '';
    container.appendChild(list);
    return;
  }

  // Case 2: A region is selected → show states breakdown
  if (activeRegion !== 'all') {
    const rIdx = +activeRegion;
    const color = REGION_COLORS[rIdx];
    title.textContent = `${regionNames[rIdx]} — States`;

    // All states in this region, sorted by client count desc
    const regionStates = ALL_STATES
      .filter(s => getStateRegionIdx(s.abbr) === rIdx)
      .map(s => {
        const count = Object.values(clientData)
          .filter(d => d.state === s.abbr)
          .reduce((sum, d) => sum + d.count, 0);
        return { abbr: s.abbr, name: s.name, count };
      })
      .sort((a,b) => b.count - a.count);

    if (!regionStates.length) {
      container.innerHTML = '<div class="ctx-empty">No states in this region.</div>';
      return;
    }

    const maxCount = regionStates[0].count || 1;
    const list = document.createElement('div');
    list.className = 'context-list';

    regionStates.forEach(s => {
      const row = document.createElement('div');
      row.className = 'context-row';
      const pct = Math.round((s.count / maxCount) * 100);
      row.innerHTML = `
        <div class="ctx-label" title="${s.name}"><strong style="color:var(--muted);font-size:0.62rem;margin-right:4px">${s.abbr}</strong>${s.name}</div>
        <div class="ctx-bar-wrap"><div class="ctx-bar" style="width:${pct}%;background:${color}"></div></div>
        <div class="ctx-count ${s.count === 0 ? 'zero' : ''}">${s.count || '—'}</div>`;
      // Clicking a state row in the legend selects that state
      row.style.cursor = 'pointer';
      row.onclick = () => { setState(s.abbr, s.name); switchTab('state'); };
      list.appendChild(row);
    });

    container.innerHTML = '';
    container.appendChild(list);
    return;
  }

  // Case 3: No filter — show dynamic density scale
  title.textContent = 'Density Scale';
  buildDensityLegend();
}

function hexToRgba(hex, alpha) {
  const r = parseInt(hex.slice(1,3),16);
  const g = parseInt(hex.slice(3,5),16);
  const b = parseInt(hex.slice(5,7),16);
  return `rgba(${r},${g},${b},${alpha})`;
}

// ─── STATS ────────────────────────────────────────────────────────
function updateStats() {
  let filtered = Object.values(clientData);
  if (activeRegion !== 'all') filtered = filtered.filter(d => d.regionIdx === +activeRegion);
  if (activeState  !== 'all') filtered = filtered.filter(d => d.state === activeState);
  const total = filtered.reduce((s,v) => s + v.count, 0);
  const counties = filtered.filter(v => v.count > 0).length;
  document.getElementById('stat-total').textContent = total;
  document.getElementById('stat-counties').textContent = counties;
}

function updateFieldDisplay(cCol, sCol, fCol, lCol, tCol) {
  const container = document.getElementById('field-map-display');
  container.innerHTML = '';
  const fields = [
    { label: `County → ${csvHeaders[cCol]}`,  mapped: true },
    { label: `State → ${csvHeaders[sCol]}`,   mapped: true },
    { label: fCol !== null ? `First Name → ${csvHeaders[fCol]}` : 'First Name → (skipped)', mapped: fCol !== null },
    { label: lCol !== null ? `Last Name → ${csvHeaders[lCol]}`  : 'Last Name → (skipped)',  mapped: lCol !== null },
    { label: tCol !== null ? `Client ID → ${csvHeaders[tCol]}`  : 'Client ID → (skipped)',  mapped: tCol !== null },
    { label: `All other fields (${csvHeaders.length - 2} cols) auto-captured`, mapped: true },
  ];
  fields.forEach(f => {
    const chip = document.createElement('span');
    chip.className = 'field-chip' + (f.mapped ? ' mapped' : '');
    chip.textContent = f.label;
    container.appendChild(chip);
  });
}

// ─── HELPERS ──────────────────────────────────────────────────────
function makeKey(county, state) {
  return county.toLowerCase().trim() + '|' + state.toUpperCase().trim();
}


// Simplified FIPS → state abbr (from first 2 digits)
const FIPS_STATE = {"01":"AL","02":"AK","04":"AZ","05":"AR","06":"CA","08":"CO","09":"CT","10":"DE","11":"DC","12":"FL","13":"GA","15":"HI","16":"ID","17":"IL","18":"IN","19":"IA","20":"KS","21":"KY","22":"LA","23":"ME","24":"MD","25":"MA","26":"MI","27":"MN","28":"MS","29":"MO","30":"MT","31":"NE","32":"NV","33":"NH","34":"NJ","35":"NM","36":"NY","37":"NC","38":"ND","39":"OH","40":"OK","41":"OR","42":"PA","44":"RI","45":"SC","46":"SD","47":"TN","48":"TX","49":"UT","50":"VT","51":"VA","53":"WA","54":"WV","55":"WI","56":"WY"};

function fipsToState(fips) {
  const prefix = String(fips).padStart(5,'0').slice(0,2);
  return FIPS_STATE[prefix] || '';
}

// CA and NV are split: Southern CA counties → Southwest (4), rest → Northwest (5)
const SOUTHERN_CA_FIPS = new Set([
  '06025','06029','06037','06059','06065','06071','06073','06079','06083','06111'
]);
const SOUTHERN_NV_FIPS = new Set(['32003']); // Clark County (Las Vegas)

function getRegionByFips(fipsStr) {
  const fips = String(fipsStr).padStart(5,'0');
  const stateFips = fips.slice(0,2);
  if (stateFips === '06') return SOUTHERN_CA_FIPS.has(fips) ? 4 : 5;
  if (stateFips === '32') return SOUTHERN_NV_FIPS.has(fips) ? 4 : 5;
  const stateAbbr = FIPS_STATE[stateFips];
  if (!stateAbbr) return 0;
  return STATE_REGION_DEFAULT[stateAbbr] ?? 0;
}

function getDefaultRegion(stateAbbr, county) {
  if (stateAbbr === 'CA' || stateAbbr === 'NV') {
    if (Object.keys(reverseLookup).length === 0) buildReverseLookup();
    const fips = reverseLookup[makeKey(county || '', stateAbbr)];
    if (fips) return getRegionByFips(fips);
    return 5; // default to Northwest
  }
  return STATE_REGION_DEFAULT[stateAbbr] ?? 0;
}

function getBaseRegionColor(regionIdx) {
  const hex = REGION_COLORS[regionIdx];
  const r = parseInt(hex.slice(1,3),16);
  const g = parseInt(hex.slice(3,5),16);
  const b = parseInt(hex.slice(5,7),16);
  const bg = [232, 240, 248];
  const t = 0.13;
  return `rgb(${Math.round(bg[0]+(r-bg[0])*t)},${Math.round(bg[1]+(g-bg[1])*t)},${Math.round(bg[2]+(b-bg[2])*t)})`;
}

// Build reverse lookup: countyName|stateAbbr → fips
let reverseLookup = {};
function buildReverseLookup() {
  reverseLookup = {};
  Object.entries(countyPaths).forEach(([fips, el]) => {
    const name  = (el.dataset.name || '').trim();
    const state = fipsToState(fips);
    reverseLookup[makeKey(name, state)] = fips;
  });
}

function findFips(county, state) {
  if (Object.keys(reverseLookup).length === 0) buildReverseLookup();
  // Try exact match first
  const exact = reverseLookup[makeKey(county, state)];
  if (exact) return exact;
  // Try stripping trailing "county", "parish", "borough", "census area" etc.
  const stripped = county.replace(/\s*(county|parish|borough|census area|municipality|city and county)\s*$/i, '').trim();
  if (stripped !== county) {
    const fallback = reverseLookup[makeKey(stripped, state)];
    if (fallback) return fallback;
  }
  return null;
}

// State name → abbr normalization
const STATE_NAME_MAP = {"alabama":"AL","alaska":"AK","arizona":"AZ","arkansas":"AR","california":"CA","colorado":"CO","connecticut":"CT","delaware":"DE","district of columbia":"DC","florida":"FL","georgia":"GA","hawaii":"HI","idaho":"ID","illinois":"IL","indiana":"IN","iowa":"IA","kansas":"KS","kentucky":"KY","louisiana":"LA","maine":"ME","maryland":"MD","massachusetts":"MA","michigan":"MI","minnesota":"MN","mississippi":"MS","missouri":"MO","montana":"MT","nebraska":"NE","nevada":"NV","new hampshire":"NH","new jersey":"NJ","new mexico":"NM","new york":"NY","north carolina":"NC","north dakota":"ND","ohio":"OH","oklahoma":"OK","oregon":"OR","pennsylvania":"PA","rhode island":"RI","south carolina":"SC","south dakota":"SD","tennessee":"TN","texas":"TX","utah":"UT","vermont":"VT","virginia":"VA","washington":"WA","west virginia":"WV","wisconsin":"WI","wyoming":"WY"};

function normalizeState(s) {
  if (!s) return '';
  if (s.length === 2) return s.toUpperCase();
  return STATE_NAME_MAP[s.toLowerCase()] || s.toUpperCase().slice(0,2);
}

// ─── CANINE COMPANIONS ORG LOCATIONS ─────────────────────────────
// Lat/lng sourced from canine.org — projected at render time via D3 geoAlbersUsa
// so markers land precisely on the map.

const CC_LOCATIONS = [
  // ─── TRAINING CENTERS ────────────────────────────────────────
  { type:'tc', name:'Northwest Training Center & HQ',  region:'Northwest',     address:'2965 Dutton Ave, Santa Rosa, CA 95407',           areas:'NorCal, NV (north), OR, WA, ID, MT, WY, AK',   lat:38.4188, lng:-122.7186 },
  { type:'tc', name:'Southwest Training Center',        region:'Southwest',     address:'124 Rancho del Oro Dr, Oceanside, CA 92057',      areas:'AZ, UT, CO, NM, SoCal, SoNV, HI',              lat:33.1959, lng:-117.3228 },
  { type:'tc', name:'Northeast Training Center',        region:'Northeast',     address:'286 Middle Island Rd, Medford, NY 11763',         areas:'NY, NJ, CT, DE, PA, MD, DC, VA, WV, MA, RI, VT, NH, ME', lat:40.8204, lng:-72.9968 },
  { type:'tc', name:'North Central Training Center',        region:'North Central', address:'7480 New Albany-Condit Rd, New Albany, OH 43054', areas:'OH, KY, MI, IN, IL, WI, MO, IA, MN, KS, NE, ND, SD, WPN', lat:40.0909, lng:-82.8068 },
  { type:'tc', name:'Southeast Training Center',        region:'Southeast',     address:'8150 Clarcona Ocoee Rd, Orlando, FL 32818',       areas:'FL, GA, TN, NC, SC, MS, AL',                    lat:28.5794, lng:-81.4874 },
  { type:'tc', name:'South Central Training Center',        region:'South Central', address:'7710 Las Colinas Ridge, Irving, TX 75063',        areas:'TX, OK, AR, LA',                                lat:32.8998, lng:-97.0390 },

  // ─── FIELD OFFICES ───────────────────────────────────────────
  { type:'fo', name:'Puget Sound Field Office',  region:'Northwest',     address:'2454 Occidental Ave S, Seattle, WA 98134',        areas:'Puget Sound area, Washington State',            lat:47.5793, lng:-122.3336 },

  // ─── NORTHWEST CHAPTERS ──────────────────────────────────────
  { type:'ch', name:'Wine Country Chapter',             region:'Northwest', areas:'Sonoma, Napa, Mendocino, Lake & Marin Counties, CA',         lat:38.4988, lng:-122.7042 },
  { type:'ch', name:'East Bay Chapter',                 region:'Northwest', areas:'San Francisco East Bay (Alameda & Contra Costa Counties, CA)', lat:37.8716, lng:-122.2727 },
  { type:'ch', name:'South Bay Chapter',                region:'Northwest', areas:'San Francisco & South Bay (Santa Clara County, CA)',           lat:37.3382, lng:-121.8863 },
  { type:'ch', name:'Gold Rush Chapter',                region:'Northwest', areas:'Greater Sacramento, CA',                                       lat:38.5816, lng:-121.4944 },
  { type:'ch', name:'Sierra Foothills Chapter',         region:'Northwest', areas:'Yuba County, CA',                                              lat:39.1388, lng:-121.6175 },
  { type:'ch', name:'Northern Nevada Comstock Chapter', region:'Northwest', areas:'Northern Nevada (Reno/Sparks area)',                            lat:39.5296, lng:-119.8138 },
  { type:'ch', name:'Cascade Chapter',                  region:'Northwest', areas:'Northern Oregon and Southern Washington',                       lat:45.5051, lng:-122.6750 },
  { type:'ch', name:'Puget Sound Chapter',              region:'Northwest', areas:'Northwest Washington (Seattle metro)',                           lat:47.6062, lng:-122.3321 },
  { type:'ch', name:'Inland Northwest Chapter',         region:'Northwest', areas:'North Idaho, Eastern Washington & Western Montana',             lat:47.6588, lng:-117.4260 },
  { type:'ch', name:'Treasure Valley Chapter',          region:'Northwest', areas:'Southwest Idaho (Boise metro)',                                  lat:43.6150, lng:-116.2023 },
  { type:'ch', name:'Big Sky Chapter',                  region:'Northwest', areas:'Northern, Southern & Eastern Montana',                          lat:46.8797, lng:-110.3626 },

  // ─── SOUTHWEST CHAPTERS ──────────────────────────────────────
  { type:'ch', name:'California Dreamin Chapter',       region:'Southwest', areas:'Orange County & Los Angeles County, CA',                        lat:33.8366, lng:-117.9143 },
  { type:'ch', name:'Greater Santa Barbara Chapter',    region:'Southwest', areas:'Ventura County, Santa Barbara County & San Luis Obispo County, CA', lat:34.4208, lng:-119.6982 },
  { type:'ch', name:'Grand Canyon State Chapter',       region:'Southwest', areas:'Arizona (Phoenix metro and statewide)',                          lat:33.4484, lng:-112.0740 },
  { type:'ch', name:'Rocky Mountain Chapter',           region:'Southwest', areas:'Denver and surrounding areas, CO',                               lat:39.7392, lng:-104.9903 },
  { type:'ch', name:'Hawaii Chapter',                   region:'Southwest', areas:'Hawaii (statewide)',                                             lat:21.3069, lng:-157.8583 },
  { type:'ch', name:'Las Vegas Chapter',                region:'Southwest', areas:'Las Vegas metro, Southern Nevada',                               lat:36.1699, lng:-115.1398 },

  // ─── NORTH CENTRAL CHAPTERS ──────────────────────────────────
  { type:'ch', name:'Greater Chicagoland Chapter', region:'North Central', areas:'Chicago, IL metro area',                                          lat:41.8781, lng:-87.6298 },
  { type:'ch', name:'Indiana Chapter',             region:'North Central', areas:'Statewide Indiana (hub in Indianapolis)',                          lat:39.7684, lng:-86.1581 },
  { type:'ch', name:'Great Lakes Chapter',         region:'North Central', areas:'Michigan (statewide)',                                             lat:44.1825, lng:-84.5068 },
  { type:'ch', name:'Badger State Chapter',        region:'North Central', areas:'Wisconsin (statewide) — founded 2025',                             lat:44.2685, lng:-89.6165 },
  { type:'ch', name:'Heartland Chapter',           region:'North Central', areas:'Kansas City metro, Kansas & Missouri',                             lat:39.0997, lng:-94.5786 },
  { type:'ch', name:'Minnesota Chapter',           region:'North Central', areas:'Twin Cities metro and greater Minnesota',                          lat:44.9778, lng:-93.2650 },
  { type:'ch', name:'Nebraska Chapter',            region:'North Central', areas:'Omaha and statewide Nebraska',                                     lat:41.2565, lng:-95.9345 },
  { type:'ch', name:'Northern Ohio Chapter',       region:'North Central', areas:'Cleveland, OH and Northern Ohio',                                  lat:41.4993, lng:-81.6944 },
  { type:'ch', name:'Buckeye Chapter',             region:'North Central', areas:'Columbus, OH and surrounding areas',                               lat:39.9612, lng:-82.9988 },
  { type:'ch', name:'Cin-Day Chapter',             region:'North Central', areas:'Greater Cincinnati & Dayton, SW Ohio',                             lat:39.1031, lng:-84.5120 },
  { type:'ch', name:'Western Pennsylvania Chapter',region:'North Central', areas:'Pittsburgh, PA metro area',                                        lat:40.4406, lng:-79.9959 },
  { type:'ch', name:'Kentucky Chapter',            region:'North Central', areas:'Statewide Kentucky (hub in Louisville); southern Indiana', lat:38.2527, lng:-85.7585 },

  // ─── NORTHEAST CHAPTERS ──────────────────────────────────────
  { type:'ch', name:'Long Island Chapter',          region:'Northeast', areas:'Long Island, NY',                                                     lat:40.7891, lng:-73.1350 },
  { type:'ch', name:'NYC Chapter',                  region:'Northeast', areas:'New York City, NY',                                                   lat:40.7128, lng:-74.0060 },
  { type:'ch', name:'Hudson Valley Chapter',        region:'Northeast', areas:'Hudson Valley, New York',                                             lat:41.7004, lng:-73.9209 },
  { type:'ch', name:'Upstate New York Chapter',     region:'Northeast', areas:'Albany and the New York Capital District',                            lat:42.6526, lng:-73.7562 },
  { type:'ch', name:'New Jersey Chapter',           region:'Northeast', areas:'New Jersey (statewide)',                                              lat:40.0583, lng:-74.4057 },
  { type:'ch', name:'Philadelphia Chapter',         region:'Northeast', areas:'Southeast Pennsylvania and Southern New Jersey',                      lat:39.9526, lng:-75.1652 },
  { type:'ch', name:'Lehigh Valley Chapter',        region:'Northeast', areas:'Lehigh Valley, Pennsylvania (Allentown/Bethlehem area)',               lat:40.6084, lng:-75.4902 },
  { type:'ch', name:'Capital Chapter',              region:'Northeast', areas:'Washington D.C., Northern Virginia & West Virginia',                  lat:38.9072, lng:-77.0369 },
  { type:'ch', name:'Chesapeake Chapter',           region:'Northeast', areas:'Maryland and Delaware',                                               lat:39.0458, lng:-76.6413 },
  { type:'ch', name:'Old Dominion Chapter',         region:'Northeast', areas:'Richmond, VA area',                                                   lat:37.5407, lng:-77.4360 },
  { type:'ch', name:'Bay State Chapter',            region:'Northeast', areas:'Massachusetts (statewide)',                                            lat:42.4072, lng:-71.3824 },
  { type:'ch', name:'Northern New England Chapter', region:'Northeast', areas:'New Hampshire, Maine & Vermont',                                      lat:43.9654, lng:-71.4179 },

  // ─── SOUTHEAST CHAPTERS ──────────────────────────────────────
  { type:'ch', name:'Central Florida Chapter',    region:'Southeast', areas:'Orlando metro, FL',                                                     lat:28.5383, lng:-81.3792 },
  { type:'ch', name:'South Florida Chapter',      region:'Southeast', areas:'Miami/Ft. Lauderdale, FL',                                              lat:25.7617, lng:-80.1918 },
  { type:'ch', name:'Florida First Coast Chapter',region:'Southeast', areas:'Jacksonville, FL area',                                                 lat:30.3322, lng:-81.6557 },
  { type:'ch', name:'Florida East Coast Chapter', region:'Southeast', areas:'Melbourne/Vero Beach area, FL',                                         lat:28.0836, lng:-80.6081 },
  { type:'ch', name:'Emerald Coast Chapter',      region:'Southeast', areas:'Pensacola/Ft. Walton Beach, FL (NW Florida panhandle)',                  lat:30.4213, lng:-87.2169 },
  { type:'ch', name:'Tampa Bay Chapter',          region:'Southeast', areas:'Tampa/St. Petersburg metro, FL',                                        lat:27.9506, lng:-82.4572 },
  { type:'ch', name:'Southwest Florida Chapter',  region:'Southeast', areas:'Southwest Florida (Naples/Ft. Myers area)',                             lat:26.1420, lng:-81.7948 },
  { type:'ch', name:'Greater Atlanta Chapter',    region:'Southeast', areas:'Atlanta and surrounding areas, GA',                                     lat:33.7490, lng:-84.3880 },
  { type:'ch', name:'Nashville Chapter',          region:'Southeast', areas:'Nashville and surrounding areas, TN',                                   lat:36.1627, lng:-86.7816 },
  { type:'ch', name:'Greater Memphis Chapter',    region:'Southeast', areas:'Memphis and surrounding area, TN',                                      lat:35.1495, lng:-90.0490 },
  { type:'ch', name:'Charlotte Chapter',          region:'Southeast', areas:'Charlotte and surrounding areas, NC/SC',                                lat:35.2271, lng:-80.8431 },
  { type:'ch', name:'Coastal Carolina Chapter',   region:'Southeast', areas:'East coast of the Carolina area (Wilmington NC / Myrtle Beach SC)',     lat:34.2257, lng:-77.9447 },
  { type:'ch', name:'Raleigh Durham Chapter',     region:'Southeast', areas:'Raleigh, Durham and surrounding areas, NC',                             lat:35.7796, lng:-78.6382 },

  // ─── SOUTH CENTRAL CHAPTERS ──────────────────────────────────
  { type:'ch', name:'DFW Chapter',          region:'South Central', areas:'Dallas/Fort Worth Metro, TX',                                             lat:32.7767, lng:-96.7970 },
  { type:'ch', name:'Gulf Coast Chapter',   region:'South Central', areas:'Houston and surrounding areas, TX',                                       lat:29.7604, lng:-95.3698 },
  { type:'ch', name:'Hill Country Chapter', region:'South Central', areas:'I-35 corridor from Waco to San Antonio (Austin, Georgetown, San Antonio, TX)', lat:30.2672, lng:-97.7431 },
  { type:'ch', name:'Oklahoma Chapter',     region:'South Central', areas:'Oklahoma (statewide, hub in Oklahoma City)',                               lat:35.4676, lng:-97.5164 },
  { type:'ch', name:'Louisiana Chapter',   region:'South Central', areas:'Louisiana (New Orleans/Baton Rouge area)',                                  lat:29.9511, lng:-90.0715 },
];

let orgLocationsVisible = false;
let albersProjection    = null; // initialized once D3 loads

function initOrgProjection() {
  // geoAlbersUsa default: scale 1070, translate [480,250] — matches us-atlas 960x600 viewBox
  albersProjection = d3.geoAlbersUsa().scale(1280).translate([480, 300]);
}

// Map region name → index for coloring for coloring
const REGION_NAME_TO_IDX = {
  'Northwest': 5, 'Southwest': 4, 'Northeast': 0,
  'North Central': 1, 'Southeast': 2, 'South Central': 3,
};

function toggleOrgLocations() {
  orgLocationsVisible = !orgLocationsVisible;
  const btn = document.getElementById('org-toggle-btn');
  const svg = document.getElementById('org-svg');
  if (orgLocationsVisible) {
    btn.textContent = '📍 Hide CC Locations';
    btn.style.background = 'rgba(0,115,255,0.08)';
    svg.style.display = '';
    document.getElementById('org-legend').style.display = '';
    renderOrgLocations();
  } else {
    btn.textContent = '📍 Show CC Locations';
    btn.style.background = '';
    svg.style.display = 'none';
    document.getElementById('org-legend').style.display = 'none';
    document.getElementById('org-popup').style.display = 'none';
  }
}

function renderOrgLocations() {
  const svg = document.getElementById('org-svg');
  svg.innerHTML = '';
  const mapVB = document.getElementById('map-svg').getAttribute('viewBox');
  svg.setAttribute('viewBox', mapVB);

  if (!albersProjection) initOrgProjection();

  const order = ['ch', 'fo', 'tc'];
  order.forEach(typeFilter => {
    CC_LOCATIONS.filter(loc => loc.type === typeFilter).forEach(loc => {
      // Project lat/lng → SVG coords using D3 geoAlbersUsa
      const projected = albersProjection([loc.lng, loc.lat]);
      if (!projected) return; // outside projection bounds (e.g. some AK/HI edge cases handled by D3)
      const [px, py] = projected;

      const g = document.createElementNS('http://www.w3.org/2000/svg', 'g');
      g.style.pointerEvents = 'all';
      const rIdx  = REGION_NAME_TO_IDX[loc.region] ?? 0;
      const rColor = REGION_COLORS[rIdx];

      if (loc.type === 'tc') {
        const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        rect.setAttribute('x', px - 6); rect.setAttribute('y', py - 6);
        rect.setAttribute('width', '12'); rect.setAttribute('height', '12');
        rect.setAttribute('transform', `rotate(45 ${px} ${py})`);
        rect.setAttribute('fill', '#0073FF');
        rect.setAttribute('stroke', 'white'); rect.setAttribute('stroke-width', '1.5');
        g.appendChild(rect);
        const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        text.setAttribute('x', px); text.setAttribute('y', py + 1);
        text.setAttribute('class', 'org-label'); text.setAttribute('fill', '#FECB00'); // CC yellow star
        text.textContent = '★';
        g.appendChild(text);
      } else if (loc.type === 'fo') {
        const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        rect.setAttribute('x', px - 5); rect.setAttribute('y', py - 5);
        rect.setAttribute('width', '10'); rect.setAttribute('height', '10');
        rect.setAttribute('rx', '2');
        rect.setAttribute('fill', '#FECB00');
        rect.setAttribute('stroke', '#c9a000'); rect.setAttribute('stroke-width', '1.5');
        g.appendChild(rect);
        const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        text.setAttribute('x', px); text.setAttribute('y', py + 1);
        text.setAttribute('class', 'org-label'); text.setAttribute('fill', '#425563');
        text.textContent = 'F';
        g.appendChild(text);
      } else {
        // Chapter: CC yellow fill, dark gold border — consistent across all regions
        const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        circle.setAttribute('cx', px); circle.setAttribute('cy', py);
        circle.setAttribute('r', '4');
        circle.setAttribute('fill', '#FECB00');
        circle.setAttribute('stroke', '#c9a000'); circle.setAttribute('stroke-width', '1.5');
        g.appendChild(circle);
      }

      g.addEventListener('mouseenter', e => showOrgPopup(e, loc));
      g.addEventListener('mouseleave', () => {
        document.getElementById('org-popup').style.display = 'none';
      });
      svg.appendChild(g);
    });
  });
}

function showOrgPopup(e, loc) {
  const popup = document.getElementById('org-popup');
  const typeLabel = loc.type === 'tc' ? 'Regional Training Center'
                  : loc.type === 'fo' ? 'Field Office'
                  : 'Volunteer Chapter';
  const rIdx = REGION_NAME_TO_IDX[loc.region] ?? 0;
  const rColor = REGION_COLORS[rIdx];

  document.getElementById('org-popup-name').textContent = loc.name;
  document.getElementById('org-popup-type').innerHTML =
    `<span style="color:${loc.type === 'tc' ? '#0073FF' : loc.type === 'fo' ? '#c9a000' : rColor}">${typeLabel}</span>`;
  document.getElementById('org-popup-region').innerHTML =
    `<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${rColor};margin-right:4px;vertical-align:middle"></span>${loc.region} Region`;
  document.getElementById('org-popup-areas').textContent = loc.areas ? `📍 ${loc.areas}` : '';
  document.getElementById('org-popup-address').textContent = loc.address || '';

  const container = document.querySelector('.map-container');
  const rect = container.getBoundingClientRect();
  let x = e.clientX - rect.left + 14;
  let y = e.clientY - rect.top - 10;
  if (x + 230 > rect.width) x -= 240;
  popup.style.left = x + 'px';
  popup.style.top  = y + 'px';
  popup.style.display = 'block';
}

// Add a legend for the org marker types to the map
function updateOrgLegendOnMap() {
  // This is handled via the button text and the popup on hover
}

function syncRouteViewBox() {
  const mapVB = document.getElementById('map-svg').getAttribute('viewBox');
  if (routeActive) {
    document.getElementById('route-svg').setAttribute('viewBox', mapVB);
  }
  if (orgLocationsVisible) {
    document.getElementById('org-svg').setAttribute('viewBox', mapVB);
  }
}

// ─── ROUTE OPTIMIZER (All Regions) ───────────────────────────────

// Home base for each region (training center location)
// fixedStops = guaranteed meeting locations pre-seeded into every route
const REGION_HOME_BASES = [
  { regionIdx:0, regionName:'Northeast',     name:'Northeast Training Center',
    lat:40.8204, lng:-72.9968,
    fixedStops:[
      { label:'Campus', name:'Northeast Training Center', lat:40.8204, lng:-72.9968 },
    ]},
  { regionIdx:1, regionName:'North Central', name:'North Central Training Center',
    lat:40.0909, lng:-82.8068,
    fixedStops:[
      { label:'Campus', name:'North Central Training Center', lat:40.0909, lng:-82.8068 },
    ]},
  { regionIdx:2, regionName:'Southeast',     name:'Southeast Training Center',
    lat:28.5794, lng:-81.4874,
    fixedStops:[
      { label:'Campus', name:'Southeast Training Center', lat:28.5794, lng:-81.4874 },
    ]},
  { regionIdx:3, regionName:'South Central', name:'South Central Training Center',
    lat:32.8998, lng:-97.0390,
    fixedStops:[
      { label:'Campus', name:'South Central Training Center', lat:32.8998, lng:-97.0390 },
    ]},
  { regionIdx:4, regionName:'Southwest',     name:'Southwest Training Center',
    lat:33.1959, lng:-117.3228,
    fixedStops:[
      { label:'Campus', name:'Southwest Training Center', lat:33.1959, lng:-117.3228 },
    ]},
  { regionIdx:5, regionName:'Northwest',     name:'Northwest Training Center',
    lat:38.4188, lng:-122.7186,
    fixedStops:[
      { label:'Campus',       name:'Northwest Training Center & HQ', lat:38.4188, lng:-122.7186 },
      { label:'Field Office', name:'Puget Sound Field Office', lat:47.5793, lng:-122.3336 },
    ]},
];

const RADIUS_SVG_PX = 60;
const RADIUS_MILES  = 150;

let routeStops       = [];
let routeActive      = false;
let activeRouteRegion = null; // index 0-5

function getHomeSvgCoords(home) {
  if (!albersProjection) return { svgX: 0, svgY: 0 };
  const pt = albersProjection([home.lng, home.lat]);
  return { svgX: pt?.[0] ?? 0, svgY: pt?.[1] ?? 0 };
}

function openRouteOptimizer() {
  if (!mapLoaded) {
    alert('Please wait for the map to finish loading.');
    return;
  }
  if (Object.keys(clientData).length === 0) {
    alert('Please upload your Salesforce CSV first.');
    return;
  }
  // Show region picker modal
  const modal = document.getElementById('route-region-modal');
  modal.style.display = 'flex';
}

function launchRouteForRegion(regionIdx) {
  document.getElementById('route-region-modal').style.display = 'none';
  activeRouteRegion = regionIdx;

  const home = REGION_HOME_BASES[regionIdx];
  routeStops = computeRoute(regionIdx, home);
  renderRouteOnMap(home);
  renderRoutePanel(home);

  // Zoom to the selected region
  activeRegion = String(regionIdx);
  activeState  = 'all';
  updateRegionFilters();
  updateStateList();
  updateActiveChips();
  updateStats();
  updateContextLegend();
  colorMap();
  zoomToSelection();

  document.getElementById('route-panel').classList.add('open');
  document.getElementById('route-svg').style.display = '';
  routeActive = true;
}

function closeRoutePanel() {
  document.getElementById('route-panel').classList.remove('open');
  document.getElementById('route-svg').style.display = 'none';
  routeActive = false;
  routeStops  = [];
  activeRouteRegion = null;
}

// ── Core algorithm ────────────────────────────────────────────────
function computeRoute(regionIdx, home) {
  const { svgX: homeX, svgY: homeY } = getHomeSvgCoords(home);

  // Collect all counties in this region that have clients
  const regionCounties = [];
  Object.entries(clientData).forEach(([key, info]) => {
    if (info.regionIdx !== regionIdx) return;
    const fips = findFips(info.county, info.state);
    if (!fips) return;
    const feature = allCountyFeatures.find(f => String(f.id) === String(fips));
    if (!feature) return;
    const [cx, cy] = geoPathFn.centroid(feature);
    if (isNaN(cx) || isNaN(cy)) return;
    regionCounties.push({ key, info, fips, cx, cy });
  });

  if (!regionCounties.length) return [];

  // Pre-seed clusters at every fixed stop location (campus + field offices)
  // Their cx/cy are locked — they never drift from the actual location
  const clusters = home.fixedStops.map(fs => {
    const pt = albersProjection([fs.lng, fs.lat]);
    const cx = pt?.[0] ?? homeX;
    const cy = pt?.[1] ?? homeY;
    return {
      counties: [],
      cx, cy,
      fixedX: cx, fixedY: cy, // never mutated
      isFixed: true,
      fixedLabel: fs.label,
      fixedName:  fs.name,
      clients: 0,
    };
  });

  // Sort counties by distance from home base outward
  regionCounties.sort((a, b) => dist(a.cx, a.cy, homeX, homeY)
                               - dist(b.cx, b.cy, homeX, homeY));

  // Greedy clustering — fixed clusters get first pick
  const assigned = new Set();
  regionCounties.forEach(county => {
    if (assigned.has(county.key)) return;
    let bestCluster = null, bestD = Infinity;
    clusters.forEach(cl => {
      // For fixed clusters use their locked coords; dynamic use current centroid
      const clX = cl.isFixed ? cl.fixedX : cl.cx;
      const clY = cl.isFixed ? cl.fixedY : cl.cy;
      const d = dist(county.cx, county.cy, clX, clY);
      if (d <= RADIUS_SVG_PX && d < bestD) { bestD = d; bestCluster = cl; }
    });
    if (bestCluster) {
      bestCluster.counties.push(county);
      assigned.add(county.key);
      if (!bestCluster.isFixed) recomputeCentroid(bestCluster);
      // Fixed clusters keep their locked coords
    } else {
      const cl = { counties: [county], cx: county.cx, cy: county.cy, clients: 0, isFixed: false };
      assigned.add(county.key);
      clusters.push(cl);
    }
  });

  // Refinement passes (only move counties between dynamic clusters, not away from fixed)
  let changed = true, passes = 0;
  while (changed && passes < 5) {
    changed = false; passes++;
    regionCounties.forEach(county => {
      const cur = clusters.find(cl => cl.counties.includes(county));
      if (!cur) return;
      clusters.forEach(cl => {
        if (cl === cur) return;
        const clX = cl.isFixed ? cl.fixedX : cl.cx;
        const clY = cl.isFixed ? cl.fixedY : cl.cy;
        const curX = cur.isFixed ? cur.fixedX : cur.cx;
        const curY = cur.isFixed ? cur.fixedY : cur.cy;
        if (dist(county.cx, county.cy, clX, clY) <= RADIUS_SVG_PX &&
            dist(county.cx, county.cy, clX, clY) < dist(county.cx, county.cy, curX, curY)) {
          cur.counties = cur.counties.filter(c => c !== county);
          cl.counties.push(county);
          if (!cur.isFixed) recomputeCentroid(cur);
          // fixed clusters keep their coords
          changed = true;
        }
      });
    });
  }

  // Include ALL clusters — fixed stops always appear even with 0 clients in range
  const allClusters = clusters.filter(cl => cl.isFixed || cl.counties.length > 0);
  allClusters.forEach(cl => { if (!cl.isFixed) recomputeCentroid(cl); });

  // Nearest-neighbor TSP from home base
  const ordered = [];
  const remaining = [...allClusters];
  let current = { cx: homeX, cy: homeY };
  while (remaining.length) {
    let nearestIdx = 0, nearestD = Infinity;
    remaining.forEach((cl, i) => {
      const clX = cl.isFixed ? cl.fixedX : cl.cx;
      const clY = cl.isFixed ? cl.fixedY : cl.cy;
      const d = dist(current.cx, current.cy, clX, clY);
      if (d < nearestD) { nearestD = d; nearestIdx = i; }
    });
    const next = remaining.splice(nearestIdx, 1)[0];
    ordered.push(next);
    current = { cx: next.isFixed ? next.fixedX : next.cx,
                cy: next.isFixed ? next.fixedY : next.cy };
  }

  return ordered.map((cl, i) => {
    const totalClients  = cl.counties.reduce((s, c) => s + c.info.count, 0);
    const countyNames   = cl.counties
      .sort((a, b) => b.info.count - a.info.count)
      .map(c => `${c.info.county}, ${c.info.state} (${c.info.count})`);
    const cx = cl.isFixed ? cl.fixedX : cl.cx;
    const cy = cl.isFixed ? cl.fixedY : cl.cy;
    const prevCl = i === 0 ? null : ordered[i-1];
    const prevX  = i === 0 ? homeX : (prevCl.isFixed ? prevCl.fixedX : prevCl.cx);
    const prevY  = i === 0 ? homeY : (prevCl.isFixed ? prevCl.fixedY : prevCl.cy);
    const hubName = cl.isFixed
      ? cl.fixedName
      : (cl.counties.length
          ? `${cl.counties.reduce((a,b) => a.info.count > b.info.count ? a : b).info.county} County, ${cl.counties.reduce((a,b) => a.info.count > b.info.count ? a : b).info.state}`
          : 'Unknown');
    return {
      stopNum: i + 1,
      cx, cy,
      totalClients,
      countyNames,
      hubName,
      driveFromPrev: estDriveHours(prevX, prevY, cx, cy),
      counties: cl.counties,
      isFixed:    cl.isFixed  || false,
      fixedLabel: cl.fixedLabel || null,
    };
  });
}

function recomputeCentroid(cl) {
  if (!cl.counties.length) return;
  const total = cl.counties.reduce((s, c) => s + c.info.count, 0) || cl.counties.length;
  cl.cx = cl.counties.reduce((s, c) => s + c.cx * c.info.count, 0) / total;
  cl.cy = cl.counties.reduce((s, c) => s + c.cy * c.info.count, 0) / total;
  cl.clients = total;
}

function dist(x1, y1, x2, y2) {
  return Math.sqrt((x2-x1)**2 + (y2-y1)**2);
}

function estDriveHours(x1, y1, x2, y2) {
  const miles = dist(x1, y1, x2, y2) * 2.5;
  return (miles / 55).toFixed(1);
}

// ── Render SVG overlay ────────────────────────────────────────────
function renderRouteOnMap(home) {
  const svg = document.getElementById('route-svg');
  svg.innerHTML = '';
  if (!routeStops.length) return;

  const mapVB = document.getElementById('map-svg').getAttribute('viewBox');
  svg.setAttribute('viewBox', mapVB);

  const { svgX: homeX, svgY: homeY } = getHomeSvgCoords(home);

  // Coverage radius circles — only for non-fixed stops
  routeStops.forEach(stop => {
    if (stop.isFixed) return;
    const circle = document.createElementNS('http://www.w3.org/2000/svg','circle');
    circle.setAttribute('cx', stop.cx); circle.setAttribute('cy', stop.cy);
    circle.setAttribute('r', RADIUS_SVG_PX);
    circle.setAttribute('class', 'route-radius');
    svg.appendChild(circle);
  });

  // Also draw radius for fixed stops (campus/field office serve nearby clients)
  routeStops.forEach(stop => {
    if (!stop.isFixed) return;
    const circle = document.createElementNS('http://www.w3.org/2000/svg','circle');
    circle.setAttribute('cx', stop.cx); circle.setAttribute('cy', stop.cy);
    circle.setAttribute('r', RADIUS_SVG_PX);
    circle.setAttribute('class', 'route-radius');
    circle.style.stroke = '#0073FF';
    circle.style.fill   = 'rgba(0,115,255,0.04)';
    svg.appendChild(circle);
  });

  // Route lines — start from home base, connect all stops in order
  const allPoints = [{ cx: homeX, cy: homeY }, ...routeStops];
  for (let i = 0; i < allPoints.length - 1; i++) {
    const line = document.createElementNS('http://www.w3.org/2000/svg','line');
    line.setAttribute('x1', allPoints[i].cx); line.setAttribute('y1', allPoints[i].cy);
    line.setAttribute('x2', allPoints[i+1].cx); line.setAttribute('y2', allPoints[i+1].cy);
    line.setAttribute('class', 'route-line');
    svg.appendChild(line);
  }

  // Draw home base departing marker (only if campus is NOT already stop 1)
  const campusIsStop1 = routeStops[0]?.isFixed && routeStops[0]?.fixedLabel === 'Campus';
  if (!campusIsStop1) {
    const homeG = document.createElementNS('http://www.w3.org/2000/svg','g');
    const homeCircle = document.createElementNS('http://www.w3.org/2000/svg','circle');
    homeCircle.setAttribute('cx', homeX); homeCircle.setAttribute('cy', homeY);
    homeCircle.setAttribute('r', 9);
    homeCircle.setAttribute('class', 'route-marker home-marker');
    const homeLabel = document.createElementNS('http://www.w3.org/2000/svg','text');
    homeLabel.setAttribute('x', homeX); homeLabel.setAttribute('y', homeY);
    homeLabel.setAttribute('class', 'route-label');
    homeLabel.textContent = '🏠';
    homeG.appendChild(homeCircle); homeG.appendChild(homeLabel);
    svg.appendChild(homeG);
  }

  // Stop markers
  routeStops.forEach(stop => {
    const g = document.createElementNS('http://www.w3.org/2000/svg','g');
    g.style.cursor = 'pointer';
    g.addEventListener('click', () => scrollToStop(stop.stopNum));

    if (stop.isFixed) {
      // Fixed stops: blue square for campus, yellow square for field office
      const isFieldOffice = stop.fixedLabel === 'Field Office';
      const fillColor   = isFieldOffice ? '#FECB00' : '#0073FF';
      const strokeColor = isFieldOffice ? '#c9a000' : 'white';
      const rect = document.createElementNS('http://www.w3.org/2000/svg','rect');
      rect.setAttribute('x', stop.cx - 9); rect.setAttribute('y', stop.cy - 9);
      rect.setAttribute('width', '18'); rect.setAttribute('height', '18');
      rect.setAttribute('rx', '4');
      rect.setAttribute('fill', fillColor);
      rect.setAttribute('stroke', strokeColor);
      rect.setAttribute('stroke-width', '2');
      const label = document.createElementNS('http://www.w3.org/2000/svg','text');
      label.setAttribute('x', stop.cx); label.setAttribute('y', stop.cy + 1);
      label.setAttribute('class', 'route-label');
      label.setAttribute('fill', isFieldOffice ? '#425563' : 'white');
      label.textContent = stop.stopNum;
      g.appendChild(rect); g.appendChild(label);
    } else {
      // Regular stops: purple circles
      const circle = document.createElementNS('http://www.w3.org/2000/svg','circle');
      circle.setAttribute('cx', stop.cx); circle.setAttribute('cy', stop.cy);
      circle.setAttribute('r', 10);
      circle.setAttribute('class', 'route-marker');
      const label = document.createElementNS('http://www.w3.org/2000/svg','text');
      label.setAttribute('x', stop.cx); label.setAttribute('y', stop.cy + 1);
      label.setAttribute('class', 'route-label');
      label.textContent = stop.stopNum;
      g.appendChild(circle); g.appendChild(label);
    }
    svg.appendChild(g);
  });
}

// ── Render panel ──────────────────────────────────────────────────
function renderRoutePanel(home) {
  const list    = document.getElementById('route-stop-list');
  const summary = document.getElementById('route-summary');
  const sub     = document.getElementById('route-panel-sub');
  const title   = document.querySelector('.route-panel-title');
  list.innerHTML = '';

  const nonFixedStops  = routeStops.filter(s => !s.isFixed);
  const totalClients   = routeStops.reduce((s, st) => s + st.totalClients, 0);
  const totalStops     = routeStops.length;
  title.textContent    = `🗺️ ${home.regionName} Route`;
  sub.textContent      = `${totalStops} stop${totalStops !== 1 ? 's' : ''} · ${totalClients} clients covered`;

  // Legend for marker types
  const legend = document.createElement('div');
  legend.style.cssText = 'display:flex;gap:12px;padding:8px 16px 6px;border-bottom:1px solid #f0f4f8;flex-wrap:wrap';
  const fixedCount = home.fixedStops.length;
  legend.innerHTML = `
    <span style="display:flex;align-items:center;gap:5px;font-size:0.64rem;color:var(--muted)">
      <span style="width:14px;height:14px;background:#0073FF;border-radius:3px;flex-shrink:0;display:inline-block"></span>Campus
    </span>
    ${fixedCount > 1 ? `<span style="display:flex;align-items:center;gap:5px;font-size:0.64rem;color:var(--muted)">
      <span style="width:14px;height:14px;background:#FECB00;border-radius:3px;flex-shrink:0;display:inline-block;border:1px solid #c9a000"></span>Field Office
    </span>` : ''}
    <span style="display:flex;align-items:center;gap:5px;font-size:0.64rem;color:var(--muted)">
      <span style="width:14px;height:14px;background:#6a0dad;border-radius:50%;flex-shrink:0;display:inline-block"></span>Meeting Stop
    </span>`;
  list.appendChild(legend);

  if (!routeStops.length) {
    const empty = document.createElement('div');
    empty.style.cssText = 'padding:24px 16px;text-align:center;color:var(--muted);font-size:0.75rem;line-height:1.8';
    empty.innerHTML = '📭 No client data found for this region.<br>Upload your Salesforce CSV to generate a route.';
    list.appendChild(empty);
    summary.innerHTML = '';
    return;
  }

  routeStops.forEach(stop => {
    const row = document.createElement('div');
    row.className = 'route-stop-row';
    row.id = `route-stop-${stop.stopNum}`;
    row.addEventListener('click', () => {
      zoomToViewBox({ x: stop.cx - 80, y: stop.cy - 60, w: 160, h: 120 });
      document.getElementById('zoom-out-btn').classList.add('visible');
    });

    const countiesHtml = stop.countyNames.length
      ? stop.countyNames.slice(0,5).join('<br>') +
        (stop.countyNames.length > 5 ? `<br><em>+${stop.countyNames.length - 5} more</em>` : '')
      : '<em style="color:var(--muted)">Clients may travel from surrounding counties</em>';

    let badge, stopLabel, stopMeta, stopExtra;

    if (stop.isFixed) {
      const isFieldOffice = stop.fixedLabel === 'Field Office';
      const bgColor = isFieldOffice ? '#FECB00' : '#0073FF';
      const txtColor = isFieldOffice ? '#425563' : 'white';
      badge     = `<div class="stop-badge" style="background:${bgColor};color:${txtColor};border-radius:4px">${stop.stopNum}</div>`;
      stopLabel = `<span style="font-size:0.6rem;font-weight:700;text-transform:uppercase;letter-spacing:0.08em;color:${isFieldOffice ? '#c9a000' : '#0073FF'};background:${isFieldOffice ? 'rgba(254,203,0,0.15)' : 'rgba(0,115,255,0.08)'};border-radius:3px;padding:1px 5px;margin-left:6px">${stop.fixedLabel}</span>`;
      stopMeta  = `~${stop.driveFromPrev}hr from ${stop.stopNum === 1 ? 'home base' : `stop ${stop.stopNum - 1}`}`;
      stopExtra = stop.totalClients > 0
        ? `<div class="stop-clients">👥 ${stop.totalClients} client${stop.totalClients !== 1 ? 's' : ''} within ~${RADIUS_MILES}mi</div>
           <div class="stop-counties">${countiesHtml}</div>`
        : `<div style="font-size:0.66rem;color:var(--muted);margin-top:3px;font-style:italic">Fixed location — clients from nearby counties</div>`;
    } else {
      badge     = `<div class="stop-badge">${stop.stopNum}</div>`;
      stopLabel = '';
      stopMeta  = `~${stop.driveFromPrev}hr from ${stop.stopNum === 1 ? 'home base' : `stop ${stop.stopNum - 1}`}`;
      stopExtra = `<div class="stop-clients">👥 ${stop.totalClients} client${stop.totalClients !== 1 ? 's' : ''} within ~${RADIUS_MILES}mi</div>
                   <div class="stop-counties">${countiesHtml}</div>`;
    }

    row.innerHTML = `
      ${badge}
      <div class="stop-details">
        <div class="stop-name">Stop ${stop.stopNum} — ${stop.hubName}${stopLabel}</div>
        <div class="stop-meta">${stopMeta}</div>
        ${stopExtra}
      </div>`;
    list.appendChild(row);
  });

  const totalDrive = routeStops.reduce((s, st) => s + parseFloat(st.driveFromPrev), 0).toFixed(1);
  const fixedLabels = home.fixedStops.map(fs => fs.label).join(' + ');
  summary.innerHTML = `
    <strong>${totalStops} stops</strong> (incl. ${fixedLabels}) · <strong>${totalClients} clients</strong><br>
    ~${totalDrive}hr total drive · All within <strong>~${RADIUS_MILES}mi</strong> of a stop<br>
    <span style="color:var(--muted);font-size:0.63rem">Click any stop to zoom · Blue/yellow = fixed locations</span>`;
}

function scrollToStop(num) {
  const el = document.getElementById(`route-stop-${num}`);
  if (el) el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}


document.addEventListener('DOMContentLoaded', () => {
  buildDensityLegend();
  // Populate route region picker
  const btnContainer = document.getElementById('route-region-buttons');
  REGION_HOME_BASES.forEach(home => {
    const rColor = REGION_COLORS[home.regionIdx];
    const btn = document.createElement('button');
    btn.style.cssText = `display:flex;align-items:center;gap:10px;padding:10px 14px;border-radius:8px;border:1px solid ${rColor};background:transparent;cursor:pointer;font-family:Poppins,sans-serif;font-size:0.78rem;font-weight:500;color:${rColor};transition:all 0.15s;text-align:left`;
    btn.innerHTML = `<span style="width:10px;height:10px;border-radius:50%;background:${rColor};flex-shrink:0"></span><span style="flex:1">${home.regionName}</span><span style="font-size:0.65rem;color:var(--muted);font-weight:400">${home.name}</span>`;
    btn.onmouseenter = () => btn.style.background = `rgba(${parseInt(rColor.slice(1,3),16)},${parseInt(rColor.slice(3,5),16)},${parseInt(rColor.slice(5,7),16)},0.08)`;
    btn.onmouseleave = () => btn.style.background = 'transparent';
    btn.onclick = () => launchRouteForRegion(home.regionIdx);
    btnContainer.appendChild(btn);
  });

  function waitAndLoad(attempts) {
    if (typeof d3 !== 'undefined' && typeof topojson !== 'undefined') {
      loadMap();
    } else if (attempts > 0) {
      setTimeout(() => waitAndLoad(attempts - 1), 100);
    } else {
      document.getElementById('loading').innerHTML = `
        <div style="text-align:center;padding:20px;max-width:340px">
          <div style="color:#1a2a3a;font-weight:600;margin-bottom:8px">Libraries failed to load</div>
          <div style="color:#6b7e8f;font-size:0.78rem">Please check your internet connection and refresh.</div>
        </div>`;
    }
  }
  waitAndLoad(30);
});