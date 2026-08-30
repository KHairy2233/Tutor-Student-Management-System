/* ==========================================================================
   ROUTE AI - LOGISTICS OPTIMIZATION & TELEMETRY ENGINE
   ========================================================================== */

// --- GLOBAL APPLICATION STATE ---
const AppState = {
  activeTab: 'view-planner',
  activeScenario: 'metropolis',
  activeAlgorithm: 'tsp-nn',
  map: null,
  mapMarkers: [],
  routePolyline: null,
  vehicleMarker: null,
  
  // Scenarios Data
  scenarios: {
    metropolis: {
      name: 'Metropolis Express (New York Core)',
      center: [40.730610, -73.935242],
      zoom: 12,
      waypoints: [
        { id: 1, name: 'Depot - Midtown Logistics Center', lat: 40.758896, lng: -73.985130 },
        { id: 2, name: 'Hub A - Wall Street Financial Plaza', lat: 40.706001, lng: -74.008827 },
        { id: 3, name: 'Hub B - Brooklyn Tech Triangle', lat: 40.693193, lng: -73.985387 },
        { id: 4, name: 'Hub C - Queens Cargo Terminal', lat: 40.742054, lng: -73.936607 },
        { id: 5, name: 'Hub D - Upper West Drop Center', lat: 40.787011, lng: -73.975368 }
      ]
    },
    coastal: {
      name: 'Coastal Logistics Hub (San Francisco)',
      center: [37.774929, -122.419416],
      zoom: 12,
      waypoints: [
        { id: 1, name: 'Depot - Port of SF Hub', lat: 37.7955, lng: -122.3937 },
        { id: 2, name: 'Hub 1 - Market St Fulfillment', lat: 37.7881, lng: -122.4075 },
        { id: 3, name: 'Hub 2 - Mission District Hub', lat: 37.7599, lng: -122.4148 },
        { id: 4, name: 'Hub 3 - Presidio Dispatch Depot', lat: 37.7989, lng: -122.4662 },
        { id: 5, name: 'Hub 4 - SOMA Commerce Center', lat: 37.7785, lng: -122.3968 }
      ]
    }
  },

  // Fleet Telemetry
  fleet: [
    { id: 'VAN-A1', name: 'Fleet Express Van A-1', driver: 'Alex Vance', type: 'Electric Van', status: 'enroute', progress: 42, speed: 52, battery: 84 },
    { id: 'TRK-B4', name: 'Heavy Freight Truck B-4', driver: 'Marcus Brody', type: 'Diesel Hauler', status: 'idle', progress: 0, speed: 0, battery: 94 },
    { id: 'BIK-C2', name: 'Cargo E-Bike C-2', driver: 'Elena Rostova', type: 'Urban E-Bike', status: 'completed', progress: 100, speed: 0, battery: 65 },
    { id: 'VAN-D9', name: 'Urban Sprinter D-9', driver: 'David Kim', type: 'Hybrid Van', status: 'enroute', progress: 75, speed: 44, battery: 72 }
  ],

  // Simulation State
  simulation: {
    isRunning: false,
    timer: null,
    speedMultiplier: 1,
    currentStep: 0,
    activeVehicleIndex: 0
  },

  // Logs Table Data
  logs: [
    { id: 'RT-9082', driver: 'Alex Vance', scenario: 'New York Core', stops: 5, distance: 28.4, duration: '42 mins', status: 'DISPATCHED', score: '99.4%' },
    { id: 'RT-9081', driver: 'Elena Rostova', scenario: 'San Francisco Bay', stops: 4, distance: 19.1, duration: '31 mins', status: 'COMPLETED', score: '98.2%' },
    { id: 'RT-9080', driver: 'Marcus Brody', scenario: 'Chicago Loop', stops: 6, distance: 41.7, duration: '68 mins', status: 'COMPLETED', score: '96.5%' },
    { id: 'RT-9079', driver: 'David Kim', scenario: 'New York Core', stops: 5, distance: 27.9, duration: '39 mins', status: 'COMPLETED', score: '99.1%' }
  ],

  // Charts
  charts: {}
};

// --- HELPER MATH: HAVERSINE DISTANCE FORMULA ---
function calculateDistanceKm(lat1, lon1, lat2, lon2) {
  const R = 6371; // Earth's radius in KM
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = 
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * 
    Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

// --- TSP NEAREST NEIGHBOR ALGORITHM ---
function solveTSPNearestNeighbor(waypoints) {
  if (waypoints.length <= 1) return waypoints;
  
  const visited = new Array(waypoints.length).fill(false);
  const resultOrder = [];
  
  let currentIndex = 0;
  visited[0] = true;
  resultOrder.push(waypoints[0]);

  for (let step = 1; step < waypoints.length; step++) {
    let nearestIndex = -1;
    let minDistance = Infinity;

    for (let i = 0; i < waypoints.length; i++) {
      if (!visited[i]) {
        const dist = calculateDistanceKm(
          waypoints[currentIndex].lat, waypoints[currentIndex].lng,
          waypoints[i].lat, waypoints[i].lng
        );
        if (dist < minDistance) {
          minDistance = dist;
          nearestIndex = i;
        }
      }
    }

    if (nearestIndex !== -1) {
      visited[nearestIndex] = true;
      resultOrder.push(waypoints[nearestIndex]);
      currentIndex = nearestIndex;
    }
  }

  return resultOrder;
}

// --- INITIALIZE APPLICATION ---
document.addEventListener('DOMContentLoaded', () => {
  initNavigation();
  initClock();
  initMap();
  loadScenario('metropolis');
  initSimulation();
  initAnalyticsCharts();
  renderVehicleCards();
  renderLogsTable();
  initEventListeners();
});

// --- NAVIGATION & TABS ---
function initNavigation() {
  const navButtons = document.querySelectorAll('.nav-menu button');
  const tabPanes = document.querySelectorAll('.tab-pane');

  const titles = {
    'view-planner': { title: 'Route Planner & Optimizer', desc: 'Interactive multi-stop pathfinding and VRP algorithm solver' },
    'view-telemetry': { title: 'Fleet Telemetry & Dispatch Center', desc: 'Live GPS location tracking and active vehicle telematics' },
    'view-analytics': { title: 'Analytics & Business Intelligence', desc: 'Fleet performance metrics, CO2 savings, and efficiency trends' },
    'view-logs': { title: 'Route History & Audit Logs', desc: 'Historical dispatch data records and downloadable CSV logs' }
  };

  navButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const targetTab = btn.getAttribute('data-tab');
      
      navButtons.forEach(b => b.classList.remove('active'));
      tabPanes.forEach(p => p.classList.remove('active'));

      btn.classList.add('active');
      const activePane = document.getElementById(targetTab);
      if (activePane) activePane.classList.add('active');

      // Update Topbar
      if (titles[targetTab]) {
        document.getElementById('current-view-title').textContent = titles[targetTab].title;
        document.getElementById('current-view-desc').textContent = titles[targetTab].desc;
      }

      // Refresh map layout if returning to planner
      if (targetTab === 'view-planner' && AppState.map) {
        setTimeout(() => AppState.map.invalidateSize(), 200);
      }
    });
  });
}

// --- TOPBAR LIVE CLOCK ---
function initClock() {
  const clockEl = document.getElementById('topbar-clock');
  function updateTime() {
    const now = new Date();
    clockEl.textContent = now.toLocaleTimeString();
  }
  setInterval(updateTime, 1000);
  updateTime();
}

// --- LEAFLET MAP INITIALIZATION ---
function initMap() {
  AppState.map = L.map('map', {
    zoomControl: true,
    attributionControl: false
  }).setView([40.730610, -73.935242], 12);

  // CartoDB Dark Matter Tiles for sleek dark aesthetics
  L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png', {
    maxZoom: 19,
    subdomains: 'abcd'
  }).addTo(AppState.map);
}

// --- LOAD SCENARIO & DRAW ROUTE ---
function loadScenario(scenarioKey) {
  AppState.activeScenario = scenarioKey;
  const scenario = AppState.scenarios[scenarioKey];
  if (!scenario) return;

  AppState.map.setView(scenario.center, scenario.zoom);
  optimizeAndRenderRoute();
}

function optimizeAndRenderRoute() {
  const currentScenario = AppState.scenarios[AppState.activeScenario];
  let waypoints = [...currentScenario.waypoints];

  // Apply Selected Optimization Algorithm
  const algorithm = document.getElementById('algorithm-select').value;
  if (algorithm === 'tsp-nn') {
    waypoints = solveTSPNearestNeighbor(waypoints);
  }

  // Clear Existing Map Markers & Polyline
  AppState.mapMarkers.forEach(m => AppState.map.removeLayer(m));
  AppState.mapMarkers = [];
  if (AppState.routePolyline) AppState.map.removeLayer(AppState.routePolyline);

  // Render Waypoint Markers
  const latLngs = [];
  const listContainer = document.getElementById('waypoint-list');
  listContainer.innerHTML = '';

  waypoints.forEach((wp, index) => {
    latLngs.push([wp.lat, wp.lng]);

    // Custom Glowing HTML Icon
    const customIcon = L.divIcon({
      className: 'custom-map-pin',
      html: `<div style="
        width: 26px; height: 26px; border-radius: 50%;
        background: linear-gradient(135deg, #6366f1, #a855f7);
        color: white; font-weight: 700; font-size: 12px;
        display: flex; align-items: center; justify-content: center;
        box-shadow: 0 0 12px rgba(99, 102, 241, 0.6);
        border: 2px solid #fff;">${index + 1}</div>`,
      iconSize: [26, 26],
      iconAnchor: [13, 13]
    });

    const marker = L.marker([wp.lat, wp.lng], { icon: customIcon })
      .bindPopup(`<strong>Stop ${index + 1}: ${wp.name}</strong><br>Lat: ${wp.lat.toFixed(4)}, Lng: ${wp.lng.toFixed(4)}`)
      .addTo(AppState.map);
    AppState.mapMarkers.push(marker);

    // Populate Sidebar Waypoint List Item
    const itemEl = document.createElement('div');
    itemEl.className = 'waypoint-item';
    itemEl.innerHTML = `
      <div style="display: flex; align-items: center;">
        <div class="waypoint-badge">${index + 1}</div>
        <span style="font-weight: 500;">${wp.name}</span>
      </div>
      <button class="btn-icon" title="Remove Node" onclick="removeWaypoint(${index})"><i class="ri-delete-bin-line"></i></button>
    `;
    listContainer.appendChild(itemEl);
  });

  document.getElementById('stop-count').textContent = waypoints.length;

  // Draw Vibrant Polyline
  if (latLngs.length > 1) {
    AppState.routePolyline = L.polyline(latLngs, {
      color: '#6366f1',
      weight: 4,
      opacity: 0.9,
      dashArray: '8, 8',
      lineJoin: 'round'
    }).addTo(AppState.map);

    AppState.map.fitBounds(AppState.routePolyline.getBounds(), { padding: [40, 40] });
  }

  // Calculate Distance & Metrics
  let totalDistanceKm = 0;
  for (let i = 0; i < latLngs.length - 1; i++) {
    totalDistanceKm += calculateDistanceKm(latLngs[i][0], latLngs[i][1], latLngs[i+1][0], latLngs[i+1][1]);
  }

  const estDurationMins = Math.round((totalDistanceKm / 40) * 60); // 40 km/h avg speed
  const estFuelLiters = (totalDistanceKm * 0.085).toFixed(1); // 8.5L / 100km

  document.getElementById('stat-distance').textContent = `${totalDistanceKm.toFixed(1)} km`;
  document.getElementById('stat-duration').textContent = `${estDurationMins} mins`;
  document.getElementById('stat-fuel').textContent = `${estFuelLiters} L`;
}

// --- TELEMETRY SIMULATION ---
function initSimulation() {
  const toggleBtn = document.getElementById('sim-toggle-btn');
  const resetBtn = document.getElementById('sim-reset-btn');

  toggleBtn.addEventListener('click', () => {
    if (AppState.simulation.isRunning) {
      pauseSimulation();
    } else {
      startSimulation();
    }
  });

  resetBtn.addEventListener('click', resetSimulation);

  document.querySelectorAll('.speed-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      document.querySelectorAll('.speed-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      AppState.simulation.speedMultiplier = parseInt(btn.getAttribute('data-speed'));
      logTelemetryEvent(`Simulation speed set to ${AppState.simulation.speedMultiplier}x`);
    });
  });
}

function startSimulation() {
  AppState.simulation.isRunning = true;
  const toggleBtn = document.getElementById('sim-toggle-btn');
  toggleBtn.innerHTML = '<i class="ri-pause-fill"></i> Pause Simulation';
  toggleBtn.classList.remove('btn-primary');
  toggleBtn.classList.add('btn-emerald');

  logTelemetryEvent('[SIMULATION STARTED] Live tracking initialized.');

  AppState.simulation.timer = setInterval(() => {
    updateSimulationStep();
  }, 1000 / AppState.simulation.speedMultiplier);
}

function pauseSimulation() {
  AppState.simulation.isRunning = false;
  clearInterval(AppState.simulation.timer);
  const toggleBtn = document.getElementById('sim-toggle-btn');
  toggleBtn.innerHTML = '<i class="ri-play-fill"></i> Resume Simulation';
  toggleBtn.classList.remove('btn-emerald');
  toggleBtn.classList.add('btn-primary');

  logTelemetryEvent('[SIMULATION PAUSED]');
}

function resetSimulation() {
  pauseSimulation();
  AppState.simulation.currentStep = 0;
  AppState.fleet[0].progress = 0;
  renderVehicleCards();
  logTelemetryEvent('[SIMULATION RESET] Vehicle returned to origin depot.');
}

function updateSimulationStep() {
  const vehicle = AppState.fleet[0];
  if (vehicle.progress < 100) {
    vehicle.progress += 2;
    if (vehicle.progress > 100) vehicle.progress = 100;
  } else {
    vehicle.progress = 0; // Loop back
  }

  // Dynamic values
  const currentSpeed = Math.floor(45 + Math.random() * 12);
  const currentBattery = Math.max(10, vehicle.battery - (vehicle.progress * 0.15)).toFixed(0);

  document.getElementById('sim-speed-val').textContent = `${currentSpeed} km/h`;
  document.getElementById('sim-battery-val').textContent = `${currentBattery}%`;

  renderVehicleCards();

  if (Math.random() > 0.65) {
    logTelemetryEvent(`[GPS PING] ${vehicle.id} - Speed: ${currentSpeed} km/h | Progress: ${vehicle.progress}%`);
  }
}

function logTelemetryEvent(message, isAlert = false) {
  const container = document.getElementById('live-feed-container');
  if (!container) return;
  const entry = document.createElement('div');
  entry.className = `log-entry ${isAlert ? 'alert' : ''}`;
  const now = new Date().toLocaleTimeString();
  entry.textContent = `[${now}] ${message}`;
  container.prepend(entry);
}

// --- RENDER FLEET CARDS ---
function renderVehicleCards() {
  const container = document.getElementById('vehicle-cards-container');
  if (!container) return;

  container.innerHTML = '';
  AppState.fleet.forEach((v, index) => {
    const card = document.createElement('div');
    card.className = `vehicle-card glass-card ${index === 0 ? 'active-vehicle' : ''}`;
    
    let statusBadgeClass = 'badge-enroute';
    if (v.status === 'idle') statusBadgeClass = 'badge-idle';
    if (v.status === 'completed') statusBadgeClass = 'badge-completed';

    card.innerHTML = `
      <div class="vehicle-header">
        <div>
          <span class="vehicle-id">${v.name}</span>
          <div style="font-size: 0.775rem; color: var(--text-muted);">${v.type} • Driver: ${v.driver}</div>
        </div>
        <span class="badge-status ${statusBadgeClass}">${v.status}</span>
      </div>
      
      <div style="display: flex; justify-content: space-between; font-size: 0.75rem; color: var(--text-muted);">
        <span>Route Progress</span>
        <span>${v.progress}%</span>
      </div>
      <div class="progress-bar-bg">
        <div class="progress-bar-fill" style="width: ${v.progress}%;"></div>
      </div>
    `;

    card.addEventListener('click', () => {
      document.querySelectorAll('.vehicle-card').forEach(c => c.classList.remove('active-vehicle'));
      card.classList.add('active-vehicle');
      document.getElementById('driver-name-display').textContent = v.driver;
      logTelemetryEvent(`Selected vehicle focus: ${v.name}`);
    });

    container.appendChild(card);
  });
}

// --- ANALYTICS CHARTS (CHART.JS) ---
function initAnalyticsCharts() {
  const trendsCtx = document.getElementById('weekly-trends-chart');
  const distCtx = document.getElementById('fleet-dist-chart');

  if (trendsCtx) {
    AppState.charts.trends = new Chart(trendsCtx, {
      type: 'line',
      data: {
        labels: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
        datasets: [
          {
            label: 'Optimized Distance (km)',
            data: [320, 290, 410, 380, 450, 210, 180],
            borderColor: '#6366f1',
            backgroundColor: 'rgba(99, 102, 241, 0.15)',
            fill: true,
            tension: 0.4
          },
          {
            label: 'Baseline Unoptimized (km)',
            data: [390, 350, 490, 460, 530, 260, 220],
            borderColor: '#f43f5e',
            borderDash: [5, 5],
            fill: false,
            tension: 0.4
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { labels: { color: '#94a3b8', font: { family: 'Inter' } } }
        },
        scales: {
          x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } },
          y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
        }
      }
    });
  }

  if (distCtx) {
    AppState.charts.dist = new Chart(distCtx, {
      type: 'doughnut',
      data: {
        labels: ['Electric Vans', 'Heavy Trucks', 'Cargo E-Bikes', 'Hybrid Sprinters'],
        datasets: [{
          data: [45, 25, 20, 10],
          backgroundColor: ['#6366f1', '#10b981', '#f59e0b', '#06b6d4'],
          borderWidth: 0
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'bottom', labels: { color: '#94a3b8', font: { family: 'Inter' } } }
        }
      }
    });
  }
}

// --- RENDER LOGS TABLE ---
function renderLogsTable() {
  const tbody = document.getElementById('logs-table-body');
  if (!tbody) return;

  tbody.innerHTML = '';
  AppState.logs.forEach(log => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="font-weight: 600; color: var(--primary);">${log.id}</td>
      <td>${log.driver}</td>
      <td>${log.scenario}</td>
      <td>${log.stops} Stops</td>
      <td>${log.distance} km</td>
      <td>${log.duration}</td>
      <td><span class="badge-status badge-enroute">${log.status}</span></td>
      <td style="color: var(--emerald); font-weight: 600;">${log.score}</td>
    `;
    tbody.appendChild(tr);
  });
}

// --- EVENT LISTENERS ---
function initEventListeners() {
  document.getElementById('scenario-select').addEventListener('change', (e) => {
    loadScenario(e.target.value);
  });

  document.getElementById('algorithm-select').addEventListener('change', () => {
    optimizeAndRenderRoute();
  });

  document.getElementById('optimize-btn').addEventListener('click', () => {
    optimizeAndRenderRoute();
  });

  document.getElementById('dispatch-route-btn').addEventListener('click', () => {
    alert('🚀 Route successfully dispatched to Fleet Van A-1!');
    document.getElementById('nav-btn-telemetry').click();
    startSimulation();
  });

  document.getElementById('add-waypoint-btn').addEventListener('click', () => {
    const input = document.getElementById('new-waypoint-input');
    const val = input.value.trim();
    if (val) {
      const scenario = AppState.scenarios[AppState.activeScenario];
      // Random offset around center
      const lat = scenario.center[0] + (Math.random() - 0.5) * 0.08;
      const lng = scenario.center[1] + (Math.random() - 0.5) * 0.08;
      scenario.waypoints.push({ id: Date.now(), name: val, lat, lng });
      input.value = '';
      optimizeAndRenderRoute();
    }
  });

  // CSV Export Listener
  document.getElementById('export-csv-btn').addEventListener('click', exportLogsCSV);
}

function removeWaypoint(index) {
  const scenario = AppState.scenarios[AppState.activeScenario];
  if (scenario.waypoints.length > 2) {
    scenario.waypoints.splice(index, 1);
    optimizeAndRenderRoute();
  } else {
    alert('Minimum 2 waypoints required for a valid route.');
  }
}

function exportLogsCSV() {
  let csvContent = "data:text/csv;charset=utf-8,Route ID,Driver,Scenario,Stops,Distance (km),Duration,Status,Efficiency Score\n";
  AppState.logs.forEach(row => {
    csvContent += `${row.id},${row.driver},${row.scenario},${row.stops},${row.distance},${row.duration},${row.status},${row.score}\n`;
  });
  const encodedUri = encodeURI(csvContent);
  const link = document.createElement("a");
  link.setAttribute("href", encodedUri);
  link.setAttribute("download", "route_audit_logs.csv");
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
