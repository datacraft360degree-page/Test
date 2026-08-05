<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Homestay Business Portal</title>
  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- FontAwesome Icons -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background-color: #F2F4F7;
    }
    ::-webkit-scrollbar { width: 5px; height: 5px; }
    ::-webkit-scrollbar-track { background: #F2F4F7; }
    ::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 10px; }
  </style>
</head>
<body class="text-slate-800 font-sans min-h-screen flex flex-col relative antialiased text-xs">

  <!-- LOGIN OVERLAY -->
  <div id="login-overlay" class="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-sm w-full p-6 space-y-4">
      <div class="text-center space-y-1">
        <div class="bg-blue-50 text-blue-600 w-12 h-12 rounded-2xl flex items-center justify-center mx-auto text-xl shadow-sm">
          <i class="fa-solid fa-lock"></i>
        </div>
        <h2 class="text-base font-bold text-slate-900 mt-2">Homestay Business Portal</h2>
        <p class="text-[11px] text-slate-500">Enter details to unlock management portal</p>
      </div>

      <form onsubmit="handleLogin(event)" class="space-y-3">
        <div>
          <label class="block text-[11px] font-semibold text-slate-700 mb-1">User ID</label>
          <input type="text" id="login-userid" required value="admin" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl px-3 py-2 focus:outline-none focus:bg-white text-xs transition">
        </div>
        <div>
          <label class="block text-[11px] font-semibold text-slate-700 mb-1">Password</label>
          <input type="password" id="login-password" required value="admin" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl px-3 py-2 focus:outline-none focus:bg-white text-xs transition">
        </div>
        <button type="submit" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 rounded-2xl shadow-sm transition text-xs flex items-center justify-center gap-1.5">
          <i class="fa-solid fa-right-to-bracket"></i> Login
        </button>
      </form>
    </div>
  </div>

  <!-- HEADER NAVIGATION -->
  <header class="bg-white/80 backdrop-blur-md border-b border-slate-200/60 sticky top-0 z-40">
    <div class="max-w-7xl mx-auto px-4 py-2.5 flex flex-col md:flex-row justify-between items-center gap-2.5">
      <div class="flex items-center space-x-2.5">
        <div class="bg-blue-600 p-2 rounded-2xl text-white shadow-sm">
          <i class="fa-solid fa-hotel text-sm"></i>
        </div>
        <div>
          <h1 class="text-sm font-bold tracking-tight text-slate-900">Homestay Business Portal 🏠</h1>
          <p class="text-[10px] text-slate-500 mt-0.5" id="sync-status-bar">Sync Status: Initializing...</p>
        </div>
      </div>
      
      <!-- PILL NAVIGATION -->
      <nav class="flex space-x-1 bg-slate-100 p-1 rounded-full text-[11px] font-medium">
        <button onclick="switchTab('dashboard')" id="btn-dashboard" class="tab-btn px-3 py-1 rounded-full transition-all bg-white text-blue-600 shadow-sm font-bold">Dashboard</button>
        <button onclick="switchTab('booking')" id="btn-booking" class="tab-btn px-3 py-1 rounded-full transition-all text-slate-600 hover:text-slate-900">Booking Details</button>
      </nav>

      <!-- ACTION BUTTONS -->
      <div class="flex items-center space-x-1.5">
        <button onclick="uploadToGoogleSheet(true)" id="btn-sync-sheets" class="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition shadow-sm">
          <i class="fa-solid fa-cloud-arrow-up text-[10px]"></i> Sync to Sheets
        </button>
        <button onclick="logoutUser()" title="Logout" class="bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-right-from-bracket text-[10px]"></i> Logout
        </button>
      </div>
    </div>
  </header>

  <!-- MAIN VIEW -->
  <main class="max-w-7xl mx-auto px-4 py-6 w-full flex-grow">
    
    <!-- DASHBOARD VIEW -->
    <div id="dashboard-view" class="space-y-4">
      <div class="bg-white p-6 rounded-3xl border border-slate-200/80 shadow-sm">
        <h2 class="text-sm font-bold text-slate-900">Dashboard & Sync Status</h2>
        <p class="text-slate-500 text-[11px] mt-1">Automatic sync triggers every 15 minutes. Remote data will automatically populate when opening on any new device.</p>
        <div class="mt-4 grid grid-cols-1 md:grid-cols-3 gap-4">
          <div class="bg-blue-50 p-4 rounded-2xl border border-blue-100">
            <span class="text-blue-500 text-[10px] font-bold uppercase tracking-wider">Total Bookings</span>
            <div class="text-xl font-bold text-slate-800 mt-1" id="dash-total-count">0</div>
          </div>
        </div>
      </div>
    </div>

    <!-- BOOKINGS VIEW -->
    <div id="booking-view" class="hidden space-y-4">
      <div class="bg-white p-6 rounded-3xl border border-slate-200/80 shadow-sm space-y-4">
        <div class="flex justify-between items-center">
          <h2 class="text-sm font-bold text-slate-900">Booking Management</h2>
          <button onclick="addNewBookingRow()" class="bg-blue-600 text-white px-3 py-1.5 rounded-xl font-medium text-xs flex items-center gap-1">
            <i class="fa-solid fa-plus"></i> Add Entry
          </button>
        </div>

        <!-- BOOKING TABLE -->
        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="border-b border-slate-200 text-slate-400 font-medium text-[10px] uppercase">
                <th class="py-2 px-2">Booking ID</th>
                <th class="py-2 px-2">Guest Name</th>
                <th class="py-2 px-2">Contact</th>
                <th class="py-2 px-2">Room</th>
                <th class="py-2 px-2">Total (₹)</th>
                <th class="py-2 px-2">Status</th>
                <th class="py-2 px-2 text-right">Action</th>
              </tr>
            </thead>
            <tbody id="booking-table-body" class="divide-y divide-slate-100">
              <!-- Dynamic Rows inserted here -->
            </tbody>
          </table>
        </div>
      </div>
    </div>

  </main>

  <!-- TOAST NOTIFICATION -->
  <div id="toast" class="hidden fixed bottom-6 right-6 bg-slate-900/90 text-white px-4 py-2.5 rounded-2xl shadow-xl z-50 flex items-center gap-2 border border-slate-800 text-[11px]">
    <i class="fa-solid fa-circle-check text-emerald-400 text-base"></i>
    <span id="toast-message">Operation Completed</span>
  </div>

  <script>
    // PASTE YOUR GOOGLE APPS SCRIPT WEB APP URL HERE
    const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyJ0apWBEOs8VaSYg86bCBnDjruKF163bKyH6T13fQeBloPZXbgA7fWUH9FvwgvLEtd/exec";

    // Application State
    let state = { bookings: [] };

    // 15 Minutes Interval Timer (in milliseconds)
    const AUTO_SAVE_INTERVAL_MS = 15 * 60 * 1000;

    // Load initial data on page boot
    window.addEventListener('DOMContentLoaded', async () => {
      await fetchFromGoogleSheet();
      
      // Start 15-Minute Auto Sync Background Loop
      setInterval(() => {
        uploadToGoogleSheet(false);
      }, AUTO_SAVE_INTERVAL_MS);
    });

    /**
     * FETCHES REMOTE DATA FROM GOOGLE SHEETS
     */
    async function fetchFromGoogleSheet() {
      const statusBar = document.getElementById('sync-status-bar');
      if (!GOOGLE_SCRIPT_URL || GOOGLE_SCRIPT_URL.includes("YOUR_GOOGLE_APPS_SCRIPT")) {
        statusBar.innerText = "Sync Status: App URL missing inside script";
        return;
      }

      statusBar.innerHTML = `<i class="fa-solid fa-spinner animate-spin"></i> Fetching latest remote data...`;

      try {
        const response = await fetch(GOOGLE_SCRIPT_URL);
        const result = await response.json();

        if (result.status === "success" && Array.isArray(result.bookings)) {
          state.bookings = result.bookings;
          renderTable();
          const lastSyncTime = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
          statusBar.innerText = `Last Fetched: ${lastSyncTime}`;
          showToast("Remote data synced successfully!");
        }
      } catch (err) {
        console.error("Fetch failed:", err);
        statusBar.innerText = "Sync Status: Offline / Using Local Mode";
      }
    }

    /**
     * SAVES LOCAL DATA BACK TO GOOGLE SHEETS
     */
    async function uploadToGoogleSheet(isManual = false) {
      const statusBar = document.getElementById('sync-status-bar');
      const btn = document.getElementById('btn-sync-sheets');

      if (!GOOGLE_SCRIPT_URL || GOOGLE_SCRIPT_URL.includes("YOUR_GOOGLE_APPS_SCRIPT")) {
        alert("⚠️ Please add your deployed Web App URL into the script.");
        return;
      }

      if (isManual) {
        btn.disabled = true;
        btn.innerHTML = `<i class="fa-solid fa-spinner animate-spin"></i> Saving...`;
      }

      try {
        await fetch(GOOGLE_SCRIPT_URL, {
          method: "POST",
          mode: "no-cors",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ bookings: state.bookings })
        });

        const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        statusBar.innerText = `Last Synced: ${timeStr} (${isManual ? 'Manual' : '15-min Auto'})`;
        showToast(isManual ? `Saved manually at ${timeStr}` : `15-Min Auto-Sync completed!`);
      } catch (err) {
        console.error("Upload error:", err);
        statusBar.innerText = "Sync Error: Couldn't upload data";
      } finally {
        if (isManual) {
          btn.disabled = false;
          btn.innerHTML = `<i class="fa-solid fa-cloud-arrow-up text-[10px]"></i> Sync to Sheets`;
        }
      }
    }

    /**
     * DOM RENDER FUNCTIONS
     */
    function renderTable() {
      const tbody = document.getElementById('booking-table-body');
      document.getElementById('dash-total-count').innerText = state.bookings.length;
      tbody.innerHTML = '';

      if(state.bookings.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center py-4 text-slate-400">No bookings available. Click 'Add Entry' to create one.</td></tr>`;
        return;
      }

      state.bookings.forEach((item, index) => {
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50 transition";
        tr.innerHTML = `
          <td class="py-2.5 px-2 font-mono font-bold">${item.bookingCode || 'HS-' + (index+100)}</td>
          <td class="py-2.5 px-2"><input type="text" value="${item.name || ''}" onchange="updateRow(${index}, 'name', this.value)" class="bg-transparent border-b border-transparent focus:border-blue-500 focus:bg-white px-1 py-0.5 w-full"></td>
          <td class="py-2.5 px-2"><input type="text" value="${item.contactNo || ''}" onchange="updateRow(${index}, 'contactNo', this.value)" class="bg-transparent border-b border-transparent focus:border-blue-500 focus:bg-white px-1 py-0.5 w-full"></td>
          <td class="py-2.5 px-2"><input type="text" value="${item.roomNo || ''}" onchange="updateRow(${index}, 'roomNo', this.value)" class="bg-transparent border-b border-transparent focus:border-blue-500 focus:bg-white px-1 py-0.5 w-20"></td>
          <td class="py-2.5 px-2"><input type="number" value="${item.totalAmount || 0}" onchange="updateRow(${index}, 'totalAmount', this.value)" class="bg-transparent border-b border-transparent focus:border-blue-500 focus:bg-white px-1 py-0.5 w-20"></td>
          <td class="py-2.5 px-2">
            <span class="px-2 py-0.5 rounded-full text-[10px] font-bold ${item.status === 'Active' ? 'bg-emerald-100 text-emerald-700' : 'bg-slate-100 text-slate-600'}">
              ${item.status || 'Active'}
            </span>
          </td>
          <td class="py-2.5 px-2 text-right">
            <button onclick="removeRow(${index})" class="text-rose-500 hover:text-rose-700 px-2 py-1"><i class="fa-solid fa-trash"></i></button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function addNewBookingRow() {
      state.bookings.push({
        bookingCode: 'HS-' + Math.floor(1000 + Math.random() * 9000),
        name: 'New Guest',
        contactNo: '',
        roomNo: '101',
        totalAmount: 0,
        status: 'Active'
      });
      renderTable();
    }

    function updateRow(index, field, value) {
      state.bookings[index][field] = value;
    }

    function removeRow(index) {
      state.bookings.splice(index, 1);
      renderTable();
    }

    function switchTab(tab) {
      document.getElementById('dashboard-view').classList.toggle('hidden', tab !== 'dashboard');
      document.getElementById('booking-view').classList.toggle('hidden', tab !== 'booking');
      
      document.getElementById('btn-dashboard').className = tab === 'dashboard' ? 'tab-btn px-3 py-1 rounded-full bg-white text-blue-600 shadow-sm font-bold' : 'tab-btn px-3 py-1 rounded-full text-slate-600';
      document.getElementById('btn-booking').className = tab === 'booking' ? 'tab-btn px-3 py-1 rounded-full bg-white text-blue-600 shadow-sm font-bold' : 'tab-btn px-3 py-1 rounded-full text-slate-600';
    }

    function handleLogin(e) {
      e.preventDefault();
      document.getElementById('login-overlay').classList.add('hidden');
    }

    function logoutUser() {
      document.getElementById('login-overlay').classList.remove('hidden');
    }

    function showToast(msg) {
      const toast = document.getElementById('toast');
      document.getElementById('toast-message').innerText = msg;
      toast.classList.remove('hidden');
      setTimeout(() => toast.classList.add('hidden'), 3500);
    }
  </script>
</body>
</html>
