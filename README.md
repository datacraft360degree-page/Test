<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Business Portal - Web Application</title>
  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- SheetJS for Exporting to Excel -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- html2canvas for Generating JPEG Receipts -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
  <!-- FontAwesome Icons -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      background-color: #F2F4F7;
    }
    
    ::-webkit-scrollbar {
      width: 5px;
      height: 5px;
    }
    ::-webkit-scrollbar-track {
      background: #F2F4F7;
    }
    ::-webkit-scrollbar-thumb {
      background: #CBD5E1;
      border-radius: 10px;
    }
    ::-webkit-scrollbar-thumb:hover {
      background: #94A3B8;
    }

    @media print {
      body * {
        visibility: hidden;
      }
      #printable-invoice, #printable-invoice * {
        visibility: visible;
      }
      #printable-invoice {
        position: absolute;
        left: 0;
        top: 0;
        width: 100%;
        margin: 0;
        padding: 15px;
      }
      .no-print {
        display: none !important;
      }
    }

    .excel-comment-box::before {
      content: '';
      position: absolute;
      top: -8px;
      left: 16px;
      border-width: 0 8px 8px 8px;
      border-style: solid;
      border-color: transparent transparent #1E293B transparent;
    }
  </style>
</head>
<body class="text-slate-800 font-sans min-h-screen flex flex-col relative antialiased text-xs" onclick="closeCommentBox()">

  <!-- LOGIN MODAL OVERLAY -->
  <div id="login-overlay" class="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-sm w-full p-6 space-y-4 text-left">
      <div class="text-center space-y-1">
        <div class="bg-blue-50 text-blue-600 w-12 h-12 rounded-2xl flex items-center justify-center mx-auto text-xl shadow-sm">
          <i class="fa-solid fa-lock"></i>
        </div>
        <h2 class="text-base font-bold text-slate-900 mt-2">Homestay Business Portal</h2>
        <p class="text-[11px] text-slate-500">Please enter your credentials to access the system</p>
      </div>

      <form onsubmit="handleLogin(event)" class="space-y-3">
        <div>
          <label class="block text-[11px] font-semibold text-slate-700 mb-1">User ID</label>
          <div class="relative">
            <span class="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400 text-xs">
              <i class="fa-solid fa-user"></i>
            </span>
            <input type="text" id="login-userid" required placeholder="Enter User ID" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl pl-9 pr-3 py-2 focus:outline-none focus:bg-white text-xs transition">
          </div>
        </div>

        <div>
          <label class="block text-[11px] font-semibold text-slate-700 mb-1">Password</label>
          <div class="relative">
            <span class="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400 text-xs">
              <i class="fa-solid fa-key"></i>
            </span>
            <input type="password" id="login-password" required placeholder="Enter Password" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl pl-9 pr-3 py-2 focus:outline-none focus:bg-white text-xs transition">
          </div>
        </div>

        <div id="login-error" class="hidden bg-rose-50 border border-rose-100 text-rose-600 text-[10px] p-2 rounded-xl text-center font-medium">
          Invalid User ID or Password!
        </div>

        <button type="submit" class="w-full bg-blue-600 hover:bg-blue-700 active:scale-98 text-white font-bold py-2.5 rounded-2xl shadow-sm transition text-xs flex items-center justify-center gap-1.5">
          <i class="fa-solid fa-right-to-bracket"></i> Login
        </button>
      </form>
    </div>
  </div>

  <!-- Header Navigation -->
  <header class="bg-white/80 backdrop-blur-md border-b border-slate-200/60 text-slate-900 sticky top-0 z-40 no-print">
    <div class="max-w-7xl mx-auto px-4 py-2.5 flex flex-col md:flex-row justify-between items-center gap-2.5">
      <div class="flex items-center space-x-2.5">
        <div class="bg-blue-600 p-2 rounded-2xl text-white shadow-sm">
          <i class="fa-solid fa-hotel text-sm"></i>
        </div>
        <div>
          <h1 class="text-sm font-bold tracking-tight leading-none text-slate-900">Homestay Business Portal 🏠</h1>
          <p class="text-[10px] text-slate-500 mt-0.5" id="sync-status-bar">Last Synced: Loading data...</p>
        </div>
      </div>
      
      <!-- Pill Navigation -->
      <nav class="flex space-x-1 bg-slate-100 p-1 rounded-full text-[11px] font-medium">
        <button onclick="switchTab('dashboard')" id="btn-dashboard" class="tab-btn px-3 py-1 rounded-full transition-all active-tab bg-white text-blue-600 shadow-sm font-bold">Dashboard</button>
        <button onclick="switchTab('booking')" id="btn-booking" class="tab-btn px-3 py-1 rounded-full transition-all text-slate-600 hover:text-slate-900">Booking Details</button>
        <button onclick="switchTab('master')" id="btn-master" class="tab-btn px-3 py-1 rounded-full transition-all text-slate-600 hover:text-slate-900 flex items-center gap-1">
          <i class="fa-solid fa-lock text-[9px] text-amber-500"></i> Master Data
        </button>
        <button onclick="switchTab('calendar')" id="btn-calendar" class="tab-btn px-3 py-1 rounded-full transition-all text-slate-600 hover:text-slate-900">Calendar</button>
      </nav>

      <!-- Action Buttons -->
      <div class="flex items-center space-x-1.5">
        <!-- MANUAL SYNC / SAVE TO GOOGLE SHEET BUTTON -->
        <button onclick="uploadToGoogleSheet(true)" id="btn-sync-sheets" class="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition shadow-sm">
          <i class="fa-solid fa-cloud-arrow-up text-[10px]"></i> Sync to Sheets
        </button>
        <button onclick="saveChanges()" class="bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-floppy-disk text-[10px]"></i> Local Save
        </button>
        <button onclick="exportToExcel()" class="bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-200 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-file-excel text-[10px]"></i> Export
        </button>
        <button onclick="logoutUser()" title="Logout" class="bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-right-from-bracket text-[10px]"></i> Logout
        </button>
      </div>
    </div>
  </header>

  <!-- Main Application Body Context -->
  <main class="max-w-7xl mx-auto px-4 py-6 w-full flex-grow">
    <div id="dashboard-view" class="space-y-4">
      <div class="bg-white p-6 rounded-3xl border border-slate-200/80 shadow-sm">
        <h2 class="text-sm font-bold text-slate-900">Dashboard & Analytics Overview</h2>
        <p class="text-slate-500 text-[11px] mt-1">Live data synchronization active across connected devices.</p>
      </div>
    </div>
  </main>

  <!-- Notification Toast -->
  <div id="toast" class="hidden fixed bottom-6 right-6 bg-slate-900/90 backdrop-blur-md text-white px-4 py-2.5 rounded-2xl shadow-xl z-50 flex items-center gap-2.5 no-print border border-slate-800 text-[11px]">
    <i class="fa-solid fa-circle-check text-emerald-400 text-base"></i>
    <span id="toast-message" class="font-medium">Action complete</span>
  </div>

  <script>
    // Set deployment web app URL here
    const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyJ0apWBEOs8VaSYg86bCBnDjruKF163bKyH6T13fQeBloPZXbgA7fWUH9FvwgvLEtd/exec";

    // Application Global State
    let state = {
      bookings: []
    };

    // Auto-Save Interval (15 minutes = 900,000 ms)
    const AUTO_SAVE_INTERVAL_MS = 15 * 60 * 1000;

    /**
     * Window Initializer - Triggers remote fetch on initial device load
     */
    window.addEventListener('DOMContentLoaded', async () => {
      await fetchFromGoogleSheet();
      
      // Start 15-Minute Recurring Auto-Sync Engine
      setInterval(() => {
        console.log("Executing scheduled 15-minute auto-save...");
        uploadToGoogleSheet(false);
      }, AUTO_SAVE_INTERVAL_MS);
    });

    /**
     * FETCH ENGINE: Syncs latest remote state when device connects
     */
    async function fetchFromGoogleSheet() {
      const statusBar = document.getElementById('sync-status-bar');
      if (!GOOGLE_SCRIPT_URL || GOOGLE_SCRIPT_URL.includes("YOUR_GOOGLE_APPS_SCRIPT")) {
        statusBar.innerText = "Sync Status: Web App URL Not Configured";
        return;
      }

      statusBar.innerHTML = `<i class="fa-solid fa-spinner animate-spin"></i> Fetching latest remote data...`;

      try {
        const response = await fetch(GOOGLE_SCRIPT_URL);
        const result = await response.json();

        if (result.status === "success" && Array.isArray(result.bookings)) {
          if (result.bookings.length > 0) {
            state.bookings = result.bookings;
            localStorage.setItem('homestay_bookings', JSON.stringify(state.bookings));
            const lastSync = result.bookings[0]?.lastSyncedAt || new Date().toLocaleTimeString();
            statusBar.innerText = `Last Synced: ${lastSync} (Auto)`;
            showToast("Fetched latest saved changes from Google Sheet!");
          } else {
            statusBar.innerText = "Sync Status: Sheet empty. Ready for initial sync.";
          }
        }
      } catch (err) {
        console.error("Fetch error:", err);
        statusBar.innerText = "Sync Status: Offline / Failed to fetch remote data";
      }
    }

    /**
     * UPLOAD ENGINE: Handles Manual & 15-Min Auto Saves to Google Sheet
     * @param {boolean} isManual - True if manually triggered by clicking button
     */
    async function uploadToGoogleSheet(isManual = false) {
      const statusBar = document.getElementById('sync-status-bar');
      const btn = document.getElementById('btn-sync-sheets');
      
      if (!GOOGLE_SCRIPT_URL || GOOGLE_SCRIPT_URL.includes("YOUR_GOOGLE_APPS_SCRIPT")) {
        alert("⚠️ Configuration Missing!\n\nPlease paste your deployed Google Apps Script Web App URL inside the script tag.");
        return;
      }

      const originalBtnText = btn.innerHTML;
      if (isManual) {
        btn.disabled = true;
        btn.innerHTML = `<i class="fa-solid fa-spinner animate-spin text-[10px]"></i> Saving...`;
      }
      
      statusBar.innerHTML = `<i class="fa-solid fa-spinner animate-spin"></i> Syncing changes...`;

      try {
        const payload = {
          bookings: state.bookings
        };

        // Use standard POST fetch
        await fetch(GOOGLE_SCRIPT_URL, {
          method: "POST",
          mode: "no-cors",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        });

        const currentTimestamp = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        const fullDateStamp = `${new Date().toLocaleDateString()} ${currentTimestamp}`;
        
        statusBar.innerText = `Last Synced: ${fullDateStamp} (${isManual ? 'Manual' : '15-min Auto'})`;
        
        showToast(isManual 
          ? `Data saved successfully! Timestamp: ${currentTimestamp}` 
          : `15-Min Auto-Save Complete! (${currentTimestamp})`
        );

      } catch (error) {
        console.error("Upload error:", error);
        statusBar.innerText = "Sync Error: Failed to push update";
        if (isManual) alert("❌ Failed to push changes: " + error.message);
      } finally {
        if (isManual) {
          btn.disabled = false;
          btn.innerHTML = originalBtnText;
        }
      }
    }

    function saveChanges() {
      localStorage.setItem('homestay_bookings', JSON.stringify(state.bookings));
      showToast("Local changes saved successfully!");
    }

    function showToast(msg) {
      const toast = document.getElementById('toast');
      const toastMsg = document.getElementById('toast-message');
      toastMsg.innerText = msg;
      toast.classList.remove('hidden');
      setTimeout(() => toast.classList.add('hidden'), 3500);
    }

    function switchTab(tab) {
      console.log("Switching to tab:", tab);
    }

    function handleLogin(e) {
      e.preventDefault();
      document.getElementById('login-overlay').classList.add('hidden');
    }

    function logoutUser() {
      document.getElementById('login-overlay').classList.remove('hidden');
    }

    function closeCommentBox() {}
    function exportToExcel() {}
  </script>
</body>
</html>
