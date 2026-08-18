<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Business Portal - Web Application</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" />
  <style>
    /* Samsung One UI Smooth Styling & Compact Scrollbar */
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

    /* Print-specific Styles */
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

    /* One UI Comment Box Arrow */
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

  <div id="login-overlay" class="fixed inset-0 z-[60] bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-sm w-full p-6 space-y-4 text-left">
      <div class="text-center space-y-1">
        <div class="bg-blue-50 text-blue-600 w-12 h-12 rounded-2xl flex items-center justify-center mx-auto text-xl shadow-sm">
          <i class="fa-solid fa-lock"></i>
        </div>
        <p class="text-[11px] text-slate-500">Please enter your credentials to access the system</p>
      </div>

      <form onsubmit="handleLogin(event)" class="space-y-3">
        <div>
          <label class="block text-[11px] font-semibold text-slate-700 mb-1">User ID</label>
          <div class="relative">
            <span class="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400 text-xs">
              <i class="fa-solid fa-user"></i>
            </span>
            <input type="text" id="login-userid" required="" placeholder="Enter User ID" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl pl-9 pr-3 py-2 focus:outline-none focus:bg-white text-xs transition" />
          </div>
        </div>

        <div>
          <label class="block text-[11px] font-semibold text-slate-700 mb-1">Password</label>
          <div class="relative">
            <span class="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400 text-xs">
              <i class="fa-solid fa-key"></i>
            </span>
            <input type="password" id="login-password" required="" placeholder="Enter Password" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl pl-9 pr-3 py-2 focus:outline-none focus:bg-white text-xs transition" />
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

  <div id="login-alert-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-blue-100 max-w-md w-full p-6 space-y-4 text-left">
      <div class="flex items-center gap-3 border-b border-slate-100 pb-3">
        <div class="bg-amber-50 text-amber-600 w-10 h-10 rounded-2xl flex items-center justify-center text-lg shadow-sm">
          <i class="fa-solid fa-circle-exclamation"></i>
        </div>
        <div>
          <h3 class="text-sm font-bold text-slate-900">Important System Guidelines</h3>
          <p class="text-[10px] text-slate-500">Please read these instructions before continuing</p>
        </div>
      </div>
      <div class="space-y-2.5 text-[11px] text-slate-700 font-medium">
         <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>1. Google Chrome/Microsoft edge is best view browser for this Portal.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>2. Take backup every day or every week.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>3. Do not force close 'The Portal'; always close it using the 'Logout' option.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>4. Do not 'Login' with multiple device/multiple browser/multiple browser tab at a same time to avoid data merge.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>5. Data save/Data fetch take little bit time so hold on ⏳.</span></p>
      </div>
      <button onclick="closeLoginAlertModal()" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 rounded-2xl shadow-sm transition text-xs mt-2">
        I Understand, Continue
      </button>
    </div>
  </div>

  <div id="master-auth-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-xs w-full p-5 space-y-3 text-left">
      <div class="text-center space-y-1">
        <div class="bg-rose-50 text-rose-600 w-10 h-10 rounded-2xl flex items-center justify-center mx-auto text-lg shadow-sm">
          <i class="fa-solid fa-shield-halved"></i>
        </div>
        <h3 class="text-xs font-bold text-slate-900 mt-1">Master Data Protected</h3>
        <p class="text-[10px] text-slate-500">Enter master password to access configuration and deletion tools.</p>
      </div>

      <form onsubmit="handleMasterAuth(event)" class="space-y-2.5">
        <div>
          <label class="block text-[10px] font-semibold text-slate-700 mb-1">Master Password</label>
          <div class="relative">
            <span class="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400 text-xs">
              <i class="fa-solid fa-key"></i>
            </span>
            <input type="password" id="master-password-input" required="" placeholder="Enter Master Password" class="w-full bg-slate-100 border border-transparent focus:border-rose-500 rounded-2xl pl-9 pr-3 py-2 focus:outline-none focus:bg-white text-xs transition" />
          </div>
        </div>

        <div id="master-auth-error" class="hidden bg-rose-50 border border-rose-100 text-rose-600 text-[10px] p-1.5 rounded-xl text-center font-medium">
          Incorrect Master Password!
        </div>

        <div class="flex space-x-2 pt-1">
          <button type="button" onclick="closeMasterAuthModal()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 rounded-xl text-[11px] transition">
            Cancel
          </button>
          <button type="submit" class="w-1/2 bg-rose-600 hover:bg-rose-700 text-white font-bold py-2 rounded-xl shadow-sm transition text-[11px] flex items-center justify-center gap-1">
            <i class="fa-solid fa-unlock text-[10px]"></i> Unlock
          </button>
        </div>
      </form>
    </div>
  </div>

  <div id="master-delete-confirm-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-rose-100 max-w-sm w-full p-5 space-y-3 text-center">
      <div class="bg-rose-50 text-rose-600 w-12 h-12 rounded-2xl flex items-center justify-center mx-auto text-xl shadow-sm">
        <i class="fa-solid fa-triangle-exclamation"></i>
      </div>
      <div>
        <h3 class="text-xs font-bold text-slate-900">Confirm Permanent Deletion</h3>
        <p id="master-delete-modal-msg" class="text-[11px] text-slate-600 mt-1">Are you sure you want to permanently delete this data from Master Tab? This action cannot be undone.</p>
      </div>
      <div class="flex space-x-2 pt-2">
        <button type="button" onclick="closeMasterDeleteModal()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 rounded-xl text-[11px] transition">
          Cancel
        </button>
        <button type="button" onclick="confirmMasterDeletion()" class="w-1/2 bg-rose-600 hover:bg-rose-700 text-white font-bold py-2 rounded-xl shadow-sm transition text-[11px] flex items-center justify-center gap-1">
          <i class="fa-solid fa-trash-can text-[10px]"></i> Delete Permanently
        </button>
      </div>
    </div>
  </div>

  <div id="logout-confirm-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-xl border border-slate-100 max-w-xs w-full p-5 space-y-3 text-center">
      <div class="bg-rose-50 text-rose-600 w-10 h-10 rounded-2xl flex items-center justify-center mx-auto text-lg">
        <i class="fa-solid fa-right-from-bracket"></i>
      </div>
      <div>
        <h3 class="text-xs font-bold text-slate-900">Confirm Logout</h3>
        <p class="text-[10px] text-slate-500 mt-1">Are you sure you want to log out? All your latest changes will be saved securely before exiting.</p>
      </div>
      <div class="flex space-x-2 pt-2">
        <button type="button" onclick="cancelLogout()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 rounded-xl text-[11px] transition">Cancel</button>
        <button type="button" onclick="processLogoutWithSave()" class="w-1/2 bg-rose-600 hover:bg-rose-700 text-white font-bold py-2 rounded-xl shadow-sm transition text-[11px]">Logout</button>
      </div>
    </div>
  </div>

  <div id="saving-lock-modal" class="hidden fixed inset-0 z-[100] bg-slate-900/60 backdrop-blur-md flex items-center justify-center p-4 no-print cursor-wait">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-sm w-full p-6 text-center space-y-4">
      <div class="text-blue-600 text-5xl animate-bounce">
        ⏳
      </div>
      <div>
        <h3 class="text-lg font-black text-slate-900">Saving & Logging Out...</h3>
        <p class="text-xs text-rose-600 mt-2 font-bold uppercase">Do not close window or shutdown!</p>
        <p class="text-[10px] text-slate-500 mt-1">Please wait while we secure your data.</p>
      </div>
    </div>
  </div>

  <div id="logout-warning-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-xl border border-slate-100 max-w-xs w-full p-5 space-y-3 text-center">
      <div class="bg-amber-50 text-amber-600 w-10 h-10 rounded-2xl flex items-center justify-center mx-auto text-lg">
        <i class="fa-solid fa-hourglass-half"></i>
      </div>
      <div>
        <h3 class="text-xs font-bold text-slate-900">Inactivity Timeout Warning</h3>
        <p class="text-[10px] text-slate-500 mt-1">You will be logged out automatically in <strong id="logout-countdown-seconds" class="text-rose-600">60</strong> seconds due to inactivity. Data will be saved securely.</p>
      </div>
      <button onclick="resetInactivityTimer()" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-2 rounded-xl text-[11px] transition shadow-sm">
        Stay Logged In
      </button>
    </div>
  </div>

  <div id="export-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-sm w-full p-5 space-y-4 text-left">
      <div class="flex justify-between items-center pb-2 border-b border-slate-100">
        <h3 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
          <i class="fa-solid fa-file-excel text-emerald-600"></i> Export to Excel
        </h3>
        <button onclick="closeExportModal()" class="text-slate-400 hover:text-slate-600 p-0.5 text-base"><i class="fa-solid fa-xmark"></i></button>
      </div>
      <div class="space-y-3 text-[11px]">
        <p class="text-slate-500">Select a specific period to download booking details. Available from 1st Aug 2026 to 31st Dec 2085.</p>
        <div>
          <label class="block font-semibold text-slate-600 mb-0.5">Start Date</label>
          <input type="date" id="export-start-date" min="2026-08-01" max="2085-12-31" onchange="validateExportDates()" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-emerald-500 font-medium" />
        </div>
        <div>
          <label class="block font-semibold text-slate-600 mb-0.5">End Date</label>
          <input type="date" id="export-end-date" min="2026-08-01" max="2085-12-31" onchange="validateExportDates()" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-emerald-500 font-medium" />
        </div>
      </div>
      <div class="flex space-x-2 pt-2 border-t border-slate-100">
        <button type="button" onclick="closeExportModal()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 rounded-xl text-[11px] transition">
          Cancel
        </button>
        <button type="button" onclick="processExport()" class="w-1/2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2 rounded-xl shadow-sm transition text-[11px] flex items-center justify-center gap-1">
          <i class="fa-solid fa-download text-[10px]"></i> Download
        </button>
      </div>
    </div>
  </div>

  <div id="excel-comment-box" onclick="event.stopPropagation()" class="excel-comment-box hidden absolute z-50 bg-slate-900 text-white text-[11px] rounded-2xl p-3 shadow-2xl border border-slate-800 space-y-2 w-64 transition-all duration-150">
    <div class="font-bold text-blue-400 border-b border-slate-800 pb-1.5 flex justify-between items-center text-[10px]">
      <span class="flex items-center gap-1.5">
        <i class="fa-solid fa-comment-dots text-blue-400"></i>
        <span id="comm-date-header">Date Overview</span>
      </span>
      <button onclick="closeCommentBox()" class="text-slate-400 hover:text-white px-1 py-0.5 rounded text-[10px]">
        <i class="fa-solid fa-xmark"></i>
      </button>
    </div>
    <div id="comm-booking-list" class="space-y-1.5 max-h-56 overflow-y-auto pr-0.5"></div>
  </div>

  <header class="bg-white/80 backdrop-blur-md border-b border-slate-200/60 text-slate-900 sticky top-0 z-40 no-print">
    <div class="max-w-7xl mx-auto px-4 py-2.5 flex flex-col md:flex-row justify-between items-center gap-2.5">
      <div class="flex items-center space-x-2.5">
        <div class="bg-blue-600 p-2 rounded-2xl text-white shadow-sm">
          <i class="fa-solid fa-hotel text-sm"></i>
        </div>
        <div>
          <p class="text-[10px] text-slate-500 mt-0.5">Management &amp; Booking Control System</p>
        </div>
      </div>
      
      <nav class="flex space-x-1 bg-slate-100 p-1 rounded-full text-[11px] font-medium">
        <button onclick="switchTab('dashboard')" id="btn-dashboard" class="tab-btn px-3 py-1 rounded-full transition-all active-tab bg-white text-blue-600 shadow-sm font-bold">Dashboard</button>
        <button onclick="switchTab('booking')" id="btn-booking" class="tab-btn px-3 py-1 rounded-full transition-all text-slate-600 hover:text-slate-900">Booking Details</button>
        <button onclick="switchTab('master')" id="btn-master" class="tab-btn px-3 py-1 rounded-full transition-all text-slate-600 hover:text-slate-900 flex items-center gap-1">
          <i class="fa-solid fa-lock text-[9px] text-amber-500"></i> Master Data
        </button>
        <button onclick="switchTab('calendar')" id="btn-calendar" class="tab-btn px-3 py-1 rounded-full transition-all text-slate-600 hover:text-slate-900">Calendar</button>
      </nav>

      <div class="flex items-center space-x-1.5">
        <button onclick="openAlertModal()" title="View Alerts" class="relative bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-200 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-bell text-[10px]"></i> Alerts
          <span id="alert-badge" class="hidden absolute -top-1 -right-1 bg-rose-600 text-white text-[9px] font-black px-1.5 py-0.2 rounded-full border border-white animate-bounce">0</span>
        </button>
        <button onclick="saveChanges()" class="bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-brands fa-google text-[10px]"></i> Save
        </button>
        
        <div id="wipe-layer-1-modal" class="hidden fixed inset-0 z-[60] bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 no-print">
          <div class="bg-white rounded-3xl shadow-2xl border border-rose-100 max-w-sm w-full p-6 text-center space-y-4">
            <div class="bg-rose-100 text-rose-600 w-16 h-16 rounded-full flex items-center justify-center mx-auto text-3xl shadow-sm">
              <i class="fa-solid fa-triangle-exclamation"></i>
            </div>
            <div>
              <h3 class="text-base font-black text-slate-900">Initiate Data Wipe?</h3>
              <p class="text-xs text-slate-600 mt-2">You are about to delete ALL data from the Google Sheet. This affects bookings, rooms, and agents.</p>
            </div>
            <div class="flex space-x-3 pt-2">
              <button type="button" onclick="closeWipeModals()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold py-2.5 rounded-xl text-xs transition">Cancel</button>
              <button type="button" onclick="proceedToWipeLayer2()" class="w-1/2 bg-rose-600 hover:bg-rose-700 text-white font-bold py-2.5 rounded-xl shadow-sm transition text-xs">Proceed</button>
            </div>
          </div>
        </div>

        <div id="wipe-layer-2-modal" class="hidden fixed inset-0 z-[70] bg-rose-900/80 backdrop-blur-md flex items-center justify-center p-4 no-print">
          <div class="bg-black rounded-3xl shadow-2xl border border-rose-600 max-w-sm w-full p-6 text-center space-y-4">
            <div class="text-rose-500 w-16 h-16 rounded-full flex items-center justify-center mx-auto text-4xl animate-pulse">
              <i class="fa-solid fa-skull-crossbones"></i>
            </div>
            <div>
              <h3 class="text-lg font-black text-white uppercase tracking-widest">Final Warning</h3>
              <p class="text-xs text-rose-200 mt-2">This action is <strong class="text-white">IRREVERSIBLE</strong>. All records will be permanently deleted from the database. Are you absolutely sure?</p>
            </div>
            <div class="flex space-x-3 pt-2">
              <button type="button" onclick="closeWipeModals()" class="w-1/2 bg-slate-800 hover:bg-slate-700 text-white font-bold py-2.5 rounded-xl text-xs transition">Cancel</button>
              <button type="button" id="btn-final-wipe" onclick="executeGoogleSheetWipe()" class="w-1/2 bg-rose-600 hover:bg-rose-700 text-white font-bold py-2.5 rounded-xl shadow-lg shadow-rose-900/50 transition text-xs">ERASE ALL DATA</button>
            </div>
          </div>
        </div>

        <button onclick="requestDataWipe()" class="bg-rose-600 hover:bg-rose-700 text-white border border-rose-800 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition shadow-sm">
          <i class="fa-solid fa-skull-crossbones text-[8px]"></i> Wipe Data
        </button>

        <button onclick="openExportModal()" class="bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-file-excel text-[10px]"></i> Export Excel
        </button>
        <button onclick="logout()" class="bg-slate-100 hover:bg-slate-200 text-slate-700 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-right-from-bracket text-[10px]"></i> Logout
        </button>
      </div>
    </div>
  </header>

  <div id="notification-banner" class="hidden bg-emerald-600 text-white text-center py-2 px-4 text-xs font-semibold shadow-md relative no-print transition-all">
    <span id="notification-text">Action completed successfully!</span>
    <button onclick="closeNotification()" class="absolute right-4 top-2 text-white/80 hover:text-white"><i class="fa-solid fa-xmark"></i></button>
  </div>

  <main class="flex-grow max-w-7xl w-full mx-auto px-4 py-4 space-y-4">

    <section id="tab-dashboard" class="tab-content space-y-4">
      <div class="grid grid-cols-1 md:grid-cols-4 gap-3">
        <div class="bg-white rounded-3xl p-3.5 border border-slate-100 shadow-sm flex items-center space-x-3">
          <div class="bg-blue-50 text-blue-600 p-2.5 rounded-2xl text-base">
            <i class="fa-solid fa-bed"></i>
          </div>
          <div>
            <p class="text-[10px] text-slate-400 font-semibold uppercase tracking-wider">Total Rooms</p>
            <h4 id="dash-total-rooms" class="text-base font-bold text-slate-800">0</h4>
          </div>
        </div>
        <div class="bg-white rounded-3xl p-3.5 border border-slate-100 shadow-sm flex items-center space-x-3">
          <div class="bg-emerald-50 text-emerald-600 p-2.5 rounded-2xl text-base">
            <i class="fa-solid fa-calendar-check"></i>
          </div>
          <div>
            <p class="text-[10px] text-slate-400 font-semibold uppercase tracking-wider">Active Bookings</p>
            <h4 id="dash-active-bookings" class="text-base font-bold text-slate-800">0</h4>
          </div>
        </div>
        <div class="bg-white rounded-3xl p-3.5 border border-slate-100 shadow-sm flex items-center space-x-3">
          <div class="bg-amber-50 text-amber-600 p-2.5 rounded-2xl text-base">
            <i class="fa-solid fa-users"></i>
          </div>
          <div>
            <p class="text-[10px] text-slate-400 font-semibold uppercase tracking-wider">Total Agents</p>
            <h4 id="dash-total-agents" class="text-base font-bold text-slate-800">0</h4>
          </div>
        </div>
        <div class="bg-white rounded-3xl p-3.5 border border-slate-100 shadow-sm flex items-center space-x-3">
          <div class="bg-indigo-50 text-indigo-600 p-2.5 rounded-2xl text-base">
            <i class="fa-solid fa-indian-rupee-sign"></i>
          </div>
          <div>
            <p class="text-[10px] text-slate-400 font-semibold uppercase tracking-wider">Total Revenue</p>
            <h4 id="dash-total-revenue" class="text-base font-bold text-slate-800">₹0</h4>
          </div>
        </div>
      </div>

      <div class="bg-white rounded-3xl p-4 border border-slate-100 shadow-sm space-y-3">
        <div class="flex justify-between items-center">
          <h3 class="text-xs font-bold text-slate-800 flex items-center gap-1.5">
            <i class="fa-solid fa-grid-2 text-blue-600"></i> Room Occupancy Status
          </h3>
          <div class="flex items-center gap-3 text-[10px] font-medium">
            <span class="flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-emerald-500"></span> Available</span>
            <span class="flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-rose-500"></span> Booked</span>
          </div>
        </div>
        <div id="dashboard-room-grid" class="grid grid-cols-2 sm:grid-cols-4 md:grid-cols-6 lg:grid-cols-8 gap-2">
          </div>
      </div>
    </section>

    <section id="tab-booking" class="tab-content hidden space-y-4">
      <div class="bg-white rounded-3xl p-4 border border-slate-100 shadow-sm space-y-3">
        <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-2">
          <div class="flex items-center gap-2">
            <h3 class="text-xs font-bold text-slate-800">Bookings List</h3>
            <span id="booking-count-badge" class="bg-blue-50 text-blue-600 px-2 py-0.5 rounded-full text-[10px] font-bold">0</span>
          </div>
          <div class="flex flex-wrap items-center gap-2 w-full md:w-auto">
            <input type="text" id="booking-search" placeholder="Search by name, room, agent..." oninput="renderBookingsTable()" class="bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl px-3 py-1.5 focus:outline-none focus:bg-white text-xs w-full md:w-48 transition" />
            <button onclick="openBookingModal()" class="bg-blue-600 hover:bg-blue-700 active:scale-98 text-white font-semibold px-3 py-1.5 rounded-2xl text-xs transition shadow-sm flex items-center gap-1">
              <i class="fa-solid fa-plus text-[10px]"></i> New Booking
            </button>
          </div>
        </div>

        <div class="overflow-x-auto rounded-2xl border border-slate-100">
          <table class="w-full text-left border-collapse text-[11px]">
            <thead>
              <tr class="bg-slate-50 text-slate-500 font-semibold border-b border-slate-100 uppercase tracking-wider">
                <th class="py-2.5 px-3">Booking ID</th>
                <th class="py-2.5 px-3">Guest Name</th>
                <th class="py-2.5 px-3">Contact No</th>
                <th class="py-2.5 px-3">Room(s)</th>
                <th class="py-2.5 px-3">Check In</th>
                <th class="py-2.5 px-3">Check Out</th>
                <th class="py-2.5 px-3">Agent</th>
                <th class="py-2.5 px-3 text-right">Total</th>
                <th class="py-2.5 px-3 text-right">Advance</th>
                <th class="py-2.5 px-3 text-right">Due</th>
                <th class="py-2.5 px-3 text-center">Actions</th>
              </tr>
            </thead>
            <tbody id="bookings-table-body" class="divide-y divide-slate-100">
              </tbody>
          </table>
        </div>
      </div>
    </section>

    <section id="tab-master" class="tab-content hidden space-y-4">
      <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div class="bg-white rounded-3xl p-4 border border-slate-100 shadow-sm space-y-3">
          <div class="flex justify-between items-center">
            <h3 class="text-xs font-bold text-slate-800 flex items-center gap-1.5">
              <i class="fa-solid fa-door-open text-blue-600"></i> Room Master Configuration
            </h3>
          </div>
          <form onsubmit="handleAddRoom(event)" class="flex gap-2">
            <input type="text" id="master-room-no" placeholder="Room No (e.g. 101)" required="" class="bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl px-3 py-1.5 focus:outline-none focus:bg-white text-xs flex-grow transition" />
            <input type="number" id="master-room-tariff" placeholder="Tariff (₹)" required="" class="bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl px-3 py-1.5 focus:outline-none focus:bg-white text-xs w-28 transition" />
            <button type="submit" class="bg-blue-600 hover:bg-blue-700 text-white font-semibold px-3 py-1.5 rounded-2xl text-xs transition shadow-sm">Add</button>
          </form>

          <div class="overflow-x-auto rounded-2xl border border-slate-100">
            <table class="w-full text-left text-[11px]">
              <thead>
                <tr class="bg-slate-50 text-slate-500 font-semibold border-b border-slate-100">
                  <th class="py-2 px-3">Room No</th>
                  <th class="py-2 px-3">Tariff (₹)</th>
                  <th class="py-2 px-3 text-center">Action</th>
                </tr>
              </thead>
              <tbody id="master-rooms-table" class="divide-y divide-slate-100">
                </tbody>
            </table>
          </div>
        </div>

        <div class="bg-white rounded-3xl p-4 border border-slate-100 shadow-sm space-y-3">
          <div class="flex justify-between items-center">
            <h3 class="text-xs font-bold text-slate-800 flex items-center gap-1.5">
              <i class="fa-solid fa-user-tie text-amber-600"></i> Agent Master Configuration
            </h3>
          </div>
          <form onsubmit="handleAddAgent(event)" class="flex gap-2">
            <input type="text" id="master-agent-name" placeholder="Agent Name" required="" class="bg-slate-100 border border-transparent focus:border-amber-500 rounded-2xl px-3 py-1.5 focus:outline-none focus:bg-white text-xs flex-grow transition" />
            <button type="submit" class="bg-amber-600 hover:bg-amber-700 text-white font-semibold px-3 py-1.5 rounded-2xl text-xs transition shadow-sm">Add</button>
          </form>

          <div class="overflow-x-auto rounded-2xl border border-slate-100">
            <table class="w-full text-left text-[11px]">
              <thead>
                <tr class="bg-slate-50 text-slate-500 font-semibold border-b border-slate-100">
                  <th class="py-2 px-3">Agent Name</th>
                  <th class="py-2 px-3 text-center">Action</th>
                </tr>
              </thead>
              <tbody id="master-agents-table" class="divide-y divide-slate-100">
                </tbody>
            </table>
          </div>
        </div>
      </div>
    </section>

    <section id="tab-calendar" class="tab-content hidden space-y-4">
      <div class="bg-white rounded-3xl p-4 border border-slate-100 shadow-sm space-y-3">
        <div class="flex flex-col sm:flex-row justify-between items-center gap-2">
          <h3 class="text-xs font-bold text-slate-800 flex items-center gap-1.5">
            <i class="fa-solid fa-calendar-days text-blue-600"></i> Monthly Booking Overview
          </h3>
          <div class="flex items-center space-x-2">
            <button onclick="changeCalMonth(-1)" class="bg-slate-100 hover:bg-slate-200 text-slate-700 px-2.5 py-1 rounded-full text-xs transition"><i class="fa-solid fa-chevron-left"></i></button>
            <span id="cal-month-year-title" class="text-xs font-bold text-slate-800 w-32 text-center">August 2026</span>
            <button onclick="changeCalMonth(1)" class="bg-slate-100 hover:bg-slate-200 text-slate-700 px-2.5 py-1 rounded-full text-xs transition"><i class="fa-solid fa-chevron-right"></i></button>
          </div>
        </div>

        <div class="grid grid-cols-7 gap-1 text-center font-bold text-slate-400 text-[10px] uppercase py-1 border-b border-slate-100">
          <div>Sun</div><div>Mon</div><div>Tue</div><div>Wed</div><div>Thu</div><div>Fri</div><div>Sat</div>
        </div>
        <div id="calendar-grid-days" class="grid grid-cols-7 gap-1 text-center text-xs">
          </div>
      </div>
    </section>

  </main>

  <div id="booking-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-3 overflow-y-auto no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-3xl w-full p-5 space-y-3 my-4 max-h-[90vh] overflow-y-auto">
      <div class="flex justify-between items-center pb-2 border-b border-slate-100">
        <h3 id="booking-modal-title" class="text-xs font-bold text-slate-900">New Booking Entry</h3>
        <button onclick="closeBookingModal()" class="text-slate-400 hover:text-slate-600 text-base p-0.5"><i class="fa-solid fa-xmark"></i></button>
      </div>

      <form onsubmit="handleSaveBooking(event)" class="space-y-3">
        <input type="hidden" id="modal-booking-id" />

        <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
          <div>
            <label class="block text-[10px] font-semibold text-slate-700 mb-0.5">Guest Full Name *</label>
            <input type="text" id="modal-guest-name" required="" placeholder="Enter Guest Name" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl px-3 py-1.5 focus:outline-none focus:bg-white text-xs transition" />
          </div>

          <div>
            <label class="block text-[10px] font-semibold text-slate-700 mb-0.5">Contact Number *</label>
            <input type="tel" id="modal-guest-phone" required="" placeholder="Enter Mobile Number" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl px-3 py-1.5 focus:outline-none focus:bg-white text-xs transition" />
          </div>

          <div>
            <label class="block text-[10px] font-semibold text-slate-700 mb-0.5">Check-In Date &amp; Time *</label>
            <input type="datetime-local" id="modal-checkin" required="" onchange="calculateTariffAndNights()" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl px-3 py-1.5 focus:outline-none focus:bg-white text-xs transition" />
          </div>

          <div>
            <label class="block text-[10px] font-semibold text-slate-700 mb-0.5">Check-Out Date &amp; Time *</label>
            <input type="datetime-local" id="modal-checkout" required="" onchange="calculateTariffAndNights()" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl px-3 py-1.5 focus:outline-none focus:bg-white text-xs transition" />
          </div>
        </div>

        <div>
          <label class="block text-[10px] font-semibold text-slate-700 mb-1">Select Room(s) *</label>
          <div id="modal-room-checkboxes" class="grid grid-cols-2 sm:grid-cols-4 gap-2 max-h-36 overflow-y-auto bg-slate-50 p-2.5 rounded-2xl border border-slate-100">
            </div>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-3">
          <div>
            <label class="block text-[10px] font-semibold text-slate-700 mb-0.5">Agent Assignment</label>
            <select id="modal-agent" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl px-3 py-1.5 focus:outline-none focus:bg-white text-xs transition">
              <option value="Direct">Direct</option>
            </select>
          </div>

          <div>
            <label class="block text-[10px] font-semibold text-slate-700 mb-0.5">Total Amount (₹) *</label>
            <input type="number" id="modal-total-amount" required="" oninput="calculateDueAmount()" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl px-3 py-1.5 focus:outline-none focus:bg-white text-xs transition font-bold text-slate-800" />
          </div>

          <div>
            <label class="block text-[10px] font-semibold text-slate-700 mb-0.5">Advance Paid (₹)</label>
            <input type="number" id="modal-advance-paid" value="0" oninput="calculateDueAmount()" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl px-3 py-1.5 focus:outline-none focus:bg-white text-xs transition font-semibold text-emerald-600" />
          </div>
        </div>

        <div>
          <label class="block text-[10px] font-semibold text-slate-700 mb-0.5">Balance Due (₹)</label>
          <input type="number" id="modal-total-due" readonly="" value="0" class="w-full bg-slate-200/60 border border-transparent rounded-2xl px-3 py-1.5 text-xs font-bold text-rose-600 cursor-not-allowed" />
        </div>

        <div class="flex space-x-2 pt-2 border-t border-slate-100">
          <button type="button" onclick="closeBookingModal()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 rounded-2xl text-xs transition">
            Cancel
          </button>
          <button type="submit" class="w-1/2 bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 rounded-2xl shadow-sm transition text-xs flex items-center justify-center gap-1">
            <i class="fa-solid fa-floppy-disk"></i> Save Booking
          </button>
        </div>
      </form>
    </div>
  </div>

  <div id="invoice-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 overflow-y-auto no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-md w-full p-5 space-y-4 my-4">
      <div id="printable-invoice" class="bg-white p-4 rounded-2xl border border-slate-200 space-y-3">
        <div class="text-center border-b border-slate-100 pb-3">
          <h2 class="text-sm font-black text-slate-900 tracking-wider uppercase">Booking Voucher</h2>
          <p class="text-[9px] text-slate-400">Official Stay Confirmation</p>
        </div>

        <div class="space-y-1.5 text-[11px]">
          <div class="flex justify-between"><span class="text-slate-500">Booking ID:</span> <span id="inv-id" class="font-bold text-slate-800"></span></div>
          <div class="flex justify-between"><span class="text-slate-500">Guest Name:</span> <span id="inv-name" class="font-bold text-slate-800"></span></div>
          <div class="flex justify-between"><span class="text-slate-500">Contact:</span> <span id="inv-phone" class="font-semibold text-slate-700"></span></div>
          <div class="flex justify-between"><span class="text-slate-500">Room(s):</span> <span id="inv-rooms" class="font-bold text-blue-600"></span></div>
          <div class="flex justify-between"><span class="text-slate-500">Check In:</span> <span id="inv-checkin" class="font-medium text-slate-700"></span></div>
          <div class="flex justify-between"><span class="text-slate-500">Check Out:</span> <span id="inv-checkout" class="font-medium text-slate-700"></span></div>
          <div class="flex justify-between"><span class="text-slate-500">Agent:</span> <span id="inv-agent" class="font-medium text-slate-700"></span></div>
        </div>

        <div class="border-t border-slate-100 pt-2 space-y-1 text-[11px]">
          <div class="flex justify-between"><span class="text-slate-500">Total Bill:</span> <span id="inv-total" class="font-bold text-slate-800">₹0</span></div>
          <div class="flex justify-between"><span class="text-slate-500">Advance Paid:</span> <span id="inv-advance" class="font-semibold text-emerald-600">₹0</span></div>
          <div class="flex justify-between text-xs font-bold pt-1 border-t border-slate-100"><span class="text-slate-800">Balance Due:</span> <span id="inv-due" class="text-rose-600">₹0</span></div>
        </div>
      </div>

      <div class="flex space-x-2 no-print">
        <button onclick="closeInvoiceModal()" class="w-1/3 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 rounded-2xl text-xs transition">Close</button>
        <button onclick="window.print()" class="w-1/3 bg-slate-800 hover:bg-slate-900 text-white font-semibold py-2 rounded-2xl text-xs transition flex items-center justify-center gap-1">
          <i class="fa-solid fa-print"></i> Print
        </button>
        <button onclick="downloadReceiptAsJPEG()" class="w-1/3 bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 rounded-2xl shadow-sm transition text-xs flex items-center justify-center gap-1">
          <i class="fa-solid fa-file-image"></i> JPEG
        </button>
      </div>
    </div>
  </div>

  <div id="alert-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-md w-full p-5 space-y-3">
      <div class="flex justify-between items-center pb-2 border-b border-slate-100">
        <h3 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
          <i class="fa-solid fa-bell text-amber-500"></i> Active Notifications & Alerts
        </h3>
        <button onclick="closeAlertModal()" class="text-slate-400 hover:text-slate-600 text-base"><i class="fa-solid fa-xmark"></i></button>
      </div>

      <div id="alert-list-container" class="space-y-2 max-h-64 overflow-y-auto text-xs">
        </div>
    </div>
  </div>

  <script>
    // System Configurations & Endpoint
    const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbz_XXXXXXXXXXXXXX/exec"; // Replace with deployed Apps Script URL if needed
    
    let state = {
      rooms: [],
      agents: [],
      bookings: []
    };

    let calendarCurrentDate = new Date();
    let masterAuthenticated = false;
    let masterDeleteTarget = null;
    let inactivityTimer = null;
    let countdownInterval = null;

    // Default System Load
    document.addEventListener("DOMContentLoaded", () => {
      loadStateFromLocalStorage();
      checkAuthStatus();
      resetInactivityTimer();
    });

    // Handle Authentication Logic
    function handleLogin(e) {
      e.preventDefault();
      const u = document.getElementById('login-userid').value.trim();
      const p = document.getElementById('login-password').value.trim();

      if ((u === "admin" && p === "admin123") || (u === "user" && p === "user123")) {
        localStorage.setItem("portal_logged_in", "true");
        document.getElementById('login-overlay').classList.add('hidden');
        document.getElementById('login-alert-modal').classList.remove('hidden');
        fetchGoogleSheetData();
      } else {
        document.getElementById('login-error').classList.remove('hidden');
      }
    }

    function checkAuthStatus() {
      if (localStorage.getItem("portal_logged_in") === "true") {
        document.getElementById('login-overlay').classList.add('hidden');
      } else {
        document.getElementById('login-overlay').classList.remove('hidden');
      }
    }

    function closeLoginAlertModal() {
      document.getElementById('login-alert-modal').classList.add('hidden');
    }

    function logout() {
      document.getElementById('logout-confirm-modal').classList.remove('hidden');
    }

    function cancelLogout() {
      document.getElementById('logout-confirm-modal').classList.add('hidden');
    }

    function processLogoutWithSave() {
      document.getElementById('logout-confirm-modal').classList.add('hidden');
      document.getElementById('saving-lock-modal').classList.remove('hidden');

      saveChangesToGoogleSheet().finally(() => {
        setTimeout(() => {
          localStorage.removeItem("portal_logged_in");
          masterAuthenticated = false;
          document.getElementById('saving-lock-modal').classList.add('hidden');
          location.reload();
        }, 1200);
      });
    }

    // Tab Navigation Logic
    function switchTab(tabId) {
      if (tabId === 'master' && !masterAuthenticated) {
        document.getElementById('master-auth-modal').classList.remove('hidden');
        return;
      }

      document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
      document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('bg-white', 'text-blue-600', 'shadow-sm', 'font-bold');
        btn.classList.add('text-slate-600');
      });

      document.getElementById(`tab-${tabId}`).classList.remove('hidden');
      const activeBtn = document.getElementById(`btn-${tabId}`);
      activeBtn.classList.add('bg-white', 'text-blue-600', 'shadow-sm', 'font-bold');
      activeBtn.classList.remove('text-slate-600');

      if (tabId === 'dashboard') renderDashboard();
      if (tabId === 'booking') renderBookingsTable();
      if (tabId === 'master') renderMasterTables();
      if (tabId === 'calendar') renderCalendar();
    }

    // Master Authentication Security
    function handleMasterAuth(e) {
      e.preventDefault();
      const pwd = document.getElementById('master-password-input').value.trim();
      if (pwd === "master123") { // Default Master Password
        masterAuthenticated = true;
        document.getElementById('master-auth-modal').classList.add('hidden');
        document.getElementById('master-password-input').value = '';
        document.getElementById('master-auth-error').classList.add('hidden');
        switchTab('master');
      } else {
        document.getElementById('master-auth-error').classList.remove('hidden');
      }
    }

    function closeMasterAuthModal() {
      document.getElementById('master-auth-modal').classList.add('hidden');
    }

    // Local Storage Management
    function loadStateFromLocalStorage() {
      const stored = localStorage.getItem('portal_state');
      if (stored) {
        try {
          state = JSON.parse(stored);
        } catch (e) {
          console.error("Local storage load error", e);
        }
      }
    }

    function saveStateToLocalStorage() {
      localStorage.setItem('portal_state', JSON.stringify(state));
    }

    // Render Dashboard Overview
    function renderDashboard() {
      document.getElementById('dash-total-rooms').innerText = state.rooms.length;
      document.getElementById('dash-active-bookings').innerText = state.bookings.length;
      document.getElementById('dash-total-agents').innerText = state.agents.length;

      const totalRev = state.bookings.reduce((sum, b) => sum + (parseFloat(b.totalAmount) || 0), 0);
      document.getElementById('dash-total-revenue').innerText = `₹${totalRev.toLocaleString('en-IN')}`;

      // Room Grid Matrix Rendering
      const grid = document.getElementById('dashboard-room-grid');
      grid.innerHTML = '';

      const now = new Date();

      state.rooms.forEach(room => {
        // Check if room is currently booked based on check-in and check-out
        const isOccupied = state.bookings.some(b => {
          const roomList = normalizeRooms(b.roomNumber || b.roomsList);
          const cIn = new Date(b.checkIn);
          const cOut = new Date(b.checkOut);
          return roomList.includes(String(room.roomNo)) && (now >= cIn && now <= cOut);
        });

        const card = document.createElement('div');
        card.className = `p-3 rounded-2xl border flex flex-col justify-between space-y-1 ${
          isOccupied ? 'bg-rose-50 border-rose-200 text-rose-800' : 'bg-emerald-50 border-emerald-200 text-emerald-800'
        }`;
        card.innerHTML = `
          <div class="flex justify-between items-center font-bold text-xs">
            <span>Room ${room.roomNo}</span>
            <span class="w-2 h-2 rounded-full ${isOccupied ? 'bg-rose-500' : 'bg-emerald-500'}"></span>
          </div>
          <div class="text-[10px] opacity-75">Tariff: ₹${room.tariff}</div>
          <div class="text-[9px] font-bold uppercase tracking-wider ${isOccupied ? 'text-rose-600' : 'text-emerald-600'}">
            ${isOccupied ? 'Occupied' : 'Vacant'}
          </div>
        `;
        grid.appendChild(card);
      });
    }

    // Normalize room numbers to array regardless of input type (FIXED MULTI-ROOM BUG)
    function normalizeRooms(roomData) {
      if (!roomData) return [];
      if (Array.isArray(roomData)) {
        return roomData.map(r => String(r).trim()).filter(Boolean);
      }
      if (typeof roomData === 'string') {
        return roomData.split(',').map(r => r.trim()).filter(Boolean);
      }
      return [String(roomData)];
    }

    // Render Bookings Data Table
    function renderBookingsTable() {
      const tbody = document.getElementById('bookings-table-body');
      tbody.innerHTML = '';

      const q = document.getElementById('booking-search').value.toLowerCase();
      const filtered = state.bookings.filter(b => {
        const roomsStr = normalizeRooms(b.roomNumber || b.roomsList).join(', ');
        return (
          (b.guestName && b.guestName.toLowerCase().includes(q)) ||
          (b.id && b.id.toLowerCase().includes(q)) ||
          (b.phone && b.phone.includes(q)) ||
          (roomsStr.toLowerCase().includes(q)) ||
          (b.agent && b.agent.toLowerCase().includes(q))
        );
      });

      document.getElementById('booking-count-badge').innerText = filtered.length;

      filtered.forEach(b => {
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50 transition border-b border-slate-100";
        
        const roomsDisplay = normalizeRooms(b.roomNumber || b.roomsList).join(', ');

        tr.innerHTML = `
          <td class="py-2.5 px-3 font-mono font-bold text-blue-600">${b.id || ''}</td>
          <td class="py-2.5 px-3 font-bold text-slate-800">${b.guestName || ''}</td>
          <td class="py-2.5 px-3 text-slate-600">${b.phone || ''}</td>
          <td class="py-2.5 px-3 font-semibold text-slate-700">${roomsDisplay}</td>
          <td class="py-2.5 px-3 text-slate-600">${formatDateTime(b.checkIn)}</td>
          <td class="py-2.5 px-3 text-slate-600">${formatDateTime(b.checkOut)}</td>
          <td class="py-2.5 px-3 text-slate-600">${b.agent || 'Direct'}</td>
          <td class="py-2.5 px-3 text-right font-semibold">₹${b.totalAmount || 0}</td>
          <td class="py-2.5 px-3 text-right font-semibold text-emerald-600">₹${b.advancePaid || 0}</td>
          <td class="py-2.5 px-3 text-right font-bold text-rose-600">₹${b.totalDue || 0}</td>
          <td class="py-2.5 px-3 text-center space-x-1">
            <button onclick="editBooking('${b.id}')" title="Edit" class="bg-blue-50 text-blue-600 hover:bg-blue-100 p-1.5 rounded-xl transition"><i class="fa-solid fa-pen-to-square"></i></button>
            <button onclick="openInvoiceModal('${b.id}')" title="Receipt" class="bg-slate-100 text-slate-700 hover:bg-slate-200 p-1.5 rounded-xl transition"><i class="fa-solid fa-receipt"></i></button>
            <button onclick="deleteBooking('${b.id}')" title="Delete" class="bg-rose-50 text-rose-600 hover:bg-rose-100 p-1.5 rounded-xl transition"><i class="fa-solid fa-trash-can"></i></button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    // Modal Booking Handlers
    function openBookingModal(editId = null) {
      document.getElementById('modal-booking-id').value = editId || '';
      document.getElementById('booking-modal-title').innerText = editId ? 'Edit Booking Record' : 'New Booking Entry';

      // Load Room Checkboxes
      const roomContainer = document.getElementById('modal-room-checkboxes');
      roomContainer.innerHTML = '';
      
      let selectedRooms = [];
      if (editId) {
        const b = state.bookings.find(item => item.id === editId);
        if (b) {
          selectedRooms = normalizeRooms(b.roomNumber || b.roomsList);
          document.getElementById('modal-guest-name').value = b.guestName || '';
          document.getElementById('modal-guest-phone').value = b.phone || '';
          document.getElementById('modal-checkin').value = b.checkIn || '';
          document.getElementById('modal-checkout').value = b.checkOut || '';
          document.getElementById('modal-total-amount').value = b.totalAmount || 0;
          document.getElementById('modal-advance-paid').value = b.advancePaid || 0;
          document.getElementById('modal-total-due').value = b.totalDue || 0;
        }
      } else {
        document.getElementById('modal-guest-name').value = '';
        document.getElementById('modal-guest-phone').value = '';
        document.getElementById('modal-checkin').value = '';
        document.getElementById('modal-checkout').value = '';
        document.getElementById('modal-total-amount').value = '';
        document.getElementById('modal-advance-paid').value = '0';
        document.getElementById('modal-total-due').value = '0';
      }

      state.rooms.forEach(r => {
        const isChecked = selectedRooms.includes(String(r.roomNo)) ? 'checked' : '';
        const div = document.createElement('div');
        div.className = "flex items-center gap-2 bg-white p-1.5 rounded-xl border border-slate-200 text-xs";
        div.innerHTML = `
          <input type="checkbox" value="${r.roomNo}" ${isChecked} onchange="calculateTariffAndNights()" class="room-checkbox rounded border-slate-300 text-blue-600 focus:ring-blue-500" />
          <span class="font-medium text-slate-700">Room ${r.roomNo}</span>
        `;
        roomContainer.appendChild(div);
      });

      // Load Agent Dropdown Options
      const agentSelect = document.getElementById('modal-agent');
      agentSelect.innerHTML = '<option value="Direct">Direct</option>';
      state.agents.forEach(a => {
        const opt = document.createElement('option');
        opt.value = a.name;
        opt.innerText = a.name;
        if (editId) {
          const b = state.bookings.find(item => item.id === editId);
          if (b && b.agent === a.name) opt.selected = true;
        }
        agentSelect.appendChild(opt);
      });

      document.getElementById('booking-modal').classList.remove('hidden');
    }

    function closeBookingModal() {
      document.getElementById('booking-modal').classList.add('hidden');
    }

    // Calculated fields based on Tariff & Room Count
    function calculateTariffAndNights() {
      const selectedBoxes = Array.from(document.querySelectorAll('.room-checkbox:checked')).map(cb => cb.value);
      const cIn = new Date(document.getElementById('modal-checkin').value);
      const cOut = new Date(document.getElementById('modal-checkout').value);

      if (selectedBoxes.length > 0 && !isNaN(cIn) && !isNaN(cOut) && cOut > cIn) {
        const diffTime = Math.abs(cOut - cIn);
        const nights = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) || 1;

        let dayTariffSum = 0;
        selectedBoxes.forEach(rNo => {
          const rm = state.rooms.find(r => String(r.roomNo) === String(rNo));
          if (rm) dayTariffSum += parseFloat(rm.tariff) || 0;
        });

        const autoTotal = dayTariffSum * nights;
        document.getElementById('modal-total-amount').value = autoTotal;
        calculateDueAmount();
      }
    }

    function calculateDueAmount() {
      const tot = parseFloat(document.getElementById('modal-total-amount').value) || 0;
      const adv = parseFloat(document.getElementById('modal-advance-paid').value) || 0;
      document.getElementById('modal-total-due').value = Math.max(0, tot - adv);
    }

    // SAVE BOOKING ENTRY (FIXED MULTI-ROOM ERASURE BUG)
    function handleSaveBooking(e) {
      e.preventDefault();

      // Gather Checked Rooms safely as an Array and String
      const selectedRoomList = Array.from(document.querySelectorAll('.room-checkbox:checked')).map(cb => cb.value.trim());
      if (selectedRoomList.length === 0) {
        alert("Please select at least one room!");
        return;
      }

      const formattedRoomString = selectedRoomList.join(', ');
      const editId = document.getElementById('modal-booking-id').value;
      const tot = parseFloat(document.getElementById('modal-total-amount').value) || 0;
      const adv = parseFloat(document.getElementById('modal-advance-paid').value) || 0;

      const bookingObj = {
        id: editId || 'BK-' + Date.now(),
        guestName: document.getElementById('modal-guest-name').value.trim(),
        phone: document.getElementById('modal-guest-phone').value.trim(),
        checkIn: document.getElementById('modal-checkin').value,
        checkOut: document.getElementById('modal-checkout').value,
        roomNumber: formattedRoomString, // Stored consistently as comma-separated string
        roomsList: selectedRoomList,     // Stored as normalized array for calculations
        agent: document.getElementById('modal-agent').value,
        totalAmount: tot,
        advancePaid: adv,
        totalDue: Math.max(0, tot - adv)
      };

      if (editId) {
        const idx = state.bookings.findIndex(b => b.id === editId);
        if (idx > -1) state.bookings[idx] = bookingObj;
      } else {
        state.bookings.push(bookingObj);
      }

      saveStateToLocalStorage();
      closeBookingModal();
      renderBookingsTable();
      renderDashboard();
      showNotification("Booking saved successfully!");
    }

    function editBooking(id) {
      openBookingModal(id);
    }

    function deleteBooking(id) {
      if (confirm("Are you sure you want to delete this booking entry?")) {
        state.bookings = state.bookings.filter(b => b.id !== id);
        saveStateToLocalStorage();
        renderBookingsTable();
        renderDashboard();
        showNotification("Booking entry deleted.");
      }
    }

    // Master Room/Agent Handlers
    function renderMasterTables() {
      // Room Table
      const roomTbody = document.getElementById('master-rooms-table');
      roomTbody.innerHTML = '';
      state.rooms.forEach(r => {
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50 border-b border-slate-100";
        tr.innerHTML = `
          <td class="py-2 px-3 font-bold text-slate-800">Room ${r.roomNo}</td>
          <td class="py-2 px-3 text-slate-600">₹${r.tariff}</td>
          <td class="py-2 px-3 text-center">
            <button onclick="requestMasterDelete('room', '${r.roomNo}')" class="text-rose-600 hover:text-rose-800 p-1"><i class="fa-solid fa-trash-can"></i></button>
          </td>
        `;
        roomTbody.appendChild(tr);
      });

      // Agent Table
      const agentTbody = document.getElementById('master-agents-table');
      agentTbody.innerHTML = '';
      state.agents.forEach(a => {
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50 border-b border-slate-100";
        tr.innerHTML = `
          <td class="py-2 px-3 font-semibold text-slate-800">${a.name}</td>
          <td class="py-2 px-3 text-center">
            <button onclick="requestMasterDelete('agent', '${a.name}')" class="text-rose-600 hover:text-rose-800 p-1"><i class="fa-solid fa-trash-can"></i></button>
          </td>
        `;
        agentTbody.appendChild(tr);
      });
    }

    function handleAddRoom(e) {
      e.preventDefault();
      const rNo = document.getElementById('master-room-no').value.trim();
      const tariff = parseFloat(document.getElementById('master-room-tariff').value) || 0;

      if (state.rooms.some(r => String(r.roomNo) === String(rNo))) {
        alert("Room number already exists!");
        return;
      }

      state.rooms.push({ roomNo: rNo, tariff: tariff });
      document.getElementById('master-room-no').value = '';
      document.getElementById('master-room-tariff').value = '';
      saveStateToLocalStorage();
      renderMasterTables();
      renderDashboard();
      showNotification(`Room ${rNo} added to Master Data.`);
    }

    function handleAddAgent(e) {
      e.preventDefault();
      const name = document.getElementById('master-agent-name').value.trim();

      if (state.agents.some(a => a.name.toLowerCase() === name.toLowerCase())) {
        alert("Agent name already exists!");
        return;
      }

      state.agents.push({ name: name });
      document.getElementById('master-agent-name').value = '';
      saveStateToLocalStorage();
      renderMasterTables();
      showNotification(`Agent ${name} added.`);
    }

    function requestMasterDelete(type, key) {
      masterDeleteTarget = { type, key };
      document.getElementById('master-delete-modal-msg').innerText = `Are you sure you want to delete ${type} "${key}" from Master Tab?`;
      document.getElementById('master-delete-confirm-modal').classList.remove('hidden');
    }

    function closeMasterDeleteModal() {
      document.getElementById('master-delete-confirm-modal').classList.add('hidden');
      masterDeleteTarget = null;
    }

    function confirmMasterDeletion() {
      if (!masterDeleteTarget) return;

      if (masterDeleteTarget.type === 'room') {
        state.rooms = state.rooms.filter(r => String(r.roomNo) !== String(masterDeleteTarget.key));
      } else if (masterDeleteTarget.type === 'agent') {
        state.agents = state.agents.filter(a => a.name !== masterDeleteTarget.key);
      }

      saveStateToLocalStorage();
      closeMasterDeleteModal();
      renderMasterTables();
      renderDashboard();
      showNotification("Master record deleted.");
    }

    // Calendar Overview Logic
    function renderCalendar() {
      const title = document.getElementById('cal-month-year-title');
      const year = calendarCurrentDate.getFullYear();
      const month = calendarCurrentDate.getMonth();

      const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
      title.innerText = `${monthNames[month]} ${year}`;

      const grid = document.getElementById('calendar-grid-days');
      grid.innerHTML = '';

      const firstDay = new Date(year, month, 1).getDay();
      const daysInMonth = new Date(year, month + 1, 0).getDate();

      // Blank slots for layout alignment
      for (let i = 0; i < firstDay; i++) {
        const emptyCell = document.createElement('div');
        emptyCell.className = "h-16 bg-slate-50/50 rounded-xl border border-slate-100";
        grid.appendChild(emptyCell);
      }

      for (let day = 1; day <= daysInMonth; day++) {
        const currentDateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        
        // Find bookings matching day
        const dayBookings = state.bookings.filter(b => {
          const cInDate = b.checkIn ? b.checkIn.split('T')[0] : '';
          const cOutDate = b.checkOut ? b.checkOut.split('T')[0] : '';
          return currentDateStr >= cInDate && currentDateStr <= cOutDate;
        });

        const dayCell = document.createElement('div');
        dayCell.className = `h-16 p-1 rounded-xl border flex flex-col justify-between text-left transition ${
          dayBookings.length > 0 ? 'bg-blue-50/60 border-blue-200' : 'bg-white border-slate-100'
        }`;

        dayCell.innerHTML = `
          <div class="font-bold text-[10px] text-slate-700">${day}</div>
          ${dayBookings.length > 0 ? `
            <div class="bg-blue-600 text-white rounded-lg px-1 py-0.5 text-[8px] font-bold truncate">
              ${dayBookings.length} Booked
            </div>
          ` : ''}
        `;

        grid.appendChild(dayCell);
      }
    }

    function changeCalMonth(offset) {
      calendarCurrentDate.setMonth(calendarCurrentDate.getMonth() + offset);
      renderCalendar();
    }

    // Invoice Print & JPEG Logic
    function openInvoiceModal(id) {
      const b = state.bookings.find(item => item.id === id);
      if (!b) return;

      const roomsDisplay = normalizeRooms(b.roomNumber || b.roomsList).join(', ');

      document.getElementById('inv-id').innerText = b.id;
      document.getElementById('inv-name').innerText = b.guestName;
      document.getElementById('inv-phone').innerText = b.phone;
      document.getElementById('inv-rooms').innerText = roomsDisplay;
      document.getElementById('inv-checkin').innerText = formatDateTime(b.checkIn);
      document.getElementById('inv-checkout').innerText = formatDateTime(b.checkOut);
      document.getElementById('inv-agent').innerText = b.agent || 'Direct';
      document.getElementById('inv-total').innerText = `₹${b.totalAmount}`;
      document.getElementById('inv-advance').innerText = `₹${b.advancePaid}`;
      document.getElementById('inv-due').innerText = `₹${b.totalDue}`;

      document.getElementById('invoice-modal').classList.remove('hidden');
    }

    function closeInvoiceModal() {
      document.getElementById('invoice-modal').classList.add('hidden');
    }

    function downloadReceiptAsJPEG() {
      const node = document.getElementById('printable-invoice');
      html2canvas(node).then(canvas => {
        const link = document.createElement('a');
        link.download = `Voucher-${document.getElementById('inv-id').innerText}.jpg`;
        link.href = canvas.toDataURL('image/jpeg');
        link.click();
      });
    }

    // Google Sheets Sync & Save Integration
    function saveChanges() {
      document.getElementById('saving-lock-modal').classList.remove('hidden');
      saveChangesToGoogleSheet().then(() => {
        showNotification("Data synchronized with Google Sheets!");
      }).catch(err => {
        alert("Google Sheet sync error: " + err.message);
      }).finally(() => {
        document.getElementById('saving-lock-modal').classList.add('hidden');
      });
    }

    async function saveChangesToGoogleSheet() {
      // Ensure bookings payload roomNumber string property is strictly clean string
      const sanitizedState = {
        ...state,
        bookings: state.bookings.map(b => ({
          ...b,
          roomNumber: normalizeRooms(b.roomNumber || b.roomsList).join(', ')
        }))
      };

      try {
        await fetch(GOOGLE_SCRIPT_URL, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ action: 'syncAll', data: sanitizedState })
        });
      } catch (err) {
        console.error("Google sync error", err);
      }
    }

    async function fetchGoogleSheetData() {
      try {
        const res = await fetch(`${GOOGLE_SCRIPT_URL}?action=fetchAll`);
        if (res.ok) {
          const data = await res.json();
          if (data && data.state) {
            state = data.state;
            saveStateToLocalStorage();
            renderDashboard();
            renderBookingsTable();
          }
        }
      } catch (e) {
        console.warn("Could not fetch remote sheet data, using local state.", e);
      }
    }

    // Wipe Data System
    function requestDataWipe() {
      document.getElementById('wipe-layer-1-modal').classList.remove('hidden');
    }

    function proceedToWipeLayer2() {
      document.getElementById('wipe-layer-1-modal').classList.add('hidden');
      document.getElementById('wipe-layer-2-modal').classList.remove('hidden');
    }

    function closeWipeModals() {
      document.getElementById('wipe-layer-1-modal').classList.add('hidden');
      document.getElementById('wipe-layer-2-modal').classList.add('hidden');
    }

    function executeGoogleSheetWipe() {
      state = { rooms: [], agents: [], bookings: [] };
      saveStateToLocalStorage();
      closeWipeModals();
      renderDashboard();
      renderBookingsTable();
      renderMasterTables();
      showNotification("All portal and sheet data erased.");
    }

    // Excel Export Logic
    function openExportModal() {
      document.getElementById('export-modal').classList.remove('hidden');
    }

    function closeExportModal() {
      document.getElementById('export-modal').classList.add('hidden');
    }

    function processExport() {
      const s = document.getElementById('export-start-date').value;
      const e = document.getElementById('export-end-date').value;

      const filteredBookings = state.bookings.filter(b => {
        const cIn = b.checkIn ? b.checkIn.split('T')[0] : '';
        if (!s || !e) return true;
        return cIn >= s && cIn <= e;
      });

      const exportData = filteredBookings.map(b => ({
        "Booking ID": b.id,
        "Guest Name": b.guestName,
        "Phone": b.phone,
        "Room(s)": normalizeRooms(b.roomNumber || b.roomsList).join(', '),
        "Check-In": b.checkIn,
        "Check-Out": b.checkOut,
        "Agent": b.agent || 'Direct',
        "Total (₹)": b.totalAmount,
        "Advance (₹)": b.advancePaid,
        "Due (₹)": b.totalDue
      }));

      const worksheet = XLSX.utils.json_to_sheet(exportData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Bookings");
      XLSX.writeFile(workbook, `Booking_Report_${s || 'All'}_to_${e || 'All'}.xlsx`);

      closeExportModal();
    }

    // Utility Helpers
    function formatDateTime(dtStr) {
      if (!dtStr) return '-';
      const dt = new Date(dtStr);
      if (isNaN(dt)) return dtStr;
      return dt.toLocaleString('en-IN', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    }

    function showNotification(msg) {
      const banner = document.getElementById('notification-banner');
      document.getElementById('notification-text').innerText = msg;
      banner.classList.remove('hidden');
      setTimeout(() => {
        banner.classList.add('hidden');
      }, 3500);
    }

    function closeNotification() {
      document.getElementById('notification-banner').classList.add('hidden');
    }

    function openAlertModal() {
      const container = document.getElementById('alert-list-container');
      container.innerHTML = '';

      const pendingDues = state.bookings.filter(b => (parseFloat(b.totalDue) || 0) > 0);

      if (pendingDues.length === 0) {
        container.innerHTML = `<p class="text-slate-500 text-center py-4">No pending alerts.</p>`;
      } else {
        pendingDues.forEach(b => {
          const item = document.createElement('div');
          item.className = "bg-amber-50 border border-amber-200 p-2.5 rounded-2xl flex justify-between items-center";
          item.innerHTML = `
            <div>
              <p class="font-bold text-amber-900">${b.guestName} (${b.id})</p>
              <p class="text-[10px] text-amber-700">Room: ${normalizeRooms(b.roomNumber || b.roomsList).join(', ')}</p>
            </div>
            <span class="font-bold text-rose-600">Due: ₹${b.totalDue}</span>
          `;
          container.appendChild(item);
        });
      }

      document.getElementById('alert-modal').classList.remove('hidden');
    }

    function closeAlertModal() {
      document.getElementById('alert-modal').classList.add('hidden');
    }

    function closeCommentBox() {
      document.getElementById('excel-comment-box').classList.add('hidden');
    }

    // Inactivity Auto-Logout Mechanism
    function resetInactivityTimer() {
      clearTimeout(inactivityTimer);
      clearInterval(countdownInterval);
      document.getElementById('logout-warning-modal').classList.add('hidden');

      // 15 Minutes Inactivity Limit
      inactivityTimer = setTimeout(() => {
        showLogoutWarning();
      }, 14 * 60 * 1000);
    }

    function showLogoutWarning() {
      let count = 60;
      document.getElementById('logout-countdown-seconds').innerText = count;
      document.getElementById('logout-warning-modal').classList.remove('hidden');

      countdownInterval = setInterval(() => {
        count--;
        document.getElementById('logout-countdown-seconds').innerText = count;
        if (count <= 0) {
          clearInterval(countdownInterval);
          processLogoutWithSave();
        }
      }, 1000);
    }

    // Global Activity Event Listeners
    ['mousemove', 'keydown', 'click', 'scroll'].forEach(evt => {
      window.addEventListener(evt, () => {
        if (localStorage.getItem("portal_logged_in") === "true") {
          resetInactivityTimer();
        }
      });
    });
  </script>
</body>
</html>
