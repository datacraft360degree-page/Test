

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

  <div id="login-overlay" class="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4">
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

  <div id="logout-warning-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4">
    <div class="bg-white rounded-3xl shadow-xl border border-slate-100 max-w-xs w-full p-5 space-y-3 text-center">
      <div class="bg-amber-50 text-amber-600 w-10 h-10 rounded-2xl flex items-center justify-center mx-auto text-lg">
        <i class="fa-solid fa-hourglass-half"></i>
      </div>
      <div>
        <h3 class="text-xs font-bold text-slate-900">Inactivity Timeout Warning</h3>
        <p class="text-[10px] text-slate-500 mt-1">You will be logged out automatically in <strong id="logout-countdown-seconds" class="text-rose-600">60</strong> seconds due to inactivity.</p>
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
        <button onclick="openAlertModal()" title="View Alerts" class="relative bg-amber-50 hover:bg

