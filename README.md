<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Business Portal - Web Application</title>
  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- SheetJS for Exporting to Excel -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- html2canvas for Generating JPEG Receipts -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
  <!-- FontAwesome Icons -->
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
    /* Floating Balloons & Ribbons Keyframe Animations */
    @keyframes floatUp {
      0% { transform: translateY(100vh) rotate(0deg); opacity: 0; }
      20% { opacity: 0.8; }
      100% { transform: translateY(-20vh) rotate(360deg); opacity: 0; }
    }

    @keyframes ribbonSway {
      0%, 100% { transform: translateY(0) rotate(-5deg); }
      50% { transform: translateY(-15px) rotate(5deg); }
    }

    .balloon {
      position: absolute;
      bottom: -100px;
      font-size: 2rem;
      animation: floatUp 8s linear infinite;
      z-index: 10;
      pointer-events: none;
    }

    .ribbon {
      position: absolute;
      font-size: 1.8rem;
      animation: ribbonSway 4s ease-in-out infinite;
      z-index: 10;
      pointer-events: none;
    }
  </style>
</head>
<body class="text-slate-800 font-sans min-h-screen flex flex-col relative antialiased text-xs" onclick="closeCommentBox()">

<!-- BIRTHDAY HURRAY ANIMATION MODAL (TRIGGERS ON 27th September) -->
  <div id="birthday-hurray-modal" class="hidden fixed inset-0 z-[110] bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print overflow-hidden">

    <!-- Modal Card (Soft Light Palette) -->
    <div class="bg-gradient-to-br from-amber-100 via-pink-100 to-sky-100 rounded-3xl p-1 shadow-2xl max-w-md w-full relative z-20 border border-white/60">
      <div class="bg-white/90 backdrop-blur-md rounded-[22px] p-6 text-center space-y-4 relative overflow-hidden border border-slate-100">
        
        <div class="absolute -top-10 -right-10 w-32 h-32 bg-amber-200/50 rounded-full blur-2xl"></div>
        <div class="absolute -bottom-10 -left-10 w-32 h-32 bg-pink-200/50 rounded-full blur-2xl"></div>
        
        <div class="text-6xl animate-bounce relative z-10">
          🎉🥳🎂
        </div>
        
        <div class="space-y-2 relative z-10">
          <h2 class="text-xl font-black text-slate-800 tracking-tight">Happy Birthday Mr. Aniruddha Sir 🥳🎂</h2>
          <p class="text-sm font-semibold text-slate-600 bg-sky-50/80 border border-sky-100 py-3 px-4 rounded-2xl shadow-inner leading-relaxed">
            "Hi Aniruddha Sir, it's your birthday. Keep smiling always and stay healthy! This is a message from Mr. Kapil: wishing you a very Happy Birthday." 😊
          </p>
        </div>

        <button onclick="closeBirthdayModal()" class="w-full bg-gradient-to-r from-sky-400 via-rose-400 to-amber-400 hover:opacity-95 text-white font-black py-3 rounded-2xl shadow-md transition text-xs relative z-10 cursor-pointer active:scale-95 tracking-wide">
          Hurray! Let's Celebrate 🚀
        </button>
      </div>
    </div>
  </div>

  <!-- LOGIN MODAL OVERLAY -->
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

  <!-- LOGIN ALERT MESSAGE MODAL (POPUP ON SUCCESSFUL LOGIN) -->
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

  <!-- MASTER DATA ACCESS PASSWORD MODAL -->
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

  <!-- MASTER DATA PERMANENT DELETION RECONFIRMATION POPUP MODAL -->
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

  <!-- MANUAL LOGOUT CONFIRM MODAL -->
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

  <!-- SAVING LOCK MODAL (SAND TIMER) -->
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

  <!-- SESSION AUTO LOGOUT WARNING MODAL -->
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

  <!-- EXPORT TO EXCEL DATE RANGE MODAL -->
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

  <!-- Excel Comment Box Popout -->
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

  <!-- Header Navigation -->
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
      
      <!-- One UI Pill Navigation -->
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
        <button onclick="openAlertModal()" title="View Alerts" class="relative bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-200 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-bell text-[10px]"></i> Alerts
          <span id="alert-badge" class="hidden absolute -top-1 -right-1 bg-rose-600 text-white text-[9px] font-black px-1.5 py-0.2 rounded-full border border-white animate-bounce">0</span>
        </button>
        <!-- Save Button -->
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
          <i class="fa-solid fa-file-excel text-[10px]"></i> Export
        </button>
        <button onclick="logoutUser()" title="Logout" class="bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-right-from-bracket text-[10px]"></i> Logout
        </button>
      </div>
    </div>
  </header>

  <!-- Notification Toast -->
  <div id="toast" class="hidden fixed bottom-6 right-6 bg-slate-900/90 backdrop-blur-md text-white px-4 py-2.5 rounded-2xl shadow-xl z-50 flex items-center gap-2.5 no-print border border-slate-800 text-[11px]">
    <i class="fa-solid fa-circle-check text-emerald-400 text-base"></i>
    <span id="toast-message" class="font-medium">Changes Auto save successfully!</span>
  </div>

  <!-- Main Content Area -->
  <main class="max-w-7xl mx-auto px-4 py-4 flex-1 w-full no-print space-y-4">

    <!-- DASHBOARD TAB -->
    <section id="tab-dashboard" class="tab-content space-y-4">
      <div class="bg-gradient-to-r from-blue-600 to-indigo-600 rounded-3xl p-5 text-white shadow-sm flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
        <div>
          <h2 class="text-base font-bold tracking-tight">Hi, Welcome to dashboard 🏠</h2>
          <p class="text-blue-100 text-[10px] mt-0.5">Quickly view, schedule, and manage room allocations and orders.</p>
        </div>
        <div class="flex items-center bg-white/10 backdrop-blur-md px-3 py-1.5 rounded-2xl border border-white/20 space-x-2">
          <label for="dash-year-select" class="text-[10px] font-bold text-blue-50 uppercase flex items-center gap-1">
            <i class="fa-solid fa-filter text-amber-300"></i> Filter Year:
          </label>
          <select id="dash-year-select" onchange="handleDashboardYearChange(this.value)" class="bg-white text-blue-900 text-[11px] font-bold rounded-xl px-2.5 py-1 focus:outline-none cursor-pointer shadow-sm">
          </select>
        </div>
      </div>

      <!-- Summary Filter Banner Indicator -->
      <div class="flex items-center justify-between bg-white px-4 py-2 rounded-2xl border border-slate-200/60 shadow-sm">
        <span class="text-[11px] font-semibold text-slate-600 flex items-center gap-2">
          <i class="fa-solid fa-chart-line text-blue-600"></i>
          Showing Summary For: <strong id="dash-filter-label" class="text-blue-600 font-bold">Consolidated (All Years)</strong>
        </span>
        <button onclick="handleDashboardYearChange('CURRENT')" class="text-[10px] bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold px-3 py-1 rounded-full transition border border-slate-200">
          Reset to Current Year
        </button>
      </div>

      <!-- One UI Rounded Cards -->
    <!-- 6-Box Grid Layout (Line 1: 3 Boxes | Line 2: 3 Boxes) -->
      <div class="grid grid-cols-1 md:grid-cols-3 gap-3">

        <!-- BOX 1: TOTAL BOOKINGS (COUNT PIE CHART) -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex flex-col justify-between space-y-3">
          <div class="flex justify-between items-center border-b border-slate-100 pb-2">
            <div>
              <p class="text-[10px] uppercase font-bold text-slate-400 tracking-wider break-words">Total Bookings</p>
              <p id="dash-total-bookings" class="text-lg font-black text-slate-900 leading-tight">0</p>
            </div>
            <div class="p-2.5 bg-amber-50 text-amber-600 rounded-2xl"><i class="fa-solid fa-bookmark text-sm"></i></div>
          </div>
          <div class="flex items-center space-x-3">
            <div id="pie-bookings" class="w-16 h-16 rounded-full shrink-0 shadow-inner border border-slate-100" style="background: conic-gradient(#E2E8F0 0% 100%);"></div>
            <div class="text-[11px] space-y-1 font-semibold text-slate-700 w-full">
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-amber-500 inline-block"></span>Live:</span> <strong id="dash-count-live" class="text-amber-600 font-bold">0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-blue-500 inline-block"></span>Upcoming:</span> <strong id="dash-count-upcoming" class="text-blue-600 font-bold">0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block"></span>Closed:</span> <strong id="dash-count-closed" class="text-emerald-600 font-bold">0</strong></div>
            </div>
          </div>
        </div>

        <!-- BOX 2: BOOKING AMOUNT (AMOUNT PIE CHART) -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex flex-col justify-between space-y-3">
          <div class="flex justify-between items-center border-b border-slate-100 pb-2">
            <div>
              <p class="text-[10px] uppercase font-bold text-slate-400 tracking-wider break-words">Booking Amount</p>
              <p id="dash-total-amount" class="text-lg font-black text-slate-900 leading-tight">₹0</p>
            </div>
            <div class="p-2.5 bg-blue-50 text-blue-600 rounded-2xl"><i class="fa-solid fa-receipt text-sm"></i></div>
          </div>
          <div class="flex items-center space-x-3">
            <div id="pie-amount" class="w-16 h-16 rounded-full shrink-0 shadow-inner border border-slate-100" style="background: conic-gradient(#E2E8F0 0% 100%);"></div>
            <div class="text-[11px] space-y-1 font-semibold text-slate-700 w-full">
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-amber-500 inline-block"></span>Live:</span> <strong id="dash-amount-live" class="text-amber-600 font-bold">₹0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-blue-500 inline-block"></span>Upcoming:</span> <strong id="dash-amount-upcoming" class="text-blue-600 font-bold">₹0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block"></span>Closed:</span> <strong id="dash-amount-closed" class="text-emerald-600 font-bold">₹0</strong></div>
            </div>
          </div>
        </div>

        <!-- BOX 3: AMOUNT RECEIVED (ADVANCE PIE CHART) -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex flex-col justify-between space-y-3">
          <div class="flex justify-between items-center border-b border-slate-100 pb-2">
            <div>
              <p class="text-[10px] uppercase font-bold text-slate-400 tracking-wider break-words">Amount Received</p>
              <p id="dash-advanced" class="text-lg font-black text-emerald-600 leading-tight">₹0</p>
            </div>
            <div class="p-2.5 bg-emerald-50 text-emerald-600 rounded-2xl"><i class="fa-solid fa-wallet text-sm"></i></div>
          </div>
          <div class="flex items-center space-x-3">
            <div id="pie-received" class="w-16 h-16 rounded-full shrink-0 shadow-inner border border-slate-100" style="background: conic-gradient(#E2E8F0 0% 100%);"></div>
            <div class="text-[11px] space-y-1 font-semibold text-slate-700 w-full">
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-amber-500 inline-block"></span>Live:</span> <strong id="dash-recv-live" class="text-amber-600 font-bold">₹0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-blue-500 inline-block"></span>Upcoming:</span> <strong id="dash-recv-upcoming" class="text-blue-600 font-bold">₹0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block"></span>Closed:</span> <strong id="dash-recv-closed" class="text-emerald-600 font-bold">₹0</strong></div>
            </div>
          </div>
        </div>

        <!-- BOX 4: TOTAL DUE AMOUNT (DUE PIE CHART) -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex flex-col justify-between space-y-3">
          <div class="flex justify-between items-center border-b border-slate-100 pb-2">
            <div>
              <p class="text-[10px] uppercase font-bold text-slate-400 tracking-wider break-words">Total Due Amount</p>
              <p id="dash-due" class="text-lg font-black text-rose-600 leading-tight">₹0</p>
            </div>
            <div class="p-2.5 bg-rose-50 text-rose-600 rounded-2xl"><i class="fa-solid fa-hand-holding-dollar text-sm"></i></div>
          </div>
          <div class="flex items-center space-x-3">
            <div id="pie-due" class="w-16 h-16 rounded-full shrink-0 shadow-inner border border-slate-100" style="background: conic-gradient(#E2E8F0 0% 100%);"></div>
            <div class="text-[11px] space-y-1 font-semibold text-slate-700 w-full">
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-amber-500 inline-block"></span>Live Due:</span> <strong id="dash-due-live" class="text-amber-600 font-bold">₹0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-blue-500 inline-block"></span>Upcoming Due:</span> <strong id="dash-due-upcoming" class="text-blue-600 font-bold">₹0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block"></span>Closed Due:</span> <strong id="dash-due-closed" class="text-emerald-600 font-bold">₹0</strong></div>
            </div>
          </div>
        </div>

        <!-- BOX 5: INACTIVE BOOKINGS SUMMARY (RED THEME) -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex flex-col justify-between space-y-3">
          <div class="flex justify-between items-center border-b border-slate-100 pb-2">
            <div>
              <p class="text-[10px] uppercase font-bold text-red-500 tracking-wider break-words">Inactive Bookings</p>
              <p id="dash-inactive-count" class="text-lg font-black text-red-600 leading-tight">0 Bookings</p>
            </div>
            <div class="p-2.5 bg-red-50 text-red-600 rounded-2xl"><i class="fa-solid fa-ban text-sm"></i></div>
          </div>
          <div class="bg-red-50/50 p-3 rounded-2xl border border-red-100 text-[11px] space-y-1">
            <p class="text-slate-500 font-medium">Inactive Total Booking Amount:</p>
            <p id="dash-inactive-amount" class="text-lg font-black text-red-600">₹0</p>
          </div>
        </div>

        <!-- BOX 6: ALL LIVE BOOKING DETAILS (YELLOW/AMBER THEME) -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex flex-col justify-between space-y-2">
          <div class="flex justify-between items-center border-b border-slate-100 pb-2">
            <p class="text-[10px] uppercase font-bold text-slate-400 tracking-wider break-words">All Live Bookings Details</p>
            <span class="px-2 py-0.5 text-[9px] font-bold bg-amber-50 text-amber-600 rounded-full border border-amber-200 flex items-center gap-1">
              <span class="w-1.5 h-1.5 bg-amber-500 rounded-full animate-ping"></span> Live Today
            </span>
          </div>
          <div id="dash-today-live-container" class="overflow-x-auto max-h-32 overflow-y-auto">
            <p class="text-[11px] text-slate-400 italic py-2 text-center">No live bookings active today.</p>
          </div>
        </div>

      </div>

      <!-- ALL BOOKING REMARKS: separate dashboard row with Add Cab Trip-style toggle -->
      <div class="bg-white rounded-3xl shadow-sm border border-indigo-200/80 p-4 space-y-3">
        <div class="flex justify-between items-center gap-3 border-b border-slate-100 pb-2">
          <div>
            <h3 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
              <i class="fa-solid fa-comment-dots text-indigo-600"></i> All Booking Remarks
            </h3>
            <p class="text-[10px] text-slate-500 mt-0.5">View saved remarks by booking.</p>
          </div>
          <button type="button" id="btn-toggle-all-remarks" onclick="toggleDashboardAllRemarks()" class="bg-indigo-100 hover:bg-indigo-200 text-indigo-700 px-3 py-1.5 rounded-full text-[10px] font-semibold inline-flex items-center gap-1.5 transition whitespace-nowrap">
            <i class="fa-solid fa-plus text-[9px]"></i> Show all remarks
          </button>
        </div>
        <div id="dash-all-remarks-container" class="hidden space-y-2">
          <p class="text-[11px] text-slate-400 italic py-2 text-center">No saved remarks found.</p>
        </div>
      </div>

      <!-- BOOKING DETAILS BY AGENT / SELF -->
      <div class="bg-white rounded-3xl shadow-sm border border-violet-200/80 p-4 space-y-3">
        <div class="flex flex-wrap justify-between items-center gap-3 border-b border-slate-100 pb-2">
          <div><h3 class="text-xs font-bold text-slate-900"><i class="fa-solid fa-user-tie text-violet-600"></i> Booking details by Agent/Self</h3><p class="text-[10px] text-slate-500 mt-0.5">Booking summary and payment details by booking agent.</p></div>
          <div class="flex items-center gap-2"><label for="dash-agent-filter" class="text-[10px] font-semibold text-slate-600">Agent</label><select id="dash-agent-filter" onchange="updateDashboardAgentDetails()" class="border border-violet-200 rounded-lg px-2 py-1.5 text-[10px] bg-white"><option value="ALL">All Agents / Self</option></select></div>
          <button type="button" id="btn-toggle-agent-details" onclick="toggleDashboardAgentDetails()" class="bg-violet-100 hover:bg-violet-200 text-violet-700 px-3 py-1.5 rounded-full text-[10px] font-semibold whitespace-nowrap"><i class="fa-solid fa-plus text-[9px]"></i> Show details</button>
        </div>
        <div id="dash-agent-details-container" class="hidden overflow-x-auto"><div id="dash-agent-details-content"></div></div>
      </div>

      <!-- Active years Directory Table Hidden -->
      <div class="hidden bg-white rounded-3xl shadow-sm border border-slate-200/60 p-4">
        <div class="mb-3 flex justify-between items-center">
          <h3 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
            <i class="fa-solid fa-calendar-days text-blue-600"></i> Active Years Directory (2026 – 2085)
          </h3>
          <span class="text-[10px] text-slate-400 font-medium">Click any year to filter dashboard &amp; open year calendar</span>
        </div>
        <div id="years-grid" class="grid grid-cols-6 sm:grid-cols-10 md:grid-cols-12 gap-2"></div>
      </div>
    </section>

    <!-- BOOKING DETAILS TAB -->
    <section id="tab-booking" class="tab-content hidden space-y-4">
      <div class="bg-white rounded-3xl shadow-sm border border-slate-200/60 p-4">
        <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-3 mb-4 pb-3 border-b border-slate-100">
          <div>
            <h2 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
              <i class="fa-solid fa-address-card text-blue-600"></i> Guest Information &amp; Reservation Directory
            </h2>
            <div class="flex items-center gap-3 mt-2 text-[10px]">
              <span class="flex items-center gap-1.5 font-semibold text-amber-800">
                <span class="relative flex h-2.5 w-2.5">
                  <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-400 opacity-75"></span>
                  <span class="relative inline-flex rounded-full h-2.5 w-2.5 bg-amber-500"></span>
                </span> Live Booking
              </span>
              <span class="flex items-center gap-1.5 font-semibold text-blue-800">
                <span class="w-2.5 h-2.5 bg-blue-500 rounded-full inline-block"></span> Upcoming Booking
              </span>
              <span class="flex items-center gap-1.5 font-semibold text-emerald-800">
                <span class="w-2.5 h-2.5 bg-emerald-500 rounded-full inline-block"></span> Closed Booking
              </span>
              <span class="flex items-center gap-1.5 font-semibold text-slate-700">
                <span class="w-2 h-2 bg-rose-600 rounded-full inline-block"></span> Inactive Booking
              </span>
            </div>
          </div>
          
          <div class="flex items-center space-x-2 w-full md:w-auto">
            <!-- Search by Date -->
            <div class="flex items-center bg-slate-100 border border-slate-200 rounded-2xl px-2 py-1 space-x-1.5">
              <label for="booking-date-search" class="text-[10px] font-bold text-slate-500 uppercase flex items-center gap-1 pl-1">
                <i class="fa-solid fa-calendar-day text-blue-600"></i> Search Date:
              </label>
              <input type="date" id="booking-date-search" onchange="searchBookingByDate()" class="bg-white text-[11px] border border-slate-200 rounded-xl px-2 py-0.5 focus:outline-none focus:ring-2 focus:ring-blue-500 font-bold text-blue-600 cursor-pointer" />
              <button onclick="clearDateSearchBooking()" class="text-slate-400 hover:text-slate-600 px-1 text-[10px]" title="Reset Filter">
                <i class="fa-solid fa-rotate-left"></i> Reset
              </button>
            </div>

            <button onclick="openBookingModal()" class="bg-blue-600 hover:bg-blue-700 text-white px-3.5 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1.5 transition shadow-sm whitespace-nowrap">
              <i class="fa-solid fa-plus text-[10px]"></i> Add Booking
            </button>
          </div>
        </div>

        <!-- Bookings Table View -->
        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200/80 text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                <th class="py-2.5 px-3">Booking ID</th>
                <th class="py-2.5 px-3">Guest Name</th>
                <th class="py-2.5 px-3">Contact No</th>
                <th class="py-2.5 px-3">ID No</th>
                <th class="py-2.5 px-3">Attached ID</th>
                <th class="py-2.5 px-3">Room</th>
                <th class="py-2.5 px-3">Capacity</th>
                <th class="py-2.5 px-3">Agent Info</th>
                <th class="py-2.5 px-3 min-w-[150px]">Stay Window</th>
                <th class="py-2.5 px-3">Tariff &amp; Extras</th>
                <th class="py-2.5 px-3">Payment/Adv</th>
                <th class="py-2.5 px-3">Due</th>
                <th class="py-2.5 px-3 text-center">Actions</th>
              </tr>
            </thead>
            <tbody id="bookings-tbody" class="divide-y divide-slate-100 text-[11px]"></tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- MASTER DATA TAB -->
    <section id="tab-master" class="tab-content hidden space-y-4">
      
      <!-- Room Capacity Table -->
      <div class="bg-white rounded-3xl shadow-sm border border-slate-200/60 p-4">
        <div class="flex justify-between items-center mb-3 pb-2 border-b border-slate-100">
          <div>
            <h2 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
              <i class="fa-solid fa-door-open text-blue-600"></i> Room Capacity Configuration
            </h2>
            <p class="text-[10px] text-slate-500 mt-0.5">Default rooms 1 to 5. Click Add Room to append new rooms anytime.</p>
          </div>
          <button type="button" onclick="addRoomCapacityRow()" class="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-full text-[11px] font-medium flex items-center gap-1.5 transition shadow-sm cursor-pointer">
            <i class="fa-solid fa-plus text-[10px]"></i> Add Room
          </button>
        </div>

        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200/80 text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                <th class="py-2.5 px-3">Room No</th>
                <th class="py-2.5 px-3">Room Capacity (Person)</th>
                <th class="py-2.5 px-3 text-center">Actions</th>
              </tr>
            </thead>
            <tbody id="room-capacity-tbody" class="divide-y divide-slate-100 text-[11px]"></tbody>
          </table>
        </div>
      </div>

      <!-- Agent Information Table -->
      <div class="bg-white rounded-3xl shadow-sm border border-slate-200/60 p-4">
        <div class="flex justify-between items-center mb-3 pb-2 border-b border-slate-100">
          <div>
            <h2 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
              <i class="fa-solid fa-users-gear text-blue-600"></i> Master Agent Directory
            </h2>
            <p class="text-[10px] text-slate-500 mt-0.5">Manage Agents linked with room allocations.</p>
          </div>
          <button type="button" onclick="addAgentRow()" class="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-full text-[11px] font-medium flex items-center gap-1.5 transition shadow-sm cursor-pointer">
            <i class="fa-solid fa-plus text-[10px]"></i> Add Agent Entry
          </button>
        </div>

        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200/80 text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                <th class="py-2.5 px-3">Agent Name</th>
                <th class="py-2.5 px-3">Agent Contact</th>
                <th class="py-2.5 px-3">Linked Room No</th>
                <th class="py-2.5 px-3 text-center">Actions</th>
              </tr>
            </thead>
            <tbody id="agent-tbody" class="divide-y divide-slate-100 text-[11px]"></tbody>
          </table>
        </div>
      </div>

      <!-- BOOKING ID TYPING SEARCH & DELETION CONTROL -->
      <div class="bg-white rounded-3xl shadow-sm border border-rose-200/80 p-4 space-y-3">
        <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2 border-b border-slate-100 pb-2.5">
          <div>
            <h2 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
              <i class="fa-solid fa-trash-can text-rose-600"></i> Booking Deletion Manager
            </h2>
            <p class="text-[10px] text-slate-500 mt-0.5">Type a Booking ID directly to safely locate and remove it from the system.</p>
          </div>
          
          <div class="flex items-center bg-slate-100 border border-slate-200 rounded-2xl px-2 py-1 space-x-1.5">
            <label for="master-booking-search-input" class="text-[10px] font-bold text-slate-600 uppercase flex items-center gap-1 pl-1">
              <i class="fa-solid fa-magnifying-glass text-blue-600"></i> Type Booking ID:
            </label>
            <input type="text" id="master-booking-search-input" oninput="searchMasterBookingById()" placeholder="e.g. BKG-2026-0000001" class="bg-white text-[11px] border border-slate-200 rounded-xl px-2 py-0.5 focus:outline-none focus:ring-2 focus:ring-blue-500 font-mono font-bold text-blue-600 uppercase w-48" />
            <button onclick="clearMasterBookingSearch()" class="text-slate-400 hover:text-slate-600 px-1 text-[10px]" title="Clear Search">
              <i class="fa-solid fa-xmark"></i>
            </button>
          </div>
        </div>

        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-rose-50/60 border-b border-rose-100 text-[10px] font-bold text-rose-800 uppercase tracking-wider">
                <th class="py-2.5 px-3">Booking ID</th>
                <th class="py-2.5 px-3">Guest Name</th>
                <th class="py-2.5 px-3">Room No</th>
                <th class="py-2.5 px-3">Stay Window</th>
                <th class="py-2.5 px-3">Total Amount</th>
                <th class="py-2.5 px-3">Due Amount</th>
                <th class="py-2.5 px-3 text-center">Delete Linked Booking</th>
              </tr>
            </thead>
            <tbody id="master-delete-tbody" class="divide-y divide-slate-100 text-[11px]">
              <tr>
                <td colspan="7" class="text-center py-4 text-slate-400">Please type a Booking ID into the search field above to view and delete details.</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- CALENDAR TAB -->
    <section id="tab-calendar" class="tab-content hidden space-y-4">
      <div class="bg-white rounded-3xl shadow-sm border border-slate-200/60 p-4">
        <div class="flex justify-between items-center mb-4">
          <div>
            <h2 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
              <i class="fa-regular fa-calendar-check text-blue-600"></i> Year Overview Calendar
            </h2>
          </div>
          
          <div class="flex items-center bg-slate-100 border border-slate-200 px-3 py-1.5 rounded-2xl space-x-2">
            <label for="cal-year-select" class="text-[10px] font-bold text-slate-600 uppercase flex items-center gap-1">
              <i class="fa-solid fa-filter text-blue-600"></i> Filter Year:
            </label>
            <select id="cal-year-select" onchange="renderCalendar(parseInt(this.value))" class="bg-white text-blue-900 text-[11px] font-bold rounded-xl px-2.5 py-1 focus:outline-none focus:ring-2 focus:ring-blue-500 cursor-pointer shadow-sm"></select>
          </div>
        </div>

        <div id="calendar-container" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-3"></div>
      </div>
    </section>

  </main>

  <!-- POPUP MODAL: CHECK-OUT ALERT LIST -->
  <div id="alert-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-lg w-full flex flex-col max-h-[85vh] overflow-hidden">
      <div class="bg-amber-500 p-4 text-white flex justify-between items-center">
        <div class="flex items-center space-x-2">
          <i class="fa-solid fa-bell text-base"></i>
          <h3 class="text-xs font-bold">Due Payment Alert</h3>
        </div>
        <button onclick="closeAlertModal()" class="text-amber-100 hover:text-white px-1 text-base">
          <i class="fa-solid fa-xmark"></i>
        </button>
      </div>

      <div id="alert-list-container" class="p-4 overflow-y-auto space-y-2 flex-1 text-[11px]"></div>

      <div class="bg-slate-50 border-t border-slate-100 p-3 flex justify-between items-center text-[11px]">
        <span id="alert-list-count-text" class="text-slate-500 font-medium">0 active warnings found</span>
        <button onclick="closeAlertModal()" class="px-4 py-1.5 bg-slate-900 text-white rounded-xl font-semibold text-[10px]">Dismiss</button>
      </div>
    </div>
  </div>

  <!-- CLOSED BOOKING CONFIRMATION MODAL -->
  <div id="closed-booking-confirm-modal" class="hidden fixed inset-0 z-[80] bg-slate-900/50 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-amber-100 max-w-md w-full p-5 space-y-4 text-left">
      <div class="flex items-start gap-3">
        <div class="bg-amber-50 text-amber-600 w-11 h-11 rounded-2xl flex items-center justify-center text-lg shadow-sm shrink-0">
          <i class="fa-solid fa-circle-exclamation"></i>
        </div>
        <div>
          <h3 class="text-sm font-bold text-slate-900">Closed Booking Confirmation</h3>
          <p class="text-[11px] text-slate-600 mt-2 leading-relaxed">
            This Booking will be considered as “Closed Booking” so make sure you have filled all the required details because after saving this Only “Extra Food/Drink, Cab Fare &amp; Billing Summary Section" will be editable.
          </p>
        </div>
      </div>
      <div class="flex space-x-2 pt-1">
        <button type="button" onclick="closeClosedBookingConfirm(false)" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2.5 rounded-xl text-[11px] transition">
          Back
        </button>
        <button type="button" onclick="closeClosedBookingConfirm(true)" class="w-1/2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2.5 rounded-xl shadow-sm transition text-[11px]">
          Go ahead
        </button>
      </div>
    </div>
  </div>

  <!-- COMPACT ADD / EDIT BOOKING MODAL -->
 <div id="booking-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-3 overflow-y-auto no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-3xl w-full p-5 space-y-3 my-4 max-h-[90vh] overflow-y-auto">
      <div class="flex justify-between items-center pb-2.5 border-b border-slate-100">
        <div>
          <h3 id="modal-title" class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
            <i class="fa-solid fa-calendar-plus text-blue-600"></i> Add New Booking
          </h3>
        </div>
        <button onclick="closeBookingModal()" class="text-slate-400 hover:text-slate-600 p-0.5 text-base"><i class="fa-solid fa-xmark"></i></button>
      </div>

      <form id="booking-form" onsubmit="handleSaveBooking(event)" class="space-y-3 text-[11px]">
        <input type="hidden" id="modal-booking-id" />

        <!-- GUEST DETAILS -->
        <div id="sec-guest-info" class="bg-slate-50 p-3 rounded-2xl border border-slate-200/60 space-y-2.5 transition-all">
          <h4 class="text-[9px] font-bold uppercase tracking-wider text-slate-400 flex items-center gap-1.5">
            <i class="fa-solid fa-user-tag text-blue-600"></i> Guest Information
          </h4>
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Guest Name <span class="text-rose-500">*</span></label>
              <input type="text" id="cust-name" required="" pattern="[A-Za-z\s]+" oninput="this.value = formatTitleCase(this.value.replace(/[^A-Za-z\s]/g, ''))" title="Please enter Guest Name using characters only" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Address</label>
              <input type="text" id="cust-address" oninput="this.value = formatTitleCase(this.value)" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">City</label>
              <input type="text" id="cust-city" placeholder="City" oninput="this.value = formatTitleCase(this.value)" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">State</label>
              <input type="text" id="cust-state" oninput="this.value = formatTitleCase(this.value); handleStateChange(this.value)" placeholder="State" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Country</label>
              <input type="text" id="cust-country" placeholder="Country" oninput="this.value = formatTitleCase(this.value)" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Pin/Zip Code</label>
              <input type="text" id="cust-zip" placeholder="Pin/Zip Code" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">ID Number</label>
              <input type="text" id="cust-id" maxlength="16" pattern="[A-Za-z0-9\s]*" oninput="this.value = this.value.replace(/[^A-Za-z0-9\s]/g, '')" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500" />
            </div>
            <!-- EDITABLE COUNTRY CODE & GUEST CONTACT NUMBER -->
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Contact No</label>
              <div class="flex gap-1">
                <input type="text" id="cust-country-code" value="+91" placeholder="+91" class="w-1/3 bg-white border border-slate-200 rounded-xl px-1.5 py-1.5 focus:outline-none focus:border-blue-500 font-bold text-center text-blue-700" />
                <input type="text" id="cust-contact" maxlength="10" pattern="[0-9]*" oninput="this.value = this.value.replace(/[^0-9]/g, '').slice(0, 10)" placeholder="Mobile No" class="w-2/3 bg-white border border-slate-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-blue-500" />
              </div>
            </div>
            <div class="sm:col-span-2">
              <label class="block font-semibold text-slate-600 mb-0.5 flex justify-between items-center">
                <span>Attached ID Proof <span class="text-[9px] text-blue-600 font-normal">(PDF, 10KB - 900KB)</span></span>
                <button type="button" id="cust-id-file-remove" onclick="removeAttachedIdProof()" class="hidden text-rose-500 hover:text-rose-700 text-[9px] font-bold">Remove</button>
              </label>
              <div class="flex items-center gap-1.5">
                <input type="file" id="cust-id-file" accept="application/pdf" onchange="handleIdProofUpload(event)" class="w-full text-[10px] text-slate-500 file:mr-2 file:py-1 file:px-2.5 file:rounded-xl file:border-0 file:text-[10px] file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 cursor-pointer bg-white border border-slate-200 rounded-xl py-1" />
                <input type="hidden" id="cust-id-file-base64" />
                <input type="hidden" id="cust-id-file-name" />
              </div>
              <p id="cust-id-file-status" class="text-[9px] text-slate-400 mt-0.5 italic">No PDF document attached.</p>
            </div>
          </div>
        </div>
<!-- Room & Stay Schedule Box -->
<div id="sec-room-dates" class="bg-slate-50 p-3 rounded-2xl border border-slate-200/60 space-y-2.5 transition-all">
  <h4 class="text-[9px] font-bold uppercase tracking-wider text-slate-400 flex items-center gap-1.5">
    <i class="fa-solid fa-bed text-blue-600"></i> Room Selection &amp; Stay Dates
  </h4>
  
  <div class="space-y-2">
    <!-- First Row: Room No, Agent Info, Main Guest Count, Extra Guest Count -->
    <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
      <div>
        <label class="block font-semibold text-slate-600 mb-0.5">Room No(s)</label>
        <div class="relative" id="room-dropdown-container">
          <button type="button" onclick="toggleRoomDropdown()" id="room-dropdown-btn" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 focus:outline-none focus:border-blue-500 font-bold text-blue-600 text-left flex justify-between items-center" style="height: 34px;">
            <span id="room-dropdown-text" class="truncate pr-2">Select Rooms...</span>
            <i class="fa-solid fa-chevron-down text-slate-400"></i>
          </button>
          <div id="room-checkboxes" class="hidden absolute z-50 w-full mt-1 bg-white border border-slate-200 rounded-xl shadow-xl max-h-48 overflow-y-auto p-2 space-y-1">
            <!-- Generated Checkboxes Go Here -->
          </div>
        </div>
      </div>

      <div>
        <label class="block font-semibold text-slate-600 mb-0.5">Agent Info</label>
        <select id="cust-agent" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500 font-bold text-slate-700" style="height: 34px;"></select>
      </div>

      <div>
        <label class="block font-semibold text-slate-600 mb-0.5">Main Guest Joined</label>
        <input type="number" id="cust-capacity" min="1" value="1" oninput="calculateModalBilling()" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500 font-bold text-slate-700" style="height: 34px;" />
      </div>

      <div>
        <label class="block font-semibold text-amber-700 mb-0.5 flex items-center gap-1">
          <i class="fa-solid fa-user-plus text-amber-600"></i> Add Extra Guest(s)
        </label>
        <input type="number" id="cust-extra-persons" min="0" value="0" placeholder="0" oninput="calculateModalBilling()" class="w-full bg-amber-50 border border-amber-300 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-amber-500 font-bold text-amber-900" style="height: 34px;" />
      </div>
    </div>
            <!-- ADDITIONAL PERSON CUSTOM CHECK-IN & CHECK-OUT WINDOW -->
            <div id="sec-extra-person-time-wrapper" class="sm:col-span-4 hidden bg-amber-50/70 p-2.5 rounded-2xl border border-amber-200/80 space-y-2">
              <label class="block font-bold text-amber-900 mb-1 flex items-center gap-1">
                <i class="fa-solid fa-clock-rotate-left text-amber-600"></i> Additional Person Stay Window (Custom Dates Required)
              </label>
              <div class="grid grid-cols-1 sm:grid-cols-2 gap-2">
                <div>
                  <label class="block font-semibold text-amber-800 text-[10px] mb-0.5">Extra Guest Check-In</label>
                  <div class="flex gap-1">
                    <input type="date" id="cust-extra-person-date" onchange="handleExtraPersonDatesChange()" class="w-2/3 bg-white border border-amber-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-amber-500 font-semibold text-amber-900" />
                    <input type="time" id="cust-extra-person-time" onchange="handleExtraPersonDatesChange()" class="w-1/3 bg-white border border-amber-200 rounded-xl px-1.5 py-1.5 focus:outline-none focus:border-amber-500 font-semibold text-amber-900" />
                  </div>
                </div>
                <div>
                  <label class="block font-semibold text-amber-800 text-[10px] mb-0.5">Extra Guest Check-Out</label>
                  <div class="flex gap-1">
                    <input type="date" id="cust-extra-person-out-date" onchange="handleExtraPersonDatesChange()" class="w-2/3 bg-white border border-amber-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-amber-500 font-semibold text-amber-900" />
                    <input type="time" id="cust-extra-person-out-time" onchange="handleExtraPersonDatesChange()" class="w-1/3 bg-white border border-amber-200 rounded-xl px-1.5 py-1.5 focus:outline-none focus:border-amber-500 font-semibold text-amber-900" />
                  </div>
                </div>
              </div>
            </div>

            <div class="sm:col-span-4 grid grid-cols-2 gap-2 pt-1 border-t border-slate-200/60">
              <div>
                <label class="block font-semibold text-slate-600 mb-0.5"><i class="fa-solid fa-plane-arrival text-emerald-600 mr-1"></i> Check In</label>
                <div class="flex gap-1">
                  <input type="date" id="cust-checkin-date" onchange="handleStayDatesChange()" required="" class="w-2/3 bg-white border border-slate-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
                  <input type="time" id="cust-checkin-time" onchange="handleStayDatesChange()" required="" class="w-1/3 bg-white border border-slate-200 rounded-xl px-1.5 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
                </div>
              </div>

              <div>
                <label class="block font-semibold text-slate-600 mb-0.5"><i class="fa-solid fa-plane-departure text-rose-500 mr-1"></i> Check Out</label>
                <div class="flex gap-1">
                  <input type="date" id="cust-checkout-date" onchange="handleStayDatesChange()" required="" class="w-2/3 bg-white border border-slate-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
                  <input type="time" id="cust-checkout-time" onchange="handleStayDatesChange()" required="" class="w-1/3 bg-white border border-slate-200 rounded-xl px-1.5 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
                </div>
              </div>
            </div>

            <!-- Extended Check Out Checkbox & Fields -->
            <div id="sec-extended-checkout-wrapper" class="sm:col-span-4 pt-2 border-t border-slate-200/60">
              <div class="flex items-center gap-2 mb-1">
                <input type="checkbox" id="cust-has-extended-checkout" onchange="toggleExtendedCheckoutFields(this.checked)" class="w-4 h-4 text-blue-600 rounded-md border-slate-300 focus:ring-blue-500 cursor-pointer" />
                <label for="cust-has-extended-checkout" id="lbl-has-extended-checkout" class="font-bold text-slate-700 cursor-pointer flex items-center gap-1 select-none text-[11px]">
                  <i class="fa-solid fa-clock-rotate-left text-blue-600"></i> Extended Check-out Date &amp; Time <span id="ext-checkout-timer-notice" class="text-[9px] text-amber-700 font-normal ml-1 hidden">(Active post check-out)</span>
                </label>
              </div>

              <div id="extended-checkout-container" class="hidden bg-blue-50/70 p-2.5 rounded-2xl border border-blue-200/80 mt-1.5">
                <label class="block font-bold text-blue-900 mb-1 flex items-center gap-1">
                  <i class="fa-solid fa-calendar-plus text-blue-600"></i> New Check-out Date &amp; Time
                </label>
                <div class="flex gap-1.5">
                  <input type="date" id="cust-ext-checkout-date" onchange="handleStayDatesChange()" class="w-2/3 bg-white border border-blue-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-blue-500 font-semibold text-blue-900" />
                  <input type="time" id="cust-ext-checkout-time" onchange="handleStayDatesChange()" class="w-1/3 bg-white border border-blue-200 rounded-xl px-1.5 py-1.5 focus:outline-none focus:border-blue-500 font-semibold text-blue-900" />
                </div>
              </div>
            </div>

            <!-- Trip Cancellation Checkbox & Dates -->
            <div id="sec-trip-cancel-wrapper" class="sm:col-span-4 pt-2 border-t border-slate-200/60">
              <div class="flex items-center gap-2 mb-1">
                <input type="checkbox" id="cust-has-trip-cancel" onchange="toggleTripCancelFields(this.checked)" class="w-4 h-4 text-rose-600 rounded-md border-slate-300 cursor-pointer" />
                <label for="cust-has-trip-cancel" class="font-bold text-rose-700 cursor-pointer flex items-center gap-1 select-none text-[11px]">
                  <i class="fa-solid fa-ban text-rose-600"></i> Trip Cancelled <span id="trip-cancel-notice" class="text-[9px] text-rose-600 font-normal ml-1">(Active for Live Booking)</span>
                </label>
              </div>
              <div id="trip-cancel-container" class="hidden bg-rose-50/70 p-2.5 rounded-2xl border border-rose-200/80 mt-1.5">
                <label class="block font-bold text-rose-900 mb-1">Cancellation Check-In &amp; Check-Out Date &amp; Time</label>
                <div class="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  <div><label class="block text-[10px] font-semibold text-rose-800">Cancel Check-In</label><div class="flex gap-1"><input type="date" id="cust-cancel-checkin-date" onchange="handleExtraPersonDatesChange()" class="w-2/3 bg-white border border-rose-200 rounded-xl px-2 py-1.5"><input type="time" id="cust-cancel-checkin-time" onchange="handleExtraPersonDatesChange()" class="w-1/3 bg-white border border-rose-200 rounded-xl px-1.5 py-1.5"></div></div>
                  <div><label class="block text-[10px] font-semibold text-rose-800">Cancel Check-Out</label><div class="flex gap-1"><input type="date" id="cust-cancel-checkout-date" onchange="handleExtraPersonDatesChange()" class="w-2/3 bg-white border border-rose-200 rounded-xl px-2 py-1.5"><input type="time" id="cust-cancel-checkout-time" onchange="handleExtraPersonDatesChange()" class="w-1/3 bg-white border border-rose-200 rounded-xl px-1.5 py-1.5"></div></div>
                </div>
              </div>
            </div>

            <!-- Meal Plan Inclusions Checkbox -->
            <div class="sm:col-span-4 pt-2 border-t border-slate-200/60">
              <div class="flex items-center gap-2">
                <input type="checkbox" id="cust-include-meals" checked="" class="w-4 h-4 text-blue-600 rounded-md border-slate-300 focus:ring-blue-500 cursor-pointer" />
                <label for="cust-include-meals" class="font-bold text-slate-700 cursor-pointer flex items-center gap-1 select-none text-[11px]">
                  <i class="fa-solid fa-utensils text-emerald-600"></i> Include Meal (*Include Breakfast, Lunch, Evening snack &amp; Dinner)
                </label>
              </div>
            </div>

          </div>
        </div>

        <!-- EXTRA FOOD SECTION WITH DATE & TIME -->
        <div id="sec-extra-food" class="bg-amber-50/60 p-3 rounded-2xl border border-amber-200/80 space-y-2.5 transition-all">
          <div class="flex justify-between items-center">
            <h4 class="text-[9px] font-bold uppercase tracking-wider text-amber-800 flex items-center gap-1.5">
              <i class="fa-solid fa-utensils text-amber-600"></i> Extra Food / Drink Orders List
            </h4>
            <button type="button" id="btn-add-food-order" onclick="addFoodOrderItem()" class="bg-amber-600 hover:bg-amber-700 text-white px-2.5 py-1 rounded-full text-[10px] font-semibold flex items-center gap-1 transition shadow-sm">
              <i class="fa-solid fa-plus text-[9px]"></i> Add Food Order
            </button>
          </div>
          
          <div id="food-orders-container" class="space-y-2 max-h-40 overflow-y-auto pr-1"></div>
        </div>

        <!-- CAB FARE SECTION -->
        <div id="sec-cab-fare" class="bg-indigo-50/40 p-3 rounded-2xl border border-indigo-200/80 space-y-2.5 transition-all">
          <div class="flex justify-between items-center">
            <h4 class="text-[9px] font-bold uppercase tracking-wider text-indigo-700 flex items-center gap-1.5">
              <i class="fa-solid fa-taxi text-indigo-600"></i> Cab Fare Details
            </h4>
            <button type="button" id="btn-add-cab-trip" onclick="addCabTripRow()" class="bg-indigo-600 hover:bg-indigo-700 text-white px-2.5 py-1 rounded-full text-[10px] font-semibold flex items-center gap-1 transition shadow-sm">
              <i class="fa-solid fa-plus text-[9px]"></i> Add Cab Trip
            </button>
          </div>
          <div id="cab-trips-container" class="space-y-2 max-h-40 overflow-y-auto pr-1 mt-2"></div>
        </div>

        <!-- BOOKING REMARK SECTION: separate row like Add Cab Trip / Add Food Order -->
        <div id="sec-booking-remark" class="bg-indigo-50/40 p-3 rounded-2xl border border-indigo-200/80 space-y-2.5 transition-all">
          <div class="flex justify-between items-center">
            <h4 class="text-[9px] font-bold uppercase tracking-wider text-indigo-700 flex items-center gap-1.5">
              <i class="fa-solid fa-comment-dots text-indigo-600"></i> Booking Remark
            </h4>
            <button type="button" id="btn-toggle-booking-remark" onclick="toggleBookingRemark()" class="bg-lime-100 hover:bg-lime-200 text-green-800 border border-green-200 px-2.5 py-1 rounded-full text-[10px] font-semibold inline-flex items-center gap-1 transition">
              <i class="fa-solid fa-plus text-[9px]"></i> Add Remark
            </button>
          </div>
          <div id="booking-remark-editor" class="hidden">
            <label for="cust-booking-remark" class="block text-[10px] font-semibold text-indigo-700 mb-1">Booking Remark <span class="font-normal text-slate-500">(maximum 1000 words)</span></label>
            <textarea id="cust-booking-remark" rows="3" maxlength="12000" oninput="updateBookingRemarkCount(); keepBookingRemarkOpenIfFilled()" placeholder="Add booking remark..." class="w-full bg-white border border-indigo-200 rounded-xl px-3 py-2 text-[11px] focus:outline-none focus:border-indigo-500 resize-y"></textarea>
            <p id="cust-booking-remark-count" class="text-[9px] text-slate-500 text-right">0 / 1000 words</p>
          </div>
        </div>

        <!-- Billing Calculation Box -->
<div id="sec-billing-summary" class="bg-blue-50/40 p-3 rounded-2xl border border-blue-100 space-y-2.5 transition-all">
  <h4 class="text-[9px] font-bold uppercase tracking-wider text-blue-700 flex items-center gap-1.5">
    <i class="fa-solid fa-calculator text-blue-600"></i> Billing Summary
  </h4>
  
  <div class="space-y-2">
    <!-- First row: Extra guest price/day, Extra guest stay days, Extra guest Total price, Total cab fare -->
    <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
      <div>
        <label class="block font-semibold text-amber-700 mb-0.5 flex items-center gap-1">
          <i class="fa-solid fa-indian-rupee-sign text-amber-600"></i> Extra Guest Price/Day (₹)
        </label>
        <input type="number" id="cust-extra-rate" min="0" value="0" placeholder="0" oninput="calculateModalBilling()" class="w-full bg-amber-50 border border-amber-300 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-amber-500 font-bold text-amber-900" />
      </div>
      <div>
        <label class="block font-semibold text-slate-600 mb-0.5">Extra Guest Stay Days</label>
        <input type="number" id="cust-extra-days" readonly class="w-full bg-slate-200/60 font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" value="0" />
      </div>
      <div>
        <label class="block font-semibold text-slate-600 mb-0.5">Extra Guest Total Price (₹)</label>
        <input type="number" id="cust-extra-total" readonly class="w-full bg-slate-200/60 font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" value="0" />
      </div>
      <div>
        <label class="block font-semibold text-slate-600 mb-0.5">Total Cab Fare (₹)</label>
        <input type="number" id="cust-cab-total" readonly class="w-full bg-slate-200/60 font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" value="0" />
      </div>
    </div>

    <!-- Second row: Main guest price/Day, Main guest Stay Days, Main guest Total price, Total extra food & drinks -->
    <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
      <div>
        <label class="block font-semibold text-slate-600 mb-0.5">Main Guest Price/Day (₹)</label>
        <input type="number" id="cust-price" value="1200" oninput="calculateModalBilling()" class="w-full bg-white font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-blue-500" />
      </div>
      <div>
        <label class="block font-semibold text-slate-600 mb-0.5">Main Guest Stay Days</label>
        <input type="number" id="cust-days" readonly class="w-full bg-slate-200/60 font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" value="0" />
      </div>
      <div>
        <label class="block font-semibold text-slate-600 mb-0.5">Main Guest Total Price (₹)</label>
        <input type="number" id="cust-main-total" readonly class="w-full bg-slate-200/60 font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" value="0" />
      </div>
      <div>
        <label class="block font-semibold text-slate-600 mb-0.5">Total Extra Food &amp; Drinks (₹)</label>
        <input type="number" id="cust-food-total" readonly class="w-full bg-slate-200/60 font-bold text-amber-700 border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" value="0" />
      </div>
    </div>

    <!-- Third row: Grand Total, Advanced, Due, Discount %, Discount Amount, Clear Due -->
    <div class="grid grid-cols-2 sm:grid-cols-6 gap-2">
      <div>
        <label class="block font-semibold text-slate-600 mb-0.5">Grand Total (₹)</label>
        <input type="number" id="cust-total" readonly class="w-full bg-slate-200/60 text-blue-700 font-bold border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" value="0" />
      </div>
      <div>
        <label class="block font-semibold text-slate-600 mb-0.5">Advanced (₹)</label>
        <input type="number" id="cust-advance" value="0" oninput="calculateModalBilling()" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500 font-semibold text-emerald-600" />
      </div>
      <div>
        <label class="block font-semibold text-slate-600 mb-0.5">Due (₹)</label>
        <input type="number" id="cust-due" readonly class="w-full bg-slate-200/60 text-rose-700 font-bold border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" value="0" />
      </div>
      <div>
        <label class="block font-semibold text-blue-700 mb-0.5">Discount by %</label>
        <input type="number" id="cust-discount-percent" min="0" max="100" value="0" oninput="handleDiscountPercent(this.value)" class="w-full bg-blue-50 border border-blue-200 font-bold text-blue-800 rounded-xl px-2 py-1.5" />
      </div>
      <div>
        <label class="block font-semibold text-blue-700 mb-0.5">Discount by Amount (₹)</label>
        <input type="number" id="cust-discount-amount" min="0" value="0" oninput="handleDiscountAmount(this.value)" class="w-full bg-blue-50 border border-blue-200 font-bold text-blue-800 rounded-xl px-2 py-1.5" />
      </div>
      <div>
        <label class="block font-semibold text-emerald-700 mb-0.5">Clear Due (₹)</label>
        <input type="number" id="cust-clear-bill" value="0" placeholder="0" oninput="handleClearBillPayment(this.value)" class="w-full bg-emerald-50 border border-emerald-300 font-bold text-emerald-800 rounded-xl px-2 py-1.5 focus:outline-none focus:border-emerald-500" title="Payment amount received against due" />
      </div>
    </div>
  </div>
</div>

        <div class="flex justify-end space-x-2 pt-1">
          <button type="button" onclick="closeBookingModal()" class="px-4 py-1.5 bg-slate-100 text-slate-700 rounded-xl font-semibold transition hover:bg-slate-200">Cancel</button>
          <button type="submit" id="btn-save-booking" class="px-5 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl font-semibold shadow-sm transition">Save Booking</button>
        </div>
      </form>
    </div>
  </div>

   <!-- FIXED & PRINTABLE INVOICE / BOOKING RECEIPT MODAL -->
  <div id="invoice-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-3 sm:p-6 overflow-y-auto">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-xl w-full p-5 sm:p-6 space-y-4 relative my-auto max-h-[92vh] overflow-y-auto" id="printable-invoice">
      
      <!-- Read-Only Notice Bar -->
      <div id="inv-readonly-notice" class="hidden bg-slate-900 text-amber-300 text-[10px] font-bold px-3.5 py-2 rounded-2xl flex items-center justify-between border border-slate-800 no-print">
        <span class="flex items-center gap-1.5">
          <i class="fa-solid fa-lock text-amber-400"></i> Read-Only View Mode (Editing Disabled)
        </span>
        <span class="text-[9px] text-slate-400 font-normal">System Protected</span>
      </div>

      <div class="flex justify-between items-start border-b border-slate-200 pb-3">
        <div class="flex items-center gap-3">
          <img alt="Anaristays logo" class="w-20 h-20 object-contain rounded-xl" src="data:image/jpeg;base64,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">
          <div>
          <h2 class="text-base sm:text-lg font-black text-blue-600 uppercase tracking-wide">Anaristays</h2>
          <p class="text-base sm:text-sm font-medium text-blue-600 tracking-normal">Sanoum Pema Homestay</p>
          <p class="text-[10px] text-slate-500 mt-0.5">Sittong, Village in West Bengal</p>
          <p class="text-[10px] text-slate-500">Phone: +91 9804396541 | Email: demo@gmail.com</p>
          </div>
        </div>
        <div class="text-right">
          <div id="e-invoice-section">
            <span id="inv-badge" class="inline-block bg-blue-50 text-blue-700 text-[9px] font-bold px-2.5 py-0.5 rounded-full uppercase mb-1 border border-blue-100">e-Invoice</span>
            <p id="inv-id-container" class="text-[10px] text-slate-500">Invoice ID: <strong id="inv-id" class="text-blue-800 font-mono">INV-2026-0000001</strong></p>
          </div>
          <p class="text-[10px] text-slate-500">Booking ID: <strong id="inv-booking-id" class="text-blue-600 font-mono">BKG-2026-0000001</strong></p>
          <p class="text-[10px] text-slate-500">Issued On: <strong id="inv-date" class="text-blue-800"></strong></p>
        </div>
      </div>

      <div class="grid grid-cols-2 gap-3 bg-slate-50 p-3 rounded-2xl border border-slate-100 text-[10px] sm:text-[11px]">
        <div class="space-y-0.5">
          <h4 class="font-bold text-slate-400 uppercase text-[9px] tracking-wider mb-1">Guest Information</h4>
          <p class="text-slate-800 font-bold" id="inv-guest-name">-</p>
          <p class="text-slate-600 leading-tight" id="inv-guest-address">Address: -</p>
          <p class="text-slate-600" id="inv-guest-contact">Contact: -</p>
          <p class="text-slate-600" id="inv-guest-id">ID No: -</p>
        </div>
        <div class="space-y-0.5">
          <h4 class="font-bold text-slate-400 uppercase text-[9px] tracking-wider mb-1">Reservation Info</h4>
          <p class="text-slate-800 font-bold" id="inv-room">Room No: -</p>
          <p class="text-slate-600" id="inv-checkin">Check-in: -</p>
          <div id="inv-checkout-container" class="space-y-0.5">
            <p class="text-slate-600 font-medium" id="inv-checkout">Check-out: -</p>
            <p id="inv-trip-cancel-details" class="hidden text-[10px] font-semibold text-rose-700"></p>
            <p class="text-blue-600 font-bold hidden" id="inv-ext-checkout"></p>
          </div>
        </div>
      </div>

      <div class="overflow-x-auto border border-slate-100 rounded-2xl">
        <table class="w-full text-left text-[10px] sm:text-[11px]">
          <thead>
            <tr class="bg-blue-50/70 text-blue-900 border-b border-blue-100">
              <th class="p-2 sm:p-2.5 font-bold">Description</th>
              <th class="p-2 sm:p-2.5 text-center font-bold">Qty / Duration</th>
              <th class="p-2 sm:p-2.5 text-right font-bold">Rate/Day</th>
              <th class="p-2 sm:p-2.5 text-right font-bold">Total Amount</th>
            </tr>
          </thead>
          <tbody id="inv-items-tbody" class="divide-y divide-slate-100"></tbody>
        </table>
      </div>

      <div class="flex justify-end pt-1">
        <div class="w-1/2 space-y-1 text-[10px] sm:text-[11px]">
          <div class="flex justify-between text-slate-600">
            <span>Total Amount:</span>
            <strong id="inv-sum-total" class="text-slate-800">₹0</strong>
          </div>
          <div class="flex justify-between text-emerald-600">
            <span>Advance Payment:</span>
            <strong id="inv-sum-advance">₹0</strong>
          </div>
          <div class="flex justify-between text-rose-600 font-bold border-t border-slate-200 pt-1">
            <span>Balance Due:</span>
            <span id="inv-sum-due">₹0</span>
          </div>
          <div id="inv-discount-percent-row" class="hidden flex justify-between text-blue-700 font-bold border-t border-slate-100 pt-1"><span>Discount by %:</span><strong id="inv-sum-discount-percent">0%</strong></div>
          <div id="inv-discount-amount-row" class="hidden flex justify-between text-blue-700 font-bold"><span>Discount by Amount:</span><strong id="inv-sum-discount-amount">₹0</strong></div>
          <div id="inv-clear-due-row" class="hidden flex justify-between text-emerald-700 font-bold border-t border-slate-100 pt-1">
            <span>Clear Due:</span>
            <strong id="inv-sum-clear-due">₹0</strong>
          </div>
        </div>
      </div>

      <div class="flex items-center justify-between no-print text-[10px]">
        <label class="inline-flex items-center gap-2 font-semibold text-slate-600 cursor-pointer">
          <input type="checkbox" id="inv-show-qr-toggle" onchange="toggleInvoiceQR(this.checked)" class="w-4 h-4 text-blue-600 rounded">
          Show payment QR on receipt
        </label>
      </div>
      <div id="inv-payment-qr" class="hidden flex flex-col items-center border-t border-slate-100 pt-3">
        <img alt="Google Pay payment QR code" class="w-36 h-48 object-contain" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAABaAAAAfQCAYAAAA0HfggAACAAElEQVR4Xuzd+Xfc9Z3n+/kH7j3n3jl9z713zszpmem+OZme9HQn6e4kBEKAsOMNvOJ9N7Yx2MYY431ftFn7vu+SZVn7VqVSSSWV9t2S5ZXFYLYABtMkoO/rfj9VEhCLdBPCF2zy/OFxvt/6Vn2/VahsODz98fv7Hz746FMBAAAAAAAAAPBN+w83HwAAAAAAAAAA4JtAgAYAAAAAAAAAOIIADQAAAAAAAABwBAEaAAAAAAAAAOAIAjQAAAAAAAAAwBEEaAAAAAAAAACAIwjQAAAAAAAAAABHEKABAAAAAAAAAI4gQAMAAAAAAAAAHEGABgAAAAAAAAA4ggANAAAAAAAAAHAEARoAAAAAAAAA4AgCNAAAAAAAAADAEQRoAAAAAAAAAIAjCNAAAAAAAAAAAEcQoAEAAAAAAAAAjiBAAwAAAAAAAAAcQYAGAAAAAAAAADiCAA0AAAAAAAAAcAQBGgAAAAAAAADgCAI0AAAAAAAAAMARBGgAAAAAAAAAgCMI0AAAAAAAAAAARxCgAQAAAAAAAACOIEADAAAAAAAAABxBgAYAAAAAAAAAOIIADQAAAAAAAABwBAEaAAAAAAAAAOAIAjQAAAAAAAAAwBEEaAAAAAAAAACAIwjQAAAAAAAAAABHEKABAAAAAAAAAI4gQAMAAAAAAAAAHEGABgAAAAAAAAA4ggANAAAAAAAAAHAEARoAAAAAAAAA4AgCNAAAAAAAAADAEQRoAAAAAAAAAIAjCNAAAAAAAAAAAEcQoAEAAAAAAAAAjiBAAwAAAAAAAAAcQYAGAAAAAAAAADiCAA0AAAAAAAAAcAQBGgAAAAAAAADgCAI0AAAAAAAAAMARBGgAAAAAAAAAgCMI0AAAAAAAAAAARxCgAQAAAAAAAACOIEADAAAAAAAAABxBgAYAAAAAAAAAOIIADQAAAAAAAABwBAEaAAAAAAAAAOAIAjQAAAAAAAAAwBEEaAAAAAAAAACAIwjQAAAAAAAAAABHEKABAAAAAAAAAI4gQAMAAAAAAAAAHEGABgAAAAAAAAA4ggANAAAAAAAAAHAEARoAAAAAAAAA4AgCNAAAAAAAAADAEQRoAAAAAAAAAIAjCNAAAAAAAAAAAEcQoAEAAAAAAAAAjiBAAwAAAAAAAAAcQYAGAAAAAAAAADiCAA0AAAAAAAAAcAQBGgAAAAAAAADgCAI0AAAAAAAAAMARBGgAAAAAAAAAgCMI0AAAAAAAAAAARxCgAQAAAAAAAACOIEADAAAAAAAAABxBgAYAAAAAAAAAOIIADQAAAAAAAABwBAEaAAAAAAAAAOAIAjQAAAAAAAAAwBEEaAAAAAAAAACAIwjQAAAAAAAAAABHEKABAAAAAAAAAI4gQAMAAAAAAAAAHEGABgAAAAAAAAA4ggANAAAAAAAAAHAEARoAAAAAAAAA4AgCNAAAAAAAAADAEQRoAAAAAAAAAIAjCNAAAAAAAAAAAEcQoAEAAAAAAAAAjiBAAwAAAAAAAAAcQYAGAAAAAAAAADiCAA0AAAAAAAAAcAQBGgAAAAAAAADgCAI0AAAAAAAAAMARBGgAAAAAAAAAgCMI0AAAAAAAAAAARxCgAQAAAAAAAACOIEADAAAAAAAAABxBgAYAAAAAAAAAOIIADQAAAAAAAABwBAEaAAAAAAAAAOAIAjQAAAAAAAAAwBEEaAAAAAAAAACAIwjQAAAAAAAAAABHEKABAAAAAAAAAI4gQAMAAAAAAAAAHEGABgAAAAAAAAA4ggANAAAAAAAAAHAEARoAAAAAAAAA4AgCNAAAAAAAAADAEQRoAAAAAAAAAIAjCNAAAAAAAAAAAEcQoAEAAAAAAAAAjiBAAwAAAAAAAAAcQYAGAAAAAAAAADiCAA0AAAAAAAAAcAQBGgAAAAAAAADgCAI0AAAAAAAAAMARBGgAAAAAAAAAgCMI0AAAAAAAAAAARxCgAQAAAAAAAACOIEADAAAAAAAAABxBgAYAAAAAAAAAOIIADQAAAAAAAABwBAEaAAAAAAAAAOAIAjQAAAAAAAAAwBEEaAAAAAAAAACAIwjQAAAAAAAAAABHEKABAAAAAAAAAI4gQAMAAAAAAAAAHEGABgAAAAAAAAA4ggANAAAAAAAAAHAEARoAAAAAAAAA4AgCNAAAAAAAAADAEQRoAAAAAAAAAIAjCNAAAAAAAAAAAEcQoAEAAAAAAAAAjiBAAwAAAAAAAAAcQYAGAAAAAAAAADiCAA0AAAAAAAAAcAQBGgAAAAAAAADgCAI0AAAAAAAAAMARBGgAAAAAAAAAgCMI0AAAAAAAAAAARxCgAQAAAAAAAACOIEADAAAAAAAAABxBgAYAAAAAAAAAOIIA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oMaxqHFMaj5OBbC8MUADy9gzcc/9emg2/uJnWCYiIiIiIqJsZxybGseoxrGqccxqHLuaj2cBLB8M0MAyYzyj/F+DMxEREREREVEu9P8P0vFjWvNxLoDcxgANLAPGX9AvnuIVzkRERERERJT7Gce2xjEuYzSwPDBAAznKOG/WyRdfTpxbi4iIiIiIiGg5ZhzzGse+nDsayF0M0ECOefaFlxNfxIGIiIiIiIjobMo4FjaOic3HyQCWNgZoIEcY//ToJYZnIiIiIiIiOsszjo05PQeQOxiggSUuMTxzng0iIiIiIiKi/5ZxrMwQDSx9DNDAEmWc34pTbRARERERERGlzjh25hzRwNLFAA0sQS+8+LKYnomIiIiIiIheW8YxtHEsbT6+BpB9DNDAEmL806GXOd0GERERERER0evKOKbmtBzA0sIADSwRJ08xPBMRERERERGlI+MY23zcDSA7GKCBLDPOU8UXGSQiIiIiIiJKb8axNueGBrKPARrIoudPvqxX2J6JiIiIiIiIMpJxzG0ce5uPxwGcOQzQQJZwyg0iIiIiIiKiMxOn5ACyhwEayIJTLzE+ExEREREREZ3JjGNx8/E5gMxjgAbOoGee43zPRERERERERNkqcV7oVzleB5A5DNDAGWJ84YOXGZ+JiIiIiIiIsppxbM4XJwTOHAZo4AxIjM9sz0RERERERERLIuMYnREaODMYoIEMY3wmIiIiIiIiWnoxQgNnBgM0kEHGeaU47QYRERERERHR0ixxOo5XOZ4HkD4M0EAG8QUHiYiIiIiIiJZ2xrG7+XgeQPowQAMZcuolxmciIiIiIiKiXMg4hjcf1wNIDwZoIANOnmJ8JiIiIiIiIsqljGN58/E9gNPHAA2k2fMnXzb/HUZEREREREREOZBxTG8+zgdwehiggTQyvnruK7z4mYiIiIiIiCgnM47pjWN78/E+gNePARpII77oIBEREREREVFuxxclBNKLARpIE877TERERERERLQ84nzQQPowQANp8NwLnPeZiIiIiIiIaDllHOubj/8B/PYYoIE0eJlTbxAREREREREtq4xjffPxP4DfHgM0cJpeeJFXPxMREREREREtx4xjfvMOAOC3wwANnAbjK+Py2mciIiIiIiKi5ZlxzG8c+5v3AACvHQM0cBpOvcT8TERERERERLScM479zXsAgNeOARp4nfjCg0RERERERERnR3xBQuD1Y4AGXqeX+MKDRERERERERGdFxgZg3gUAvDYM0MDrwKufiYiIiIiIiM6ueBU08PowQAOvw0uc+5mIiIiIiIjorMrYAsz7AABrDNDAb+lZXv1MREREREREdFZmbALmnQBAagzQwG/J+Oq3RERERERERHT2ZWwC5p0AQGoM0MBv4dnnXzL/3UNEREREREREZ1HGNmDeCwD8ZgzQwG/h5IucfoOIiIiIiIjobM7YBsx7AYDfjAEa+C28zNk3iIiIiIiIiM7qjG3AvBcA+M0YoIHX6Dm++CARERERERERxTM2AvNuAODVMUADr9GLp3j5MxEREREREREpsRGYdwMAr44BGniNXmF/JiIiIiIiIqJ4xkZg3g0AvDoGaOA14PQbRERERERERPT/xmk4gNeGARp4DU5y+g0iIiIiIiIi+n8ytgLzfgAgGQM08Bq8ZHyJWyIiIiIiIiKiX2dsBeb9AEAyBmjAwjNxRERERERERETmjM3AvCMA+O8YoAELz53k/M9ERERERERElJyxGZh3BAD/HQM0YOFFzv9MRERERERERK+SsRmYdwQA/x0DNGCB8z8TERERERER0avFeaABawzQgIVX2J+JiIiIiIiI6FUyNgPzjgDgv2OABlJ45nm+ACERERERERER/eaM7cC8JwBYxAANpPDcC3wBQiIiIiIiIiL6zRnbgXlPALCIARpI4fmTDNBERERERERE9JsztgPzngBgEQM0kMLJU5wAmoiIiIiIiIh+c8Z2YN4TACxigAZSOPUSAzQRERERERER/eaM7cC8JwBYxAANpPASAzQRERERERERpcjYDsx7AoBFDNBACi+9zABNRERERERERL85Yzsw7wkAFjFAAym8zABNRERERERERCkytgPzngBgEQM0kAL7MxERERERERGlytgOzHsCgEUM0EAKrzBAExEREREREVGKjO3AvCcAWMQADaRARERERERERGSVeU8AsIgBGkiBiIiIiIiIiMgq854AYBEDNJACEREREREREZFV5j0BwCIGaCAFIiIiIiIiIiKrzHsCgEUM0EAKRERERERERERWmfcEAIsYoIEUiIiIiIiIiIisMu8JABYxQAMpUHKPPPKIDh48qEgkopGREWBJ2r59u+68804999xz5k9hOkM988wzuv3227Vt27ak+ydX7Ny5U3ff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ArQOATq9BGgdAnR6CNA6BGgdAnRa3J6AiIUEaMSA4EOArr4E6PQQoHUI0PUhAVqHAJ1eArQOATo9BGgdArQOATotbk9AxEICNGJA8CFAV18CdHoI0DoE6PqQAK1DgE4vAVqHAJ0eArQOAVqHAJ0WtycgYiEBGjEg+BCgqy8BOj0EaB0CdH1IgNYhQKeXAK1DgE4PAVqHAK1DgE6L2xMQsZAAjRgQfHIP0LNnzzaPP/646d27t12olKFMitddd12z3HLLeceXg7UI0NOnTzevvvqqefjhh81DDz2EjrLwueKKK0z37t29+yuWp512mrn00kvNXXfd5Y0fwyeffNK8/fbb9jVXBhLpX3jhBbvIco8tlrvttpvZcMMNTbNmzbzXSQxbtGhhNttsMxut3LFj2LFjR7P99tubtm3bemPHsuwAPXfuXPPOO++YESNGePdYDB955BHz+uuvm48//tgdOhq5B+gVV1zRtG/f3kYx9x6L5c4772zWX39907RpU2/8GG6wwQbm2GOPNddcc413D8RQHhjK3OX000/3ji2W3bp1M4MGDTL333+/N34Mb7rpJtOnTx9z/PHHe2PnonxmymfnAw884B1fDg4ZMsSccMIJZvPNN/fu4RhKgJbXWs+ePe08wx0/B+UhjDxM6tKli3f9YygP/nv06GEGDx7sjR1LmbtIJJYH6WUwY8YMO78+77zzvOOL4dFHH23Pj8zjwcftCYhYSIBGDAg+uQdo2Tn03nvvmaefftrukilDibcywS9rIV22ZQfo77//3owaNcruvu3cubONAvj/KuflnHPOsQtp9/6KpYSYCy+80C523fFjKPfQPffcYyZPnuzeAlFYtGiR3QklCzn32GJ50UUX2cDaqlUr73USw3bt2tldjfIgwB07hrfeeqs59dRTzbbbbuuNHcuyA7Q8rJKF9Pnnn+/dYzE86aSTbPR566233KGjkXuAlgcl8pBEdse691gszz33XPutDNlZ544fw9atW5sf/ehH5uCDD/bugRj+8pe/NL169bLfsHKPLZZXXnmlPU+y89AdP4YSz/v3729DtDt2Lspnpnx25jq3kG/0bLfddmaNNdbw7uFYrrfeembXXXe1u3zd8XNQ3rNlE4l848C9/jGUh/Kyw1oitDt2LOV1Jg/py3pA/+WXX5oJEybYMdzji+G9995rXnnlFbtjH3zcnoCIhQRoxIDgk3uAlvgp4erbb7+1X9EsQ/lq4zHHHGN3jbnHl4O1CNCy80Ym4RIb5Dzh/+tKK61kz4/sdHPvr1g+++yzdked7Dp0x4+h7OCSwC1f6y4L+QrxwoULvWOL5fDhw+1O8bJiwCabbGLDquwicseOoSxub775ZnPggQd6Y8ey7AA9ceJEu9tNdnK791gM5drKTjpZqJdF7gF61VVXtdFHvj3k3mOxXDK3kNjtjh9D2fm5/PLL26/Xu/dADNdaay1zxhln2PcM99hiKQ9ijjvuOLPyyit748dwhx12sPMv+caBO3Yuys5nic9yH7nHl4Nyf8p9Kverew/HUr6dJxsk3LFzUX4iRh4yyBzVvf4x/Oqrr2xklZ8qcceOpTzQu/baa+3PApZBLdY6MvfK9WdcysbtCYhYSIBGDAg+uQfoWiBfr5N4KJNM9/hysBYBWsLqz372s2x/pqRsa/EbhLJzWAK0hBN3/BhutNFGNq6+9tpr7tDZIJFeotKaa67pHV8MN910Uxvp33jjDXfoKMhC+rbbbrO7Pt2xY1l2gJZdXP369TM//OEPvbFjKDFP4uoTTzzhDh2NKgRoifRPPfWUe2jReOyxx+wu4rICdNnKg4yuXbuaZ555xj20aDD/0sl9/oW68nM6EqBHjhzpXv4oSLTl3+CAZcHtCYhYSIBGDAg+LIB0cl8AEaDTS4CuDwjQugRoHQK0DgFah/mXTu7zL9QlQEO94/YERCwkQCMGBB8WQDq5L4AI0OklQNcHBGhdArQOAVqHAK3D/Esn9/kX6hKgod5xewIiFhKgEQOCDwsgndwXQATo9BKg6wMCtC4BWocArUOA1mH+pZP7/At1CdBQ77g9ARELCdCIAcGHBZBO7gsgAnR6CdD1AQFalwCtQ4DWIUDrMP/SyX3+hboEaKh33J6AiIUEaMSA4MMCSCf3BRABOr0E6PqAAK1LgNYhQOsQoHWYf+nkPv9CXQI01DtuT0DEQgI0YkDwYQGkk/sCiACdXgJ0fUCA1iVA6xCgdQjQOsy/dHKff6EuARrqHbcnIGIhARoxIPiwANLJfQFEgE4vAbo+IEDrEqB1CNA6BGgd5l86uc+/UJcADfWO2xMQsZAAjRgQfFgA6eS+ACJAp5cAXR8QoHUJ0DoEaB0CtA7zL53c51+oS4CGesftCYhYSIBGDAg+LIB0cl8AEaDTS4CuDwjQugRoHQK0DgFah/mXTu7zL9QlQEO94/YERCwkQCMGBB8WQDq5L4AI0OklQNcHBGhdArQOAVqHAK3D/Esn9/kX6hKgod5xewIiFhKgEQOCDwsgndwXQATo9BKg6wMCtC4BWocArUOA1mH+pZP7/At1CdBQ77g9ARELCdCIAcGHBZBO7gsgAnR6CdD1AQFalwCtQ4DWIUDrMP/SyX3+hboEaKh33J6AiIUEaMSA4MMCSCf3BRABOr0E6PqAAK1LgNYhQOsQoHWYf+nkPv9CXQI01DtuT0DEQgI0YkDwYQGkk/sCiACdXgJ0fUCA1iVA6xCgdQjQOsy/dHKff6EuARrqHbcnIGIhARoxIPiwANLJfQFEgE4vAbo+IEDrEqB1CNA6BGgd5l86uc+/UJcADfWO2xMQsZAAjRgQfFgA6eS+ACJAp5cAXR8QoHUJ0DoEaB0CtA7zL53c51+oS4CGesftCYhYSIBGDAg+LIB0cl8AEaDTS4CuDwjQugRoHQK0DgFah/mXTu7zL9QlQEO94/YERCwkQCMGBB8WQDq5L4AI0OklQNcHBGhdArQOAVqHAK3D/Esn9/kX6hKgod5xewIiFhKgEQOCDwsgndwXQATo9BKg6wMCtC4BWocArUOA1mH+pZP7/At1CdBQ77g9ARELCdCIAcGHBZBO7gsgAnR6CdD1AQFalwCtQ4DWIUDrMP/SyX3+hboEaKh33J6AiIUEaMSA4CPR6pJLLrGTpx122CG6xx9/vPnjH/9opk2b5g6dDc8995zp06eP2WWXXbzjy8FDDjnEXHnllWbs2LHuoUVBArQs0iW87bjjjt74ObjNNtuY9dZbr7RY0rRpU7PvvvvaGDBq1KhSHDp0qDn33HPtOO7xxfCII44w1157rZk4caJ7C0Thu+++M7NmzTLvvfeed2yxvPPOO+1Cd6+99vKOL4YShnv37m0fNLhjx/DFF180V199tTnxxBO9sWN53HHH2fcLeU2748dQopsE6MMPP9wbO4ZybQcMGGBeeeUV9xaLRu4BulWrVuawww6zwcS9PrGUe+iggw4qLa6WbS0CdNnzr06dOtnXgjwMcK9PDOVB25QpU8zcuXPdQ4tG2QFaPvPls1/mAO75w8V26NDBtG3btrRrUIUALfdPjx49SnutvfXWW+aDDz6wD6Gh9rg9ARELCdCIAcFHdsY8/vjj5qqrrjKXX355dO+44w47efr888/dobNh0qRJ5tFHHzVXXHGFd3w5eMMNN9hF9PTp091Di8a7775rHnnkEfO73/3OGz8He/XqZeOhLETdhUUMZWf4FltsYQ499FAbictQHgBI2Pvtb3/rHV8MhwwZYoPJjBkz3MsfhYULF5rXX3/d3H777d6xxXLJObrsssu844uhPGCQoCT3kzt2LOVhmEQld+xYDhw40I5x3nnneWPH8IILLjB9+/a158odO4YSVZ988knz/vvvu7dYNHIP0D/4wQ/sDvSjjjrKuz6xlAdWW221lf32hzt+DtYiQJc9/7r00kvta01ec+71iWH//v3Nn/70JzN58mT30KJRdoCWz3x5UCLv2e75w8XKt4Z222230t7vqhCg5ZuGe+65pzn99NO910kM5fNS3iskQkPtcXsCIhYSoBEDgs+8efPsrkPZxfL3v/89uh999JH54osvbFzKFTlHM2fOLO0cla3sPv/ss8/MN9984x5aNGRXhpwjd+xc/Mtf/mLj5LbbbustLGLYpEkTs9JKK9mffpDFVhl27NjRPiQZM2aMd3wxlPto9uzZpd1H8udKzJCfBnCPLZay61Niz5tvvukdXwzlPpI4vPfee3tjx3DzzTc3p512ml1Mu2PHUnZvn3nmmTZQuuPHcPfddzcXX3yxDXvu2DGcOnWq/UybP3++e4tFI/cALQ/EWrZsadq0aeNdn1jKny27S+W9zx0/B2sRoMuef8mDGPlck29vudcnhj/+8Y9t4JZNBmVRdoCWnasSn+W92z1/uFi5BvKtm3bt2nnnL4ZyL+UeoOWhnvy00brrruu9TmIo35KQzQVlfZMRwrg9ARELCdCIAQEAloYssmT3tvyEiLuwyMXcf4Pw66+/tj+RIZHYPbZY7rHHHvZnRD755BN3+CjIz4fIDuKyHmTIQwz5XV35RkZZDB8+3Jx66qlmtdVW88aPoXydWwJ0zgvp3AM06tYiQJfN6NGj7e7njTfe2Du+GEqQlIdVzz//vDt0NMoO0PITE7LLV+YAsHRefvllc/bZZ9u46p6/GEpgzT1Al20V/g2OnHF7AiIWEqARAwIALA0CdHoI0LoE6PqAAF19CdC6BOjGgACdXgJ0WtyegIiFBGjEgAAAS4MAnR4CtC4Buj4gQFdfArQuAboxIECnlwCdFrcnIGIhARoxIADA0iBAp4cArUuArg8I0NWXAK1LgG4MCNDpJUCnxe0JiFhIgEYMCACwNAjQ6SFA6xKg6wMCdPUlQOsSoBsDAnR6CdBpcXsCIhYSoBEDAgAsDQJ0egjQugTo+oAAXX0J0LoE6MaAAJ1eAnRa3J6AiIUEaMSAAABLgwCdHgK0LgG6PiBAV18CtC4BujEgQKeXAJ0WtycgYiEBGjEgAMDSIECnhwCtS4CuDwjQ1ZcArUuAbgwI0OklQKfF7QmIWEiARgwIALA0CNDpIUDrEqDrAwJ09SVA6xKgGwMCdHoJ0GlxewIiFhKgEQMCACwNAnR6CNC6BOj6gABdfQnQugToxoAAnV4CdFrcnoCIhQRoxIAAAEuDAJ0eArQuAbo+IEBXXwK0LgG6MSBAp5cAnRa3JyBiIQEaMSAAwNIgQKeHAK1LgK4PCNDVlwCtS4BuDAjQ6SVAp8XtCYhYSIBGDAgAsDQI0OkhQOsSoOsDAnT1JUDrEqAbAwJ0egnQaXF7AiIWEqARAwIALA0CdHoI0LoE6PqAAF19CdC6BOjGgACdXgJ0WtyegIiFBGjEgAAAS4MAnR4CtC4Buj4gQFdfArQuAboxIECnlwCdFrcnIGIhARoxIADA0iBAp4cArUuArg8I0NWXAK1LgG4MCNDpJUCnxe0JiFhIgEYMCACwNAjQ6SFA6xKg6wMCdPUlQOsSoBsDAnR6CdBpcXsCIhYSoBEDAgAsDQJ0egjQugTo+oAAXX0J0LoE6MaAAJ1eAnRa3J6AiIUEaMSAAABLgwCdHgK0LgG6PiBAV18CtC4BujEgQKeXAJ0WtycgYiEBGjEgAMDSIECnhwCtS4CuDwjQ1ZcArUuAbgwI0OklQKfF7QmIWEiARgwIALA0CNDpIUDrEqDrAwJ09SVA6xKgGwMCdHoJ0GlxewIiFhKgEQOCz/z5881nn31mpk+fbqZNm4ZYl0ow/PLLL82iReW8jmWMG2+80XTs2NGss8460V177bXNqquuav7t3/7NW1jEkgCtW3aAnjx5srnmmmvMfvvt590DMZSQJHH43nvv9V4jsZSFeufOnU3r1q298xfDsgP0d999Z7766isza9Ys79hi+cADD9g4ufnmm3vXKIZt2rQxrVq1MiussIJ3/mK43HLL2YcZq6++ujd2LOXPbt68uWnSpIk3fgzl3Mg5atu2rTd2DLfaaivTrVs381//9V/e9Y/lp59+aubNm2e+//579zaOgrzGLr30UrPrrrt6xxfD7bbbzvTp06e0cCjkHqBlziJzl5kzZ3rXP5ZyH8lcviwI0LrNmjWzDyTloYz7Oomh3KfyMEkeELvXP4ayBpS1YJn3Uc64PQERCwnQiAHBR4LJk08+aaPJVVddhViX3nfffeaNN94wc+fOdW/hKMjE+9lnn7VjnXXWWdGVXWIHHHBAaTvRRAK0btkBWiLD008/bXfTu/dADCUC9O7d20Yl9zUSy9NPP93svPPONlC65y+GZQfob775xowfP9786U9/8o4tlnL+5Tr06NHDu0Yx/MUvfmGDQ1m70H/wgx/YwHr44Yd7Y8fy0EMPtde6rIguO5TlPj3++OO9sWMo11bi6m9+8xvv+sfykUceMRMnTjQLFy50b+MofPDBB/bbEgMGDPCOL4byOpZA/Le//c0dOhq5B2iJzzJ3kTmMe/1jOHjwYPPYY4+ZSZMmlfYggwCtK5sMfvrTn9pvxrivkxjK+ZHX22WXXebdAzGUNaCsBct8LeeM2xMQsZAAjRgQfF588UXTv39/s+eee5pddtkFsS5dsuuzrHAo0Up2gEgUe+mll6IrX1G+5JJL7ALFXbjEkgCtW3aAlt1DspvozTff9O6BGMpDErnGssh1XyOx3HLLLe2u0qZNm3rnL4ZlB2jZ/SzRrWfPnt6xxfKEE06wi/YRI0Z41yiGd9xxhx2jffv23vmLoewclm97yDG4Y8dSHsIceOCBpX3rY5NNNjFdunSxYckdO4ZPPfWUPYbjjjvOu/6xlK/US/SRz58ykPg5ZcoUM2bMGO/4Yig/ByCbGObMmeMOHY3cA7Q8lAjvdAQAAIAASURBVJT4fNppp3nXP4ayu10ehslPxci3P8qAAK27zTbb2FAs35hwXycx/POf/2wfJB199NHePRBDmRv169fPrgnBx+0JiFhIgEYMCD4yWfrP//zP0haJiDGUnW5XXnml3dGVI7VYABGgdcsO0GUjcfW2224zBx98sHdsuVh2gP7iiy/MzTffbL9x4I4dy/3339/cdNNN5vPPP3eHj8K4ceNM37597cMAd+wYys8BSbyVyFoWsitTfq+8RYsW3vgxlOAjQebdd991h46ChMPrrrvOPpx3x46l7BKXhw187f1/JvcALQ8k5UGPRD537BjKT9wcddRRNnIToJdOFeZfsjP5t7/9rfmP//gPb+wYyrdi5Js38s0h8HF7AiIWEqARA4IPARpzkACtW/YCqGwI0DoEaB0CtC4BWocAXR8QoMMSoHWqMP8iQKfF7QmIWEiARgwIPgRozEECtG7ZC6CyIUDrEKB1CNC6BGgdAnR9QIAOS4DWqcL8iwCdFrcnIGIhARoxIPgQoDEHCdC6ZS+AyoYArUOA1iFA6xKgdQjQ9QEBOiwBWqcK8y8CdFrcnoCIhQRoxIDgQ4DGHCRA65a9ACobArQOAVqHAK1LgNYhQNcHBOiwBGidKsy/CNBpcXsCIhYSoBEDgg8BGnOQAK1b9gKobAjQOgRoHQK0LgFahwBdHxCgwxKgdaow/yJAp8XtCYhYSIBGDAg+BGjMQQK0btkLoLIhQOsQoHUI0LoEaB0CdH1AgA5LgNapwvyLAJ0WtycgYiEBGjEg+BCgMQcJ0LplL4DKhgCtQ4DWIUDrEqB1CND1AQE6LAFapwrzLwJ0WtyegIiFBGjEgOBDgMYcJEDrlr0AKhsCtA4BWocArUuA1iFA1wcE6LAEaJ0qzL8I0GlxewIiFhKgEQOCDwEac5AArVv2AqhsCNA6BGgdArQuAVqHAF0fEKDDEqB1qjD/IkCnxe0JiFhIgEYMCD4EaMxBArRu2QugsiFA6xCgdQjQugRoHQJ0fUCADkuA1qnC/IsAnRa3JyBiIQEaMSD4EKAxBwnQumUvgMqGAK1DgNYhQOsSoHUI0PUBATosAVqnCvMvAnRa3J6AiIUEaMSA4EOAxhwkQOuWvQAqGwK0DgFahwCtS4DWIUDXBwTosARonSrMvwjQaXF7AiIWEqARA4IPARpzkACtW/YCqGwI0DoEaB0CtC4BWocAXR8QoMMSoHWqMP8iQKfF7QmIWEiARgwIPgRozEECtG7ZC6CyIUDrEKB1CNC6BGgdAnR9QIAOS4DWqcL8iwCdFrcnIGIhARoxIPgQoDEHCdC6ZS+AyoYArUOA1iFA6xKgdQjQ9QEBOiwBWqcK8y8CdFrcnoCIhQRoxIDgQ4DGHCRA65a9ACobArQOAVqHAK1LgNYhQNcHBOiwBGidKsy/CNBpcXsCIhYSoBEDgg8BGnOQAK1b9gKobAjQOgRoHQK0LgFahwBdHxCgwxKgdaow/yJAp8XtCYhYSIBGDAg+BGjMQQK0btkLoLIhQOsQoHUI0LoEaB0CdH1AgA5LgNapwvyLAJ0WtycgYiEBGjEg+BCgMQcJ0LplL4DKhgCtQ4DWIUDrEqB1CND1AQE6LAFapwrzLwJ0WtyegIiFBGjEgOBDgMYcJEDrlr0AKhsCtA4BWocArUuA1iFA1wcE6LAEaJ0qzL8I0GlxewIiFhKgEQOCTy0CtEyQsfq61z2mtQrQ33//fSkuWLDgHwsg97zFcvfddzdXX311aQugJbjHFkuJMEsCtHtssVwSoD/++GNv/Bz88ssvbYA+6KCDvNdITN3zFlOJqmUH6FtuucUGaHfsWEqAlshddoDeaqutvLFjuNpqq/0jQLv3WCyXBOiWLVt648dwSYB+5513vLFjKA+plgRod+xYSoCW97wyA7R7XLkpAfrnP/951gF68ODBZtddd/WufwyXW245G6Dvv//+rAN0z549zSuvvOJd/xhWYf4lAfqyyy4zP/7xj72xYyhrQAnQsiYEH7cnIGIhARoxIPiUHaBbt25ttt56axsDZJKM1bNTp05mp512Muuss46dyLr3QAzLDtCys/S9994zzzzzjF3IxVZ2J11xxRXmrLPO8s5fLCXqPfHEE2b27Nnu4UVB4qfsNhwxYoR3fDG89957zaBBg0z37t29Y4tl165dzW9+8xsbfdzxc1D+3nIM22+/vfcaiWWbNm3sn9+xY0fv/MVQQoa8HqZMmeLeYlGYN2+efR0PHDjQGzuW8mfL60DeN8pg6tSp9hzJuXLHjmHnzp3N+eefb6655hrvHoul7AY899xzbTx0x4/hKaecYvr162cfNrhjx1Bea/JeccYZZ3hjx1KC0gsvvGADWRnIZ4E8zJDPBff4clF2xkq8bdq0qfdeFcOyA/Snn35qH8b07t3bu/6xlLnFknhbBmUH6DXWWMMccsgh9n3Vvf4xrML8Sx6ay3rtggsu8MaOoXzLQAL6a6+95g4NhgCNGJIAjRgQfMoO0BtuuOE/JjYyCcRqKosTidC5Bmj5uvUjjzxiLrroIvOzn/0sukcffbTd4SPH4J67WD777LNm0qRJpe2mW7IA6tWrl3d8MTzmmGNsbJCvK7vHFkvZ/Sx/f3nPc8fPwSOOOMLugGrXrp33GonlD3/4Qxv3JOy55y+Gw4YNM2+//bbdqVwG8nXr999/37z44ove2LGUaChjyFhlIOdm/Pjx5sknn/TGjqHEVfk695lnnundY7Hs1q2bDXtDhw71xo/hkCFDbPQ58cQTvbFjKO8REnvkPcMdO5YvvfSSfRCzaFE589PJkyfb81/mdS5b+exfb7317E5f970qhmUHaPk8ls9leSjmXv8YSmCV+Cxzo1wD9EorrWQ222wzs9dee3nXP4ZVmH/JBgD5tsfTTz/tjR3DBx980IwaNcpMnz7dHRoMARoxJAEaMSD4lB2gZSfdpZdeaneXytcDsXrK4lkmsLKLoqxFYtkBWhafsvtWIrocQ2ybNWtmdwLKJN89fzEtawEq/PNXQN3ji2Hz5s3/8RuE7nHFUiKA7Ghs27atN34uLvnKrPsaieXee+9tbrjhBvtQxj1/sSzzPhWWfPXaHTeWS/78Minz7z9r1izzhz/8wey3337e/RVL+Z1y+bmYOXPmeOPH8I033rAPDDfffHNv7BjKNwHk2way090dO5Zl30djxoyxD9w23XRT7/hysez3u7IDtFDma1ks+z4qO0CLco3dax/LKsy/hNzvo5xxewIiFhKgEQOCT9kB+t///d/t11j/+te/ukNDRZAJq+zCkZ0mMtl374EY1iJA/+53vzM77rijN3YMV1hhhX8sgHKlCv8IjuxSkgC95ppreuPjYvfZZx9z44032q+OQzX57LPPbIDed999vesfS/mdcgnQsnOvDN58800boGXnpDt2DOVnASRAy0OrXBk9erTdxb3xxht7x4eLrUWAzp1aBOgyrcL8C9Li9gRELCRAIwYEHwI0LCsEaN0qLIAI0I0hAbr6EKB1CdCNIQFahwANjY7bExCxkACNGBB8CNCwrBCgdauwACJAN4YE6OpDgNYlQDeGBGgdAjQ0Om5PQMRCAjRiQPAhQMOyQoDWrcICiADdGBKgqw8BWpcA3RgSoHUI0NDouD0BEQsJ0IgBwYcADcsKAVq3CgsgAnRjSICuPgRoXQJ0Y0iA1iFAQ6Pj9gRELCRAIwYEHwI0LCsEaN0qLIAI0I0hAbr6EKB1CdCNIQFahwANjY7bExCxkACNGBB8CNCwrBCgdauwACJAN4YE6OpDgNYlQDeGBGgdAjQ0Om5PQMRCAjRiQPAhQMOyQoDWrcICiADdGBKgqw8BWpcA3RgSoHUI0NDouD0BEQsJ0IgBwYcADcsKAVq3CgsgAnRjSICuPgRoXQJ0Y0iA1iFAQ6Pj9gRELCRAIwYEHwI0LCsEaN0qLIAI0I0hAbr6EKB1CdCNIQFahwANjY7bExCxkACNGBB8CNCwrBCgdauwACJAN4YE6OpDgNYlQDeGBGgdAjQ0Om5PQMRCAjRiQPAhQMOyQoDWrcICiADdGBKgqw8BWpcA3RgSoHUI0NDouD0BEQsJ0IgBwYcADcsKAVq3CgsgAnRjSICuPgRoXQJ0Y0iA1iFAQ6Pj9gRELCRAIwYEHwI0LCsEaN0qLIAI0I0hAbr6EKB1CdCNIQFahwANjY7bExCxkACNGBB8CNCwrBCgdauwACJAN4YE6OpDgNYlQDeGBGgdAjQ0Om5PQMRCAjRiQPAhQMOyQoDWrcICiADdGBKgqw8BWpcA3RgSoHUI0NDouD0BEQsJ0IgBwYcADcsKAVq3CgsgAnRjSICuPgRoXQJ0Y0iA1iFAQ6Pj9gRELCRAIwYEHwI0LCsEaN0qLIAI0I0hAbr6EKB1CdCNIQFahwANjY7bExCxkACNGBB8CNCwrBCgdauwACJAN4YE6OpDgNYlQDeGBGgdAjQ0Om5PQMRCAjRiQPAhQMOyQoDWrcICiADdGBKgqw8BWpcA3RgSoHUI0NDouD0BEQsJ0IgBwYcADcsKAVq3CgsgAnRjSICuPgRoXQJ0Y0iA1iFAQ6Pj9gRELCRAIwYEn9wD9IIFC8xHH31k3n77bTtJRt8xY8bYxdXcuXPd0xcFArRuLRZAc+bMMe+//76NDu49EMNHHnnEDBgwwBx99NH2esT2pz/9qenXr5954YUX3EOLRu4BWl5f8nffYostvPMXy+7du9vPhS+++MI9fVmwaNEi88knn5iJEyd693AuSlyV97p58+a5hxeFWgRouZckEI8YMcI7vhjefvvt5vjjjzfrrbeeN3YMV1llFXPEEUeYq6++2hs7F2+55Rb7UG+dddbxjg8X26FDB/uZ8NBDD3nnDxcr7xXyuZ/r52bTpk3tg9WBAwd6xxZLWYPIWkTWJGXw9ddfm+nTp5uxY8d6Y8dw5MiRZtKkSfazAXzcnoCIhQRoxIDgk3uAluj24osvmmuvvdbu0EDfSy65xO4qlThZBgRo3VoE6MmTJ5uHH37Y9O3b17sHYnjuueeaiy++2O6ClmsRWwk9w4YNszutyyL3AL3iiivaHegnn3yyd/5iKSFm3LhxdsGbI/L3HjVqlI0m7j2ci7Ij88knnzQzZsxwDy8KtQjQ7du3t3++vN7c44th586dzY9//GPTunVrb+wYypzoRz/6kenUqZM3di4ee+yx9v1CYrp7fLjYtm3bmt12282ceOKJ3vnDxUp83m677UzLli2985eDMi+Vb0ocfPDB3rHF8pprrrFrEVmTlIF8I0m+jSGf0e7YMZT53V133WU/+8HH7QmIWEiARgwIPrkHaNnpJl/zPeqoo+zXA9FXFukSoWVnbBkQoHVrEaAlukkgltecew/EUM5Nr169zFNPPWWvQ2ynTZtmF1nz5893Dy0auQfo5s2b2/c6ec9zz18sZ86caX824bvvvnNPXxbINz3uvfde+7nm3sO5KLv1rrrqKvPee++5hxeFWgRo+UmdVVdd1e6+dY8vhm3atDGtWrWyuxvdsWMon2UtWrSwP8Xhjp2LZZ+jKigP9STQt2vXzjt/uFj5vJT4vPzyy3vnLwebNGliPztXX31179hiKQ+qbr31VvPxxx+7b7dRmDJlirn++uttRHfHjqH8TE+3bt3sg0/wcXsCIhYSoBEDgk/uAVq+8iY7J3fffXdvbFysBICzzjrLvPTSS+7piwIBWrcWAVp+ukIWEGuttZY3fgw32mgjc/7555vXXnvNHTobcg/QK620kvnlL39pHn30UffQ4P8iPx1y8803mwMOOMA7f7m49dZb24eG48ePdw8vCrUI0IiIjeJPfvIT8/vf/958+OGH7tttFKrwb3DkjNsTELGQAI0YEHwI0NWXAK1DgNYlQKeXAK1DgNYhQCMixpMAXW3cnoCIhQRoxIDgQ4CuvgRoHQK0LgE6vQRoHQK0DgEaETGeBOhq4/YERCwkQCMGBB8CdPUlQOsQoHUJ0OklQOsQoHUI0IiI8SRAVxu3JyBiIQEaMSD4EKCrLwFahwCtS4BOLwFahwCtQ4BGRIwnAbrauD0BEQsJ0IgBwYcAXX0J0DoEaF0CdHoJ0DoEaB0CNCJiPAnQ1cbtCYhYSIBGDAg+BOjqS4DWIUDrEqDTS4DWIUDrEKAREeNJgK42bk9AxEICNGJA8CFAV18CtA4BWpcAnV4CtA4BWocAjYgYTwJ0tXF7AiIWEqARA4IPAbr6EqB1CNC6BOj0EqB1CNA6BGhExHgSoKuN2xMQsZAAjRgQfAjQ1ZcArUOA1iVAp5cArUOA1iFAIyLGkwBdbdyegIiFBGjEgOBDgK6+BGgdArQuATq9BGgdArQOARoRMZ4E6Grj9gRELCRAIwYEHwJ09SVA6xCgdQnQ6SVA6xCgdQjQiIjxJEBXG7cnIGIhARoxIPgQoKsvAVqHAK1LgE4vAVqHAK1DgEZEjCcButq4PQERCwnQiAHBhwBdfQnQOgRoXQJ0egnQOgRoHQI0ImI8CdDVxu0JiFhIgEYMCD4E6OpLgNYhQOsSoNNLgNYhQOsQoBER40mArjZuT0DEQgI0YkDwIUBXXwK0DgFalwCdXgK0DgFahwCNiBhPAnS1cXsCIhYSoBEDgg8BuvoSoHUI0LoE6PQSoHUI0DoEaETEeBKgq43bExCxkACNGBB8CNDVlwCtQ4DWJUCnlwCtQ4DWIUAjIsaTAF1t3J6AiIUEaMSA4EOArr4EaB0CtC4BOr0EaB0CtA4BGhExngToauP2BEQsJEAjBgQfAnT1JUDrEKB1CdDpJUDrEKB1CNCIiPEkQFcbtycgYiEBGjEg+BCgqy8BWocArUuATi8BWocArUOARkSMJwG62rg9ARELCdCIAcGHAF19CdA6BGjd9ddf35x22mnmz3/+sz1fsZ0yZYqZNWuWmTdvnnto0cg9QDdv3tx06tTJ3HLLLd75y8Xp06ebzz//3CxYsMC9PFGYO3euue+++2yo32CDDbJ0v/32M4MHDzaTJk1yDy8Kcv7vvvtu+9nvjo21U96HWrRoYZo0aeK91mO44oormtatW5u1117bGxsX265dO7Pyyiubpk2beucPq6G8vuThrbze3OsfS5n/3nbbbebjjz92326jQIBOi9sTELGQAI0YEHwI0NWXAK0jYYwAHXa11VYze+yxh+nevbs9V7G96qqrzOOPP17ae4WQe4CWoCTvqSeeeKJ3/nLx1ltvte9FM2fOdC9PFL7++mvz6quvmptuusmcc845WTpo0CAzfPjw0mKGPOQZOXKkGTJkiDc21s7DDjvMdOjQobT42aZNG7PbbruZk046yRsbF9u5c2cb9Vq1auWdP6yGMi/dfPPNTceOHb3rH8vrrrvO/OUvfzFz5sxx326jQIBOi9sTELGQAI0YEHwI0NWXAK1DgNaVBYr82VtttZU9T7GVr7D27dvXHkdZ5B6g5fW1xhprmM0228w7f7l4zDHH2PBZ1mfCokWLzCeffGLeeecdG1lzdOzYsWbatGlm/vz57uFFYeHChfYcTZw40Rsba+fll19u9t9/f9OsWTPvtR5DiW5du3Y1DzzwgDc2LlZ+rkcitOyEds8fVkN5wLPPPvvYtYh7/WM5YcIEM2PGjNK+2UOATovbExCxkACNGBB8CNDVlwCtQ4BOby0WQLkH6CooOz4vvvhiG1kBGpnc519VYNSoUebcc8+1P6Pgnj+shrWYf5UNATotbk9A/D/snXewHeV9v8djJ/8kcdxSnMROtzOJ29gzTjx2PHEmP4MLVabbAoEx3QIhmihCIAGiV4GMEZgihCk2iA4C0ZEQiK6C6U0gJASogsT++Kw42nO/73ved8+5e86e3fs8M8/EQXv23Lt3T9ln330XMwnQiAHBpeoHQATouAToOATo8u3FARABunwJ0ADrqfr3rzpAgK6/vfj+1W0I0OViewIiZhKgEQOCS9UPgAjQcQnQcQjQ5duLAyACdPkSoAHWU/XvX3WAAF1/e/H9q9sQoMvF9gREzCRAIwYEl6ofABGg4xKg4xCgy7cXB0AE6PIlQAOsp+rfv+oAAbr+9uL7V7chQJeL7QmImEmARgwILlU/ACJAxyVAxyFAl28vDoAI0OVLgAZYT9W/f9UBAnT97cX3r25DgC4X2xMQMZMAjRgQXKp+AESAjkuAjkOALt9eHAARoMuXAA2wnqp//6oDBOj624vvX92GAF0uticgYiYBGjEguFT9AIgAHZcAHYcAXb69OAAiQJcvARpgPVX//lUHCND1txffv7oNAbpcbE9AxEwCNGJAcKn6ARABOi4BOg4Bunx7cQBEgC5fAjTAeqr+/asOEKDrby++f3UbAnS52J6AiJkEaMSA4FL1AyACdFwCdBwCdPn24gCIAF2+BGiA9VT9+1cdIEDX3158/+o2BOhysT0BETMJ0IgBwaXqB0AE6LgE6DgE6PLtxQEQAbp8CdAA66n69686QICuv734/tVtCNDlYnsCImYSoBEDgkvVD4AI0HEJ0HEI0OXbiwMgAnT5EqAB1lP17191gABdf3vx/avbEKDLxfYERMwkQCMGBJeqHwARoOMSoOMQoMu3FwdABOjyJUADrKfq37/qAAG6/vbi+1e3IUCXi+0JiJhJgEYMCC5VPwAiQMclQMchQJdvLw6ACNDlS4AGWE/Vv3/VAQJ0/e3F969uQ4AuF9sTEDGTAI0YEFyqfgBEgI5LgI5DgC7fXhwAEaDLlwANsJ6qf/+qAwTo+tuL71/dhgBdLrYnIGImARoxILhU/QCIAB2XAB2HAF2+vTgAIkCXLwEaYD1V//5VBwjQ9bcX37+6DQG6XGxPQMRMAjRiQHCp+gEQATouAToOAbp8e3EARIAuXwI0wHqq/v2rDhCg628vvn91GwJ0udiegIiZBGjEgOBS9QMgAnRcAnQcAnT59uIAiABdvgRogPVU/ftXHSBA199efP/qNgTocrE9AREzCdCIAcGl6gdABOi4BOg4BOjy7cUBEAG6fAnQAOup+vevOkCArr+9+P7VbQjQ5WJ7AiJmEqARA4JL1Q+ACNBxCdBxCNDl24sDIAJ0+RKgAdZT9e9fdYAAXX978f2r2xCgy8X2BETMJEAjBgSXqh8AEaDjEqDjEKDLtxcHQATo8iVAA6yn6t+/6gABuv724vtXtyFAl4vtCYiYSYBGDAguVT8AIkDHJUDHIUCXby8OgAjQ5UuABlhP1b9/1QECdP3txfevbkOALhfbExAxkwCNGBBcqn4ARICOS4CO04sAvc022ySXXXZZ8u6773bF22+/PQ3Qn//859PnK9qPfexjyUc/+tHkIx/5iPP7FaEOgLbffvv0INH+bkV56623pgFakd7+ftgbv/KVryRHHnlk8uijj9qXYWGsW7cuee+995y/f1Fq3XqObqH31LVr13b1d+i23d5GdeCaa65Jdtppp+TjH/+48zopQn2eHX/88Wm86hbdfq3pddDN/WjWrFnJQQcdlPzLv/yLs/2qoj6Xu/XdS+ozX8+h7wD2uYuw298t9Bzd/v7VbefPn58cc8wx6TGV/f2KkAAdxvYERMwkQCMGBBcCdP0lQMfpdoDWwZV+B/0dpk6d2hXPPPPMZMyYMcmIESPS0T5Fu8kmmyRf/epXk0996lPO71eEOkj8zne+k4waNcr53YryjDPOSA455JBkxx13dH4/7I2jR49OLr/88uT555+3L8NC0MH6008/ncycOdP5+xelTvYo6q1Zs8Y+fSG8+eabaaC/7rrrnOeuitr+zzzzTBonwc+cOXOSs846q2vvR/o8uPrqq5NFixbZpy6ElStXJgsWLEhP7Nm/f1HeddddyXPPPde1CK3vppdcckmyzz77ONuvCm677bbJt7/97fTEc7cC7mc+85nk61//erLZZps5z1+Em266afK1r30t+fSnP+08dxH24vtXtz31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">
        <p class="text-[9px] text-slate-500 font-semibold">Scan to pay using any UPI app</p>
      </div>
      <div id="inv-booking-remark" class="hidden mb-3 p-3 border border-slate-200 rounded-xl text-[10px] whitespace-pre-wrap break-words"></div>
      <div class="pt-4 border-t border-slate-200/80 flex justify-between items-end text-[10px] text-slate-400">
        <div>
          <p class="font-bold text-slate-600">Thank you for staying with us!</p>
          <p>For inquiries, please contact hotel management.</p>
        </div>
        <div class="text-center border-t border-slate-300 pt-1 w-28">
          <p class="font-semibold text-slate-600">Authorized Signature</p>
        </div>
      </div>

      <div class="flex flex-wrap justify-end space-x-2 gap-y-2 pt-2 no-print border-t border-slate-100">
        <button type="button" onclick="closeInvoiceModal()" class="px-4 py-1.5 bg-slate-100 text-slate-700 rounded-xl font-semibold transition hover:bg-slate-200">Close</button>
        <button type="button" id="inv-whatsapp-btn" onclick="sendReceiptViaWhatsApp()" class="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-semibold shadow-sm flex items-center gap-1.5 transition cursor-pointer">
          <i class="fa-brands fa-whatsapp text-sm"></i> Send receipt via WhatsApp
        </button>
        <button type="button" onclick="window.print()" id="inv-print-btn" class="px-4 py-1.5 bg-blue-600 text-white rounded-xl font-semibold shadow-sm flex items-center gap-1 transition hover:bg-blue-700 cursor-pointer">
          <i class="fa-solid fa-print"></i> Print Invoice
        </button>
      </div>
    </div>
  </div>

  <script>

// Function to check and display the birthday animation on September 27th
function checkBirthdayTrigger() {
      const now = new Date();
      const currentMonth = now.getMonth() + 1; // Month 9 is September
      const currentDay = now.getDate();        // 27th day

      if (currentMonth === 9 && currentDay === 27) {
        const lastShownYear = sessionStorage.getItem('birthday_shown_year');
        const currentYear = now.getFullYear();

        if (lastShownYear !== String(currentYear)) {
          const bModal = document.getElementById('birthday-hurray-modal');
          if (bModal) {
            bModal.classList.remove('hidden');
            sessionStorage.setItem('birthday_shown_year', String(currentYear));
          }
        }
      }
    }

    function closeBirthdayModal() {
      const bModal = document.getElementById('birthday-hurray-modal');
      if (bModal) {
        bModal.classList.add('hidden');
      }
    }

    // System Constants & Robust Helpers
    const MAX_SHEET_ROWS = 10000000;

    function isTrue(val) {
      return val === true || val === 'true' || val === 'TRUE' || val === 1 || val === '1';
    }

    function isInactiveBooking(b) {
      return isTrue(b && b.inactive);
    }

    function getBookingRooms(b) {
      if (!b || b.roomNo === undefined || b.roomNo === null) return [];
      if (Array.isArray(b.roomNo)) return b.roomNo.map(r => String(r).trim());
      return String(b.roomNo).split(/[,|]/).map(s => s.trim());
    }

    function parseDateMs(dtStr) {
      if (!dtStr) return NaN;
      if (typeof dtStr === 'number') return dtStr;
      let sanitized = String(dtStr).trim().replace(' ', 'T');
      let d = new Date(sanitized);
      if (isNaN(d.getTime())) {
        d = new Date(dtStr);
      }
      return d.getTime();
    }
    
    // Extract IST local parts robustly for input fields ensuring UTC+05:30 offset
    function extractISTDateParts(dtStr) {
      if (!dtStr) return { date: '', time: '' };
      let d = new Date(typeof dtStr === 'string' ? dtStr.replace(' ', 'T') : dtStr);
      if (isNaN(d.getTime())) {
        const parts = String(dtStr).replace(' ', 'T').split('T');
        return { date: parts[0] || '', time: parts[1] ? parts[1].substring(0, 5) : '' };
      }
      const utcMs = d.getTime(); 
      const istDate = new Date(utcMs + (330 * 60000));
      
      const yyyy = istDate.getUTCFullYear();
      const mm = String(istDate.getUTCMonth() + 1).padStart(2, '0');
      const dd = String(istDate.getUTCDate()).padStart(2, '0');
      const hh = String(istDate.getUTCHours()).padStart(2, '0');
      const min = String(istDate.getUTCMinutes()).padStart(2, '0');
      
      return { date: `${yyyy}-${mm}-${dd}`, time: `${hh}:${min}` };
    }

    // Convert a Date object to local ISO string adhering to IST (UTC+05:30)
    function toLocalISOString(date) {
        if (!(date instanceof Date) || isNaN(date)) return '';
        const utcMs = date.getTime();
        const istDate = new Date(utcMs + (330 * 60000));
        
        const yyyy = istDate.getUTCFullYear();
        const mm = String(istDate.getUTCMonth() + 1).padStart(2, '0');
        const dd = String(istDate.getUTCDate()).padStart(2, '0');
        const hh = String(istDate.getUTCHours()).padStart(2, '0');
        const min = String(istDate.getUTCMinutes()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}T${hh}:${min}:00+05:30`;
    }

    // NEW helper logic to format dates strictly into 24-hour style "dd/mm/yy hh:mm" for Excel exports (Enforcing IST)
    function format24hDate(dtStr) {
      if (!dtStr) return '';
      const d = new Date(typeof dtStr === 'string' ? dtStr.replace(' ', 'T') : dtStr);
      if (isNaN(d.getTime())) return String(dtStr);
      
      const utcMs = d.getTime();
      const istDate = new Date(utcMs + (330 * 60000));
      
      const dd = String(istDate.getUTCDate()).padStart(2, '0');
      const mm = String(istDate.getUTCMonth() + 1).padStart(2, '0');
      const yy = String(istDate.getUTCFullYear()).slice(-2);
      const hh = String(istDate.getUTCHours()).padStart(2, '0');
      const min = String(istDate.getUTCMinutes()).padStart(2, '0');
      return `${dd}/${mm}/${yy} ${hh}:${min}`;
    }
    
    // Robust parsing for generic JSON array fields (Food Orders, Cab Trips)
    function parseJSONField(fieldData) {
      if (!fieldData) return [];
      if (Array.isArray(fieldData)) return fieldData;
      if (typeof fieldData === 'string' && fieldData.length > 5) {
        try { return JSON.parse(fieldData); } catch (e) {}
      }
      return [];
    }
    
    function checkSheetRowLimits() {
      const currentRowCount = state.bookings.length; 
      
      if (currentRowCount >= (MAX_SHEET_ROWS - 20)) {
        alert("CRITICAL NOTIFICATION: Google Sheet has reached its row limit (less than 20 rows left). Saving has been stopped. Please generate a new sheet and link it to continue.");
        return false;
      } else if (currentRowCount >= (MAX_SHEET_ROWS - 100)) {
        alert("WARNING: Google Sheet is approaching its 10 million row limit (less than 100 rows left). Please prepare a new sheet soon.");
      }
      return true;
    }

    // Double Layer Data Wipe
    function requestDataWipe() {
      if (!isMasterUnlocked) {
        alert("You must unlock Master Data access first to perform a data wipe.");
        openMasterAuthModal();
        return;
      }
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

    async function executeGoogleSheetWipe() {
      const btn = document.getElementById('btn-final-wipe');
      btn.innerText = "WIPING DATA...";
      btn.disabled = true;

      try {
        const payload = { action: "wipeData" };
        const response = await fetch(GAS_API_URL, {
          method: "POST",
          body: JSON.stringify(payload)
        });

        // Try reading error text first safely to avoid JSON parse errors
        const textResult = await response.text();
        try {
            JSON.parse(textResult);
        } catch(e) {
            console.error("Wipe format warning:", textResult);
        }

        // Reset local state
        state.bookings = [];
        state.yearlyCounters = {}; // Clears the sequence counter to start fresh IDs from 01
        state.roomsCapacity = [
          { roomNo: 1, capacity: 4 },
          { roomNo: 2, capacity: 2 },
          { roomNo: 3, capacity: 4 },
          { roomNo: 4, capacity: 4 },
          { roomNo: 5, capacity: 4 }
        ];
        state.masterAgents = [{ agentName: "Self", phone: "Direct", roomNo: "All Rooms" }];
        
        refreshAllUI();
        closeWipeModals();
        alert("Database has been completely wiped.");
      } catch (error) {
        console.error("Wipe error:", error);
        alert("Failed to wipe database. Please check your connection.");
      } finally {
        btn.innerText = "ERASE ALL DATA";
        btn.disabled = false;
      }
    }

    const GAS_API_URL = "https://script.google.com/macros/s/AKfycbx1y0ZQcPX9v66ddWlU8B5xCnOpgGvld39iY3EVNzKQ9tcNcod2onajvq0fM2p6pqExqQ/exec"; 
    
    const ONE_HOUR_MS = 1 * 60 * 60 * 1000;
    let activeModalBooking = null;

    window.addEventListener('beforeunload', function (e) {
      if (isLoggedIn) {
        e.preventDefault();
        e.returnValue = 'Please click "Save Changes" button to save the history.'; 
        return e.returnValue;
      }
    });

    function formatTitleCase(text) {
      if (!text) return '';
      return String(text).replace(/\w\S*/g, function(txt) {
        return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
      });
    }

    function handleStateChange(stateValue) {
      if (stateValue && stateValue.trim().toLowerCase() === 'west bengal') {
        const countryInput = document.getElementById('cust-country');
        if (countryInput) countryInput.value = 'India';
      }
    }

    function getEffectiveCheckoutTime(b) {
      if (!b) return 0;
      if (isTrue(b.hasTripCancel) && b.tripCancelCheckOut) return parseDateMs(b.tripCancelCheckOut);
      if (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) return parseDateMs(b.extendedCheckOut);
      return parseDateMs(b.checkOut);
    }

    function getModalFoodWindow() {
      const hasTripCancel = document.getElementById('cust-has-trip-cancel')?.checked;
      const inDate = hasTripCancel ? (document.getElementById('cust-cancel-checkin-date')?.value || document.getElementById('cust-checkin-date')?.value) : document.getElementById('cust-checkin-date')?.value;
      const inTime = (hasTripCancel ? document.getElementById('cust-cancel-checkin-time')?.value : '') || document.getElementById('cust-checkin-time')?.value || '12:00';
      const hasExtCheckout = document.getElementById('cust-has-extended-checkout')?.checked;
      
      let outDate = document.getElementById('cust-checkout-date')?.value;
      let outTime = document.getElementById('cust-checkout-time')?.value || '11:00';

      if (hasExtCheckout) {
        const extDate = document.getElementById('cust-ext-checkout-date')?.value;
        const extTime = document.getElementById('cust-ext-checkout-time')?.value;
        if (extDate) outDate = extDate;
        if (extTime) outTime = extTime;
      }

      if (!inDate || !outDate) return null;

      const checkInDt = new Date(`${inDate}T${inTime}:00+05:30`);
      const checkOutDt = new Date(`${outDate}T${outTime}:00+05:30`);

      if (isNaN(checkInDt.getTime()) || isNaN(checkOutDt.getTime())) return null;

      const minFoodDt = new Date(checkInDt.getTime() + 15 * 60 * 1000);  
      const maxFoodDt = new Date(checkOutDt.getTime() - 30 * 60 * 1000); 

      return { checkInDt, checkOutDt, minFoodDt, maxFoodDt };
    }

    function validateFoodRowDateTime(inputElem) {
      const row = inputElem.closest('.food-order-row');
      if (!row) return;

      const fDate = row.querySelector('.cust-food-date').value;
      const fTime = row.querySelector('.cust-food-time').value || '00:00';

      if (!fDate) return;

      const foodWin = getModalFoodWindow();
      if (!foodWin) return;

      const selectedDt = new Date(`${fDate}T${fTime}:00+05:30`);

      if (selectedDt < foodWin.minFoodDt || selectedDt > foodWin.maxFoodDt) {
        const minStr = formatDateTime(foodWin.minFoodDt);
        const maxStr = formatDateTime(foodWin.maxFoodDt);
        alert(`⚠️ Extra Food Order time must be after 15 mins of Check-In (${minStr}) and at least 30 mins before Check-Out (${maxStr})!`);
        
        const targetDt = selectedDt < foodWin.minFoodDt ? foodWin.minFoodDt : foodWin.maxFoodDt;
        
        const utcMs = targetDt.getTime();
        const istDate = new Date(utcMs + (330 * 60000));
        
        const yyyy = istDate.getUTCFullYear();
        const mm = String(istDate.getUTCMonth() + 1).padStart(2, '0');
        const dd = String(istDate.getUTCDate()).padStart(2, '0');
        const hh = String(istDate.getUTCHours()).padStart(2, '0');
        const min = String(istDate.getUTCMinutes()).padStart(2, '0');

        row.querySelector('.cust-food-date').value = `${yyyy}-${mm}-${dd}`;
        row.querySelector('.cust-food-time').value = `${hh}:${min}`;
      }
    }

    function handleIdProofUpload(e) {
      const fileInput = e.target;
      const file = fileInput.files[0];
      const statusText = document.getElementById('cust-id-file-status');
      const base64Input = document.getElementById('cust-id-file-base64');
      const fileNameInput = document.getElementById('cust-id-file-name');
      const removeBtn = document.getElementById('cust-id-file-remove');

      if (!file) return;

      if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith('.pdf')) {
        alert("⚠️ Invalid file format! Only PDF files are allowed.");
        fileInput.value = '';
        return;
      }

      const minSize = 10 * 1024;  
      const maxSize = 900 * 1024; 

      if (file.size < minSize || file.size > maxSize) {
        const fileSizeKB = (file.size / 1024).toFixed(1);
        alert(`⚠️ Invalid file size (${fileSizeKB} KB)!\n\nThe attached ID proof PDF must be between 10 KB and 900 KB.`);
        fileInput.value = '';
        return;
      }

      const reader = new FileReader();
      reader.onload = function(evt) {
        base64Input.value = evt.target.result;
        fileNameInput.value = file.name;
        statusText.innerHTML = `<span class="text-emerald-600 font-semibold"><i class="fa-solid fa-circle-check"></i> Attached: ${file.name} (${(file.size / 1024).toFixed(1)} KB)</span>`;
        removeBtn.classList.remove('hidden');
      };
      reader.readAsDataURL(file);
    }

    function removeAttachedIdProof() {
      document.getElementById('cust-id-file').value = '';
      document.getElementById('cust-id-file-base64').value = '';
      document.getElementById('cust-id-file-name').value = '';
      document.getElementById('cust-id-file-status').innerText = 'No PDF document attached.';
      document.getElementById('cust-id-file-remove').classList.add('hidden');
    }

    function openPdfAttachment(base64Data) {
      if (!base64Data) {
        alert("No ID Proof attached!");
        return;
      }
      const win = window.open();
      if (win) {
        win.document.write(`<iframe src="${base64Data}" frameborder="0" style="border:0; top:0px; left:0px; bottom:0px; right:0px; width:100%; height:100%;" allowfullscreen></iframe>`);
      } else {
        alert("Please allow popups to view attached PDF document.");
      }
    }

    function toggleInvoiceQR(show) {
      const qr = document.getElementById('inv-payment-qr');
      if (qr) qr.classList.toggle('hidden', !show);
    }

    function toggleTripCancelFields(checked) {
      const container = document.getElementById('trip-cancel-container');
      if (container) container.classList.toggle('hidden', !checked);
      const extToggle = document.getElementById('cust-has-extended-checkout');
      if (extToggle) {
        if (checked) {
          extToggle.checked = false;
          extToggle.disabled = true;
          document.getElementById('extended-checkout-container')?.classList.add('hidden');
        } else {
          const id = document.getElementById('modal-booking-id')?.value;
          const booking = id ? state.bookings.find(x => String(x.id) === String(id)) : null;
          if (!booking) extToggle.disabled = false;
          else {
            const now = Date.now();
            const checkout = parseDateMs(booking.checkOut);
            const live = now >= parseDateMs(booking.checkIn) && now <= getEffectiveCheckoutTime(booking);
            extToggle.disabled = !(live || (now > checkout && now <= checkout + ONE_HOUR_MS));
          }
        }
      }
      handleExtraPersonDatesChange();
      calculateModalBilling();
    }

    function updateDiscountPercent() {
      const dueBeforeDiscount = Math.max(0, (parseFloat(document.getElementById('cust-total')?.value) || 0) - (parseFloat(document.getElementById('cust-advance')?.value) || 0));
      const pctEl = document.getElementById('cust-discount-percent');
      const amtEl = document.getElementById('cust-discount-amount');
      if (!pctEl || !amtEl) return;
      const amount = Math.max(0, Math.min(dueBeforeDiscount, parseFloat(amtEl.value) || 0));
      amtEl.value = amount.toFixed(2).replace(/\.?0+$/, '');
      pctEl.value = dueBeforeDiscount > 0 ? (amount / dueBeforeDiscount * 100).toFixed(2).replace(/\.?0+$/, '') : '0';
    }

    function handleDiscountPercent(value) {
      const due = Math.max(0, (parseFloat(document.getElementById('cust-total')?.value) || 0) - (parseFloat(document.getElementById('cust-advance')?.value) || 0));
      const pct = Math.max(0, Math.min(100, parseFloat(value) || 0));
      document.getElementById('cust-discount-amount').value = (due * pct / 100).toFixed(2).replace(/\.?0+$/, '');
      recalculateDueAfterDiscount();
    }

    function handleDiscountAmount(value) {
      const due = Math.max(0, (parseFloat(document.getElementById('cust-total')?.value) || 0) - (parseFloat(document.getElementById('cust-advance')?.value) || 0));
      const amount = Math.max(0, Math.min(due, parseFloat(value) || 0));
      document.getElementById('cust-discount-amount').value = amount;
      document.getElementById('cust-discount-percent').value = due > 0 ? (amount / due * 100).toFixed(2).replace(/\.?0+$/, '') : '0';
      recalculateDueAfterDiscount();
    }

    function recalculateDueAfterDiscount() {
      const total = parseFloat(document.getElementById('cust-total')?.value) || 0;
      const advance = parseFloat(document.getElementById('cust-advance')?.value) || 0;
      const discount = parseFloat(document.getElementById('cust-discount-amount')?.value) || 0;
      const clear = parseFloat(document.getElementById('cust-clear-bill')?.value) || 0;
      document.getElementById('cust-due').value = Math.max(0, total - advance - discount - clear);
    }

    function toggleExtendedCheckoutFields(checked) {
      const container = document.getElementById('extended-checkout-container');
      if (!container) return;
      const normalOutDate = document.getElementById('cust-checkout-date').value;
      const normalOutTime = document.getElementById('cust-checkout-time').value;
      const extOutDate = document.getElementById('cust-ext-checkout-date');
      const extOutTime = document.getElementById('cust-ext-checkout-time');

      if (checked) {
        container.classList.remove('hidden');
        if (extOutDate) {
          extOutDate.min = normalOutDate;
          if (!extOutDate.value || extOutDate.value < normalOutDate) {
            extOutDate.value = normalOutDate;
          }
        }
        if (extOutTime && !extOutTime.value) extOutTime.value = normalOutTime || "12:00";
      } else {
        container.classList.add('hidden');
      }
      
      calculateModalBilling();
    }

    let isLoggedIn = false;
    let isMasterUnlocked = false; 
    let inactivityTimer = null;
    let warningTimer = null;
    let countdownInterval = null;
    const INACTIVITY_LIMIT_MS = 10 * 60 * 1000; 
    const WARNING_BUFFER_MS = 1 * 60 * 1000;   

    const DEFAULT_USER_ID = "Admin";
    const DEFAULT_PASSWORD = "Aadmin123";

    let pendingMasterDeleteType = null; 
    let pendingMasterDeleteTarget = null; 

    function openMasterDeleteModal(type, target) {
      pendingMasterDeleteType = type;
      pendingMasterDeleteTarget = target;

      const msgElem = document.getElementById('master-delete-modal-msg');
      if (type === 'booking') {
        const b = state.bookings.find(item => String(item.id) === String(target));
        const bCode = b ? b.bookingCode : 'this booking';
        msgElem.innerText = `Are you sure you want to permanently delete booking ${bCode} from the Master Tab? This action cannot be undone.`;
      } else if (type === 'room') {
        msgElem.innerText = `Are you sure you want to permanently delete this Room Capacity record? This action cannot be undone.`;
      } else if (type === 'agent') {
        msgElem.innerText = `Are you sure you want to permanently delete this Agent record? This action cannot be undone.`;
      }

      document.getElementById('master-delete-confirm-modal').classList.remove('hidden');
    }

    function closeMasterDeleteModal() {
      pendingMasterDeleteType = null;
      pendingMasterDeleteTarget = null;
      document.getElementById('master-delete-confirm-modal').classList.add('hidden');
    }

    function confirmMasterDeletion() {
      if (pendingMasterDeleteType === 'booking') {
        const id = pendingMasterDeleteTarget;
        const idx = state.bookings.findIndex(b => String(b.id) === String(id));
        if (idx !== -1) {
          state.bookings[idx].inactive = true;
        }
        searchMasterBookingById();
        renderBookingsTable();
        updateDashboardCards();
        renderCalendar(defaultAppYear);
        checkUpcomingCheckoutsWithDue();
        saveChanges(false, false);
      } else if (pendingMasterDeleteType === 'room') {
        const index = pendingMasterDeleteTarget;
        state.roomsCapacity.splice(index, 1);
        renderRoomCapacityTable();
        populateRoomDropdown();
        populateAgentDropdown();
        saveChanges(false, false);
      } else if (pendingMasterDeleteType === 'agent') {
        const index = pendingMasterDeleteTarget;
        state.masterAgents.splice(index, 1);
        renderMasterAgentTable();
        populateAgentDropdown();
        saveChanges(false, false);
      }
      closeMasterDeleteModal();
    }

    function closeLoginAlertModal() {
      document.getElementById('login-alert-modal').classList.add('hidden');
    }

    function checkAuthStatus() {
      const sessionAuth = sessionStorage.getItem('app_authenticated');
      if (sessionAuth === 'true') {
        isLoggedIn = true;
        document.getElementById('login-overlay').classList.add('hidden');
        startInactivityMonitoring();
      } else {
        isLoggedIn = false;
        document.getElementById('login-overlay').classList.remove('hidden');
      }
    }

 function handleLogin(e) {
      e.preventDefault();
      const user = document.getElementById('login-userid').value.trim();
      const pass = document.getElementById('login-password').value.trim();

      if (user === DEFAULT_USER_ID && pass === DEFAULT_PASSWORD) {
        isLoggedIn = true;
        sessionStorage.setItem('app_authenticated', 'true');
        document.getElementById('login-overlay').classList.add('hidden');
        document.getElementById('login-error').classList.add('hidden');
        startInactivityMonitoring();
        document.getElementById('login-alert-modal').classList.remove('hidden');
        
        // ---> ADD THIS LINE HERE TO TRIGGER BIRTHDAY CHECK ON LOGIN <---
        checkBirthdayTrigger();

      } else {
        document.getElementById('login-error').classList.remove('hidden');
      }
    }

    function logoutUser(isAuto = false) {
      if (isAuto) {
         processLogoutWithSave();
      } else {
         document.getElementById('logout-confirm-modal').classList.remove('hidden');
      }
    }
    
    function cancelLogout() {
      document.getElementById('logout-confirm-modal').classList.add('hidden');
      resetInactivityTimer();
    }
    
    async function processLogoutWithSave() {
      document.getElementById('logout-warning-modal').classList.add('hidden');
      document.getElementById('logout-confirm-modal').classList.add('hidden');
      
      if (isLoggedIn) {
        document.getElementById('saving-lock-modal').classList.remove('hidden');
        try {
          await saveChanges(true, true);
        } catch (e) {
          console.error("Save on logout error", e);
        }
      }
      
      isLoggedIn = false;
      isMasterUnlocked = false;
      sessionStorage.removeItem('app_authenticated');
      stopInactivityMonitoring();
      
      window.location.reload();
    }

    function openMasterAuthModal() {
      document.getElementById('master-password-input').value = '';
      document.getElementById('master-auth-error').classList.add('hidden');
      document.getElementById('master-auth-modal').classList.remove('hidden');
    }

    function closeMasterAuthModal() {
      document.getElementById('master-auth-modal').classList.add('hidden');
    }

    function handleMasterAuth(e) {
      e.preventDefault();
      const enteredPass = document.getElementById('master-password-input').value.trim();

      if (enteredPass === DEFAULT_PASSWORD) {
        isMasterUnlocked = true;
        closeMasterAuthModal();
        performSwitchTab('master');
      } else {
        document.getElementById('master-auth-error').classList.remove('hidden');
      }
    }

    function startInactivityMonitoring() {
      stopInactivityMonitoring();
      
      const activityEvents = ['mousemove', 'keydown', 'mousedown', 'touchstart', 'scroll'];
      activityEvents.forEach(evt => {
        window.addEventListener(evt, resetInactivityTimer);
      });

      resetInactivityTimer();
    }

    function stopInactivityMonitoring() {
      if (inactivityTimer) clearTimeout(inactivityTimer);
      if (warningTimer) clearTimeout(warningTimer);
      if (countdownInterval) clearInterval(countdownInterval);
      
      const activityEvents = ['mousemove', 'keydown', 'mousedown', 'touchstart', 'scroll'];
      activityEvents.forEach(evt => {
        window.removeEventListener(evt, resetInactivityTimer);
      });
    }

    function resetInactivityTimer() {
      if (!isLoggedIn) return;

      if (inactivityTimer) clearTimeout(inactivityTimer);
      if (warningTimer) clearTimeout(warningTimer);
      if (countdownInterval) clearInterval(countdownInterval);

      document.getElementById('logout-warning-modal').classList.add('hidden');

      warningTimer = setTimeout(showInactivityWarning, INACTIVITY_LIMIT_MS - WARNING_BUFFER_MS);
      inactivityTimer = setTimeout(() => logoutUser(true), INACTIVITY_LIMIT_MS);
    }

    function showInactivityWarning() {
      if (!isLoggedIn) return;

      let secondsLeft = 60;
      document.getElementById('logout-countdown-seconds').innerText = secondsLeft;
      document.getElementById('logout-warning-modal').classList.remove('hidden');

      countdownInterval = setInterval(() => {
        secondsLeft--;
        if (secondsLeft >= 0) {
          document.getElementById('logout-countdown-seconds').innerText = secondsLeft;
        } else {
          clearInterval(countdownInterval);
        }
      }, 1000);
    }

    function formatDateTime(dtStr) {
      if (!dtStr) return '-';
      const d = new Date(typeof dtStr === 'string' ? dtStr.replace(' ', 'T') : dtStr);
      if (isNaN(d.getTime())) {
        const parts = String(dtStr).split('T');
        if (parts.length === 2) {
          const dateParts = parts[0].split('-');
          if (dateParts.length === 3) {
            return `${dateParts[2]}-${dateParts[1]}-${dateParts[0]} ${parts[1].substring(0, 5)}`;
          }
        }
        return String(dtStr).replace('T', ' ');
      }
      
      const utcMs = d.getTime();
      const istDate = new Date(utcMs + (330 * 60000));
      
      const day = String(istDate.getUTCDate()).padStart(2, '0');
      const month = String(istDate.getUTCMonth() + 1).padStart(2, '0');
      const year = istDate.getUTCFullYear();
      const hours = String(istDate.getUTCHours()).padStart(2, '0');
      const minutes = String(istDate.getUTCMinutes()).padStart(2, '0');
      return `${day}-${month}-${year} ${hours}:${minutes}`;
    }

    function formatDate(d) {
      if (!d) return '-';
      const dateObj = typeof d === 'string' ? new Date(d.replace(' ', 'T')) : d;
      if (isNaN(dateObj.getTime())) return d;
      const day = String(dateObj.getDate()).padStart(2, '0');
      const month = String(dateObj.getMonth() + 1).padStart(2, '0');
      const year = dateObj.getFullYear();
      return `${day}-${month}-${year}`;
    }

    const currentRealYear = new Date().getFullYear();
    const defaultAppYear = currentRealYear >= 2026 && currentRealYear <= 2085 ? currentRealYear : 2026;

    let state = {
      yearlyCounters: { [defaultAppYear]: 0 },
      bookings: [],
      roomsCapacity: [
        { roomNo: 1, capacity: 4 },
        { roomNo: 2, capacity: 2 },
        { roomNo: 3, capacity: 4 },
        { roomNo: 4, capacity: 4 },
        { roomNo: 5, capacity: 4 }
      ],
      masterAgents: [
        { agentName: "Self", phone: "Direct", roomNo: "All Rooms" }
      ],
      selectedYear: defaultAppYear,
      dashSelectedYear: defaultAppYear
    };

    function isRoomInMaster(roomNo) {
      if (!state.roomsCapacity) return true;
      let rooms = getBookingRooms({ roomNo });
      return rooms.every(r => state.roomsCapacity.some(m => String(m.roomNo) === String(r)));
    }

    function openExportModal() {
      if (!state.bookings || state.bookings.length === 0) {
        alert("No booking records available to export!");
        return;
      }
      document.getElementById('export-start-date').value = '';
      document.getElementById('export-end-date').value = '';
      document.getElementById('export-modal').classList.remove('hidden');
    }

    function closeExportModal() {
      document.getElementById('export-modal').classList.add('hidden');
    }

    function validateExportDates() {
      const minDate = "2026-08-01";
      const maxDate = "2085-12-31";
      const startInput = document.getElementById('export-start-date');
      const endInput = document.getElementById('export-end-date');

      if (startInput.value && (startInput.value < minDate || startInput.value > maxDate)) {
        alert(`⚠️ Please select a Start Date between ${formatDate(minDate)} and ${formatDate(maxDate)}.`);
        startInput.value = "";
      }
      if (endInput.value && (endInput.value < minDate || endInput.value > maxDate)) {
        alert(`⚠️ Please select an End Date between ${formatDate(minDate)} and ${formatDate(maxDate)}.`);
        endInput.value = "";
      }
      if (startInput.value && endInput.value && startInput.value > endInput.value) {
        alert("⚠️ Start Date cannot be after End Date.");
        endInput.value = "";
      }
    }

    function processExport() {
      const startDateStr = document.getElementById('export-start-date').value;
      const endDateStr = document.getElementById('export-end-date').value;

      if (!startDateStr || !endDateStr) {
        alert("Please select both Start and End dates.");
        return;
      }

      exportToExcel(startDateStr, endDateStr);
    }

    function exportToExcel(startDateStr, endDateStr) {
      if (!state.bookings || state.bookings.length === 0) {
        alert("No booking records available to export!");
        return;
      }

      const filteredBookings = state.bookings.filter(b => {
        if(!b.checkIn) return false;
        
        const bIn = String(b.checkIn).replace(' ', 'T').split('T')[0];
        return (bIn >= startDateStr) && (bIn <= endDateStr);
      });

      if (filteredBookings.length === 0) {
        alert(`No booking records found with a Check-In date between ${formatDate(startDateStr)} and ${formatDate(endDateStr)}!`);
        return;
      }

      const now = new Date().getTime();

      const exportData = filteredBookings.map(b => {
        let bStatus = "Unknown";
        if (isInactiveBooking(b)) {
            bStatus = "Inactive";
        } else {
            const cIn = parseDateMs(b.checkIn);
            const cOut = getEffectiveCheckoutTime(b);
            if (now > cOut) bStatus = "Closed";
            else if (now >= cIn && now <= cOut) bStatus = "Live";
            else bStatus = "Upcoming";
        }

        const foodList = parseJSONField(b.foodOrders);
        const cabList = parseJSONField(b.cabTrips);
        
      return {
          "Booking ID (System)": b.id || "",
          "Booking ID": b.bookingCode || "",
          "Invoice ID": b.invoiceNo || "",
          "Booking Status": bStatus,
          "Guest Name": b.name || "",
          "Address": b.address || "",
          "City": b.city || "",
          "State": b.state || "",
          "Country": b.country || "",
          "Pin/Zip Code": b.zipCode || "",
          "ID Number": b.idNo || "",
          "Country Code": b.countryCode || "",
          "Contact No": b.contactNo || "",
          "Attached ID File Name": b.idProofFileName || "",
          "Room No(s)": getBookingRooms(b).join(" | "),
          "Agent Info": b.agentInfo || "",
          "Main Guest Joined": b.capacity || 1,
          "Add Extra Guest": b.extraPersons || 0,
          "Extra Guest Check-In": format24hDate(b.extraPersonJoined),
          "Extra Guest Check-Out": format24hDate(b.extraPersonOut),
          "Check-In": format24hDate(b.checkIn),
          "Check-Out": format24hDate(b.checkOut),
          "Has Extended Check-Out": isTrue(b.hasExtendedCheckout) ? "Yes" : "No",
          "Extended Check-Out": format24hDate(b.extendedCheckOut),
          "Trip Cancelled": isTrue(b.hasTripCancel) ? "Yes" : "No",
          "Trip Cancel Check-In": format24hDate(b.tripCancelCheckIn),
          "Trip Cancel Check-Out": format24hDate(b.tripCancelCheckOut),
          "Include Meals": (b.includeMeals !== false && b.includeMeals !== 'false') ? "Yes" : "No",
          "Food Orders Details": foodList.map(f => `${f.foodDesc} (${format24hDate(f.foodDateTime)}): ${f.plates} pl @ ₹${f.itemPrice} = ₹${f.foodCharge}`).join('\n'),
          "Cab Trips Details": cabList.map(c => `${c.tripName} (${format24hDate(c.dateTime)}): ₹${c.rate} ${c.remark ? `[${c.remark}]` : ''}`).join('\n'),
	  "Booking Remarks": b.remark || "",
          "Extra Guest Price / Day": b.ExtraperDayPrice || 0,
          "Extra Guest Stay Days": b.extraPersonDays || 0,
          "Extra Guest Total Price": b.custextratotal || 0,
          "Total Cab Fare": b.custcabtotal || 0,
          "Main Guest Price / Day": b.perDayPrice || 0,
          "Main Guest Stay Days": b.noOfDays || 0,
          "Main Guest Total Price": b.custmaintotal || 0,
          "Total Extra Food & Drinks": b.custfoodtotal || 0,
          "Grand Total": b.totalAmount || 0,
          "Advanced": b.initialAdv || 0,
          "Balance Due": b.totalDue || 0,
          "Cleared Bill": b.clearedDue || 0,
          "Discount %": Number(b.discountPercent || 0),
          "Discount Amount": Number(b.discountAmount || 0)
        };
      });

      const worksheet = XLSX.utils.json_to_sheet(exportData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Bookings");

      XLSX.writeFile(workbook, `Booking_Report_${startDateStr}_to_${endDateStr}.xlsx`);
      closeExportModal();
    }

    function searchBookingByDate() {
      const dateVal = document.getElementById('booking-date-search').value;
      renderBookingsTable(dateVal);
    }

    function clearDateSearchBooking() {
      const input = document.getElementById('booking-date-search');
      if (input) input.value = "";
      renderBookingsTable();
    }

    function searchMasterBookingById() {
      const inputElem = document.getElementById('master-booking-search-input');
      if (!inputElem) return;

      const query = inputElem.value.trim().toUpperCase();
      const tbody = document.getElementById('master-delete-tbody');
      if (!tbody) return;

      tbody.innerHTML = '';

      if (!query) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center py-4 text-slate-400">Please type a Booking ID into the search field above to view and delete details.</td></tr>`;
        return;
      }

      const matchedBookings = state.bookings.filter(item => 
        !isInactiveBooking(item) && (item.bookingCode || '').toUpperCase().includes(query)
      );

      if (matchedBookings.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center py-4 text-rose-500 font-semibold">No active booking found matching "${query}".</td></tr>`;
        return;
      }

      matchedBookings.forEach(b => {
        const effectiveOutStr = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
        const tr = document.createElement('tr');
        tr.className = "bg-white hover:bg-slate-50 transition border-b border-slate-100";
        const roomsDisplay = getBookingRooms(b).join(', ');
        
        tr.innerHTML = `
          <td class="py-2.5 px-3 font-mono font-bold text-blue-600">${b.bookingCode}</td>
          <td class="py-2.5 px-3 font-bold text-slate-800">${b.name}</td>
          <td class="py-2.5 px-3"><span class="bg-blue-50 text-blue-700 font-bold px-2 py-0.5 rounded-full text-[10px]">Room ${roomsDisplay}</span></td>
          <td class="py-2.5 px-3 text-[10px] text-slate-600">${formatDateTime(b.checkIn)} to ${formatDateTime(effectiveOutStr)}</td>
          <td class="py-2.5 px-3 font-semibold text-slate-800">₹${b.totalAmount}</td>
          <td class="py-2.5 px-3 font-bold text-rose-600">₹${b.totalDue}</td>
          <td class="py-2.5 px-3 text-center">
            <button onclick="deleteBooking('${b.id}')" class="bg-rose-600 hover:bg-rose-700 text-white px-3 py-1 rounded-full text-[10px] font-semibold flex items-center gap-1 mx-auto transition shadow-xs">
              <i class="fa-solid fa-trash-can text-[9px]"></i> Delete Linked Booking
            </button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function clearMasterBookingSearch() {
      const input = document.getElementById('master-booking-search-input');
      if (input) input.value = '';
      searchMasterBookingById();
    }

    function generateIDsForYear(checkInDateStr) {
      let targetYear = defaultAppYear;
      if (checkInDateStr) {
        targetYear = new Date(checkInDateStr.replace(' ', 'T')).getFullYear() || defaultAppYear;
      }

      if (!state.yearlyCounters) state.yearlyCounters = {};

      if (!state.yearlyCounters[targetYear]) {
        const countForYear = state.bookings.filter(b => {
          return b.checkIn && new Date(b.checkIn.replace(' ', 'T')).getFullYear() === targetYear;
        }).length;
        state.yearlyCounters[targetYear] = countForYear;
      }

      state.yearlyCounters[targetYear] += 1;
      const seq = state.yearlyCounters[targetYear];
      const paddedSeq = String(seq).padStart(7, '0');

      return {
        bookingCode: `BKG-${targetYear}-${paddedSeq}`,
        invoiceNo: `INV-${targetYear}-${paddedSeq}`
      };
    }
    
    function refreshAllUI() {
      if (!state.roomsCapacity || state.roomsCapacity.length === 0) {
        state.roomsCapacity = [
          { roomNo: 1, capacity: 4 },
          { roomNo: 2, capacity: 2 },
          { roomNo: 3, capacity: 4 },
          { roomNo: 4, capacity: 4 },
          { roomNo: 5, capacity: 4 }
        ];
      }
      if (!state.masterAgents || state.masterAgents.length === 0) {
        state.masterAgents = [{ agentName: "Self", phone: "Direct", roomNo: "All Rooms" }];
      }
      
      state.selectedYear = defaultAppYear;
      state.dashSelectedYear = defaultAppYear;
      
      if (!state.yearlyCounters || Object.keys(state.yearlyCounters).length === 0) {
        state.yearlyCounters = { [defaultAppYear]: state.bookings.length || 0 };
      }

      populateRoomDropdown();
      populateAgentDropdown();
      searchMasterBookingById();
      renderBookingsTable();
      renderRoomCapacityTable();
      renderMasterAgentTable();
      renderCalendar(defaultAppYear);
      updateDashboardCards();
    }

    async function loadSavedData() {
      const toast = document.getElementById('toast');
      const msg = document.getElementById('toast-message');
      
      msg.innerText = 'Syncing database...';
      toast.classList.remove('hidden');

      try {
        const response = await fetch(GAS_API_URL + "?action=fetchData");
        const textData = await response.text();
        
        let sheetData;
        try {
            sheetData = JSON.parse(textData);
        } catch(err) {
            console.error("JSON Error: Sync issue.", textData);
            throw new Error("Invalid response format from server (Sync Failed).");
        }
        
        if (sheetData && sheetData.bookings) {
          state = sheetData;
          refreshAllUI(); 
          msg.innerText = 'Database synced successfully!';
        }
      } catch (error) {
        console.error("Error loading data from Google Sheets:", error);
        msg.innerText = 'Failed to connect to Database.';
      }
      
      checkSheetRowLimits();
      setTimeout(() => toast.classList.add('hidden'), 2000);
    }
    
    function setMinBookingDates() {
      const checkInInput = document.getElementById('cust-checkin-date');
      const checkOutInput = document.getElementById('cust-checkout-date');
      const extDateInput = document.getElementById('cust-ext-checkout-date');
      
      if (checkInInput) checkInInput.removeAttribute('min');
      if (checkOutInput) checkOutInput.removeAttribute('min');

      if (checkOutInput && extDateInput) {
        extDateInput.min = checkOutInput.value;
      }
    }

    document.addEventListener("DOMContentLoaded", () => {
      checkAuthStatus();
      loadSavedData();
      setMinBookingDates();
      populateDashboardYearDropdown();
      initDashboard();
      populateCalendarYearDropdown();
      
      document.addEventListener('click', function(e) {
        const container = document.getElementById('room-dropdown-container');
        if (container && !container.contains(e.target)) {
          const boxes = document.getElementById('room-checkboxes');
          if (boxes) boxes.classList.add('hidden');
        }
      });

      checkUpcomingCheckoutsWithDue();
      setInterval(checkUpcomingCheckoutsWithDue, 60000);
      setInterval(triggerPeriodicAutoSave, 300000);
    });

    function refreshDynamicUI() {
      if (document.getElementById('tab-booking') && !document.getElementById('tab-booking').classList.contains('hidden')) {
        renderBookingsTable(document.getElementById('booking-date-search').value);
      }
      if (document.getElementById('tab-dashboard') && !document.getElementById('tab-dashboard').classList.contains('hidden')) {
        updateDashboardCards();
      }
    }

    function triggerPeriodicAutoSave() {
      saveChanges(true, true);
    }

    async function saveChanges(isAutoSave = false, quiet = false) {
      if (!checkSheetRowLimits()) return;
      
      if (!quiet) {
        const toast = document.getElementById('toast');
        const msg = document.getElementById('toast-message');
       	msg.innerText = isAutoSave ? 'Auto-saving to cloud...' : 'Saving to cloud storage...';
        toast.classList.remove('hidden');
      }

      try {
        const payload = {
          action: "saveData",
          state: state
        };

        const response = await fetch(GAS_API_URL, {
          method: "POST",
          headers: {
            "Content-Type": "text/plain;charset=utf-8"
          },
          body: JSON.stringify(payload)
        });

        const textResult = await response.text();
        let result;
        try {
            result = JSON.parse(textResult);
        } catch(e) {
            console.error("Save JSON error", textResult);
            throw new Error("Invalid response format received from server.");
        }

        if (result.status === "success") {
          if (!quiet) {
            const msg = document.getElementById('toast-message');
            msg.innerText = isAutoSave ? 'Changes Auto saved successfully!' : 'Data synced with Cloud Storage!';
            setTimeout(() => document.getElementById('toast').classList.add('hidden'), 3000);
          }
        } else {
          throw new Error(result.message || "Server Error");
        }
      } catch (error) {
        console.error("Error saving to Google Sheets:", error);
        if (!quiet) {
          alert("Saving Error: " + error.message + "\n\nChecks:\n1. Ensure 'Who has access' is set to 'Anyone' in Web App deployment.\n2. Ensure URL in GAS_API_URL is correct.");
          document.getElementById('toast').classList.add('hidden');
        }
      }
    }

    function populateDashboardYearDropdown() {
      const yearSelect = document.getElementById('dash-year-select');
      if (!yearSelect) return;
      yearSelect.innerHTML = '';

      const optConsolidated = document.createElement('option');
      optConsolidated.value = "ALL";
      optConsolidated.text = "All Years (Consolidated)";
      yearSelect.appendChild(optConsolidated);

      for (let y = 2026; y <= 2085; y++) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.text = y === defaultAppYear ? `${y} (Current Year)` : `Year ${y}`;
        yearSelect.appendChild(opt);
      }

      yearSelect.value = defaultAppYear;
      state.dashSelectedYear = defaultAppYear;
    }

    function handleDashboardYearChange(val) {
      if (val === 'CURRENT') {
        val = defaultAppYear;
      }
      
      const select = document.getElementById('dash-year-select');
      if (select) select.value = val;

      if (val === 'ALL') {
        state.dashSelectedYear = 'ALL';
      } else {
        state.dashSelectedYear = parseInt(val);
      }

      initDashboard();
    }

    function checkUpcomingCheckoutsWithDue() {
      const alertBookings = state.bookings.filter(b => {
        if (!isRoomInMaster(b.roomNo) || isInactiveBooking(b)) return false;
        
        const hasDue = (b.totalDue || 0) > 0;
        return hasDue;
      });

      const badge = document.getElementById('alert-badge');
      if (alertBookings.length > 0) {
        badge.innerText = alertBookings.length;
        badge.classList.remove('hidden');
      } else {
        badge.classList.add('hidden');
      }

      renderAlertModalList(alertBookings);
      refreshDynamicUI();
    }

    function renderAlertModalList(alertList) {
      const container = document.getElementById('alert-list-container');
      const textCount = document.getElementById('alert-list-count-text');
      container.innerHTML = '';

      textCount.innerText = `${alertList.length} active warnings found`;

      if (alertList.length === 0) {
        container.innerHTML = `
          <div class="text-center py-8 space-y-1">
            <div class="bg-emerald-50 text-emerald-600 w-10 h-10 rounded-2xl flex items-center justify-center mx-auto text-base">
              <i class="fa-solid fa-circle-check"></i>
            </div>
            <p class="font-bold text-slate-800">No Due Payment Alerts</p>
            <p class="text-slate-400 text-[10px]">All bookings have clear payments with no pending dues.</p>
          </div>
        `;
        return;
      }

      alertList.forEach((b, i) => {
        const effectiveOut = isTrue(b.hasTripCancel) && b.tripCancelCheckOut ? b.tripCancelCheckOut : ((isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut);
        const timeFormatted = formatDateTime(effectiveOut);
        const roomsDisplay = getBookingRooms(b).join(', ');

        const card = document.createElement('div');
        card.className = "bg-amber-50/60 border border-amber-200/80 rounded-2xl overflow-hidden shadow-xs";
        
        const alertMessageText = `Checkout: <strong>${timeFormatted}</strong> | Room ${roomsDisplay} | Guest: <strong>${b.name}</strong> | Total: ₹${b.totalAmount} | Due: ₹${b.totalDue}`;
        const alertBadgeHtml = `<span class="text-[11px] font-black text-rose-600 bg-rose-50 border border-rose-200 px-2.5 py-0.5 rounded-full">₹${b.totalDue.toLocaleString('en-IN')} Due</span>`;

        card.innerHTML = `
          <div class="p-3 flex justify-between items-center cursor-pointer hover:bg-amber-100/50 transition" onclick="toggleAlertDetails('alert-details-${i}')">
            <div class="flex items-center space-x-2.5">
              <span class="bg-amber-500 text-white p-2 rounded-xl text-[10px] font-bold shadow-xs"><i class="fa-solid fa-clock"></i></span>
              <div>
                <h4 class="font-bold text-slate-900 text-[11px] flex items-center gap-1.5">
                  ${b.name} <span class="bg-blue-50 text-blue-700 text-[9px] px-2 py-0.5 rounded-full font-mono">${b.bookingCode || 'N/A'}</span>
                  <span class="bg-slate-100 text-slate-700 text-[9px] px-2 py-0.5 rounded-full font-medium">Room ${roomsDisplay}</span>
                </h4>
                <p class="text-[10px] text-slate-600 mt-0.5">${alertMessageText}</p>
              </div>
            </div>
            <div class="flex items-center space-x-1.5">
              ${alertBadgeHtml}
              <i class="fa-solid fa-chevron-down text-slate-400 text-[10px]"></i>
            </div>
          </div>

          <div id="alert-details-${i}" class="hidden bg-white border-t border-amber-200/60 p-3 space-y-2 text-[10px]">
            <div class="grid grid-cols-2 gap-1 text-slate-600">
              <div>Total Charges: <strong>₹${b.totalAmount}</strong></div>
              <div>Advance Paid: <strong class="text-emerald-600">₹${b.initialAdv || 0}</strong></div>
            </div>
            <div class="flex justify-end pt-1 border-t border-slate-100">
              <button onclick="closeAlertModal(); openBookingModal('${b.id}')" class="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-full font-bold text-[10px] flex items-center gap-1 transition shadow-xs">
                <i class="fa-solid fa-wallet"></i> View / Edit Booking
              </button>
            </div>
          </div>
        `;
        container.appendChild(card);
      });
    }

    function toggleAlertDetails(elemId) {
      const detailsBox = document.getElementById(elemId);
      if (detailsBox) detailsBox.classList.toggle('hidden');
    }

    function openAlertModal() {
      checkUpcomingCheckoutsWithDue();
      document.getElementById('alert-modal').classList.remove('hidden');
    }

    function closeAlertModal() {
      document.getElementById('alert-modal').classList.add('hidden');
    }

    function populateCalendarYearDropdown() {
      const yearSelect = document.getElementById('cal-year-select');
      if (!yearSelect) return;
      yearSelect.innerHTML = '';
      for (let y = 2026; y <= 2085; y++) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.text = y === defaultAppYear ? `${y} (Current Year)` : `Year ${y}`;
        if (y === defaultAppYear) opt.selected = true;
        yearSelect.appendChild(opt);
      }
    }

    function toggleRoomDropdown() {
      document.getElementById('room-checkboxes').classList.toggle('hidden');
    }

    function populateRoomDropdown(selectedRoomNos = []) {
      const container = document.getElementById('room-checkboxes');
      if (!container) return;
      container.innerHTML = '';

      let selArr = [];
      if (Array.isArray(selectedRoomNos)) selArr = selectedRoomNos.map(String);
      else if (selectedRoomNos) selArr = String(selectedRoomNos).split(/[,|]/).map(s => s.trim());

      const allDiv = document.createElement('div');
      allDiv.className = "flex items-center gap-2 mb-1.5 pb-1.5 border-b border-slate-100";
      allDiv.innerHTML = `
        <input type="checkbox" id="room-all" value="ALL" onchange="handleRoomSelection(this)" class="w-3.5 h-3.5 text-blue-600 rounded border-slate-300 focus:ring-blue-500 cursor-pointer room-chk">
        <label for="room-all" class="text-[11px] font-bold text-slate-700 cursor-pointer flex-1">Select all rooms</label>
      `;
      container.appendChild(allDiv);

      state.roomsCapacity.forEach(m => {
        const isChecked = selArr.includes(String(m.roomNo)) ? 'checked' : '';
        const div = document.createElement('div');
        div.className = "flex items-center gap-2 py-1";
        div.innerHTML = `
          <input type="checkbox" id="room-${m.roomNo}" value="${m.roomNo}" ${isChecked} onchange="handleRoomSelection(this)" class="w-3.5 h-3.5 text-blue-600 rounded border-slate-300 focus:ring-blue-500 cursor-pointer room-chk item-chk">
          <label for="room-${m.roomNo}" class="text-[11px] font-bold text-slate-700 cursor-pointer flex-1">Room ${m.roomNo}</label>
        `;
        container.appendChild(div);
      });

      updateRoomDropdownText();
      autoCaptureRoomDetails();
    }

    function handleRoomSelection(chk) {
      const allChk = document.getElementById('room-all');
      const itemChks = document.querySelectorAll('.item-chk');

      if (chk.value === 'ALL') {
        itemChks.forEach(c => c.checked = chk.checked);
      } else {
        const allSelected = Array.from(itemChks).every(c => c.checked);
        if (allChk) allChk.checked = allSelected;
      }
      
      updateRoomDropdownText();
      autoCaptureRoomDetails();
    }

    function updateRoomDropdownText() {
      const itemChks = document.querySelectorAll('.item-chk');
      const checkedVals = Array.from(itemChks).filter(c => c.checked).map(c => c.value);
      const textSpan = document.getElementById('room-dropdown-text');
      
      if (checkedVals.length === 0) {
        textSpan.innerText = "Select Rooms...";
      } else if (checkedVals.length === itemChks.length) {
        textSpan.innerText = "All Rooms Selected";
        const allChk = document.getElementById('room-all');
        if(allChk) allChk.checked = true;
      } else {
        textSpan.innerText = checkedVals.map(r => `Room ${r}`).join(', ');
      }
    }
    
    function getSelectedRooms() {
      const itemChks = document.querySelectorAll('.item-chk');
      if(!itemChks.length) return [];
      const checkedVals = Array.from(itemChks).filter(c => c.checked).map(c => c.value);
      if (checkedVals.length === itemChks.length) return ["ALL"];
      return checkedVals;
    }

    function populateAgentDropdown(selectedAgentName = "") {
      const agentSelect = document.getElementById('cust-agent');
      if (!agentSelect) return;
      agentSelect.innerHTML = '';

      state.masterAgents.forEach(a => {
        const opt = document.createElement('option');
        opt.value = `${a.agentName} (${a.phone})`;
        opt.text = `${a.agentName} (${a.phone})`;
        if (selectedAgentName && opt.value.includes(selectedAgentName)) {
          opt.selected = true;
        }
        agentSelect.appendChild(opt);
      });
    }

    function autoCaptureRoomDetails() {
      let totalCap = 0;
      let allSelected = false;

      const allChk = document.getElementById('room-all');
      if (allChk && allChk.checked) {
         allSelected = true;
      }

      if (allSelected) {
         totalCap = state.roomsCapacity.reduce((sum, m) => sum + (m.capacity || 1), 0);
      } else {
         const itemChks = document.querySelectorAll('.item-chk');
         itemChks.forEach(chk => {
           if (chk.checked) {
             const matched = state.roomsCapacity.find(m => String(m.roomNo) === chk.value);
             if (matched) totalCap += (matched.capacity || 1);
           }
         });
      }
      
      document.getElementById('cust-capacity').value = totalCap > 0 ? totalCap : 1;
      calculateModalBilling();
    }

    function switchTab(tabId) {
      if (tabId === 'master' && !isMasterUnlocked) {
        openMasterAuthModal();
        return;
      }
      performSwitchTab(tabId);
    }

    function performSwitchTab(tabId) {
      document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
      document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active-tab', 'bg-white', 'text-blue-600', 'shadow-sm', 'font-bold');
        btn.classList.add('text-slate-600', 'hover:text-slate-900');
      });

      document.getElementById(`tab-${tabId}`).classList.remove('hidden');
      const activeBtn = document.getElementById(`btn-${tabId}`);
      activeBtn.classList.add('active-tab', 'bg-white', 'text-blue-600', 'shadow-sm', 'font-bold');
      closeCommentBox();

      if (tabId === 'dashboard') {
        handleDashboardYearChange(defaultAppYear);
      }
    }

    function selectDashboardYear(year) {
      handleDashboardYearChange(year);
      renderCalendar(year);
      switchTab('calendar');
    }

    function initDashboard() {
      const grid = document.getElementById('years-grid');
      grid.innerHTML = '';
      
      for (let y = 2026; y <= 2085; y++) {
        const item = document.createElement('div');
        const isSelectedYear = state.dashSelectedYear !== 'ALL' && parseInt(state.dashSelectedYear) === y;
        const isCurrentRealYear = y === defaultAppYear;

        item.className = `text-center py-1.5 px-1 rounded-2xl text-[10px] font-bold cursor-pointer transition ${
          isSelectedYear
            ? 'bg-blue-600 text-white shadow-xs' 
            : (isCurrentRealYear ? 'bg-amber-100 text-amber-900 font-extrabold hover:bg-amber-200' : 'bg-slate-100 text-slate-600 hover:bg-blue-50 hover:text-blue-600')
        }`;
        
        item.innerText = y;
        if (isCurrentRealYear) {
          item.title = "Current Active Year";
        }

        item.onclick = () => selectDashboardYear(y);
        grid.appendChild(item);
      }

      updateDashboardCards();
      const remarksContainer = document.getElementById('dash-all-remarks-container');
      if (remarksContainer && !remarksContainer.classList.contains('hidden')) {
        updateDashboardAllRemarks();
      }
    }
  // Helper to render conic-gradient pie charts (Live = Amber, Upcoming = Blue, Closed = Emerald, Inactive = Slate)
function updatePieChart(elementId, liveVal = null, upcomingVal = null, closedVal = null) {
  const elem = document.getElementById(elementId);
  if (!elem) return;

  // 1. Until data is fetched (null/undefined), display grey placeholder
  if (liveVal === null || upcomingVal === null || closedVal === null) {
    elem.style.background = 'conic-gradient(#e2e8f0 0% 100%)';
    return;
  }

  const live = Number(liveVal) || 0;
  const upcoming = Number(upcomingVal) || 0;
  const closed = Number(closedVal) || 0;
  const total = live + upcoming + closed;

  // 2. If data is fetched but blank (total is 0), keep grey placeholder
  if (total <= 0) {
    elem.style.background = 'conic-gradient(#e2e8f0 0% 100%)';
    return;
  }

  // 3. Active state: Render exclusively Yellow, Blue, and Green
  const p1 = ((live / total) * 100).toFixed(2);
  const p2 = (((live + upcoming) / total) * 100).toFixed(2);

  elem.style.background = `conic-gradient(
    #f59e0b 0% ${p1}%,
    #3b82f6 ${p1}% ${p2}%,
    #10b981 ${p2}% 100%
  )`;
}

// Helper to update Box 6: All Live Booking Details Table (Moved to Top-Level Scope)
function updateDashboardLiveBookings(liveBookings) {
  const container = document.getElementById('dash-today-live-container');
  if (!container) return;

  if (!liveBookings || !Array.isArray(liveBookings) || liveBookings.length === 0) {
    container.innerHTML = '<p class="text-[11px] text-slate-400 italic py-2 text-center">No live bookings active today.</p>';
    return;
  }

  let html = `
    <table class="w-full text-left text-[10px] divide-y divide-slate-100">
      <thead class="bg-amber-50/60 text-amber-900 font-bold sticky top-0 bg-white">
        <tr>
          <th class="py-1 px-1.5">Guest Name</th>
          <th class="py-1 px-1.5">Room No</th>
          <th class="py-1 px-1.5">Total Amount</th>
          <th class="py-1 px-1.5">Due Amount</th>
          <th class="py-1 px-1.5">Remark</th>
        </tr>
      </thead>
      <tbody class="divide-y divide-slate-100 font-medium text-slate-700">
  `;

  liveBookings.forEach(b => {
    // Safe extraction for room numbers and formatted dates
    const roomList = typeof getBookingRooms === 'function' ? getBookingRooms(b) : (b.rooms || []);
    const roomsStr = Array.isArray(roomList) && roomList.length > 0 ? roomList.join(', ') : '-';

    const rawOutDate = b.hasExtendedCheckout && b.extendedCheckOut ? b.extendedCheckOut : b.checkOut;
    const outStr = typeof format24hDate === 'function' ? format24hDate(rawOutDate) : (rawOutDate || '-');

    html += `
      <tr class="hover:bg-amber-50/40 transition">
        <td class="py-1 px-1.5 font-mono text-blue-600 font-bold">${b.name || b.id || 'N/A'}</td>
        <td class="py-1 px-1.5 font-bold text-slate-800">${roomsStr}</td>
        <td class="py-1 px-1.5"><span class="bg-amber-100 text-amber-800 font-bold px-1.5 py-0.5 rounded-md text-[9px]">₹ ${Number(b.totalAmount || 0).toLocaleString('en-IN')}</span></td>
        <td class="py-1 px-1.5 text-slate-600">₹ ${Number(b.totalDue || 0).toLocaleString('en-IN')}</td>
        <td class="py-1 px-1.5 whitespace-pre-wrap break-words">${(b.remark || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</td>
      </tr>
    `;
  });

  html += `</tbody></table>`;
  container.innerHTML = html;
}

// Toggle the dashboard's complete saved-remarks list.
function toggleDashboardAllRemarks() {
  const container = document.getElementById('dash-all-remarks-container');
  const button = document.getElementById('btn-toggle-all-remarks');
  if (!container || !button) return;
  const opening = container.classList.contains('hidden');
  container.classList.toggle('hidden', !opening);
  button.innerHTML = opening
    ? '<i class="fa-solid fa-minus text-[9px]"></i> Hide all remarks'
    : '<i class="fa-solid fa-plus text-[9px]"></i> Show all remarks';
  if (opening) updateDashboardAllRemarks();
}

function escapeRemarkHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getDashboardRemarkBookingStatus(booking) {
  if (typeof isInactiveBooking === 'function' && isInactiveBooking(booking)) return 'Inactive';
  const checkIn = booking.checkIn ? new Date(String(booking.checkIn).replace(' ', 'T')).getTime() : NaN;
  const checkoutRaw = typeof getEffectiveCheckoutTime === 'function'
    ? getEffectiveCheckoutTime(booking)
    : (booking.checkOut ? new Date(String(booking.checkOut).replace(' ', 'T')).getTime() : NaN);
  const now = Date.now();
  if (!Number.isNaN(checkIn) && !Number.isNaN(checkoutRaw)) {
    if (now >= checkIn && now <= checkoutRaw) return 'Live';
    if (now < checkIn) return 'Upcoming';
    return 'Closed';
  }
  return booking.status || booking.bookingStatus || 'Unknown';
}

function updateDashboardAllRemarks() {
  const container = document.getElementById('dash-all-remarks-container');
  if (!container) return;
  const bookings = (typeof state !== 'undefined' && Array.isArray(state.bookings)) ? state.bookings : [];
  const selectedFilter = typeof state !== 'undefined' ? state.dashSelectedYear : 'ALL';
  const targetYear = selectedFilter && selectedFilter !== 'ALL' ? parseInt(selectedFilter, 10) : null;
  const statusOrder = { Live: 0, Upcoming: 1, Closed: 2, Inactive: 3 };
  const remarkBookings = bookings
    .filter(b => {
      if (!String(b.remark || '').trim()) return false;
      // Match the dashboard pie-chart cards: filter by booking check-in year.
      if (targetYear) {
        const checkInMs = b.checkIn ? new Date(String(b.checkIn).replace(' ', 'T')).getTime() : NaN;
        if (!Number.isNaN(checkInMs) && new Date(checkInMs).getFullYear() !== targetYear) return false;
        // Records without a usable check-in date are not assigned to a specific year.
        if (Number.isNaN(checkInMs)) return false;
      }
      return true;
    })
    .map((booking, originalIndex) => ({
      booking,
      originalIndex,
      status: getDashboardRemarkBookingStatus(booking)
    }))
    .sort((a, b) =>
      (statusOrder[a.status] ?? 4) - (statusOrder[b.status] ?? 4) ||
      a.originalIndex - b.originalIndex
    )
    .map(item => item.booking);

  if (!remarkBookings.length) {
    container.innerHTML = '<p class="text-[11px] text-slate-400 italic py-2 text-center">No saved remarks found.</p>';
    return;
  }

  const parseStayDate = value => {
    if (!value) return '-';
    if (typeof format24hDate === 'function') return format24hDate(value);
    return String(value);
  };

  container.innerHTML = remarkBookings.map((b, index) => {
    const checkIn = parseStayDate(b.checkIn);
    const checkOutValue = b.hasExtendedCheckout && b.extendedCheckOut ? b.extendedCheckOut : b.checkOut;
    const checkOut = parseStayDate(checkOutValue);
    const status = getDashboardRemarkBookingStatus(b);
    const statusClass = status === 'Live' ? 'bg-amber-100 text-amber-800'
      : status === 'Upcoming' ? 'bg-blue-100 text-blue-800'
      : status === 'Closed' ? 'bg-emerald-100 text-emerald-800'
      : status === 'Inactive' ? 'bg-rose-100 text-rose-800'
      : 'bg-slate-100 text-slate-700';

    return `
      <article class="rounded-2xl border border-indigo-100 bg-indigo-50/40 p-3 space-y-1.5">
        <div class="flex flex-wrap items-center justify-between gap-2">
          <h4 class="text-[11px] font-bold text-indigo-800">${String(index + 1).padStart(2, '0')}. Booking ID: ${escapeRemarkHtml(b.bookingCode || b.booking_code || (String(b.id || '').startsWith('BKG-') ? b.id : '') || 'N/A')}</h4>
          <span class="text-[9px] font-bold px-2 py-0.5 rounded-full ${statusClass}">Booking status: ${escapeRemarkHtml(status)}</span>
        </div>
        <p class="text-[10px] text-slate-700"><strong>Guest Name:</strong> ${escapeRemarkHtml(b.name || b.guestName || 'N/A')}</p>
        <p class="text-[10px] text-slate-700"><strong>Stay window:</strong> ${escapeRemarkHtml(checkIn)} – ${escapeRemarkHtml(checkOut)}</p>
        <div class="text-[10px] text-slate-700 whitespace-pre-wrap break-words"><strong>Remarks:</strong> ${escapeRemarkHtml(String(b.remark || '').trim())}</div>
      </article>`;
  }).join('');
}

function toggleDashboardAgentDetails() {
  const c = document.getElementById('dash-agent-details-container');
  const b = document.getElementById('btn-toggle-agent-details');
  if (!c || !b) return;
  const opening = c.classList.contains('hidden');
  c.classList.toggle('hidden', !opening);
  b.innerHTML = opening ? '<i class="fa-solid fa-minus text-[9px]"></i> Hide details' : '<i class="fa-solid fa-plus text-[9px]"></i> Show details';
  if (opening) updateDashboardAgentDetails();
}

function updateDashboardAgentDetails() {
  const select = document.getElementById('dash-agent-filter');
  const container = document.getElementById('dash-agent-details-content');
  if (!select || !container) return;
  const current = select.value || 'ALL';
  const agents = ['ALL', ...new Set([...(state.masterAgents || []).map(a => `${a.agentName} (${a.phone})`), 'Self'])];
  select.innerHTML = agents.map(a => `<option value="${escapeRemarkHtml(a)}">${a === 'ALL' ? 'All Agents / Self' : escapeRemarkHtml(a)}</option>`).join('');
  select.value = agents.includes(current) ? current : 'ALL';
  const targetYear = state.dashSelectedYear && state.dashSelectedYear !== 'ALL' ? Number(state.dashSelectedYear) : null;
  const rows = (state.bookings || []).filter(b => {
    const ci = b.checkIn ? new Date(String(b.checkIn).replace(' ', 'T')) : null;
    if (targetYear && (!ci || Number.isNaN(ci.getTime()) || ci.getFullYear() !== targetYear)) return false;
    const agent = String(b.agentInfo || 'Self');
    return select.value === 'ALL' || agent === select.value || agent.startsWith(select.value.split(' (')[0] + ' (');
  });
  const grouped = rows.map(b => {
    const status = getDashboardRemarkBookingStatus(b);
    const out = isTrue(b.hasTripCancel) && b.tripCancelCheckOut ? b.tripCancelCheckOut : ((isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut);
    return {b,status,stay:`${format24hDate(b.checkIn)} – ${format24hDate(out)}`};
  });
  const counts = grouped.reduce((a,x) => { a[x.status]=(a[x.status]||0)+1; return a; },{});
  const statusSummary = ['Upcoming','Live','Closed','Inactive'].map(k => `${k}: ${counts[k]||0}`).join(' | ');
  if (!grouped.length) { container.innerHTML = '<p class="text-[11px] text-slate-400 italic py-3 text-center">No bookings found for this agent and year filter.</p>'; return; }
  container.innerHTML = `<p class="text-[10px] font-semibold text-violet-700 py-2">No. of bookings — ${statusSummary}</p><table class="min-w-full text-[10px] text-left"><thead class="bg-violet-50 text-violet-800"><tr><th class="p-2">Booking Status</th><th class="p-2">Booking ID</th><th class="p-2">Stay Window</th><th class="p-2 text-right">Total Amount</th><th class="p-2 text-right">Advance</th><th class="p-2 text-right">Due</th><th class="p-2 text-right">Clear Due</th></tr></thead><tbody>${grouped.map(({b,status,stay}) => `<tr class="border-b border-slate-100"><td class="p-2">${escapeRemarkHtml(status)}</td><td class="p-2">${escapeRemarkHtml(b.bookingCode || b.booking_code || b.id || 'N/A')}</td><td class="p-2 whitespace-nowrap">${escapeRemarkHtml(stay)}</td><td class="p-2 text-right">₹${Number(b.totalAmount||0).toLocaleString('en-IN')}</td><td class="p-2 text-right">₹${Number(b.initialAdv||0).toLocaleString('en-IN')}</td><td class="p-2 text-right">₹${Number(b.totalDue||0).toLocaleString('en-IN')}</td><td class="p-2 text-right">₹${Number(b.clearedDue||0).toLocaleString('en-IN')}</td></tr>`).join('')}</tbody></table>`;
}

// Main Dashboard Update Function
function updateDashboardCards() {
  const selectedFilter = typeof state !== 'undefined' ? state.dashSelectedYear : null;
  const label = document.getElementById('dash-filter-label');
  const targetYear = (selectedFilter && selectedFilter !== 'ALL') ? parseInt(selectedFilter, 10) : null;

  // 1. Update Filter Header Label
  if (label) {
    if (!targetYear) {
      label.innerText = "Consolidated Summary (All Years)";
    } else {
      const defaultYr = typeof defaultAppYear !== 'undefined' ? defaultAppYear : new Date().getFullYear();
      label.innerText = targetYear === defaultYr
        ? `Year ${targetYear} (Current Year)`
        : `Year ${targetYear}`;
    }
  }

  const now = Date.now();
  const liveBookingsList = []; // Array to store live bookings for Box 6

  // 2. Initialize Metrics Object
  const stats = {
    live: { count: 0, amount: 0, adv: 0, due: 0 },
    upcoming: { count: 0, amount: 0, adv: 0, due: 0 },
    closed: { count: 0, amount: 0, adv: 0, due: 0 },
    inactive: { count: 0, amount: 0, adv: 0, due: 0 }
  };

  const parseDate = (dtStr) => {
    if (!dtStr) return NaN;
    return new Date(String(dtStr).replace(' ', 'T')).getTime();
  };

  // 3. Loop and Categorize Bookings
  const bookings = (typeof state !== 'undefined' && Array.isArray(state.bookings)) ? state.bookings : [];

  bookings.forEach(b => {
    const checkInMs = parseDate(b.checkIn);
    const checkOutMs = getEffectiveCheckoutTime(b);

    // Year Filter Validation
    if (targetYear && !isNaN(checkInMs)) {
      const yr = new Date(checkInMs).getFullYear();
      if (yr !== targetYear) return;
    }

    // Amount extractions
    const amt = Number(b.totalAmount || 0);
    const adv = Number(b.initialAdv || 0) + Number(b.clearedDue || 0);
    const due = Number(b.totalDue || 0);

    // Check Inactive Status
    if (typeof isInactiveBooking === 'function' && isInactiveBooking(b)) {
      stats.inactive.count++;
      stats.inactive.amount += amt;
      stats.inactive.adv += adv;
      stats.inactive.due += due;
      return;
    }

    // Active Status Categorization (Live / Upcoming / Closed)
    if (!isNaN(checkInMs) && !isNaN(checkOutMs)) {
      if (now >= checkInMs && now <= checkOutMs) {
        stats.live.count++;
        stats.live.amount += amt;
        stats.live.adv += adv;
        stats.live.due += due;
        liveBookingsList.push(b); // Collect live booking item
      } else if (now < checkInMs) {
        stats.upcoming.count++;
        stats.upcoming.amount += amt;
        stats.upcoming.adv += adv;
        stats.upcoming.due += due;
      } else {
        stats.closed.count++;
        stats.closed.amount += amt;
        stats.closed.adv += adv;
        stats.closed.due += due;
      }
    }
  });

  // Calculate Active Totals
  const totalBookings = stats.live.count + stats.upcoming.count + stats.closed.count;
  const totalAmt = stats.live.amount + stats.upcoming.amount + stats.closed.amount;
  const totalAdv = stats.live.adv + stats.upcoming.adv + stats.closed.adv;
  const totalDue = stats.live.due + stats.upcoming.due + stats.closed.due;

  // DOM Safe Setters
  const setTxt = (id, text) => {
    const el = document.getElementById(id);
    if (el) el.innerText = text;
  };

  const fmt = num => `₹${Number(num || 0).toLocaleString('en-IN')}`;

  // 4. Update UI Elements & Pie Charts

  // --- Total Bookings Card ---
  setTxt('dash-total-bookings', totalBookings);
  setTxt('dash-count-live', stats.live.count);
  setTxt('dash-count-upcoming', stats.upcoming.count);
  setTxt('dash-count-closed', stats.closed.count);
  updatePieChart('pie-bookings', stats.live.count, stats.upcoming.count, stats.closed.count, stats.inactive.count);

  // --- Total Booking Amount Card ---
  setTxt('dash-total-amount', fmt(totalAmt));
  setTxt('dash-amount-live', fmt(stats.live.amount));
  setTxt('dash-amount-upcoming', fmt(stats.upcoming.amount));
  setTxt('dash-amount-closed', fmt(stats.closed.amount));
  updatePieChart('pie-amount', stats.live.amount, stats.upcoming.amount, stats.closed.amount, stats.inactive.amount);

  // --- Box 3: Total Advanced / Received Card ---
  setTxt('dash-advanced', fmt(totalAdv));
  setTxt('dash-recv-live', fmt(stats.live.adv));
  setTxt('dash-recv-upcoming', fmt(stats.upcoming.adv));
  setTxt('dash-recv-closed', fmt(stats.closed.adv));
  updatePieChart('pie-received', stats.live.adv, stats.upcoming.adv, stats.closed.adv, stats.inactive.adv);

  // --- Total Due Card ---
  setTxt('dash-due', fmt(totalDue));
  setTxt('dash-due-live', fmt(stats.live.due));
  setTxt('dash-due-upcoming', fmt(stats.upcoming.due));
  setTxt('dash-due-closed', fmt(stats.closed.due));
  updatePieChart('pie-due', stats.live.due, stats.upcoming.due, stats.closed.due, stats.inactive.due);

  // --- Inactive Card ---
  setTxt('dash-inactive-count', `${stats.inactive.count} Bookings`);
  setTxt('dash-inactive-amount', fmt(stats.inactive.amount));

  // --- Box 6: Live Bookings Table Update ---
  updateDashboardLiveBookings(liveBookingsList);
  const allRemarksContainer = document.getElementById('dash-all-remarks-container');
  if (allRemarksContainer && !allRemarksContainer.classList.contains('hidden')) {
    updateDashboardAllRemarks();
  }
  updateDashboardAgentDetails();
}
    
    function sendReceiptViaWhatsApp() {
      if (!activeModalBooking) {
        alert("⚠️ Booking information not found!");
        return;
      }

      const b = activeModalBooking;
      let rawCountryCode = b.countryCode ? String(b.countryCode).replace(/\D/g, '') : '91';
      let phone = b.contactNo ? String(b.contactNo).replace(/\D/g, '') : '';
      let initialAdvanceVal = b.initialAdv !== undefined ? parseFloat(b.initialAdv) : 0;
      
      let validationErrors = [];
      if (!phone || phone.length !== 10) {
        validationErrors.push("• Guest contact number must be exactly 10 digits.");
      }
      if (!(initialAdvanceVal > 0)) {
        validationErrors.push("• Advanced payment must be greater than 0.");
      }

      if (validationErrors.length > 0) {
        alert("⚠️ Cannot send via WhatsApp:\n\n" + validationErrors.join("\n"));
        return;
      }

      const fullPhoneNumber = rawCountryCode + phone;
      const upiId = "aniruddha.e@oksbi";

      const effectiveOut = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
      const roomsDisplay = getBookingRooms(b).join(', ');

      const messageText = `*Sanoum Pema Homestay-by Anaristays - Booking Receipt*\n\n` +
        `Dear *${b.name}*,\n` +
        `Thank you for booking with us! Here are your booking details:\n\n` +
        `*Reservation Details:*\n` +
        `• Booking ID: *${b.bookingCode}*\n` +
        `• Room No: ${roomsDisplay}\n` +
        `• Check-In: ${formatDateTime(b.checkIn)}\n` +
        `• Check-Out: ${formatDateTime(effectiveOut)}\n\n` +
        `*Billing Summary:*\n` +
        `• Total Amount: ₹${b.totalAmount}\n` +
        `• Advance Amount: ₹${initialAdvanceVal}\n` +
        `• Balance Due: ₹${b.totalDue}\n\n` +
        `*UPI Payment Details:*\n` +
        `• UPI ID: *${upiId}*\n\n` +
        `We look forward to hosting you! 🏠`;

      const encodedMessage = encodeURIComponent(messageText);
      const whatsappUrl = `https://api.whatsapp.com/send?phone=${fullPhoneNumber}&text=${encodedMessage}`;

      window.open(whatsappUrl, '_blank');
    }

    function printInvoice(bookingId) {
      const bIndex = state.bookings.findIndex(item => String(item.id) === String(bookingId));
      if (bIndex === -1) {
        alert("Booking details not found!");
        return;
      }

      const b = state.bookings[bIndex];
      activeModalBooking = b;
      
      const today = formatDate(new Date());

      const readOnlyNotice = document.getElementById('inv-readonly-notice');
      const invPrintBtn = document.getElementById('inv-print-btn');
      const waBtn = document.getElementById('inv-whatsapp-btn');
      const eInvoiceSection = document.getElementById('e-invoice-section');
      const invIdContainer = document.getElementById('inv-id-container');

      const now = new Date().getTime();
      const cIn = parseDateMs(b.checkIn);
      const cOut = getEffectiveCheckoutTime(b);
      const isClosed = now > cOut;
      const isInactive = isInactiveBooking(b);
      const isLive = now >= cIn && now <= cOut;
      const isUpcoming = now < cIn;

      if (waBtn) {
        if (isClosed || isInactive) {
          waBtn.classList.add('hidden');
        } else {
          waBtn.classList.remove('hidden');
          const contactDigits = b.contactNo ? String(b.contactNo).replace(/\D/g, '') : '';
          const hasValidContact = contactDigits.length === 10;
          const hasAdvanced = (parseFloat(b.initialAdv) || 0) > 0;

          if (hasValidContact && hasAdvanced) {
            waBtn.classList.remove('opacity-50', 'cursor-not-allowed');
            waBtn.title = "Send Receipt via WhatsApp";
          } else {
            waBtn.classList.add('opacity-50', 'cursor-not-allowed');
            waBtn.title = "Inactive: Advance payment must be > 0 and contact number must be 10 digits.";
          }
        }
      }

      const hasDue = (b.totalDue || 0) > 0;
      if (eInvoiceSection) {
        if (hasDue) {
          eInvoiceSection.classList.add('hidden');
        } else {
          eInvoiceSection.classList.remove('hidden');
        }
      }

      if (invIdContainer) {
        invIdContainer.querySelector('strong').innerText = b.invoiceNo || 'INV-2026-0000001';
      }

      if (invPrintBtn) {
        invPrintBtn.classList.remove('hidden');
        invPrintBtn.disabled = false;
        invPrintBtn.className = "px-4 py-1.5 bg-blue-600 text-white rounded-xl font-semibold shadow-sm flex items-center gap-1 transition hover:bg-blue-700 cursor-pointer";
        invPrintBtn.innerHTML = `<i class="fa-solid fa-print"></i> Print Invoice`;
      }

      document.getElementById('inv-booking-id').innerText = b.bookingCode || 'N/A';
      document.getElementById('inv-date').innerText = today;

      const fullLocation = [b.address, b.city, b.state, b.country, b.zipCode].filter(Boolean).map(formatTitleCase).join(', ');
      document.getElementById('inv-guest-name').innerText = formatTitleCase(b.name) || 'N/A';
      document.getElementById('inv-guest-address').innerText = `Address: ${fullLocation || 'N/A'}`;
      
      const fullGuestPhone = b.contactNo ? `${b.countryCode || '+91'} ${b.contactNo}`.trim() : '-';
      document.getElementById('inv-guest-contact').innerText = `Contact: ${fullGuestPhone}`;
      document.getElementById('inv-guest-id').innerText = `ID No: ${b.idNo || 'N/A'}`;
      const invRemark = document.getElementById('inv-booking-remark');
      if (invRemark) {
        invRemark.textContent = b.remark ? 'Remark: ' + b.remark : '';
        invRemark.classList.toggle('hidden', !b.remark);
      }


      const roomsDisplay = getBookingRooms(b).join(', ');
      document.getElementById('inv-room').innerText = `Room No: ${roomsDisplay}`;
      document.getElementById('inv-checkin').innerText = `Check-in: ${formatDateTime(isTrue(b.hasTripCancel) && b.tripCancelCheckIn ? b.tripCancelCheckIn : b.checkIn)}`;
      document.getElementById('inv-checkout').innerText = `Check-out: ${formatDateTime(isTrue(b.hasTripCancel) && b.tripCancelCheckOut ? b.tripCancelCheckOut : b.checkOut)}`;
      const tripCancelInvoiceNote = document.getElementById('inv-trip-cancel-details');
      if (tripCancelInvoiceNote) { tripCancelInvoiceNote.innerText = isTrue(b.hasTripCancel) ? `Trip cancellation: ${formatDateTime(b.tripCancelCheckIn)} – ${formatDateTime(b.tripCancelCheckOut)}` : ''; tripCancelInvoiceNote.classList.toggle('hidden', !isTrue(b.hasTripCancel)); }

      const extCheckoutElem = document.getElementById('inv-ext-checkout');
      if (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) {
        extCheckoutElem.innerHTML = `Extended Check-out: ${formatDateTime(b.extendedCheckOut)} <span class="text-[9px] text-blue-700 bg-blue-50 border border-blue-200 px-1.5 py-0.2 rounded-full font-bold ml-1">Extended</span>`;
        extCheckoutElem.classList.remove('hidden');
      } else {
        extCheckoutElem.innerText = '';
        extCheckoutElem.classList.add('hidden');
      }

      const tbody = document.getElementById('inv-items-tbody');
      tbody.innerHTML = '';

      const roomTotal = (b.noOfDays || 0) * (b.perDayPrice || 0) * (b.capacity || 1);
      const showMealsNote = b.includeMeals !== false && b.includeMeals !== 'false';
      const mealNotesStr = showMealsNote ? '<span class="text-[9px] text-slate-500 block font-normal">(*Include Breakfast,Lunch,Evening snack & Dinner)</span>' : '';

      const stayDaysCount = parseInt(b.noOfDays) || 0;
      const daysFormattedStr = stayDaysCount === 1 ? '1 Day' : `${stayDaysCount} Days`;

      const roomCapacityCount = parseInt(b.capacity) || 1;
      const roomCapacityLabel = roomCapacityCount === 1 ? 'Person' : 'Persons';

      const roomTr = document.createElement('tr');
      roomTr.innerHTML = `
        <td class="p-2.5 font-semibold text-slate-800">
          Room ${roomsDisplay} Accommodation (${roomCapacityCount} ${roomCapacityLabel}) ${isTrue(b.hasExtendedCheckout) ? '<span class="text-[9px] text-blue-600 block font-normal">(Includes extended stay duration)</span>' : ''}
          ${mealNotesStr}
        </td>
        <td class="p-2.5 text-center">${daysFormattedStr}</td>
        <td class="p-2.5 text-right">₹${(b.perDayPrice || 0).toLocaleString('en-IN')}</td>
        <td class="p-2.5 text-right font-semibold text-slate-800">₹${roomTotal.toLocaleString('en-IN')}</td>
      `;
      tbody.appendChild(roomTr);

      if (b.extraPersons && b.extraPersons > 0 && b.extraPersonDays > 0) {
        const extraPersonTotal = b.extraPersons * b.extraPersonDays * (b.ExtraperDayPrice || 0);
        const extraJoinedFmt = b.extraPersonJoined ? formatDateTime(b.extraPersonJoined) : '';
        const extraOutFmt = b.extraPersonOut ? formatDateTime(b.extraPersonOut) : '';
        const extraDaysCount = parseInt(b.extraPersonDays) || 0;
        const extraDaysFormattedStr = extraDaysCount === 1 ? '1 Day' : `${extraDaysCount} Days`;

        const extraTr = document.createElement('tr');
        extraTr.innerHTML = `
          <td class="p-2.5 font-semibold text-amber-900">
            Additional Guest Accommodation (${b.extraPersons} ${b.extraPersons === 1 ? 'Person' : 'Persons'})
            <span class="text-[9px] text-amber-700 font-normal block">Stay: ${extraJoinedFmt} to ${extraOutFmt || 'Check-Out'}</span>
          </td>
          <td class="p-2.5 text-center">${extraDaysFormattedStr}</td>
          <td class="p-2.5 text-right">₹${(b.ExtraperDayPrice || 0).toLocaleString('en-IN')}</td>
          <td class="p-2.5 text-right font-semibold text-amber-900">₹${extraPersonTotal.toLocaleString('en-IN')}</td>
        `;
        tbody.appendChild(extraTr);
      }

      const foodList = parseJSONField(b.foodOrders);
      if (foodList.length > 0) {
        foodList.forEach(fo => {
          if (fo.foodCharge > 0) {
            const foodTr = document.createElement('tr');
            const foodDateTimeFmt = fo.foodDateTime ? ` (${formatDateTime(fo.foodDateTime)})` : '';
            const plateCount = parseInt(fo.plates) || 1;
            const plateLabel = plateCount === 1 ? 'Plate' : 'Plates';

            foodTr.innerHTML = `
              <td class="p-2.5 font-semibold text-slate-800">Extra Food <span class="text-[9px] text-slate-500 font-normal block">${fo.foodDesc || 'Food Item'}${foodDateTimeFmt}</span></td>
              <td class="p-2.5 text-center">${plateCount} ${plateLabel}</td>
              <td class="p-2.5 text-right">₹${(fo.itemPrice || 0).toLocaleString('en-IN')}</td>
              <td class="p-2.5 text-right font-semibold text-slate-800">₹${(fo.foodCharge || 0).toLocaleString('en-IN')}</td>
            `;
            tbody.appendChild(foodTr);
          }
        });
      }
      
      const cabList = parseJSONField(b.cabTrips);
      if (cabList.length > 0) {
        cabList.forEach(trip => {
          if (trip.rate > 0) {
             const cabTr = document.createElement('tr');
             const dtFormat = trip.dateTime ? ` (${formatDateTime(trip.dateTime)})` : '';
             cabTr.innerHTML = `
              <td class="p-2.5 font-semibold text-indigo-900">
                Cab Fare - ${trip.tripName}${dtFormat}
                ${trip.remark ? `<span class="text-[9px] text-indigo-700 font-normal block">Remark: ${trip.remark}</span>` : ''}
              </td>
              <td class="p-2.5 text-center">1 Trip</td>
              <td class="p-2.5 text-right">₹${(trip.rate || 0).toLocaleString('en-IN')}</td>
              <td class="p-2.5 text-right font-semibold text-indigo-900">₹${(trip.rate || 0).toLocaleString('en-IN')}</td>
            `;
            tbody.appendChild(cabTr);
          }
        });
      }

      const initialAdv = b.initialAdv || 0;
      const clearDueAmt = b.clearedDue || 0;

      document.getElementById('inv-sum-total').innerText = `₹${(b.totalAmount || 0).toLocaleString('en-IN')}`;
      document.getElementById('inv-sum-advance').innerText = `₹${(initialAdv || 0).toLocaleString('en-IN')}`;
      document.getElementById('inv-sum-due').innerText = `₹${(b.totalDue || 0).toLocaleString('en-IN')}`;

      const discountPct = Number(b.discountPercent || 0);
      const discountAmount = Number(b.discountAmount || 0);
      const discountPctRow = document.getElementById('inv-discount-percent-row');
      const discountAmountRow = document.getElementById('inv-discount-amount-row');
      if (discountPctRow) { document.getElementById('inv-sum-discount-percent').innerText = `${discountPct.toFixed(2).replace(/\.?0+$/, '')}%`; discountPctRow.classList.toggle('hidden', discountPct <= 0); }
      if (discountAmountRow) { document.getElementById('inv-sum-discount-amount').innerText = `₹${discountAmount.toLocaleString('en-IN')}`; discountAmountRow.classList.toggle('hidden', discountAmount <= 0); }
      const clearDueRow = document.getElementById('inv-clear-due-row');
      if (clearDueAmt > 0) {
        document.getElementById('inv-sum-clear-due').innerText = `₹${clearDueAmt.toLocaleString('en-IN')}`;
        clearDueRow.classList.remove('hidden');
      } else {
        clearDueRow.classList.add('hidden');
      }

      const qrToggle = document.getElementById('inv-show-qr-toggle');
      if (qrToggle) { qrToggle.checked = false; toggleInvoiceQR(false); }
      document.getElementById('invoice-modal').classList.remove('hidden');
    }

    function closeInvoiceModal() {
      activeModalBooking = null;
      document.getElementById('invoice-modal').classList.add('hidden');
    }

    function addFoodOrderItem(desc = '', plates = 1, itemPrice = 0, charge = 0, dateStr = '', timeStr = '', disabled = false) {
      const foodWin = getModalFoodWindow();

      if (!foodWin && !disabled) {
        alert("⚠️ Please enter valid Check-In and Check-Out date & time first before adding extra food!");
        return;
      }

      if (foodWin && foodWin.minFoodDt >= foodWin.maxFoodDt && !disabled) {
        alert("⚠️ Invalid stay window! The duration between Check-In (+15m) and Check-Out (-30m) is too short to order food.");
        return;
      }

      if (!dateStr || !timeStr) {
        if (foodWin) {
          const now = new Date();
          let defaultDt = now;
          if (now < foodWin.minFoodDt || now > foodWin.maxFoodDt) {
            defaultDt = foodWin.minFoodDt;
          }
          
          const utcMs = defaultDt.getTime();
          const istDate = new Date(utcMs + (330 * 60000));
          
          const yyyy = istDate.getUTCFullYear();
          const mm = String(istDate.getUTCMonth() + 1).padStart(2, '0');
          const dd = String(istDate.getUTCDate()).padStart(2, '0');
          const hh = String(istDate.getUTCHours()).padStart(2, '0');
          const min = String(istDate.getUTCMinutes()).padStart(2, '0');

          if (!dateStr) dateStr = `${yyyy}-${mm}-${dd}`;
          if (!timeStr) timeStr = `${hh}:${min}`;
        }
      }

      const container = document.getElementById('food-orders-container');
      const itemRow = document.createElement('div');
      itemRow.className = "food-order-row grid grid-cols-1 sm:grid-cols-12 gap-1.5 items-end bg-white p-2.5 rounded-2xl border border-amber-200/80 shadow-xs";
      
      const disabledAttr = disabled ? 'disabled' : '';
      const bgClass = disabled ? 'bg-slate-100 cursor-not-allowed text-slate-500' : 'bg-white';

      itemRow.innerHTML = `
        <div class="sm:col-span-3">
          <label class="block font-semibold text-slate-600 mb-0.5">Item Name</label>
          <input type="text" value="${desc}" ${disabledAttr} placeholder="e.g. Thali / Tea" class="cust-food-desc w-full ${bgClass} border border-slate-200 rounded-xl px-2.5 py-1 focus:outline-none focus:border-amber-500">
        </div>

        <div class="sm:col-span-3">
          <label class="block font-semibold text-slate-600 mb-0.5"><i class="fa-regular fa-clock text-amber-600 mr-1"></i> Date & Time</label>
          <div class="flex gap-1">
            <input type="date" value="${dateStr}" ${disabledAttr} onchange="validateFoodRowDateTime(this)" class="cust-food-date w-3/5 ${bgClass} border border-slate-200 rounded-xl px-1.5 py-1 focus:outline-none focus:border-amber-500 font-medium text-[10px]">
            <input type="time" value="${timeStr}" ${disabledAttr} onchange="validateFoodRowDateTime(this)" class="cust-food-time w-2/5 ${bgClass} border border-slate-200 rounded-xl px-1 py-1 focus:outline-none focus:border-amber-500 font-medium text-[10px]">
          </div>
        </div>

        <div class="sm:col-span-2">
          <label class="block font-semibold text-slate-600 mb-0.5">Price/Plate (₹)</label>
          <input type="number" value="${itemPrice}" min="0" ${disabledAttr} oninput="calculateFoodRowTotal(this)" class="cust-food-price w-full ${bgClass} border border-slate-200 rounded-xl px-2 py-1 focus:outline-none focus:border-amber-500">
        </div>
        
        <div class="sm:col-span-1">
          <label class="block font-semibold text-slate-600 mb-0.5">Plates</label>
          <input type="number" value="${plates}" min="1" ${disabledAttr} oninput="calculateFoodRowTotal(this)" class="cust-food-plates w-full ${bgClass} border border-slate-200 rounded-xl px-2 py-1 focus:outline-none focus:border-amber-500 font-bold">
        </div>

        <div class="sm:col-span-2">
          <label class="block font-semibold text-slate-600 mb-0.5">Total (₹)</label>
          <input type="number" value="${charge}" readonly class="cust-food-charge w-full bg-slate-100 font-bold text-amber-700 border border-slate-200 rounded-xl px-2 py-1 cursor-not-allowed">
        </div>

        <div class="sm:col-span-1 flex justify-end">
          <button type="button" onclick="removeFoodOrderItem(this)" ${disabledAttr} class="btn-remove-food-item text-rose-500 hover:text-rose-700 p-1.5 ${disabled ? 'hidden' : ''}" title="Remove Order">
            <i class="fa-solid fa-trash-can"></i>
          </button>
        </div>
      `;

      container.appendChild(itemRow);
      calculateModalBilling();
    }

    function calculateFoodRowTotal(inputElem) {
      const row = inputElem.closest('.food-order-row');
      if (!row) return;

      const price = parseFloat(row.querySelector('.cust-food-price').value) || 0;
      const plates = parseInt(row.querySelector('.cust-food-plates').value) || 0;
      const totalCharge = price * plates;

      row.querySelector('.cust-food-charge').value = totalCharge;
      calculateModalBilling();
    }

    function removeFoodOrderItem(btn) {
      const row = btn.closest('.food-order-row');
      if (row) {
        row.remove();
        calculateModalBilling();
      }
    }

    function updateBookingRemarkCount() {
      const field = document.getElementById('cust-booking-remark');
      const counter = document.getElementById('cust-booking-remark-count');
      if (!field || !counter) return;
      const words = field.value.trim() ? field.value.trim().split(/\s+/).length : 0;
      counter.innerText = `${words} / 1000 words`;
      counter.classList.toggle('text-rose-600', words > 1000);
    }

    function toggleBookingRemark() {
      const editor = document.getElementById('booking-remark-editor');
      const field = document.getElementById('cust-booking-remark');
      if (!editor || !field) return;
      // Once a remark has been entered, keep its editor visible.
      if (!field.value.trim()) editor.classList.toggle('hidden');
      else editor.classList.remove('hidden');
      const isOpen = !editor.classList.contains('hidden');
      const button = document.getElementById('btn-toggle-booking-remark');
      if (button) button.innerHTML = isOpen
        ? '<i class="fa-solid fa-check text-[9px]"></i> Remark Added'
        : '<i class="fa-solid fa-plus text-[9px]"></i> Add Remark';
    }

    function keepBookingRemarkOpenIfFilled() {
      const field = document.getElementById('cust-booking-remark');
      const editor = document.getElementById('booking-remark-editor');
      const button = document.getElementById('btn-toggle-booking-remark');
      if (field && editor && field.value.trim()) {
        editor.classList.remove('hidden');
        if (button) button.innerHTML = '<i class="fa-solid fa-check text-[9px]"></i> Remark Added';
      }
    }

    function setBookingRemarkEditorOpen(open) {
      const editor = document.getElementById('booking-remark-editor');
      const button = document.getElementById('btn-toggle-booking-remark');
      const field = document.getElementById('cust-booking-remark');
      if (!editor || !field) return;
      const shouldOpen = open || !!field.value.trim();
      editor.classList.toggle('hidden', !shouldOpen);
      if (button) button.innerHTML = shouldOpen
        ? '<i class="fa-solid fa-check text-[9px]"></i> Remark Added'
        : '<i class="fa-solid fa-plus text-[9px]"></i> Add Remark';
    }

    function addCabTripRow(rate = 0, dateStr = '', timeStr = '', remark = '', disabled = false) {
      const container = document.getElementById('cab-trips-container');
      const tripCount = container.children.length + 1;
      const itemRow = document.createElement('div');
      itemRow.className = "cab-trip-row grid grid-cols-1 sm:grid-cols-12 gap-1.5 items-end bg-white p-2.5 rounded-2xl border border-indigo-200/80 shadow-xs";

      const disabledAttr = disabled ? 'disabled' : '';
      const bgClass = disabled ? 'bg-slate-100 cursor-not-allowed text-slate-500' : 'bg-white';

      if (!dateStr) {
         const inDate = document.getElementById('cust-checkin-date')?.value;
         if (inDate) dateStr = inDate;
      }
      if (!timeStr) timeStr = '12:00';

      itemRow.innerHTML = `
        <div class="sm:col-span-2">
          <label class="block font-semibold text-slate-600 mb-0.5">Trip Name</label>
          <input type="text" value="Trip ${tripCount}" readonly class="w-full bg-slate-50 border border-slate-200 rounded-xl px-2.5 py-1 text-slate-500 font-bold cursor-not-allowed text-[10px]">
        </div>
        <div class="sm:col-span-4">
          <label class="block font-semibold text-slate-600 mb-0.5"><i class="fa-regular fa-clock text-indigo-600 mr-1"></i> Date & Time</label>
          <div class="flex gap-1">
            <input type="date" value="${dateStr}" ${disabledAttr} class="cust-cab-date w-3/5 ${bgClass} border border-slate-200 rounded-xl px-1.5 py-1 focus:outline-none focus:border-indigo-500 font-medium text-[10px]">
            <input type="time" value="${timeStr}" ${disabledAttr} class="cust-cab-time w-2/5 ${bgClass} border border-slate-200 rounded-xl px-1 py-1 focus:outline-none focus:border-indigo-500 font-medium text-[10px]">
          </div>
        </div>
        <div class="sm:col-span-2">
          <label class="block font-semibold text-slate-600 mb-0.5">Rate/Trip (₹)</label>
          <input type="number" value="${rate}" min="0" ${disabledAttr} oninput="calculateModalBilling()" class="cust-cab-rate w-full ${bgClass} border border-slate-200 rounded-xl px-2 py-1 focus:outline-none focus:border-indigo-500 font-bold text-indigo-700">
        </div>
        <div class="sm:col-span-3">
          <label class="block font-semibold text-slate-600 mb-0.5">Remark</label>
          <input type="text" value="${remark}" ${disabledAttr} placeholder="e.g. Airport drop" class="cust-cab-remark w-full ${bgClass} border border-slate-200 rounded-xl px-2.5 py-1 focus:outline-none focus:border-indigo-500 text-[10px]">
        </div>
        <div class="sm:col-span-1 flex justify-end">
          <button type="button" onclick="removeCabTripRow(this)" ${disabledAttr} class="text-rose-500 hover:text-rose-700 p-1.5 ${disabled ? 'hidden' : ''}" title="Remove Trip">
            <i class="fa-solid fa-trash-can"></i>
          </button>
        </div>
      `;
      container.appendChild(itemRow);
      calculateModalBilling();
    }

    function removeCabTripRow(btn) {
      const row = btn.closest('.cab-trip-row');
      if (row) {
        row.remove();
        calculateModalBilling();
      }
    }

    function setInputEnabled(elem, isEnabled) {
      if (!elem) return;
      elem.disabled = !isEnabled;
      if (!isEnabled) {
        elem.classList.add('bg-slate-100', 'cursor-not-allowed', 'text-slate-500');
        elem.classList.remove('bg-white', 'bg-amber-50');
      } else {
        elem.classList.remove('bg-slate-100', 'cursor-not-allowed', 'text-slate-500');
      }
    }

    function openBookingModal(bookingId = null) {
      const now = new Date().getTime();
      let isLiveBooking = false;
      let isClosedBooking = false;
      let isUpcomingBooking = false;
      let isPast730Days = false;

      let b = null;
      if (bookingId) {
        b = state.bookings.find(item => String(item.id) === String(bookingId));
        if (b) {
          if (isInactiveBooking(b)) {
            // "just receipt view will be enabled."
            printInvoice(bookingId);
            return;
          }

          if (!isRoomInMaster(b.roomNo)) {
            alert("This booking details were deleted from Master Data and cannot be opened or edited.");
            return;
          }

          const effectiveOutTime = getEffectiveCheckoutTime(b);
          const checkInTime = parseDateMs(b.checkIn);

          if (now > effectiveOutTime) {
            isClosedBooking = true;
          } else if (now >= checkInTime && now <= effectiveOutTime) {
            isLiveBooking = true;
          } else {
            isUpcomingBooking = true;
          }
        }
      } else {
        isUpcomingBooking = true;
      }

      document.getElementById('modal-booking-id').value = bookingId || '';
      setMinBookingDates();
      
      const form = document.getElementById('booking-form');
      form.reset();
      removeAttachedIdProof();
      
      const clearBillInput = document.getElementById('cust-clear-bill');
      if (clearBillInput) clearBillInput.value = 0;
      const discountPctInput = document.getElementById('cust-discount-percent'); if (discountPctInput) discountPctInput.value = 0;
      const discountAmountInput = document.getElementById('cust-discount-amount'); if (discountAmountInput) discountAmountInput.value = 0;

      document.getElementById('food-orders-container').innerHTML = '';
      document.getElementById('cab-trips-container').innerHTML = '';
      document.getElementById('cust-booking-remark').value = '';
      updateBookingRemarkCount();
      setBookingRemarkEditorOpen(false);
      populateAgentDropdown();

      setSectionEditability('sec-guest-info', true);
      setSectionEditability('sec-room-dates', true);

      const extChkBox = document.getElementById('cust-has-extended-checkout');
      const extDateInput = document.getElementById('cust-ext-checkout-date');
      const extTimeInput = document.getElementById('cust-ext-checkout-time');
      const timerNotice = document.getElementById('ext-checkout-timer-notice');

      let canToggleExtendedCheckout = false;
      if (b) {
        const initialCheckOutTime = parseDateMs(b.checkOut);
        const isWithin1HrPastCheckout = now > initialCheckOutTime && now <= (initialCheckOutTime + ONE_HOUR_MS);

        if (isLiveBooking || isWithin1HrPastCheckout) {
          canToggleExtendedCheckout = true;
        }
      }

      if (extChkBox) {
        extChkBox.disabled = !canToggleExtendedCheckout;
        if (canToggleExtendedCheckout) {
          if (timerNotice) {
            timerNotice.innerText = isLiveBooking ? "(Active for Live Booking)" : "(Active up to 1hr post check-out)";
            timerNotice.classList.remove('hidden', 'text-rose-600');
          }
        } else {
          if (timerNotice) {
            timerNotice.innerText = isClosedBooking ? "(Closed Booking - Inactive)" : "(Selectable only for Live Booking)";
            timerNotice.classList.remove('hidden');
            timerNotice.classList.add('text-rose-600');
          }
        }
      }

      const tripCancelChk = document.getElementById('cust-has-trip-cancel');
      if (tripCancelChk) {
        tripCancelChk.disabled = !isLiveBooking;
        if (!isLiveBooking) tripCancelChk.checked = false;
        const tripNotice = document.getElementById('trip-cancel-notice');
        if (tripNotice) tripNotice.innerText = isLiveBooking ? '(Available while booking is live)' : '(Selectable only for Live Booking)';
      }

      const mealsChkBox = document.getElementById('cust-include-meals');
      if (mealsChkBox) {
        mealsChkBox.disabled = false;
      }

      const addFoodBtn = document.getElementById('btn-add-food-order');
      if (addFoodBtn) {
        addFoodBtn.disabled = isPast730Days;
        if (isPast730Days) {
          addFoodBtn.classList.add('opacity-50', 'cursor-not-allowed');
        } else {
          addFoodBtn.classList.remove('opacity-50', 'cursor-not-allowed');
        }
      }
      
      const addCabBtn = document.getElementById('btn-add-cab-trip');
      if (addCabBtn) {
        addCabBtn.disabled = isPast730Days;
        if (isPast730Days) {
          addCabBtn.classList.add('opacity-50', 'cursor-not-allowed');
        } else {
          addCabBtn.classList.remove('opacity-50', 'cursor-not-allowed');
        }
      }
      setSectionEditability('sec-cab-fare', !isPast730Days);
      setSectionEditability('sec-billing-summary', !isPast730Days);

      const btnSave = document.getElementById('btn-save-booking');
      if (btnSave) {
         if (isPast730Days) {
            btnSave.disabled = true;
            btnSave.classList.add('opacity-50', 'cursor-not-allowed', 'bg-slate-400');
            btnSave.classList.remove('bg-blue-600', 'hover:bg-blue-700');
         } else {
            btnSave.disabled = false;
            btnSave.classList.remove('opacity-50', 'cursor-not-allowed', 'bg-slate-400');
            btnSave.classList.add('bg-blue-600', 'hover:bg-blue-700');
         }
      }

      const extraPersonsInput = document.getElementById('cust-extra-persons');
      const extraPersonTimeWrapper = document.getElementById('sec-extra-person-time-wrapper');
      const extraPersonDateInput = document.getElementById('cust-extra-person-date');
      const extraPersonTimeInput = document.getElementById('cust-extra-person-time');
      const extraPersonOutDateInput = document.getElementById('cust-extra-person-out-date');
      const extraPersonOutTimeInput = document.getElementById('cust-extra-person-out-time');

      const canEditExtras = true;
      if (extraPersonsInput) setInputEnabled(extraPersonsInput, canEditExtras);
      if (extraPersonDateInput) setInputEnabled(extraPersonDateInput, canEditExtras);
      if (extraPersonTimeInput) setInputEnabled(extraPersonTimeInput, canEditExtras);
      if (extraPersonOutDateInput) setInputEnabled(extraPersonOutDateInput, canEditExtras);
      if (extraPersonOutTimeInput) setInputEnabled(extraPersonOutTimeInput, canEditExtras);

      if (canEditExtras) {
        if (extraPersonTimeWrapper) extraPersonTimeWrapper.classList.remove('hidden');
      } else {
        if (extraPersonTimeWrapper && (!b || !b.extraPersons || b.extraPersons <= 0)) {
          extraPersonTimeWrapper.classList.add('hidden');
        } else if (extraPersonTimeWrapper) {
          extraPersonTimeWrapper.classList.remove('hidden');
        }
      }

      if (b) {
        document.getElementById('modal-title').innerText = isPast730Days ? 'Closed Booking (Read-Only)' : (isClosedBooking ? 'Closed Booking (Billing Active)' : 'Edit Booking Details');
        
        // ** ONLY ALLOW EDITING OF MAIN CHECK-IN AND CHECK-OUT DATES IF BOOKING IS UPCOMING **
        setInputEnabled(document.getElementById('cust-checkin-date'), isUpcomingBooking);
        setInputEnabled(document.getElementById('cust-checkin-time'), isUpcomingBooking);
        setInputEnabled(document.getElementById('cust-checkout-date'), isUpcomingBooking);
        setInputEnabled(document.getElementById('cust-checkout-time'), isUpcomingBooking);
        
        document.getElementById('modal-booking-id').value = b.id;
        document.getElementById('cust-name').value = formatTitleCase(b.name);
        document.getElementById('cust-booking-remark').value = b.remark || '';
        updateBookingRemarkCount();
        setBookingRemarkEditorOpen(!!(b.remark || '').trim());
        document.getElementById('cust-address').value = formatTitleCase(b.address || '');
        document.getElementById('cust-city').value = formatTitleCase(b.city || '');
        document.getElementById('cust-state').value = formatTitleCase(b.state || '');
        document.getElementById('cust-country').value = formatTitleCase(b.country || '');
        document.getElementById('cust-zip').value = b.zipCode || '';
        document.getElementById('cust-id').value = b.idNo || '';
        document.getElementById('cust-country-code').value = b.countryCode || '+91';
        document.getElementById('cust-contact').value = b.contactNo || '';

        if (b.idProofBase64) {
          document.getElementById('cust-id-file-base64').value = b.idProofBase64;
          document.getElementById('cust-id-file-name').value = b.idProofFileName || 'Attached_ID_Proof.pdf';
          document.getElementById('cust-id-file-status').innerHTML = `<span class="text-emerald-600 font-semibold"><i class="fa-solid fa-circle-check"></i> Attached: ${b.idProofFileName || 'Attached_ID_Proof.pdf'}</span>`;
          document.getElementById('cust-id-file-remove').classList.remove('hidden');
        }
        
        populateRoomDropdown(b.roomNo);
        populateAgentDropdown(b.agentInfo);
        document.getElementById('cust-capacity').value = b.capacity || 1;

        if (extraPersonsInput) extraPersonsInput.value = b.extraPersons || 0;
        
        if (b.extraPersonJoined) {
          const parts = extractISTDateParts(b.extraPersonJoined);
          if (extraPersonDateInput) extraPersonDateInput.value = parts.date || '';
          if (extraPersonTimeInput) extraPersonTimeInput.value = parts.time || '';
        } else {
          if (extraPersonDateInput) extraPersonDateInput.value = '';
          if (extraPersonTimeInput) extraPersonTimeInput.value = '';
        }

        if (b.extraPersonOut) {
          const parts = extractISTDateParts(b.extraPersonOut);
          if (extraPersonOutDateInput) extraPersonOutDateInput.value = parts.date || '';
          if (extraPersonOutTimeInput) extraPersonOutTimeInput.value = parts.time || '';
        } else {
          if (extraPersonOutDateInput) extraPersonOutDateInput.value = '';
          if (extraPersonOutTimeInput) extraPersonOutTimeInput.value = '';
        }

        if (b.checkIn) {
          const parts = extractISTDateParts(b.checkIn);
          document.getElementById('cust-checkin-date').value = parts.date || '';
          document.getElementById('cust-checkin-time').value = parts.time || '';
        }
        if (b.checkOut) {
          const parts = extractISTDateParts(b.checkOut);
          document.getElementById('cust-checkout-date').value = parts.date || '';
          document.getElementById('cust-checkout-time').value = parts.time || '';
          if (extDateInput) extDateInput.min = parts.date || '';
        }

        extChkBox.checked = isTrue(b.hasExtendedCheckout);
        toggleExtendedCheckoutFields(extChkBox.checked);
        if (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) {
          const parts = extractISTDateParts(b.extendedCheckOut);
          extDateInput.value = parts.date || '';
          extTimeInput.value = parts.time || '';
        }

        const tripCancelChkBox = document.getElementById('cust-has-trip-cancel');
        if (tripCancelChkBox) {
          tripCancelChkBox.checked = isTrue(b.hasTripCancel);
          toggleTripCancelFields(tripCancelChkBox.checked);
        }
        [['tripCancelCheckIn','cust-cancel-checkin-date','cust-cancel-checkin-time'],['tripCancelCheckOut','cust-cancel-checkout-date','cust-cancel-checkout-time']].forEach(([key,dateId,timeId]) => {
          if (b[key]) { const cp = extractISTDateParts(b[key]); document.getElementById(dateId).value = cp.date || ''; document.getElementById(timeId).value = cp.time || ''; }
        });

        if (mealsChkBox) {
          mealsChkBox.checked = b.includeMeals !== undefined ? (b.includeMeals !== false && b.includeMeals !== 'false') : true;
        }

        const foList = parseJSONField(b.foodOrders);
        foList.forEach(fo => {
          let fDate = '', fTime = '';
          if (fo.foodDateTime) {
            const parts = extractISTDateParts(fo.foodDateTime);
            fDate = parts.date || '';
            fTime = parts.time || '';
          }
          addFoodOrderItem(fo.foodDesc || '', fo.plates || 1, fo.itemPrice || 0, fo.foodCharge || 0, fDate, fTime, false);
        });
        
        const tripsList = parseJSONField(b.cabTrips);
        if (tripsList.length > 0) {
          tripsList.forEach(trip => {
            let cDate = '', cTime = '';
            if (trip.dateTime) {
               const parts = extractISTDateParts(trip.dateTime);
               cDate = parts.date;
               cTime = parts.time;
            } else {
               cDate = trip.dateStr || '';
               cTime = trip.timeStr || '';
            }
            addCabTripRow(trip.rate || 0, cDate, cTime, trip.remark || '', false);
          });
        }

        document.getElementById('cust-price').value = b.perDayPrice;
        document.getElementById('cust-extra-rate').value = b.ExtraperDayPrice;
        
        const advanceElem = document.getElementById('cust-advance');
        const baseAdv = b.initialAdv || 0;
        advanceElem.value = baseAdv;
        advanceElem.setAttribute('data-initial-adv', baseAdv);

        if (b.clearedDue) {
          document.getElementById('cust-clear-bill').value = b.clearedDue;
        }
        document.getElementById('cust-discount-amount').value = b.discountAmount || 0;
        document.getElementById('cust-discount-percent').value = b.discountPercent || 0;

        calculateModalBilling();
      } else {
        document.getElementById('modal-title').innerText = 'Add New Booking';
        const newTripCancel = document.getElementById('cust-has-trip-cancel');
        if (newTripCancel) { newTripCancel.checked = false; newTripCancel.disabled = true; }
        toggleTripCancelFields(false);
        ['cust-cancel-checkin-date','cust-cancel-checkin-time','cust-cancel-checkout-date','cust-cancel-checkout-time'].forEach(id => document.getElementById(id).value = '');
        document.getElementById('modal-booking-id').value = '';
        
        // DO NOT lock the fields if adding a new booking
        setInputEnabled(document.getElementById('cust-checkin-date'), isUpcomingBooking);
        setInputEnabled(document.getElementById('cust-checkin-time'), isUpcomingBooking);
        setInputEnabled(document.getElementById('cust-checkout-date'), isUpcomingBooking);
        setInputEnabled(document.getElementById('cust-checkout-time'), isUpcomingBooking);

        // Calculate and set today's date as min for new bookings strictly using IST standard
        const todayDt = new Date();
        const utcMs = todayDt.getTime();
        const istDate = new Date(utcMs + (330 * 60000));
        
        const yyyy = istDate.getUTCFullYear();
        const mm = String(istDate.getUTCMonth() + 1).padStart(2, '0');
        const dd = String(istDate.getUTCDate()).padStart(2, '0');
        const todayStr = `${yyyy}-${mm}-${dd}`;

        const tomorrowDt = new Date(todayDt);
        tomorrowDt.setDate(tomorrowDt.getDate() + 1);
        const t_utcMs = tomorrowDt.getTime();
        const t_istDate = new Date(t_utcMs + (330 * 60000));
        const t_yyyy = t_istDate.getUTCFullYear();
        const t_mm = String(t_istDate.getUTCMonth() + 1).padStart(2, '0');
        const t_dd = String(t_istDate.getUTCDate()).padStart(2, '0');
        const tomorrowStr = `${t_yyyy}-${t_mm}-${t_dd}`;

        const checkInElem = document.getElementById('cust-checkin-date');
        checkInElem.value = todayStr;

        const checkOutElem = document.getElementById('cust-checkout-date');
        checkOutElem.value = tomorrowStr;

        populateRoomDropdown(state.roomsCapacity.length > 0 ? [state.roomsCapacity[0].roomNo] : []);

        document.getElementById('cust-country-code').value = "+91";

        // SET DEFAULT CHECK-IN AND CHECK-OUT TIME TO 11:00 AM FOR NEW BOOKING
        document.getElementById('cust-checkin-time').value = "11:00";
        document.getElementById('cust-checkout-time').value = "11:00";

        if (extraPersonsInput) extraPersonsInput.value = 0;
        
        if (extraPersonDateInput) extraPersonDateInput.value = "";
        
        // SET EXTRA PERSON DEFAULT TIMES TO 11:00 AM FOR NEW BOOKING
        if (extraPersonTimeInput) extraPersonTimeInput.value = "11:00";
        if (extraPersonOutDateInput) extraPersonOutDateInput.value = "";
        if (extraPersonOutTimeInput) extraPersonOutTimeInput.value = "11:00";

        extChkBox.checked = false;
        extChkBox.disabled = true;
        
        if (timerNotice) {
          timerNotice.innerText = "(Selectable only for Live Booking)";
          timerNotice.classList.remove('hidden');
          timerNotice.classList.add('text-rose-600');
        }

        toggleExtendedCheckoutFields(false);

        if (mealsChkBox) {
          mealsChkBox.checked = true;
          mealsChkBox.disabled = false;
        }

        document.getElementById('cust-price').value = 1200;
        
        const advanceElem = document.getElementById('cust-advance');
        advanceElem.value = 0;
        advanceElem.setAttribute('data-initial-adv', 0);

        calculateModalBilling();
      }

      document.getElementById('booking-modal').classList.remove('hidden');
    }

    function setSectionEditability(sectionId, isEditable) {
      const container = document.getElementById(sectionId);
      if (!container) return;

      const inputs = container.querySelectorAll('input, select, button');
      inputs.forEach(el => {
        el.disabled = !isEditable;
        if (!isEditable) {
          el.classList.add('bg-slate-100', 'cursor-not-allowed', 'text-slate-500');
        } else {
          el.classList.remove('bg-slate-100', 'cursor-not-allowed', 'text-slate-500');
        }
      });

      if (!isEditable) {
        container.classList.add('opacity-75', 'bg-slate-100/60');
      } else {
        container.classList.remove('opacity-75', 'bg-slate-100/60');
      }
    }

    function closeBookingModal() {
      document.getElementById('booking-modal').classList.add('hidden');
    }

    function handleExtraPersonDatesChange() {
      const mainInDate = document.getElementById('cust-checkin-date')?.value;
      const mainInTime = document.getElementById('cust-checkin-time')?.value || '12:00';
      const mainOutDate = document.getElementById('cust-checkout-date')?.value;
      const mainOutTime = document.getElementById('cust-checkout-time')?.value || '11:00';
      const hasExt = document.getElementById('cust-has-extended-checkout')?.checked;
      
      let latestOutD = mainOutDate;
      let latestOutT = mainOutTime;

      if (hasExt) {
        const extD = document.getElementById('cust-ext-checkout-date')?.value;
        const extT = document.getElementById('cust-ext-checkout-time')?.value;
        if (extD) {
          latestOutD = extD;
          latestOutT = extT || '12:00';
        }
      }
      
      const epInDateElem = document.getElementById('cust-extra-person-date');
      const epInTimeElem = document.getElementById('cust-extra-person-time');
      const epOutDateElem = document.getElementById('cust-extra-person-out-date');
      const epOutTimeElem = document.getElementById('cust-extra-person-out-time');
      const hasTripCancel = document.getElementById('cust-has-trip-cancel')?.checked;
      if (hasTripCancel) {
        const cancelIn = document.getElementById('cust-cancel-checkin-date')?.value || mainInDate;
        const cancelOut = document.getElementById('cust-cancel-checkout-date')?.value || mainOutDate;
        if (epInDateElem) { epInDateElem.min = cancelIn || mainInDate || ''; epInDateElem.max = cancelOut || ''; if (epInDateElem.value && cancelIn && epInDateElem.value < cancelIn) { alert('Extra guest check-in cannot be earlier than Trip Cancel check-in.'); epInDateElem.value = cancelIn; } if (epInDateElem.value && cancelOut && epInDateElem.value > cancelOut) { alert('Extra guest check-in cannot be later than Trip Cancel check-out.'); epInDateElem.value = cancelOut; } }
        if (epOutDateElem) { epOutDateElem.min = cancelIn || mainInDate || ''; epOutDateElem.max = cancelOut || ''; if (epOutDateElem.value && cancelIn && epOutDateElem.value < cancelIn) { alert('Extra guest check-out cannot be earlier than Trip Cancel check-in.'); epOutDateElem.value = cancelIn; } if (epOutDateElem.value && cancelOut && epOutDateElem.value > cancelOut) { alert('Extra guest check-out cannot be later than Trip Cancel check-out.'); epOutDateElem.value = cancelOut; } }
      } else {
        if (epInDateElem) { epInDateElem.removeAttribute('min'); epInDateElem.removeAttribute('max'); }
        if (epOutDateElem) { epOutDateElem.removeAttribute('min'); epOutDateElem.removeAttribute('max'); }
      }

      if (epInDateElem && epInDateElem.value && mainInDate) {
        const epInFull = new Date(`${epInDateElem.value}T${epInTimeElem.value || '12:00'}:00+05:30`);
        const mainInFull = new Date(`${mainInDate}T${mainInTime}:00+05:30`);

        if (epInFull < mainInFull) {
          alert(`⚠️ Additional person check-in cannot be earlier than the main check-in (${formatDateTime(`${mainInDate}T${mainInTime}`)}). Please correct it.`);
        }
      }

      if (epOutDateElem && epOutDateElem.value && latestOutD) {
        const epOutFull = new Date(`${epOutDateElem.value}T${epOutTimeElem.value || '11:00'}:00+05:30`);
        const mainOutFull = new Date(`${latestOutD}T${latestOutT}:00+05:30`);

        if (epOutFull > mainOutFull) {
          alert(`⚠️ Additional person check-out date cannot be later than the main/extended check-out date (${formatDateTime(`${latestOutD}T${latestOutT}`)}). Please correct it.`);
        }
      }
      
      calculateModalBilling();
    }

    function handleStayDatesChange() {
      const inDateInput = document.getElementById('cust-checkin-date');
      const outDateInput = document.getElementById('cust-checkout-date');
      const extDateInput = document.getElementById('cust-ext-checkout-date');

      if (inDateInput && outDateInput) {
        outDateInput.min = inDateInput.value;
        if (outDateInput.value && outDateInput.value < inDateInput.value) {
          outDateInput.value = inDateInput.value;
        }
      }

      if (outDateInput && extDateInput) {
        if (document.getElementById('cust-has-trip-cancel')?.checked) { handleExtraPersonDatesChange(); return; }
        extDateInput.min = outDateInput.value;
        if (extDateInput.value && extDateInput.value < outDateInput.value) {
          alert("⚠️ Extended Check-Out date cannot be prior to the initial Check-Out date!");
          extDateInput.value = outDateInput.value;
        }
      }

      handleExtraPersonDatesChange(); 
    }

    function handleClearBillPayment(clearAmountVal) {
      const clearVal = parseFloat(clearAmountVal) || 0;
      const total = parseFloat(document.getElementById('cust-total').value) || 0;
      
      const advanceElem = document.getElementById('cust-advance');
      let initialAdvance = parseFloat(advanceElem.getAttribute('data-initial-adv'));
      if (isNaN(initialAdvance)) {
        initialAdvance = parseFloat(advanceElem.value) || 0;
        advanceElem.setAttribute('data-initial-adv', initialAdvance);
      }

      const discountVal = parseFloat(document.getElementById('cust-discount-amount')?.value) || 0;
      if (clearVal + initialAdvance + discountVal > total) {
        alert("⚠️ Payment amount exceeds the remaining balance due!");
        document.getElementById('cust-clear-bill').value = 0;
        recalculateDueAfterDiscount();
        return;
      }

      recalculateDueAfterDiscount();
    }

    function calculateModalBilling() {
      const inDate = document.getElementById('cust-checkin-date').value;
      const inTime = document.getElementById('cust-checkin-time').value || '00:00';
      
      const hasTripCancel = document.getElementById('cust-has-trip-cancel')?.checked;
      const hasExtCheckout = document.getElementById('cust-has-extended-checkout')?.checked;
      let outDate = document.getElementById('cust-checkout-date').value;
      let outTime = document.getElementById('cust-checkout-time').value || '00:00';
      let billingInDate = inDate;
      let billingInTime = inTime;
      if (hasTripCancel) {
        billingInDate = document.getElementById('cust-cancel-checkin-date')?.value || inDate;
        billingInTime = document.getElementById('cust-cancel-checkin-time')?.value || inTime;
        outDate = document.getElementById('cust-cancel-checkout-date')?.value || outDate;
        outTime = document.getElementById('cust-cancel-checkout-time')?.value || outTime;
      } else if (hasExtCheckout) {
        const extDate = document.getElementById('cust-ext-checkout-date')?.value;
        const extTime = document.getElementById('cust-ext-checkout-time')?.value;
        if (extDate) outDate = extDate;
        if (extTime) outTime = extTime;
      }

      let days = 0;
      let latestMainCheckoutDt = null;

      if (billingInDate && outDate) {
        latestMainCheckoutDt = new Date(`${outDate}T${outTime}:00+05:30`);
        const inDateOnly = new Date(billingInDate);
        const outDateOnly = new Date(outDate);
        days = Math.max(1, Math.round((outDateOnly - inDateOnly) / (1000 * 60 * 60 * 24)));
      }

      const price = parseFloat(document.getElementById('cust-price').value) || 0;
      const Extraprice = parseFloat(document.getElementById('cust-extra-rate').value) || 0;
      const capacity = parseFloat(document.getElementById('cust-capacity').value) || 1;

      const extraPersons = parseInt(document.getElementById('cust-extra-persons')?.value) || 0;
      let extraPersonDays = 0;

      if (extraPersons > 0 && latestMainCheckoutDt) {
        const epInDate = document.getElementById('cust-extra-person-date')?.value;
        const epInTime = document.getElementById('cust-extra-person-time')?.value;
        let epOutDate = document.getElementById('cust-extra-person-out-date')?.value;
        let epOutTime = document.getElementById('cust-extra-person-out-time')?.value;

        if (epInDate && epOutDate && epInTime && epOutTime) {
          const epInDt = new Date(`${epInDate}T${epInTime}:00+05:30`);
          let epOutDt = new Date(`${epOutDate}T${epOutTime}:00+05:30`);

          if (epOutDt > latestMainCheckoutDt) {
            epOutDt = new Date(latestMainCheckoutDt.getTime());
            epOutDate = toLocalISOString(epOutDt).split('T')[0];
          }

          if (epInDt < epOutDt) {
            const epInOnly = new Date(epInDate);
            const epOutOnly = new Date(epOutDate);
            const diffDays = Math.round((epOutOnly - epInOnly) / (1000 * 60 * 60 * 24));
            extraPersonDays = Math.max(1, diffDays);
          } else {
            extraPersonDays = 0;
          }
        } else {
           extraPersonDays = 0; 
        }
      }

      const roomTotal = days * price * capacity;
      const extraPersonTotal = extraPersons * extraPersonDays * Extraprice;

      let foodTotalCharge = 0;
      document.querySelectorAll('.cust-food-charge').forEach(input => {
        foodTotalCharge += parseFloat(input.value) || 0;
      });
      
      let cabFare = 0;
      document.querySelectorAll('.cust-cab-rate').forEach(input => {
        cabFare += parseFloat(input.value) || 0;
      });

      const total = roomTotal + extraPersonTotal + foodTotalCharge + cabFare;

      const advanceInput = document.getElementById('cust-advance');
      let currentAdvVal = parseFloat(advanceInput.value) || 0;

      if (currentAdvVal > total) {
        alert(`⚠️ Advance payment (₹${currentAdvVal}) cannot exceed the total bill amount (₹${total})!`);
        currentAdvVal = total;
        advanceInput.value = total;
      }

      advanceInput.setAttribute('data-initial-adv', currentAdvVal);

      const clearBillVal = parseFloat(document.getElementById('cust-clear-bill')?.value) || 0;
      const discountAmt = parseFloat(document.getElementById('cust-discount-amount')?.value) || 0;
      const due = Math.max(0, total - currentAdvVal - discountAmt - clearBillVal);

      document.getElementById('cust-days').value = days;
      document.getElementById('cust-total').value = total;
      updateDiscountPercent();
      recalculateDueAfterDiscount();
      
      // Update Billing Summary fields.
      const extraPersonsCount = parseFloat(document.getElementById('cust-extra-persons').value) || 0;
      const extraPersonRate = parseFloat(document.getElementById('cust-extra-rate').value) || 0;
      const extrapersonTotal = extraPersonsCount * extraPersonRate * extraPersonDays;
      document.getElementById('cust-extra-days').value = extraPersonDays;
      document.getElementById('cust-extra-total').value = Math.round(extrapersonTotal);
      document.getElementById('cust-main-total').value = Math.round(roomTotal);
      
      const cabTotalInput = document.getElementById('cust-cab-total');
      if (cabTotalInput) cabTotalInput.value = cabFare;

         const foodTotalInput = document.getElementById('cust-food-total');
if (foodTotalInput) foodTotalInput.value = foodTotalCharge;
  
    }

    let closedBookingConfirmResolver = null;

    function closeClosedBookingConfirm(goAhead) {
      const modal = document.getElementById('closed-booking-confirm-modal');
      if (modal) modal.classList.add('hidden');

      if (closedBookingConfirmResolver) {
        const resolver = closedBookingConfirmResolver;
        closedBookingConfirmResolver = null;
        resolver(goAhead);
      }
    }

    function askClosedBookingConfirmation() {
      return new Promise((resolve) => {
        closedBookingConfirmResolver = resolve;
        const modal = document.getElementById('closed-booking-confirm-modal');
        if (modal) {
          modal.classList.remove('hidden');
        } else {
          closedBookingConfirmResolver = null;
          resolve(false);
        }
      });
    }

    async function handleSaveBooking(e) {
      e.preventDefault();

      const guestName = formatTitleCase(document.getElementById('cust-name').value.trim());

      if (!guestName) {
        alert("⚠️ Guest Name is a mandatory field!");
        return;
      }

      const contactNoVal = document.getElementById('cust-contact').value.trim();
      if (contactNoVal && contactNoVal.length !== 10) {
        alert("⚠️ Please provide a valid 10-digit guest contact number.");
        return;
      }

      const inDate = document.getElementById('cust-checkin-date').value;
      const outDate = document.getElementById('cust-checkout-date').value;
      const inTime = document.getElementById('cust-checkin-time').value || '00:00';
      const outTime = document.getElementById('cust-checkout-time').value || '00:00';

      const bookingModalId = document.getElementById('modal-booking-id').value;
      const id = bookingModalId;

      let selectedRooms = getSelectedRooms();
      if (selectedRooms.includes("ALL")) {
        selectedRooms = state.roomsCapacity.map(m => String(m.roomNo));
      }

      if (selectedRooms.length === 0) {
        alert("⚠️ Please select at least one Room No.");
        return;
      }

      const checkIn = `${inDate}T${inTime}:00+05:30`;
      const checkOut = `${outDate}T${outTime}:00+05:30`;

      // If both main Check-In and Check-Out dates are earlier than today,
      // ask for confirmation before continuing with the normal save process.
      const todayLocal = new Date();
      const todayDateOnly = new Date(
        todayLocal.getFullYear(),
        todayLocal.getMonth(),
        todayLocal.getDate()
      );
      const checkInDateOnly = inDate ? new Date(`${inDate}T00:00:00`) : null;
      const checkOutDateOnly = outDate ? new Date(`${outDate}T00:00:00`) : null;

      // Show the Closed Booking confirmation ONLY when creating a NEW booking.
      // Existing bookings being edited must not show this confirmation again.
      const isNewBooking = !bookingModalId;

      if (
        isNewBooking &&
        checkInDateOnly &&
        checkOutDateOnly &&
        checkInDateOnly < todayDateOnly &&
        checkOutDateOnly < todayDateOnly
      ) {
        const goAhead = await askClosedBookingConfirmation();
        if (!goAhead) {
          return;
        }
      }

      const extraPersons = parseInt(document.getElementById('cust-extra-persons')?.value) || 0;
      const hasTripCancelForSave = document.getElementById('cust-has-trip-cancel')?.checked || false;
      if (hasTripCancelForSave) {
        const cancelInDate = document.getElementById('cust-cancel-checkin-date').value;
        const cancelInTime = document.getElementById('cust-cancel-checkin-time').value;
        const cancelOutDate = document.getElementById('cust-cancel-checkout-date').value;
        const cancelOutTime = document.getElementById('cust-cancel-checkout-time').value;
        if (!cancelInDate || !cancelInTime || !cancelOutDate || !cancelOutTime ||
            new Date(`${cancelOutDate}T${cancelOutTime}:00`) <= new Date(`${cancelInDate}T${cancelInTime}:00`)) {
          alert('Please enter a valid Trip Cancellation check-in and check-out date/time. Check-out must be after check-in.');
          return;
        }
      }

      if (hasTripCancelForSave && extraPersons > 0) {
        const cancelIn = document.getElementById('cust-cancel-checkin-date').value;
        const cancelOut = document.getElementById('cust-cancel-checkout-date').value;
        const extraIn = document.getElementById('cust-extra-person-date').value;
        const extraOut = document.getElementById('cust-extra-person-out-date').value;
        if ((extraIn && (extraIn < cancelIn || extraIn > cancelOut)) || (extraOut && (extraOut < cancelIn || extraOut > cancelOut)) || (extraIn && extraOut && extraOut < extraIn)) {
          alert('Extra guest check-in and check-out must remain within the Trip Cancel check-in/check-out dates.');
          return;
        }
      }

      const hasExtendedCheckout = document.getElementById('cust-has-extended-checkout')?.checked || false;
      let extendedCheckOut = null;

      if (hasExtendedCheckout) {
        const extDate = document.getElementById('cust-ext-checkout-date').value;
        const extTime = document.getElementById('cust-ext-checkout-time').value || '00:00';
        if (extDate) {
          extendedCheckOut = `${extDate}T${extTime}:00+05:30`;
        }
      }

      const includeMeals = document.getElementById('cust-include-meals')?.checked ?? true;

      const epDateCheck = document.getElementById('cust-extra-person-date')?.value;
      const epOutDateCheck = document.getElementById('cust-extra-person-out-date')?.value;
      
      if (extraPersons === 0 && (epDateCheck || epOutDateCheck)) {
        alert("⚠️ You have entered Additional Person stay dates, but the 'Add Extra Person(s)' count is 0. Please update the person count or clear the dates.");
        return;
      }

      let extraPersonJoined = null;
      let extraPersonOut = null;
      let extraPersonDays = 0;

      const latestCheckoutStr = (hasExtendedCheckout && extendedCheckOut) ? extendedCheckOut : checkOut;
      const latestCheckoutDt = new Date(latestCheckoutStr);
      const mainCheckInDt = new Date(checkIn);

      if (extraPersons > 0) {
        const epDate = document.getElementById('cust-extra-person-date')?.value;
        const epTime = document.getElementById('cust-extra-person-time')?.value;
        let epOutDate = document.getElementById('cust-extra-person-out-date')?.value;
        let epOutTime = document.getElementById('cust-extra-person-out-time')?.value;

        if (!epDate || !epTime || !epOutDate || !epOutTime) {
          alert("⚠️ Please specify custom Check-In and Check-Out dates & times for the Additional Person(s).");
          return;
        }

        extraPersonJoined = `${epDate}T${epTime}:00+05:30`;
        extraPersonOut = `${epOutDate}T${epOutTime}:00+05:30`;

        const epInDt = new Date(extraPersonJoined);
        let epOutDt = new Date(extraPersonOut);
        
        if (epInDt < mainCheckInDt) {
           alert(`⚠️ Additional Person Check-In date & time cannot be earlier than the main Check-In date & time (${formatDateTime(checkIn)}).`);
           return;
        }

        if (epOutDt > latestCheckoutDt) {
          alert(`⚠️ Additional Person Check-Out date & time cannot exceed main/extended Check-Out date & time (${formatDateTime(latestCheckoutStr)}).`);
          extraPersonOut = latestCheckoutStr;
          epOutDt = latestCheckoutDt;
          epOutDate = toLocalISOString(epOutDt).split('T')[0];
        }

        if (epInDt < epOutDt) {
          const epInOnly = new Date(epDate);
          const epOutOnly = new Date(epOutDate);
          extraPersonDays = Math.max(1, Math.round((epOutOnly - epInOnly) / (1000 * 60 * 60 * 24)));
        }
      }

      const foodWin = getModalFoodWindow();
      const foodOrdersList = [];
      let foodValidationError = false;

      document.querySelectorAll('.food-order-row').forEach(row => {
        const desc = row.querySelector('.cust-food-desc').value || '';
        const itemPrice = parseFloat(row.querySelector('.cust-food-price').value) || 0;
        const plates = parseInt(row.querySelector('.cust-food-plates').value) || 0;
        const charge = parseFloat(row.querySelector('.cust-food-charge').value) || 0;
        const fDate = row.querySelector('.cust-food-date').value || '';
        const fTime = row.querySelector('.cust-food-time').value || '';

        const foodDateTime = (fDate && fTime) ? `${fDate}T${fTime}:00+05:30` : (fDate ? `${fDate}T00:00:00+05:30` : '');

        if (desc || charge > 0) {
          if (foodWin && foodDateTime) {
            const fDt = new Date(foodDateTime);
            if (fDt < foodWin.minFoodDt || fDt > foodWin.maxFoodDt) {
              foodValidationError = true;
            }
          }
          foodOrdersList.push({
            foodDesc: desc,
            itemPrice: itemPrice,
            plates: plates,
            foodCharge: charge,
            foodDateTime: foodDateTime
          });
        }
      });

      if (foodValidationError && foodWin) {
        const minStr = formatDateTime(foodWin.minFoodDt);
        const maxStr = formatDateTime(foodWin.maxFoodDt);
        alert(`❌ Extra Food Order Validation Error!\n\nAll Extra Food order times must be strictly after 15 minutes of Check-In (${minStr}) and at least 30 minutes before Check-Out / Extended Check-Out (${maxStr}).`);
        return;
      }
      
      const bookingRemark = document.getElementById('cust-booking-remark').value.trim();
      const remarkWords = bookingRemark ? bookingRemark.split(/\s+/).length : 0;
      if (remarkWords > 1000) {
        alert('Remark cannot exceed 1000 words.');
        document.getElementById('cust-booking-remark').focus();
        return;
      }
      const cabTripsList = [];

      document.querySelectorAll('.cab-trip-row').forEach((row, index) => {
        const rate = parseFloat(row.querySelector('.cust-cab-rate').value) || 0;
        const dateVal = row.querySelector('.cust-cab-date').value || '';
        const timeVal = row.querySelector('.cust-cab-time').value || '';
        const remark = row.querySelector('.cust-cab-remark').value || '';
        const dt = (dateVal && timeVal) ? `${dateVal}T${timeVal}:00+05:30` : '';

        if (rate > 0 || remark) {
          cabTripsList.push({
            tripName: `Trip ${index + 1}`,
            dateStr: dateVal,
            timeStr: timeVal,
            dateTime: dt,
            rate: rate,
            remark: remark
          });
        }
      });

      const effectiveCheckout = (hasExtendedCheckout && extendedCheckOut) ? extendedCheckOut : checkOut;
      const newIn = parseDateMs(checkIn);
      const newOut = parseDateMs(effectiveCheckout);

      if (isNaN(newIn) || isNaN(newOut) || newIn >= newOut) {
        alert("Check-Out date & time must be strictly after Check-In date & time.");
        return;
      }

      // Check for same room / date conflict
      const conflict = state.bookings.find(b => {
        if (isInactiveBooking(b)) return false;
        if (id && String(b.id) === String(id)) return false;
        
        const bRooms = getBookingRooms(b);
        const hasRoomOverlap = selectedRooms.some(r => bRooms.includes(String(r)));
        if (!hasRoomOverlap) return false;

        const existingIn = parseDateMs(b.checkIn);
        const existingOutVal = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
        const existingOut = parseDateMs(existingOutVal);

        if (isNaN(existingIn) || isNaN(existingOut)) return false;

        return (newIn < existingOut && newOut > existingIn);
      });

      if (conflict) {
        const confOutVal = (isTrue(conflict.hasExtendedCheckout) && conflict.extendedCheckOut) ? conflict.extendedCheckOut : conflict.checkOut;
        const confOutFormatted = formatDateTime(confOutVal);
        const conflictRooms = getBookingRooms(conflict).join(', ');
        alert(`❌ Booking Conflict Alert!\n\nRoom(s) ${conflictRooms} is already occupied by ${conflict.name} until ${confOutFormatted}.\n\nPlease select a check-in time after ${confOutFormatted} or assign different rooms.`);
        return;
      }

      let existingCode = null;
      let existingInv = null;

      if (id) {
        const existing = state.bookings.find(b => String(b.id) === String(id));
        if (existing) {
          existingCode = existing.bookingCode;
          existingInv = existing.invoiceNo;
        }
      }

      if (!existingCode) {
        const generated = generateIDsForYear(checkIn);
        existingCode = generated.bookingCode;
        existingInv = generated.invoiceNo;
      }

      const totalAmt = parseFloat(document.getElementById('cust-total').value) || 0;
      const initialAdvAmt = parseFloat(document.getElementById('cust-advance').getAttribute('data-initial-adv')) || parseFloat(document.getElementById('cust-advance').value) || 0;
      const clearedDueAmt = parseFloat(document.getElementById('cust-clear-bill').value) || 0;

      const discountAmountForSave = parseFloat(document.getElementById('cust-discount-amount').value) || 0;
      const totalPaid = initialAdvAmt + clearedDueAmt + discountAmountForSave;
      const countryCodeVal = document.getElementById('cust-country-code').value.trim() || '+91';

      const newBooking = {
        id: id || `bk_${Date.now()}`,
        bookingCode: existingCode,
        invoiceNo: existingInv,
        name: guestName,
        address: formatTitleCase(document.getElementById('cust-address').value.trim()),
        city: formatTitleCase(document.getElementById('cust-city').value.trim()),
        state: formatTitleCase(document.getElementById('cust-state').value.trim()),
        country: formatTitleCase(document.getElementById('cust-country').value.trim()),
        zipCode: document.getElementById('cust-zip').value.trim(),
        idNo: document.getElementById('cust-id').value.trim(),
        countryCode: countryCodeVal,
        contactNo: contactNoVal,
        idProofBase64: document.getElementById('cust-id-file-base64').value,
        idProofFileName: document.getElementById('cust-id-file-name').value,
        roomNo: selectedRooms.join('|'),
        agentInfo: document.getElementById('cust-agent').value,
        capacity: parseInt(document.getElementById('cust-capacity').value) || 1,
        extraPersons: extraPersons,
        extraPersonJoined: extraPersonJoined,
        extraPersonOut: extraPersonOut,
        checkIn: checkIn,
        checkOut: checkOut,
        hasExtendedCheckout: hasExtendedCheckout,
        extendedCheckOut: extendedCheckOut,
        hasTripCancel: document.getElementById('cust-has-trip-cancel').checked,
        tripCancelCheckIn: (document.getElementById('cust-cancel-checkin-date').value && document.getElementById('cust-cancel-checkin-time').value) ? `${document.getElementById('cust-cancel-checkin-date').value}T${document.getElementById('cust-cancel-checkin-time').value}:00+05:30` : '',
        tripCancelCheckOut: (document.getElementById('cust-cancel-checkout-date').value && document.getElementById('cust-cancel-checkout-time').value) ? `${document.getElementById('cust-cancel-checkout-date').value}T${document.getElementById('cust-cancel-checkout-time').value}:00+05:30` : '',
        includeMeals: includeMeals,
        foodOrders: foodOrdersList,
        cabTrips: cabTripsList,
        remark: (document.getElementById('cust-booking-remark').value || '').trim(),
        ExtraperDayPrice: parseFloat(document.getElementById('cust-extra-rate').value) || 0,
        extraPersonDays: parseInt(document.getElementById('cust-extra-days').value) || 0,
        custextratotal: parseInt(document.getElementById('cust-extra-total').value) || 0,
        custcabtotal: parseInt(document.getElementById('cust-cab-total').value) || 0,
        perDayPrice: parseFloat(document.getElementById('cust-price').value) || 0,
        noOfDays: parseInt(document.getElementById('cust-days').value) || 0,
        custmaintotal: parseInt(document.getElementById('cust-main-total').value) || 0,
        custfoodtotal: parseInt(document.getElementById('cust-food-total').value) || 0,
        totalAmount: totalAmt,
        initialAdv: initialAdvAmt,
        totalDue: Math.max(0, totalAmt - totalPaid),
        clearedDue: clearedDueAmt,
        discountPercent: parseFloat(document.getElementById('cust-discount-percent').value) || 0,
        discountAmount: parseFloat(document.getElementById('cust-discount-amount').value) || 0,
        inactive: false
      };

      if (id) {
        const idx = state.bookings.findIndex(b => String(b.id) === String(id));
        if (idx !== -1) state.bookings[idx] = newBooking;
      } else {
        state.bookings.push(newBooking);
      }

      closeBookingModal();
      searchMasterBookingById();
      renderBookingsTable();
      updateDashboardCards();
      renderCalendar(defaultAppYear);
      checkUpcomingCheckoutsWithDue();
      saveChanges(false, false);

      document.getElementById('cust-checkin-date').value = "";
      document.getElementById('cust-checkin-time').value = "";
      document.getElementById('cust-checkout-date').value = "";
      document.getElementById('cust-checkout-time').value = "";
      document.getElementById('cust-ext-checkout-date').value = "";
      document.getElementById('cust-ext-checkout-time').value = "";
    }

    function deleteBooking(id) {
      openMasterDeleteModal('booking', id);
    }

    function renderBookingsTable(dateFilter = "") {
      const tbody = document.getElementById('bookings-tbody');
      tbody.innerHTML = '';

      let listToRender = [...state.bookings];

      if (dateFilter) {
        listToRender = listToRender.filter(b => {
          if (!b.checkIn || !b.checkOut) return false;
          const bIn = String(b.checkIn).replace(' ', 'T').split('T')[0];
          const bOutVal = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
          const bOut = String(bOutVal).replace(' ', 'T').split('T')[0];
          return (dateFilter >= bIn && dateFilter <= bOut);
        });
      }

      const now = new Date().getTime();

      const getStatusPriority = (b) => {
        const isMasterValid = isRoomInMaster(b.roomNo);
        if (!isMasterValid || isInactiveBooking(b)) {
          return 4; 
        }

        const cIn = parseDateMs(b.checkIn);
        const cOut = getEffectiveCheckoutTime(b);

        if (now > cOut) {
          return 3; 
        } else if (now >= cIn && now <= cOut) {
          return 1; 
        } else {
          return 2; 
        }
      };

      listToRender.sort((a, b) => {
        const priorityA = getStatusPriority(a);
        const priorityB = getStatusPriority(b);

        if (priorityA !== priorityB) {
          return priorityA - priorityB;
        }

        return (b.bookingCode || '').localeCompare(a.bookingCode || '');
      });

      if (listToRender.length === 0) {
        tbody.innerHTML = `<tr><td colspan="13" class="text-center py-6 text-slate-400">No bookings found for the selected date.</td></tr>`;
        return;
      }

      listToRender.forEach((b) => {
        const isMasterValid = isRoomInMaster(b.roomNo);
        const checkInFmt = formatDateTime(b.checkIn);
        
        const effectiveOut = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
        const checkOutFmt = formatDateTime(effectiveOut);

        const checkInTime = parseDateMs(b.checkIn);
        const checkOutTime = getEffectiveCheckoutTime(b);

        const isClosed = now > checkOutTime;
        const isInactive = isInactiveBooking(b);
        const roomsDisplay = getBookingRooms(b).join(', ');

        let statusBgClass = "hover:bg-slate-50";
        let statusDotHtml = "";

        if (!isMasterValid) {
          statusBgClass = "bg-rose-100/70 hover:bg-rose-200/60 text-rose-900";
        } else if (isInactive) {
          statusBgClass = "bg-slate-100 hover:bg-slate-200 text-slate-500 opacity-75";
        } else if (isClosed) {
          statusDotHtml = `<span class="w-2.5 h-2.5 bg-emerald-500 rounded-full inline-block flex-shrink-0" title="Closed Booking"></span>`;
        } else if (now >= checkInTime && now <= checkOutTime) {
          statusDotHtml = `
            <span class="relative flex h-2.5 w-2.5 flex-shrink-0" title="Live Booking">
              <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-400 opacity-75"></span>
              <span class="relative inline-flex rounded-full h-2.5 w-2.5 bg-amber-500"></span>
            </span>
          `;
        } else if (now < checkInTime) {
          statusDotHtml = `<span class="w-2.5 h-2.5 bg-blue-500 rounded-full inline-block flex-shrink-0" title="Upcoming Booking"></span>`;
        }

        let foodSummaryHtml = '';
        const parseFood = parseJSONField(b.foodOrders);
        if (parseFood.length > 0) {
          const totalFoodCharge = parseFood.reduce((acc, fo) => acc + (fo.foodCharge || 0), 0);
          if (totalFoodCharge > 0) {
            foodSummaryHtml = `<div class="text-[9px] ${!isMasterValid ? 'text-rose-950 font-bold' : 'text-amber-800 font-semibold'}"><i class="fa-solid fa-utensils text-[8px] mr-0.5"></i>Food (${parseFood.length}): +₹${totalFoodCharge}</div>`;
          }
        }
        
        let cabSummaryHtml = '';
        const parseCab = parseJSONField(b.cabTrips);
        const totalCab = parseCab.reduce((acc, t) => acc + (t.rate || 0), 0);
        
        if (totalCab > 0) {
          cabSummaryHtml = `<div class="text-[9px] ${!isMasterValid ? 'text-rose-950 font-bold' : 'text-indigo-800 font-semibold'}"><i class="fa-solid fa-taxi text-[8px] mr-0.5"></i>Cab: +₹${totalCab}</div>`;
        }

        const printOnClick = `printInvoice('${b.id}')`;
        
        let actionButtonsHtml = `
          <div class="flex items-center justify-center space-x-1">
            <button onclick="openBookingModal('${b.id}')" class="text-blue-600 hover:text-blue-800 p-1 text-sm" title="Edit Booking Details">
              <i class="fa-solid fa-pen-to-square"></i>
            </button>              
            <button onclick="${printOnClick}" class="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1 rounded-full text-[11px] font-bold transition shadow-xs">Print</button>
          </div>
        `;

        let idProofCellHtml = `<span class="text-slate-400 italic text-[10px]">None</span>`;
        if (b.idProofBase64) {
          idProofCellHtml = `
            <button onclick="openPdfAttachment('${b.idProofBase64}')" class="bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 px-2 py-0.5 rounded-full text-[10px] font-bold flex items-center gap-1 transition">
              <i class="fa-solid fa-file-pdf text-rose-600"></i> View PDF
            </button>
          `;
        }

        const tableCap = parseInt(b.capacity) || 1;
        const tableCapLabel = tableCap === 1 ? 'Person' : 'Persons';
        const extraPersonsText = (b.extraPersons && b.extraPersons > 0) ? `<span class="text-amber-700 font-bold block text-[9px]">(+${b.extraPersons} Extra)</span>` : '';
        const contactDisplay = b.contactNo ? `${b.countryCode || '+91'} ${b.contactNo}`.trim() : '-';

        const totalReceived = (b.initialAdv || 0) + (b.clearedDue || 0);

        const tr = document.createElement('tr');
        tr.className = `${statusBgClass} transition border-b border-slate-100`;
        tr.innerHTML = `
          <td class="py-2.5 px-3">
            <div class="flex items-center gap-1.5">
              ${statusDotHtml}
              ${isInactive ? '<span class="w-2.5 h-2.5 bg-rose-600 rounded-full inline-block flex-shrink-0" title="Inactive Booking"></span>' : ''}
              <span class="bg-blue-50 border border-blue-200 text-blue-700 font-mono font-bold px-2 py-0.5 rounded-full text-[9px] block w-max">${b.bookingCode}</span>
            </div>
            ${isInactive ? '<span class="bg-slate-600 text-white font-bold px-1.5 py-0.2 rounded-full text-[8px] uppercase block mt-0.5 w-max">Inactive</span>' : (!isMasterValid ? '<span class="bg-rose-700 text-white font-bold px-1.5 py-0.2 rounded-full text-[8px] uppercase block mt-0.5 w-max">Master Removed</span>' : '')}
          </td>
          <td class="py-2.5 px-3 font-bold ${!isMasterValid ? 'text-rose-950' : 'text-slate-800'}">${b.name}</td>
          <td class="py-2.5 px-3 font-medium whitespace-nowrap">${contactDisplay}</td>
          <td class="py-2.5 px-3 font-mono text-[10px]">${b.idNo || '-'}</td>
          <td class="py-2.5 px-3">${idProofCellHtml}</td>
          <td class="py-2.5 px-3"><span class="bg-blue-50 text-blue-700 font-bold px-2 py-0.5 rounded-full text-[10px] break-all">Room ${roomsDisplay}</span></td>
          <td class="py-2.5 px-3 font-bold text-slate-700">${tableCap} ${tableCapLabel} ${extraPersonsText}</td>
          <td class="py-2.5 px-3 ${!isMasterValid ? 'text-rose-900' : 'text-slate-600'} text-[10px]">${b.agentInfo || '-'}</td>
          <td class="py-2.5 px-3 text-[10px]">
            <div class="font-bold text-blue-600">${checkInFmt}</div>
            <div class="text-blue-600 font-bold text-[10px]">to ${checkOutFmt} ${isTrue(b.hasExtendedCheckout) ? '<span class="text-blue-600 font-bold">(Ext)</span>' : ''}</div>
          </td>
          <td class="py-2.5 px-3">
            <div class="font-bold ${!isMasterValid ? 'text-rose-950' : 'text-slate-800'}">₹${b.perDayPrice}/day (${b.noOfDays}d)</div>
            ${foodSummaryHtml}
            ${cabSummaryHtml}
          </td>
          <td class="py-2.5 px-3 font-bold ${!isMasterValid ? 'text-rose-950' : 'text-slate-800'}">
            ₹${b.totalAmount}
            <span class="block text-[9px] text-emerald-600 font-medium">Adv: ₹${totalReceived}</span>
          </td>
          <td class="py-2.5 px-3 font-bold ${b.totalDue > 0 ? 'text-rose-600' : 'text-emerald-600'}">₹${b.totalDue}</td>
          <td class="py-2.5 px-3 text-center">
            ${actionButtonsHtml}
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function renderRoomCapacityTable() {
      const tbody = document.getElementById('room-capacity-tbody');
      tbody.innerHTML = '';

      if (!state.roomsCapacity || state.roomsCapacity.length === 0) {
        tbody.innerHTML = `<tr><td colspan="3" class="text-center py-4 text-slate-400">No room capacity data available.</td></tr>`;
        return;
      }

      state.roomsCapacity.forEach((r, idx) => {
        const tr = document.createElement('tr');
        tr.className = "bg-white hover:bg-slate-50 transition border-b border-slate-100";
        tr.innerHTML = `
          <td class="py-2.5 px-3">
            <input type="number" value="${r.roomNo}" min="1" oninput="state.roomsCapacity[${idx}].roomNo = parseInt(this.value) || 1" onchange="populateRoomDropdown(); renderBookingsTable(); saveChanges(false, true)" class="w-24 bg-transparent font-bold text-blue-600 focus:bg-white focus:border focus:border-blue-300 rounded-xl px-2 py-1">
          </td>
          <td class="py-2.5 px-3">
            <input type="number" value="${r.capacity}" min="1" oninput="state.roomsCapacity[${idx}].capacity = parseInt(this.value) || 1" onchange="saveChanges(false, true)" class="w-24 bg-transparent font-semibold text-slate-800 focus:bg-white focus:border focus:border-blue-300 rounded-xl px-2 py-1">
          </td>
          <td class="py-2.5 px-3 text-center">
            <button type="button" onclick="removeRoomCapacityRow(${idx})" class="text-rose-500 hover:text-rose-700 p-1 text-xs" title="Delete Room Entry">
              <i class="fa-solid fa-trash-can"></i>
            </button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function renderMasterAgentTable() {
      const tbody = document.getElementById('agent-tbody');
      tbody.innerHTML = '';

      if (!state.masterAgents || state.masterAgents.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" class="text-center py-4 text-slate-400">No agent data available.</td></tr>`;
        return;
      }

      state.masterAgents.forEach((a, idx) => {
        const tr = document.createElement('tr');
        tr.className = "bg-white hover:bg-slate-50 transition border-b border-slate-100";
        tr.innerHTML = `
          <td class="py-2.5 px-3">
            <input type="text" value="${a.agentName}" oninput="state.masterAgents[${idx}].agentName = formatTitleCase(this.value)" onchange="populateAgentDropdown(); saveChanges(false, true)" class="w-full bg-transparent font-semibold text-slate-800 focus:bg-white focus:border focus:border-blue-300 rounded-xl px-2 py-1">
          </td>
          <td class="py-2.5 px-3">
            <input type="text" value="${a.phone}" oninput="state.masterAgents[${idx}].phone = this.value" onchange="populateAgentDropdown(); saveChanges(false, true)" class="w-full bg-transparent text-slate-600 focus:bg-white focus:border focus:border-blue-300 rounded-xl px-2 py-1">
          </td>
          <td class="py-2.5 px-3 font-bold text-blue-600">
            ${a.roomNo || 'All Rooms'}
          </td>
          <td class="py-2.5 px-3 text-center">
            <button type="button" onclick="removeAgentRow(${idx})" class="text-rose-500 hover:text-rose-700 p-1 text-xs" title="Delete Agent Entry">
              <i class="fa-solid fa-trash-can"></i>
            </button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function addRoomCapacityRow() {
      if (!state.roomsCapacity) state.roomsCapacity = [];
      const existingRoomNos = state.roomsCapacity.map(r => parseInt(r.roomNo) || 0);
      let nextRoom = 1;
      if (existingRoomNos.length > 0) {
        nextRoom = Math.max(...existingRoomNos) + 1;
      }

      state.roomsCapacity.push({
        roomNo: nextRoom,
        capacity: 4
      });

      renderRoomCapacityTable();
      populateRoomDropdown();
      saveChanges(false, false);
    }

    function removeRoomCapacityRow(index) {
      openMasterDeleteModal('room', index);
    }

    function addAgentRow() {
      if (!state.masterAgents) state.masterAgents = [];
      const nextNum = state.masterAgents.length;
      state.masterAgents.push({
        agentName: `Agent ${nextNum + 1}`,
        phone: "1234567890",
        roomNo: "All Rooms"
      });

      renderMasterAgentTable();
      populateAgentDropdown();
      saveChanges(false, false);
    }

    function removeAgentRow(index) {
      openMasterDeleteModal('agent', index);
    }

    function renderCalendar(year) {
      state.selectedYear = year;
      const calSelect = document.getElementById('cal-year-select');
      if (calSelect) calSelect.value = year;

      const container = document.getElementById('calendar-container');
      container.innerHTML = '';

      const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

      months.forEach((monthName, monthIndex) => {
        const monthBox = document.createElement('div');
        monthBox.className = "bg-slate-50/80 rounded-3xl p-3 border border-slate-200/60 shadow-xs flex flex-col justify-between";

        const title = document.createElement('h4');
        title.className = "font-bold text-slate-800 text-[11px] mb-2 pb-1 border-b border-slate-200/60 flex justify-between items-center px-1";
        title.innerHTML = `<span>${monthName}</span> <span class="text-[9px] text-blue-600 font-mono font-normal">${year}</span>`;
        monthBox.appendChild(title);

        const grid = document.createElement('div');
        grid.className = "grid grid-cols-7 gap-1 text-center text-[9px] font-medium text-slate-500 mb-1";
        
        ['S', 'M', 'T', 'W', 'T', 'F', 'S'].forEach(d => {
          const dh = document.createElement('div');
          dh.innerText = d;
          dh.className = "font-bold text-slate-400";
          grid.appendChild(dh);
        });

        const firstDay = new Date(year, monthIndex, 1).getDay();
        const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();

        for (let i = 0; i < firstDay; i++) {
          const empty = document.createElement('div');
          grid.appendChild(empty);
        }

        const todayObj = new Date();
        const isCurrentYearAndMonth = todayObj.getFullYear() === year && todayObj.getMonth() === monthIndex;

        for (let day = 1; day <= daysInMonth; day++) {
          const cell = document.createElement('div');
          const dateStr = `${year}-${String(monthIndex + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
          
          const matchingBookings = state.bookings.filter(b => {
            if (isInactiveBooking(b) || !b.checkIn || !b.checkOut) return false;
            if (!isRoomInMaster(b.roomNo)) return false;

            const bIn = String(b.checkIn).replace(' ', 'T').split('T')[0];
            const bOutVal = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
            const bOut = String(bOutVal).replace(' ', 'T').split('T')[0];

            return (dateStr >= bIn && dateStr <= bOut);
          });

          const isBooked = matchingBookings.length > 0;
          const isToday = isCurrentYearAndMonth && todayObj.getDate() === day;

          cell.className = `py-1 rounded-xl text-[10px] font-bold cursor-pointer transition relative flex items-center justify-center ${
            isToday 
              ? 'ring-2 ring-blue-600 ring-offset-1 z-10' 
              : ''
          } ${
            isBooked 
              ? 'bg-amber-400 text-slate-900 hover:bg-amber-500 shadow-xs' 
              : 'bg-white text-slate-700 hover:bg-blue-50 hover:text-blue-600 border border-slate-100'
          }`;

          cell.innerText = day;

          if (isBooked) {
            cell.onclick = (e) => {
              e.stopPropagation();
              showExcelCommentBox(e, dateStr, matchingBookings);
            };
          }

          grid.appendChild(cell);
        }

        monthBox.appendChild(grid);
        container.appendChild(monthBox);
      });
    }

    function showExcelCommentBox(e, dateStr, bookings) {
      const box = document.getElementById('excel-comment-box');
      const dateHeader = document.getElementById('comm-date-header');
      const listContainer = document.getElementById('comm-booking-list');

      dateHeader.innerText = formatDate(dateStr);
      listContainer.innerHTML = '';

      const now = new Date().getTime();

      bookings.forEach(b => {
        const item = document.createElement('div');
        item.className = "bg-slate-800 p-2.5 rounded-2xl border border-slate-700 space-y-1 hover:border-blue-400 transition cursor-pointer";
        item.onclick = () => {
          closeCommentBox();
          openBookingModal(b.id);
        };

        const cInMs = parseDateMs(b.checkIn);
        const cOutMs = getEffectiveCheckoutTime(b);
        const effectiveCheckout = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
        const roomsDisplay = getBookingRooms(b).join(', ');

        let statusText = "Upcoming";
        let statusColorClass = "text-blue-400 font-bold";

        if (now > cOutMs) {
          statusText = "Closed";
          statusColorClass = "text-emerald-400 font-bold";
        } else if (now >= cInMs && now <= cOutMs) {
          statusText = "Live";
          statusColorClass = "text-amber-400 font-bold";
        }

        item.innerHTML = `
          <div class="flex justify-between items-center">
            <span class="font-bold text-blue-400 text-[11px]">${b.name}</span>
            <span class="bg-blue-900/60 text-blue-200 px-2 py-0.5 rounded-full text-[9px] font-mono">Room ${roomsDisplay}</span>
          </div>
          <div class="text-[9px] text-slate-300">
            Check-In: ${formatDateTime(b.checkIn)}<br>
            Check-Out: ${formatDateTime(effectiveCheckout)}<br>
            Status: <span class="${statusColorClass}">${statusText}</span>
          </div>
          <div class="flex justify-between items-center text-[9px] pt-1 border-t border-slate-700/60">
            <span class="text-emerald-400 font-semibold">Total: ₹${b.totalAmount}</span>
            <span class="${b.totalDue > 0 ? 'text-rose-400 font-bold' : 'text-emerald-400 font-bold'}">
              ${b.totalDue > 0 ? `Due: ₹${b.totalDue}` : 'Paid'}
            </span>
          </div>
        `;
        listContainer.appendChild(item);
      });

      const rect = e.target.getBoundingClientRect();
      const scrollY = window.scrollY || window.pageYOffset;
      const scrollX = window.scrollX || window.pageXOffset;

      box.style.top = `${rect.bottom + scrollY + 5}px`;
      
      let leftPos = rect.left + scrollX - 20;
      if (leftPos + 260 > window.innerWidth) {
        leftPos = window.innerWidth - 270;
      }
      box.style.left = `${Math.max(10, leftPos)}px`;

      box.classList.remove('hidden');
    }

    function closeCommentBox() {
      const box = document.getElementById('excel-comment-box');
      if (box) box.classList.add('hidden');
    }
  </script>
</body>
</html>
