<!DOCTYPE html>
<html lang="en-US">
  <head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">

<title>Test</title>
<meta name="generator" content="Jekyll v3.10.0" />
<meta property="og:title" content="Test" />
<meta property="og:locale" content="en_US" />
<link rel="canonical" href="https://datacraft360degree-page.github.io/Test/" />
<meta property="og:url" content="https://datacraft360degree-page.github.io/Test/" />
<meta property="og:site_name" content="Test" />
<meta property="og:type" content="website" />
<meta name="twitter:card" content="summary" />
<meta property="twitter:title" content="Test" />
<script type="application/ld+json">
{"@context":"https://schema.org","@type":"WebSite","headline":"Test","name":"Test","url":"https://datacraft360degree-page.github.io/Test/"}</script>
<link rel="stylesheet" href="/Test/assets/css/style.css?v=1a46f3fa793c940ba908cd2cc195662dd0bbab15">
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
      /* Custom scrollbars and grid styling */
      .grid-container::-webkit-scrollbar { height: 8px; width: 8px; }
      .grid-container::-webkit-scrollbar-track { background: #0f172a; }
      .grid-container::-webkit-scrollbar-thumb { background: #334155; border-radius: 4px; }
      .grid-container::-webkit-scrollbar-thumb:hover { background: #475569; }
      
      .excel-cell {
        transition: all 0.15s ease;
      }
      .excel-cell:hover {
        outline: 2px solid #38bdf8;
        z-index: 10;
      }
    </style>
  </head>
  <body class="bg-slate-950 text-slate-100 font-sans min-h-screen">
    <div class="container-lg px-3 my-5 markdown-body max-w-7xl mx-auto">
      
      <h1><a href="https://datacraft360degree-page.github.io/Test/" class="text-2xl font-bold text-sky-400 no-underline">Hotel Booking Matrix</a></h1>

      <div class="bg-slate-900 border border-slate-800 rounded-xl p-4 my-4 shadow-xl flex flex-wrap justify-between items-center gap-4">
        <div class="flex items-center gap-3">
          <label for="month-picker" class="text-sm font-semibold text-slate-300"><i class="fa-regular font-bold fa-calendar-days text-sky-400 mr-1"></i> Month:</label>
          <input type="month" id="month-picker" class="bg-slate-800 border border-slate-700 text-slate-100 px-3 py-1.5 rounded-lg text-sm focus:outline-none focus:border-sky-500">
        </div>

        <div class="flex items-center gap-3 flex-1 max-w-md">
          <div class="relative w-full">
            <i class="fa-solid fa-magnifying-glass absolute left-3 top-2.5 text-slate-500 text-sm"></i>
            <input type="text" id="search-input" placeholder="Search Guest / Room / Booking ID..." class="w-full bg-slate-800 border border-slate-700 text-slate-100 pl-9 pr-3 py-1.5 rounded-lg text-sm focus:outline-none focus:border-sky-500">
          </div>
        </div>

        <div class="flex items-center gap-2">
          <div id="sync-status" class="text-xs px-3 py-1.5 rounded-full bg-slate-800 text-slate-400 border border-slate-700 flex items-center gap-1.5">
            <span class="w-2 h-2 rounded-full bg-slate-500 inline-block" id="sync-dot"></span>
            <span id="sync-text">Idle</span>
          </div>

          <button onclick="fetchFromGoogleSheet()" class="bg-sky-600 hover:bg-sky-500 text-white text-xs font-semibold px-3 py-2 rounded-lg transition flex items-center gap-1.5 shadow">
            <i class="fa-solid fa-arrows-rotate"></i> Refresh Sheet
          </button>

          <button onclick="openBookingModal()" class="bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold px-3 py-2 rounded-lg transition flex items-center gap-1.5 shadow">
            <i class="fa-solid fa-plus"></i> New Booking
          </button>

          <button onclick="wipeData()" class="bg-rose-600 hover:bg-rose-500 text-white text-xs font-semibold px-3 py-2 rounded-lg transition flex items-center gap-1.5 shadow">
            <i class="fa-solid fa-trash-can"></i> Wipe Data
          </button>
        </div>
      </div>

      <div class="grid-container overflow-x-auto bg-slate-900 border border-slate-800 rounded-xl shadow-2xl relative min-h-[450px]">
        <table class="w-full border-collapse text-left text-xs" id="matrix-table">
          <thead id="matrix-head" class="bg-slate-800 text-slate-300 sticky top-0 z-20">
            </thead>
          <tbody id="matrix-body" class="divide-y divide-slate-800 text-slate-300">
            </tbody>
        </table>
      </div>

      <div id="excel-comment-box" class="hidden absolute z-50 bg-slate-900 border border-sky-500 text-slate-100 rounded-lg p-3 shadow-2xl w-64 text-xs pointer-events-none transition-all duration-150">
        <div class="flex justify-between items-center border-b border-slate-800 pb-1 mb-2">
          <span class="font-bold text-sky-400" id="comment-booking-id">BK-0000</span>
          <span class="text-[10px] text-slate-400" id="comment-room">Room</span>
        </div>
        <div id="comment-details-list" class="space-y-1.5">
          </div>
      </div>

      <div id="booking-modal" class="fixed inset-0 bg-slate-950/80 backdrop-blur-sm z-50 hidden flex justify-center items-center p-4">
        <div class="bg-slate-900 border border-slate-800 rounded-xl shadow-2xl w-full max-w-lg overflow-hidden">
          <div class="bg-slate-800/80 px-5 py-3 border-b border-slate-700 flex justify-between items-center">
            <h3 class="font-semibold text-slate-100 text-sm flex items-center gap-2">
              <i class="fa-solid fa-hotel text-sky-400"></i> <span id="modal-title">Booking Details</span>
            </h3>
            <button onclick="closeBookingModal()" class="text-slate-400 hover:text-white text-lg">&times;</button>
          </div>

          <form id="booking-form" onsubmit="handleFormSubmit(event)" class="p-5 space-y-4">
            <input type="hidden" id="booking-index">
            
            <div class="grid grid-cols-2 gap-3">
              <div>
                <label class="block text-[11px] font-medium text-slate-400 mb-1">Booking ID</label>
                <input type="text" id="booking-id" readonly class="w-full bg-slate-800/50 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-sky-400 font-mono font-bold focus:outline-none">
              </div>
              <div>
                <label class="block text-[11px] font-medium text-slate-400 mb-1">Status</label>
                <select id="booking-status" class="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-slate-100 focus:outline-none focus:border-sky-500">
                  <option value="Confirmed">Confirmed</option>
                  <option value="Checked-In">Checked-In</option>
                  <option value="Checked-Out">Checked-Out</option>
                  <option value="Blocked">Blocked</option>
                  <option value="Cancelled">Cancelled</option>
                </select>
              </div>
            </div>

            <div class="grid grid-cols-2 gap-3">
              <div>
                <label class="block text-[11px] font-medium text-slate-400 mb-1">Guest Name</label>
                <input type="text" id="guest-name" required class="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-slate-100 focus:outline-none focus:border-sky-500">
              </div>
              <div>
                <label class="block text-[11px] font-medium text-slate-400 mb-1">Phone Number</label>
                <input type="text" id="guest-phone" class="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-slate-100 focus:outline-none focus:border-sky-500">
              </div>
            </div>

            <div class="grid grid-cols-3 gap-3">
              <div>
                <label class="block text-[11px] font-medium text-slate-400 mb-1">Room</label>
                <select id="room-select" required class="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-slate-100 focus:outline-none focus:border-sky-500">
                  </select>
              </div>
              <div>
                <label class="block text-[11px] font-medium text-slate-400 mb-1">Check-In Date</label>
                <input type="date" id="checkin-date" required class="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-slate-100 focus:outline-none focus:border-sky-500">
              </div>
              <div>
                <label class="block text-[11px] font-medium text-slate-400 mb-1">Check-Out Date</label>
                <input type="date" id="checkout-date" required class="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-slate-100 focus:outline-none focus:border-sky-500">
              </div>
            </div>

            <div class="grid grid-cols-3 gap-3">
              <div>
                <label class="block text-[11px] font-medium text-slate-400 mb-1">Total Amount (₹)</label>
                <input type="number" id="total-amount" value="0" min="0" oninput="calcDue()" class="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-slate-100 focus:outline-none focus:border-sky-500">
              </div>
              <div>
                <label class="block text-[11px] font-medium text-slate-400 mb-1">Advance / Paid (₹)</label>
                <input type="number" id="paid-amount" value="0" min="0" oninput="calcDue()" class="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-slate-100 focus:outline-none focus:border-sky-500">
              </div>
              <div>
                <label class="block text-[11px] font-medium text-slate-400 mb-1">Total Due (₹)</label>
                <input type="number" id="total-due" value="0" readonly class="w-full bg-slate-800/50 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-rose-400 font-bold focus:outline-none">
              </div>
            </div>

            <div>
              <label class="block text-[11px] font-medium text-slate-400 mb-1">Notes / Special Requests</label>
              <textarea id="booking-notes" rows="2" class="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-slate-100 focus:outline-none focus:border-sky-500"></textarea>
            </div>

            <div class="flex justify-between items-center pt-2 border-t border-slate-800">
              <button type="button" id="delete-btn" onclick="deleteCurrentBooking()" class="bg-rose-600/20 hover:bg-rose-600 border border-rose-500/40 text-rose-300 hover:text-white text-xs font-semibold px-3 py-1.5 rounded-lg transition hidden">
                Delete
              </button>
              <div class="flex gap-2 ml-auto">
                <button type="button" onclick="closeBookingModal()" class="bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs px-3 py-1.5 rounded-lg transition">Cancel</button>
                <button type="submit" class="bg-sky-600 hover:bg-sky-500 text-white text-xs font-semibold px-4 py-1.5 rounded-lg transition shadow">Save Booking</button>
              </div>
            </div>
          </form>
        </div>
      </div>

    </div>

  <script>
    // Configuration & State
    const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx/exec"; // Replace with your Web App Deployment URL
    const LOCAL_STORAGE_KEY = "hotel_draft_bookings";

    const ROOMS = ["101", "102", "103", "104", "105", "201", "202", "203", "204", "205"];
    let bookings = [];
    let currentYear, currentMonth;

    // Initialization
    window.addEventListener('DOMContentLoaded', () => {
      const today = new Date();
      currentYear = today.getFullYear();
      currentMonth = today.getMonth();

      const monthPicker = document.getElementById('month-picker');
      monthPicker.value = `${currentYear}-${String(currentMonth + 1).padStart(2, '0')}`;
      monthPicker.addEventListener('change', (e) => {
        if (e.target.value) {
          const [y, m] = e.target.value.split('-');
          currentYear = parseInt(y);
          currentMonth = parseInt(m) - 1;
          renderGrid();
        }
      });

      document.getElementById('search-input').addEventListener('input', renderGrid);

      populateRoomSelect();
      loadInitialData();
    });

    function populateRoomSelect() {
      const select = document.getElementById('room-select');
      select.innerHTML = ROOMS.map(r => `<option value="${r}">Room ${r}</option>`).join('');
    }

    // Storage Management:
    // If a booking has a generated booking ID, it syncs directly to Google Sheet and is NOT stored in localStorage.
    function saveToLocalStorage() {
      const draftsOnly = bookings.filter(b => !b.bookingId || b.bookingId === "");
      localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(draftsOnly));
    }

    function loadInitialData() {
      // Load drafts from local storage first
      const localData = localStorage.getItem(LOCAL_STORAGE_KEY);
      if (localData) {
        try {
          bookings = JSON.parse(localData);
        } catch (e) {
          bookings = [];
        }
      }
      renderGrid();
      // Sync/Fetch existing bookings from Google Sheet
      fetchFromGoogleSheet();
    }

    function generateBookingId() {
      const prefix = "BK";
      const timestamp = Date.now().toString().slice(-5);
      const random = Math.floor(100 + Math.random() * 900);
      return `${prefix}-${timestamp}${random}`;
    }

    // Google Sheets Integration
    function updateSyncStatus(text, state) {
      const dot = document.getElementById('sync-dot');
      const textSpan = document.getElementById('sync-text');
      textSpan.textContent = text;

      dot.className = "w-2 h-2 rounded-full inline-block ";
      if (state === 'syncing') {
        dot.classList.add('bg-amber-400', 'animate-ping');
      } else if (state === 'success') {
        dot.classList.add('bg-emerald-400');
      } else if (state === 'error') {
        dot.classList.add('bg-rose-500');
      } else {
        dot.classList.add('bg-slate-500');
      }
    }

    function syncBookingToGoogleSheet(booking) {
      updateSyncStatus("Syncing to Sheet...", "syncing");
      
      fetch(SCRIPT_URL, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          action: "saveBooking",
          booking: booking
        })
      })
      .then(() => {
        updateSyncStatus("Synced to Google Sheet", "success");
      })
      .catch((err) => {
        console.error("Sync error:", err);
        updateSyncStatus("Sync error (Offline)", "error");
      });
    }

    function fetchFromGoogleSheet() {
      if (!SCRIPT_URL || SCRIPT_URL.includes("xxxxxxxx")) {
        updateSyncStatus("Sheet URL pending setup", "idle");
        return;
      }

      updateSyncStatus("Fetching from Sheet...", "syncing");
      fetch(`${SCRIPT_URL}?action=getBookings`)
        .then(res => res.json())
        .then(data => {
          if (Array.isArray(data)) {
            // Keep local drafts (without ID) and overwrite/append generated sheet bookings
            const drafts = bookings.filter(b => !b.bookingId);
            bookings = [...drafts, ...data];
            renderGrid();
            updateSyncStatus("Synced with Google Sheet", "success");
          }
        })
        .catch(err => {
          console.warn("Could not fetch remote sheet data:", err);
          updateSyncStatus("Sheet fetch failed", "error");
        });
    }

    // Form & Booking Operations
    function calcDue() {
      const total = parseFloat(document.getElementById('total-amount').value) || 0;
      const paid = parseFloat(document.getElementById('paid-amount').value) || 0;
      document.getElementById('total-due').value = Math.max(0, total - paid);
    }

    function openBookingModal(bookingId = null, defaultRoom = null, defaultDate = null) {
      const modal = document.getElementById('booking-modal');
      const form = document.getElementById('booking-form');
      form.reset();

      const deleteBtn = document.getElementById('delete-btn');
      let booking = null;

      if (bookingId) {
        booking = bookings.find(b => b.bookingId === bookingId);
      }

      if (booking) {
        document.getElementById('modal-title').textContent = "Edit Booking";
        document.getElementById('booking-id').value = booking.bookingId;
        document.getElementById('guest-name').value = booking.guestName || '';
        document.getElementById('guest-phone').value = booking.guestPhone || '';
        document.getElementById('room-select').value = booking.room || ROOMS[0];
        document.getElementById('checkin-date').value = booking.checkIn || '';
        document.getElementById('checkout-date').value = booking.checkOut || '';
        document.getElementById('booking-status').value = booking.status || 'Confirmed';
        document.getElementById('total-amount').value = booking.totalAmount || 0;
        document.getElementById('paid-amount').value = booking.paidAmount || 0;
        document.getElementById('total-due').value = booking.totalDue || 0;
        document.getElementById('booking-notes').value = booking.notes || '';
        deleteBtn.classList.remove('hidden');
      } else {
        document.getElementById('modal-title').textContent = "New Booking";
        document.getElementById('booking-id').value = generateBookingId(); // Generate ID immediately
        if (defaultRoom) document.getElementById('room-select').value = defaultRoom;
        if (defaultDate) {
          document.getElementById('checkin-date').value = defaultDate;
          const nextDay = new Date(defaultDate);
          nextDay.setDate(nextDay.getDate() + 1);
          document.getElementById('checkout-date').value = nextDay.toISOString().split('T')[0];
        }
        deleteBtn.classList.add('hidden');
      }

      calcDue();
      modal.classList.remove('hidden');
    }

    function closeBookingModal() {
      document.getElementById('booking-modal').classList.add('hidden');
    }

    function handleFormSubmit(e) {
      e.preventDefault();

      const bookingId = document.getElementById('booking-id').value || generateBookingId();
      const updatedBooking = {
        bookingId: bookingId,
        guestName: document.getElementById('guest-name').value,
        guestPhone: document.getElementById('guest-phone').value,
        room: document.getElementById('room-select').value,
        checkIn: document.getElementById('checkin-date').value,
        checkOut: document.getElementById('checkout-date').value,
        status: document.getElementById('booking-status').value,
        totalAmount: parseFloat(document.getElementById('total-amount').value) || 0,
        paidAmount: parseFloat(document.getElementById('paid-amount').value) || 0,
        totalDue: parseFloat(document.getElementById('total-due').value) || 0,
        notes: document.getElementById('booking-notes').value,
        updatedAt: new Date().toISOString()
      };

      const existingIndex = bookings.findIndex(b => b.bookingId === bookingId);
      if (existingIndex !== -1) {
        bookings[existingIndex] = updatedBooking;
      } else {
        bookings.push(updatedBooking);
      }

      // Requirement: If user edits a cell/booking after booking ID is generated,
      // it syncs with Google Sheet, NOT local storage.
      syncBookingToGoogleSheet(updatedBooking);
      saveToLocalStorage(); // Local storage excludes items with bookingId

      closeBookingModal();
      renderGrid();
    }

    function deleteCurrentBooking() {
      const bookingId = document.getElementById('booking-id').value;
      if (!bookingId) return;

      if (confirm("Are you sure you want to delete this booking?")) {
        bookings = bookings.filter(b => b.bookingId !== bookingId);

        // Notify Google Sheet of deletion if applicable
        fetch(SCRIPT_URL, {
          method: "POST",
          mode: "no-cors",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ action: "deleteBooking", bookingId: bookingId })
        }).catch(err => console.error("Deletion sync error:", err));

        saveToLocalStorage();
        closeBookingModal();
        renderGrid();
      }
    }

    function wipeData() {
      if (confirm("WARNING: This will wipe out all local and synced booking data. Continue?")) {
        localStorage.removeItem(LOCAL_STORAGE_KEY);
        bookings = [];

        fetch(SCRIPT_URL, {
          method: "POST",
          mode: "no-cors",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ action: "wipeAll" })
        }).catch(err => console.error("Wipe sync error:", err));

        updateSyncStatus("Data wiped", "idle");
        renderGrid();
      }
    }

    // Grid Rendering & Visual Matrix
    function renderGrid() {
      const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
      const headRow = document.getElementById('matrix-head');
      const bodyRows = document.getElementById('matrix-body');
      const searchQuery = document.getElementById('search-input').value.toLowerCase().trim();

      // Render Header Row
      let headHTML = `<tr class="border-b border-slate-800">
        <th class="p-3 w-24 sticky left-0 bg-slate-800 text-slate-300 font-bold z-30 border-r border-slate-700">Room</th>`;
      
      for (let day = 1; day <= daysInMonth; day++) {
        const dateObj = new Date(currentYear, currentMonth, day);
        const dayOfWeek = dateObj.toLocaleDateString('en-US', { weekday: 'short' });
        const isWeekend = dateObj.getDay() === 0 || dateObj.getDay() === 6;

        headHTML += `
          <th class="p-2 text-center border-r border-slate-800 min-w-[50px] ${isWeekend ? 'bg-slate-800/80 text-sky-400' : ''}">
            <div class="text-[10px] uppercase font-bold text-slate-400">${dayOfWeek}</div>
            <div class="text-xs font-extrabold">${day}</div>
          </th>`;
      }
      headHTML += `</tr>`;
      headRow.innerHTML = headHTML;

      // Filter bookings if search query is present
      const filteredBookings = bookings.filter(b => {
        if (!searchQuery) return true;
        return (
          (b.guestName && b.guestName.toLowerCase().includes(searchQuery)) ||
          (b.bookingId && b.bookingId.toLowerCase().includes(searchQuery)) ||
          (b.room && b.room.toLowerCase().includes(searchQuery))
        );
      });

      // Render Body Rows (Rooms)
      let bodyHTML = '';
      ROOMS.forEach(room => {
        bodyHTML += `<tr class="hover:bg-slate-900/50 transition">
          <td class="p-3 font-bold text-sky-400 sticky left-0 bg-slate-900 z-10 border-r border-slate-800 border-b shadow-sm">
            Room ${room}
          </td>`;

        for (let day = 1; day <= daysInMonth; day++) {
          const dateStr = `${currentYear}-${String(currentMonth + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
          
          // Find active booking for this room & date
          const cellBookings = filteredBookings.filter(b => {
            return b.room === room && dateStr >= b.checkIn && dateStr < b.checkOut;
          });

          if (cellBookings.length > 0) {
            const b = cellBookings[0];
            const statusColor = getStatusColorClass(b.status);

            bodyHTML += `
              <td class="excel-cell p-1 border-r border-b border-slate-800 cursor-pointer relative ${statusColor}"
                  onclick="openBookingModal('${b.bookingId}')"
                  onmouseenter="showCommentBox(event, '${b.bookingId}')"
                  onmouseleave="closeCommentBox()">
                <div class="truncate text-[10px] font-bold text-center px-1">
                  ${b.guestName || b.bookingId}
                </div>
              </td>`;
          } else {
            bodyHTML += `
              <td class="excel-cell p-1 border-r border-b border-slate-800 text-center hover:bg-slate-800/60 cursor-pointer text-slate-600"
                  onclick="openBookingModal(null, '${room}', '${dateStr}')">
              </td>`;
          }
        }

        bodyHTML += `</tr>`;
      });

      bodyRows.innerHTML = bodyHTML;
    }

    function getStatusColorClass(status) {
      switch (status) {
        case 'Confirmed': return 'bg-emerald-950/80 border-emerald-600 text-emerald-200';
        case 'Checked-In': return 'bg-sky-950/80 border-sky-600 text-sky-200';
        case 'Checked-Out': return 'bg-slate-800/80 border-slate-600 text-slate-300';
        case 'Blocked': return 'bg-amber-950/80 border-amber-600 text-amber-200';
        case 'Cancelled': return 'bg-rose-950/80 border-rose-600 text-rose-200';
        default: return 'bg-slate-800 text-slate-200';
      }
    }

    // Excel Hover Comment Box logic
    function showCommentBox(e, bookingId) {
      const b = bookings.find(item => item.bookingId === bookingId);
      if (!b) return;

      const box = document.getElementById('excel-comment-box');
      document.getElementById('comment-booking-id').textContent = b.bookingId;
      document.getElementById('comment-room').textContent = `Room ${b.room}`;

      const listContainer = document.getElementById('comment-details-list');
      const statusColorClass = getStatusColorClass(b.status);

      listContainer.innerHTML = `
        <div class="font-semibold text-slate-200">${b.guestName || 'N/A'}</div>
        <div class="text-slate-400 text-[10px]"><i class="fa-solid fa-phone text-slate-500 mr-1"></i>${b.guestPhone || 'No Phone'}</div>
        <div class="text-slate-400 text-[10px]"><i class="fa-regular fa-calendar mr-1"></i>${b.checkIn} →${b.checkOut}</div>
        <div class="pt-1">
