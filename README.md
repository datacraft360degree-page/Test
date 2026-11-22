<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Business Portal - Maintenance</title>
  <style>
    * { box-sizing: border-box; }
    body {
      margin: 0; min-height: 100vh; display: grid; place-items: center;
      padding: 20px; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: #f2f4f7; color: #1e293b;
    }
    .panel {
      width: min(100%, 390px); background: #fff; border: 1px solid #e2e8f0;
      border-radius: 24px; padding: 28px; box-shadow: 0 18px 55px #0f172a18;
    }
    .brand {
      width: 52px; height: 52px; display: grid; place-items: center;
      margin: 0 auto 14px; border-radius: 17px; background: #eff6ff;
      color: #2563eb; font-size: 23px;
    }
    h1 { margin: 0; text-align: center; font-size: 20px; }
    .sub { text-align: center; color: #64748b; font-size: 13px; margin: 8px 0 24px; }
    label { display: block; margin: 14px 0 6px; font-size: 12px; font-weight: 650; }
    input {
      width: 100%; border: 1px solid #cbd5e1; border-radius: 12px;
      padding: 12px 13px; font-size: 14px; outline: none;
    }
    input:focus { border-color: #3b82f6; box-shadow: 0 0 0 3px #3b82f622; }
    button {
      width: 100%; border: 0; border-radius: 12px; padding: 12px 14px;
      font-size: 14px; font-weight: 700; cursor: pointer; margin-top: 18px;
    }
    .primary { background: #2563eb; color: white; }
    .logout { background: #fff1f2; color: #be123c; border: 1px solid #fecdd3; }
    .error {
      display: none; color: #be123c; background: #fff1f2; border-radius: 10px;
      padding: 10px; margin-top: 12px; text-align: center; font-size: 12px;
    }
    #portal { display: none; width: min(100%, 520px); }
    .portal-head {
      display: flex; justify-content: space-between; align-items: center; gap: 12px;
      background: white; border: 1px solid #e2e8f0; border-radius: 18px;
      padding: 12px 14px; box-shadow: 0 8px 25px #0f172a0d;
    }
    .portal-head strong { font-size: 13px; }
    .portal-head button { width: auto; margin: 0; padding: 9px 14px; }
    .notice {
      margin-top: 18px; padding: 25px; text-align: center; background: white;
      border: 1px solid #dbeafe; border-radius: 22px; box-shadow: 0 18px 55px #0f172a12;
    }
    .notice .symbol { font-size: 34px; margin-bottom: 12px; }
    .notice p { font-size: 15px; line-height: 1.65; margin: 0; font-weight: 600; }
    .notice button { margin-top: 22px; }
    [hidden] { display: none !important; }
  </style>
</head>
<body>
  <section class="panel" id="loginPanel" aria-labelledby="loginTitle">
    <div class="brand" aria-hidden="true">🔒</div>
    <h1 id="loginTitle">Business Portal</h1>
    <p class="sub">Please sign in with your existing portal credentials.</p>
    <form id="loginForm">
      <label for="userId">User ID</label>
      <input id="userId" name="userId" autocomplete="username" required>
      <label for="password">Password</label>
      <input id="password" name="password" type="password" autocomplete="current-password" required>
      <div class="error" id="loginError" role="alert">Invalid User ID or Password!</div>
      <button class="primary" type="submit">Login</button>
    </form>
  </section>

  <main id="portal" aria-live="polite">
    <header class="portal-head">
      <strong>Business Portal</strong>
    </header>
    <section class="notice">
      <div class="symbol" aria-hidden="true">🛠️</div>
      <p>Portal is updating, you saved data is save &amp; secure please visit later.</p>
      <button class="logout" type="button" id="noticeLogout">Logout</button>
    </section>
  </main>

  <script>
    // Matches the User ID and password configured in the supplied original HTML.
    // This is a client-side gate only; credentials embedded in an HTML file are
    // visible to anyone who can inspect the source.
    const DEFAULT_USER_ID = "Admin";
    const DEFAULT_PASSWORD = "Aadmin123";
    const SESSION_KEY = "portal_maintenance_authenticated";

    const loginPanel = document.getElementById("loginPanel");
    const portal = document.getElementById("portal");
    const loginError = document.getElementById("loginError");

    function showPortal() {
      loginPanel.hidden = true;
      portal.style.display = "block";
    }

    function showLogin() {
      portal.style.display = "none";
      loginPanel.hidden = false;
      document.getElementById("loginForm").reset();
      sessionStorage.removeItem(SESSION_KEY);
    }

    document.getElementById("loginForm").addEventListener("submit", function (event) {
      event.preventDefault();
      const user = document.getElementById("userId").value.trim();
      const pass = document.getElementById("password").value.trim();

      if (user === DEFAULT_USER_ID && pass === DEFAULT_PASSWORD) {
        sessionStorage.setItem(SESSION_KEY, "true");
        loginError.style.display = "none";
        showPortal();
      } else {
        loginError.style.display = "block";
      }
    });

    document.getElementById("headerLogout").addEventListener("click", showLogin);
    document.getElementById("noticeLogout").addEventListener("click", showLogin);

    if (sessionStorage.getItem(SESSION_KEY) === "true") showPortal();
  </script>
</body>
</html>
