(() => {
  "use strict";

  const CFG = window.LEAVE_TRACKER_CONFIG;
  const SHEETS_BASE = "https://sheets.googleapis.com/v4/spreadsheets";
  const SCOPE = "https://www.googleapis.com/auth/spreadsheets";

  const OL = CFG.SHEET_OTHER_LEAVES; // "Other Leaves"
  const CO = CFG.SHEET_COMP_OFF;     // "Comp Off"

  let accessToken = null;
  let tokenClient = null;

  // ── DOM refs ─────────────────────────────────────────────────────
  const $ = (id) => document.getElementById(id);

  const screenLogin = $("screen-login");
  const appEl = $("app");
  const btnSignin = $("btn-signin");
  const btnSignout = $("btn-signout");
  const loginError = $("login-error");
  const toastEl = $("toast");

  const pageOverview = $("page-overview");
  const pageAdd = $("page-add");
  const tabOverview = $("tab-overview");
  const tabAdd = $("tab-add");
  const topbarTitle = $("topbar-title");

  const segLeave = $("seg-leave");
  const segExtra = $("seg-extra");
  const formLeave = $("form-leave");
  const formExtra = $("form-extra");
  const leaveType = $("leave-type");
  const fieldCompDaytype = $("field-comp-daytype");
  const leaveCompDaytype = $("leave-comp-daytype");
  const addError = $("add-error");

  leaveType.addEventListener("change", () => {
    fieldCompDaytype.hidden = leaveType.value !== "Comp. Off";
  });

  // ── Toast ────────────────────────────────────────────────────────
  let toastTimer = null;
  function showToast(message, isError = false) {
    clearTimeout(toastTimer);
    toastEl.textContent = message;
    toastEl.classList.toggle("is-error", isError);
    toastEl.hidden = false;
    requestAnimationFrame(() => toastEl.classList.add("is-visible"));
    toastTimer = setTimeout(() => {
      toastEl.classList.remove("is-visible");
      setTimeout(() => { toastEl.hidden = true; }, 300);
    }, 2600);
  }

  // ── Auth ─────────────────────────────────────────────────────────
  function initAuth() {
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CFG.CLIENT_ID,
      scope: SCOPE,
      callback: (resp) => {
        if (resp.error) {
          showLoginError("Sign-in failed. Please try again.");
          return;
        }
        accessToken = resp.access_token;
        enterApp();
      },
    });
  }

  function showLoginError(msg) {
    loginError.textContent = msg;
    loginError.hidden = false;
  }

  btnSignin.addEventListener("click", () => {
    loginError.hidden = true;
    tokenClient.requestAccessToken({ prompt: "consent" });
  });

  btnSignout.addEventListener("click", () => {
    if (accessToken) {
      google.accounts.oauth2.revoke(accessToken, () => {});
    }
    accessToken = null;
    appEl.hidden = true;
    screenLogin.hidden = false;
  });

  function enterApp() {
    screenLogin.hidden = true;
    appEl.hidden = false;
    goToOverview();
    loadOverview();
  }

  // ── Sheets API helpers ───────────────────────────────────────────
  async function sheetsFetch(path, options = {}) {
    const res = await fetch(`${SHEETS_BASE}/${CFG.SPREADSHEET_ID}${path}`, {
      ...options,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
        ...(options.headers || {}),
      },
    });
    if (!res.ok) {
      const body = await res.text();
      throw new Error(`Sheets API error (${res.status}): ${body.slice(0, 200)}`);
    }
    return res.json();
  }

  function batchGet(ranges) {
    const q = ranges.map((r) => `ranges=${encodeURIComponent(r)}`).join("&");
    return sheetsFetch(`/values:batchGet?${q}`);
  }

  function appendValues(range, rows) {
    return sheetsFetch(
      `/values/${encodeURIComponent(range)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
      { method: "POST", body: JSON.stringify({ values: rows }) }
    );
  }

  function updateValues(range, rows) {
    return sheetsFetch(
      `/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,
      { method: "PUT", body: JSON.stringify({ values: rows }) }
    );
  }

  function getValues(range) {
    return sheetsFetch(`/values/${encodeURIComponent(range)}`);
  }

  // ── Tab / page navigation ────────────────────────────────────────
  function goToOverview() {
    pageOverview.hidden = false;
    pageAdd.hidden = true;
    tabOverview.classList.add("is-active");
    tabAdd.classList.remove("is-active");
    topbarTitle.textContent = "Overview";
  }

  function goToAdd() {
    pageOverview.hidden = true;
    pageAdd.hidden = false;
    tabAdd.classList.add("is-active");
    tabOverview.classList.remove("is-active");
    topbarTitle.textContent = "Add entry";
  }

  tabOverview.addEventListener("click", () => { goToOverview(); loadOverview(); });
  tabAdd.addEventListener("click", goToAdd);

  segLeave.addEventListener("click", () => {
    segLeave.classList.add("is-active");
    segExtra.classList.remove("is-active");
    formLeave.hidden = false;
    formExtra.hidden = true;
  });

  segExtra.addEventListener("click", () => {
    segExtra.classList.add("is-active");
    segLeave.classList.remove("is-active");
    formExtra.hidden = false;
    formLeave.hidden = true;
  });

  // ── Overview rendering ───────────────────────────────────────────
  const BAR_TYPES = ["Sick Leave", "Casual Leave", "Priviledge Leave"];
  const BAR_TYPES_DISPLAY = { "Priviledge Leave": "Privilege Leave" };

  async function loadOverview() {
    const overviewError = $("overview-error");
    overviewError.hidden = true;

    try {
      const data = await batchGet([
        `'${OL}'!E2:H6`,
        `'${CO}'!B1:D1`,
        `'${CO}'!F3:F2000`,
      ]);

      const [olRange, coHeaderRange, coExpiryRange] = data.valueRanges;
      renderBalance(olRange.values || []);
      renderCompOff(coHeaderRange.values || [], olRange.values || []);
      renderExpiring(coExpiryRange.values || []);
    } catch (err) {
      console.error(err);
      overviewError.textContent = "Couldn't load your sheet. Check your connection and try again.";
      overviewError.hidden = false;
    }
  }

  function renderBalance(rows) {
    // rows correspond to Other Leaves!E2:H6, i.e. index 0 = Sick, 1 = Casual, 2 = Privilege, 3 = Comp.Off, 4 = Total
    const container = $("bars");
    container.innerHTML = "";

    BAR_TYPES.forEach((type, i) => {
      const row = rows[i] || [];
      const alloted = Number(row[1] ?? 0);
      const consumed = Number(row[2] ?? 0);
      const pct = alloted > 0 ? Math.min(100, (consumed / alloted) * 100) : 0;

      const wrap = document.createElement("div");
      wrap.className = "bar-row";
      wrap.innerHTML = `
        <span class="bar-row-label">${BAR_TYPES_DISPLAY[type] || type}</span>
        <span class="bar-track"><span class="bar-fill" style="width:${pct}%"></span></span>
        <span class="bar-row-value">${consumed}/${alloted}</span>
      `;
      container.appendChild(wrap);
    });
  }

  function renderCompOff(headerRow, olRows) {
    const row = headerRow[0] || [];
    const worked = row[0] ?? "0";
    const taken = row[2] ?? "0";
    // Other Leaves!H5 = Comp. Off available (index 3 within E2:H6 range, column H is index 3)
    const compRow = olRows[3] || [];
    const available = compRow[3] ?? "0";

    $("comp-worked").textContent = worked;
    $("comp-taken").textContent = taken;
    $("comp-available").textContent = available;
  }

  function renderExpiring(rows) {
    const count = rows.reduce((n, r) => n + ((r[0] || "").trim() === "EXPIRING SOON" ? 1 : 0), 0);
    const el = $("expiring-count");
    el.textContent = count;
    el.classList.toggle("is-warm", count > 0);
    $("expiring-label").textContent = count === 1 ? "comp-off expiring soon" : "comp-offs expiring soon";
  }

  // ── Add-entry: leave taken ───────────────────────────────────────
  formLeave.addEventListener("submit", async (e) => {
    e.preventDefault();
    addError.hidden = true;
    const submitBtn = formLeave.querySelector("button[type=submit]");
    const date = $("leave-date").value;
    const type = leaveType.value;

    if (!date) return;

    submitBtn.disabled = true;
    try {
      if (type === "Comp. Off") {
        await redeemCompOff(date, leaveCompDaytype.value);
      } else {
        await appendValues(`'${OL}'!A:B`, [[date, type]]);
      }
      formLeave.reset();
      fieldCompDaytype.hidden = true;
      showToast("Entry sent to your sheet.");
      loadOverview();
    } catch (err) {
      console.error(err);
      addError.textContent = err.message || "Couldn't save. Please try again.";
      addError.hidden = false;
      showToast("Couldn't save entry.", true);
    } finally {
      submitBtn.disabled = false;
    }
  });

  async function redeemCompOff(dateStr, dayType) {
    // Find the oldest comp-off row that has an "extra day worked" but hasn't been redeemed yet.
    const data = await getValues(`'${CO}'!A3:D2000`);
    const rows = data.values || [];
    let targetIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      const a = rows[i][0];
      const c = rows[i][2];
      if (a && !c) { targetIndex = i; break; }
    }
    if (targetIndex === -1) {
      throw new Error("No comp-off available to redeem right now.");
    }
    const sheetRow = targetIndex + 3;
    await updateValues(`'${CO}'!C${sheetRow}:D${sheetRow}`, [[dateStr, dayType]]);
  }

  // ── Add-entry: extra day worked ──────────────────────────────────
  formExtra.addEventListener("submit", async (e) => {
    e.preventDefault();
    addError.hidden = true;
    const submitBtn = formExtra.querySelector("button[type=submit]");
    const date = $("extra-date").value;
    const dayType = $("extra-daytype").value;

    if (!date) return;

    submitBtn.disabled = true;
    try {
      await appendValues(`'${CO}'!A:B`, [[date, dayType]]);
      formExtra.reset();
      showToast("Entry sent to your sheet.");
      loadOverview();
    } catch (err) {
      console.error(err);
      addError.textContent = err.message || "Couldn't save. Please try again.";
      addError.hidden = false;
      showToast("Couldn't save entry.", true);
    } finally {
      submitBtn.disabled = false;
    }
  });

  // ── Init ─────────────────────────────────────────────────────────
  window.addEventListener("load", () => {
    if (window.google && google.accounts && google.accounts.oauth2) {
      initAuth();
    } else {
      // GIS script can arrive slightly after load in rare cases; poll briefly.
      const t = setInterval(() => {
        if (window.google && google.accounts && google.accounts.oauth2) {
          clearInterval(t);
          initAuth();
        }
      }, 100);
    }

    if ("serviceWorker" in navigator) {
      navigator.serviceWorker.register("sw.js").catch(() => {});
    }
  });
})();
