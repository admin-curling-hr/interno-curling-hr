/* Zajednička ljuska internog dijela Saveza (interno.curling.hr)
   – header i izbornik kao na curling.hr
   – prijava preko Microsoft Entra ID (MSAL), Graph pomoćne funkcije
   Svaka stranica učitava redom: config.js, MSAL, assets/shell.js, pa svoj kod,
   i čeka Shell.ready prije nego što počne raditi s podacima. */
(function () {
  "use strict";

  /* Korijen web mjesta izračunava se iz lokacije ove datoteke (assets/shell.js),
     pa isti kod radi na interno.curling.hr, na github.io adresi i na localhostu. */
  const ROOT = new URL("../", document.currentScript.src).href;
  const CFG = window.CONFIG || {};
  const SCOPES = ["User.Read", "Sites.ReadWrite.All"];
  const GRAPH = "https://graph.microsoft.com/v1.0";

  /* ===== Izbornik =====
     Jedna razina, kao na curling.hr. "href" je relativan na korijen internog dijela,
     osim ako počinje s https:// (vanjska poveznica). Opcionalno "role": stavka se
     prikazuje samo korisnicima s tom ulogom aplikacije (Entra app roles); dok uloge
     nisu definirane, prikazuju se sve stavke. Podatke ionako štite dozvole na SharePointu.
     "desc" se prikazuje na pločici na početnoj stranici. */
  const MENU = [
    { label: "Početna", href: "" },
    { label: "Registar igrača", href: "registri/#igraci", desc: "Registrirani igrači, registracije, prestanci i zdravstvena uvjerenja." },
    { label: "Registar klubova", href: "registri/#klubovi", desc: "Klubovi članovi Saveza i njihovi podaci." },
    { label: "Registar ekipa", href: "registri/#ekipe", desc: "Ekipe klubova po kategorijama i disciplinama." },
    { label: "Uređivanje vijesti", href: "https://curling.hr/edit/", desc: "Unos i uređivanje vijesti na javnoj stranici Saveza." }
  ];

  const ICON_HOME = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 10.5 12 3l9 7.5"/><path d="M5 9.5V21h5v-6h4v6h5V9.5"/></svg>';
  const ICON_LOGOUT = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><path d="M16 17l5-5-5-5"/><path d="M21 12H9"/></svg>';
  const ICON_SUN = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M6.34 17.66l-1.41 1.41M19.07 4.93l-1.41 1.41"/></svg>';
  const ICON_MOON = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/></svg>';
  const ICON_MONITOR = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="2" y="3" width="20" height="14" rx="2"/><path d="M8 21h8M12 17v4"/></svg>';
  /* Kao na curling.hr: ikona pokazuje temu koja će se uključiti KLIKOM */
  const THEME_UI = {
    auto:  { next: ICON_SUN,     title: "Klikni za SVIJETLI prikaz boja." },
    light: { next: ICON_MOON,    title: "Klikni za TAMNI prikaz boja." },
    dark:  { next: ICON_MONITOR, title: "Klikni za prikaz boja prema postavkama SUSTAVA." }
  };
  const ICON_USER = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="8" r="4"/><path d="M4 21c0-4.4 3.6-7 8-7s8 2.6 8 7"/></svg>';

  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const $ = (s) => document.querySelector(s);
  const moduleId = document.body.dataset.module || "";

  let msalApp = null, account = null, roles = null, resolveReady;
  const ready = new Promise((r) => { resolveReady = r; });

  /* ===== Header ===== */
  function renderHeader() {
    const h = document.createElement("header");
    h.className = "site-header";
    h.innerHTML =
      '<div class="header-inner">' +
        '<a class="header-brand" href="' + esc(ROOT) + '">' +
          '<img src="https://curling.hr/assets/images/hcs-logo.png" alt="HCS Logo" class="header-logo">' +
          '<span class="header-titles"><span class="header-title">Hrvatski curling savez</span><span class="header-subtitle">Interne stranice</span></span>' +
        "</a>" +
        '<div class="header-controls">' +
          '<a href="https://curling.hr/" class="hdr-btn" title="Javna stranica Saveza">' + ICON_HOME + '<span class="sr-only">Javna stranica Saveza</span></a>' +
          '<button id="shellTheme" class="hdr-btn"></button>' +
          '<button id="shellLogin" class="hdr-btn hidden" title="Prijava">' + ICON_USER + '<span class="sr-only">Prijava</span></button>' +
          '<button id="shellLogout" class="hdr-btn hidden" title="Odjava">' + ICON_LOGOUT + '<span class="sr-only">Odjava</span></button>' +
          '<button id="shellNavToggle" class="nav-toggle hidden" aria-label="Izbornik" aria-expanded="false" aria-controls="shellNav"><span></span><span></span><span></span></button>' +
        "</div>" +
      "</div>" +
      '<nav class="site-nav hidden" id="shellNav" aria-label="Glavni izbornik"><ul></ul></nav>';
    document.body.prepend(h);

    const main = document.querySelector("main");
    const sign = document.createElement("section");
    sign.id = "shellSignin";
    sign.className = "card signin hidden";
    sign.innerHTML = "<h2>Prijava</h2><p>Interni dio web mjesta Hrvatskog curling saveza dostupan je samo ovlaštenim osobama. " +
      "Prijavite se računom Saveza – prijava se odvija na Microsoftovoj stranici, uz drugi faktor ako ga imate postavljenog.</p>" +
      '<p id="shellErr" class="formerr hidden" role="alert"></p>' +
      '<button id="shellLogin2" class="btn btn-primary">Prijavi se</button>';
    main.prepend(sign);

    const toast = document.createElement("div");
    toast.id = "shellToast"; toast.className = "toast hidden"; toast.setAttribute("role", "status");
    document.body.append(toast);

    $("#shellLogin").addEventListener("click", login);
    $("#shellLogin2").addEventListener("click", login);
    $("#shellLogout").addEventListener("click", () => msalApp && msalApp.logoutRedirect({ account, postLogoutRedirectUri: ROOT }));

    const th = $("#shellTheme");
    const paintTheme = () => {
      const ui = THEME_UI[window.HcsTheme ? HcsTheme.get() : "auto"];
      th.innerHTML = ui.next; th.title = ui.title; th.setAttribute("aria-label", ui.title);
    };
    th.addEventListener("click", () => { if (window.HcsTheme) { HcsTheme.apply(HcsTheme.next(HcsTheme.get())); paintTheme(); } });
    paintTheme();

    const nav = $("#shellNav"), tgl = $("#shellNavToggle");
    tgl.addEventListener("click", () => { const open = nav.classList.toggle("open"); tgl.setAttribute("aria-expanded", String(open)); });
    nav.addEventListener("click", (e) => { if (e.target.closest("a")) { nav.classList.remove("open"); tgl.setAttribute("aria-expanded", "false"); } });
  }

  const itemUrl = (m) => /^https?:\/\//.test(m.href) ? m.href : ROOT + m.href;
  function visibleMenu() {
    return MENU.filter((m) => !m.role || !roles || roles.includes(m.role));
  }
  /* Aktivna stavka: ista stranica i isti hash; bez hasha vrijedi prva stavka te stranice */
  function markActive() {
    const here = location.origin + location.pathname;
    const links = [...document.querySelectorAll("#shellNav a")];
    const samePage = links.filter((a) => { const u = new URL(a.href); return u.origin + u.pathname === here; });
    let active = samePage.find((a) => new URL(a.href).hash === location.hash) || (location.hash ? null : samePage[0]) || null;
    if (!active && samePage.length && location.hash) active = samePage.find((a) => !new URL(a.href).hash) || null;
    links.forEach((a) => {
      const on = a === active;
      a.classList.toggle("active", on);
      if (on) a.setAttribute("aria-current", "page"); else a.removeAttribute("aria-current");
    });
  }
  function renderNav() {
    const nav = $("#shellNav");
    nav.querySelector("ul").innerHTML = visibleMenu().map((m) => '<li><a href="' + esc(itemUrl(m)) + '">' + esc(m.label) + "</a></li>").join("");
    nav.classList.remove("hidden");
    $("#shellNavToggle").classList.remove("hidden");
    markActive();
    window.addEventListener("hashchange", markActive);
  }

  function fail(msg) {
    $("#shellSignin").classList.remove("hidden");
    const e = $("#shellErr"); e.textContent = msg; e.classList.remove("hidden");
  }
  function signedOut() {
    $("#shellSignin").classList.remove("hidden");
    $("#shellLogin").classList.remove("hidden");
  }
  function signedIn() {
    const who = [account.name, account.username].filter(Boolean).join(", ");
    const lo = $("#shellLogout"); lo.classList.remove("hidden");
    lo.title = "Prijavljen: " + who + "\nKlikni za odjavu.";
    lo.querySelector(".sr-only").textContent = "Odjava (" + who + ")";
    roles = (account.idTokenClaims && Array.isArray(account.idTokenClaims.roles)) ? account.idTokenClaims.roles : null;
    renderNav();
    const mod = document.getElementById("module");
    if (mod) mod.hidden = false;
    resolveReady({ account, roles });
  }
  function login() { if (msalApp) msalApp.loginRedirect({ scopes: SCOPES }); }

  /* ===== Prijava ===== */
  async function init() {
    renderHeader();
    if (typeof msal === "undefined") { fail("Knjižnica za prijavu (MSAL) se nije učitala. Provjerite internetsku vezu ili blokator sadržaja."); return; }
    if (!CFG.clientId || String(CFG.clientId).startsWith("UPISATI") || String(CFG.tenantId).startsWith("UPISATI")) {
      fail("Datoteka config.js nedostaje ili u njoj nisu upisani tenantId i clientId."); return;
    }
    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: CFG.clientId,
        authority: "https://login.microsoftonline.com/" + CFG.tenantId,
        redirectUri: ROOT,                 // jedna registrirana adresa za sve stranice
        navigateToLoginRequestUrl: true    // nakon prijave povratak na stranicu s koje je pokrenuta
      },
      cache: { cacheLocation: "sessionStorage" }
    });
    await msalApp.initialize();
    try {
      const resp = await msalApp.handleRedirectPromise();
      account = resp ? resp.account : (msalApp.getAllAccounts()[0] || null);
    } catch (e) { fail("Prijava nije uspjela:\n" + e.message); return; }
    if (!account) { signedOut(); return; }
    signedIn();
  }

  async function getToken() {
    try { return (await msalApp.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
    catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) { await msalApp.acquireTokenRedirect({ scopes: SCOPES, account }); return null; }
      throw e;
    }
  }

  /* ===== Graph ===== */
  async function graph(path, opts = {}) {
    const token = await getToken();
    if (!token) throw new Error("Potrebna je ponovna prijava.");
    const url = path.startsWith("https://") ? path : GRAPH + path;
    const headers = { Authorization: "Bearer " + token };
    if (opts.body !== undefined) headers["Content-Type"] = "application/json";
    const res = await fetch(url, { method: opts.method || "GET", headers, body: opts.body !== undefined ? JSON.stringify(opts.body) : undefined });
    if (res.status === 204) return null;
    if (!res.ok) {
      let detail = "";
      try { const j = await res.json(); detail = j.error && j.error.message ? j.error.message : ""; } catch (_) {}
      if (res.status === 403) detail = "Nemate dozvolu za ovu radnju na SharePoint listi." + (detail ? "\n" + detail : "");
      throw new Error(detail || ("Greška " + res.status));
    }
    return res.json();
  }

  let toastTimer = null;
  function toast(msg) {
    const t = $("#shellToast"); t.textContent = msg; t.classList.remove("hidden");
    clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.add("hidden"), 2600);
  }

  window.Shell = { ready, graph, toast, esc, root: ROOT, menu: () => visibleMenu().filter((m) => m.href).map((m) => Object.assign({ url: itemUrl(m) }, m)) };
  init();
})();
