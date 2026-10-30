/* Tema prikaza (sustav / svijetla / tamna) – isto ponašanje kao na curling.hr.
   Učitava se u <head>, prije iscrtavanja stranice, da nema bljeska krive teme.
   Izbor se pamti u localStorage pod istim ključem kao na javnoj stranici ("hcs-theme"). */
(function () {
  "use strict";
  const KEY = "hcs-theme";
  const THEMES = ["auto", "light", "dark"];
  const mq = window.matchMedia("(prefers-color-scheme: dark)");
  function get() { let t = "auto"; try { t = localStorage.getItem(KEY) || "auto"; } catch (e) {} return THEMES.includes(t) ? t : "auto"; }
  function apply(t) {
    document.documentElement.dataset.theme = t === "auto" ? (mq.matches ? "dark" : "light") : t;
    try { localStorage.setItem(KEY, t); } catch (e) {}
  }
  apply(get());
  mq.addEventListener("change", () => { if (get() === "auto") apply("auto"); });
  window.HcsTheme = {
    get, apply,
    next: (t) => THEMES[(THEMES.indexOf(t) + 1) % THEMES.length]
  };
})();
