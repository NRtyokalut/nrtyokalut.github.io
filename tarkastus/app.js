/*! © 2026 Lämpöpumppu Mafia. Kaikki oikeudet pidätetään. All rights reserved.
 * Tämän ohjelmiston kopioiminen, muokkaaminen tai jakaminen ilman lupaa on kielletty. */
/* global ShiftCalc, XLSX */
/**
 * Vuorotaulun tarkastus — the pre-publish checker. Planned shifts only: no toteuma, no lisävuorot,
 * no euros, no junat, no palkkiot. All storage under "tarkastus.*" (separate from the päiväkirja).
 */
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const views = { home: $("view-home"), overview: $("view-overview"), detail: $("view-detail"), vertailu: $("view-vertailu") };
  const NS = "tarkastus.";
  const OT_STORE_PREFIX = NS + "otOverride.";
  const LASK_PREFIX = NS + "laskenta.";
  const MARK_PREFIX = NS + "marks.";
  const STEWARD_KEY = NS + "lm";
  const HISTORY_KEY = NS + "history";
  const SETTINGS_KEY = NS + "settings";
  const NONE = {}; // no lisävuorot / toteumat in this app

  let state = null;
  let prevView = "overview";
  const WD_SHORT = { Maanantai: "Ma", Tiistai: "Ti", Keskiviikko: "Ke", Torstai: "To", Perjantai: "Pe", Lauantai: "La", Sunnuntai: "Su" };
  const WD_JS = ["Su", "Ma", "Ti", "Ke", "To", "Pe", "La"];

  function show(view) {
    Object.keys(views).forEach((k) => (views[k].hidden = k !== view));
    window.scrollTo(0, 0);
  }
  const fmt = (m) => ShiftCalc.formatHM(m);
  function dateFi(iso) {
    if (!iso) return "";
    const p = iso.split("-");
    return p.length === 3 ? p[2] + "." + p[1] + "." + p[0] : iso;
  }
  function dayLabel(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
    if (!m) return iso || "";
    return WD_JS[new Date(+m[1], +m[2] - 1, +m[3]).getDay()] + " " + +m[3] + "." + +m[2] + ".";
  }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
  }
  function getJSON(key, def) {
    try {
      const v = JSON.parse(localStorage.getItem(key) || "null");
      return v == null ? def : v;
    } catch (e) {
      return def;
    }
  }
  function setJSON(key, v) {
    try {
      if (v == null || (typeof v === "object" && !Object.keys(v).length)) localStorage.removeItem(key);
      else localStorage.setItem(key, JSON.stringify(v));
    } catch (e) {}
  }

  // --- Overtime thresholds (defaults + per-jakso override) ---
  function loadOtOverride() {
    if (!state) return null;
    const o = getJSON(OT_STORE_PREFIX + state.startDate, null);
    if (!o || o.v !== 2 || (o.lisaK == null && o.yliK == null && o.sataK == null)) return null;
    return o;
  }
  function saveOtOverride(o) {
    if (!o || (o.lisaK == null && o.yliK == null && o.sataK == null)) setJSON(OT_STORE_PREFIX + state.startDate, null);
    else setJSON(OT_STORE_PREFIX + state.startDate, Object.assign({ v: 2 }, o));
  }
  function effectiveOt() {
    return ShiftCalc.applyOvertimeOverride(state.overtime, loadOtOverride());
  }

  // --- Per person: laskentapäivät, keskeytyspäivät, LM ---
  function laskKey(p) {
    return p.col + ":" + p.name;
  }
  const loadLaskMap = () => (state ? getJSON(LASK_PREFIX + state.startDate, {}) : {});
  const loadMarkMap = () => (state ? getJSON(MARK_PREFIX + state.startDate, {}) : {});
  function personMarks(p) {
    return loadMarkMap()[laskKey(p)] || {};
  }
  function isSickDay(p, d) {
    return !!(d && ShiftCalc.isKeskeytys(personMarks(p)[d.date]));
  }
  function sickCount(p) {
    return ShiftCalc.sickDates(p.days, personMarks(p)).length;
  }
  function personNInfo(p) {
    const raw = loadLaskMap()[laskKey(p)];
    const manual = raw == null || raw === "" ? null : parseInt(raw, 10);
    return ShiftCalc.resolveLaskenta(state.dayCount, sickCount(p), Number.isFinite(manual) ? manual : null);
  }
  function personN(p) {
    const info = personNInfo(p);
    return info.n == null ? 21 : info.n;
  }
  function savePersonN(p, n) {
    const map = loadLaskMap();
    map[laskKey(p)] = n;
    setJSON(LASK_PREFIX + state.startDate, map);
  }
  function clearPersonN(p) {
    const map = loadLaskMap();
    delete map[laskKey(p)];
    setJSON(LASK_PREFIX + state.startDate, map);
  }
  function toggleSick(p, date) {
    const map = loadMarkMap();
    const key = laskKey(p);
    const cur = Object.assign({}, map[key] || {});
    if (ShiftCalc.isKeskeytys(cur[date])) delete cur[date];
    else cur[date] = "keskeytys";
    if (Object.keys(cur).length) map[key] = cur;
    else delete map[key];
    setJSON(MARK_PREFIX + state.startDate, map);
  }
  function personRest(p) {
    return ShiftCalc.restWithExtras(p.days, NONE, personMarks(p));
  }
  function sumPerson(p, key) {
    return ShiftCalc.sumActive(p.days, personMarks(p), key, NONE, state.holidayMap, NONE);
  }
  function isSteward(p) {
    return !!getJSON(STEWARD_KEY, {})[p.name];
  }
  function setSteward(p, on) {
    setJSON(STEWARD_KEY, ShiftCalc.stewardMapSet(getJSON(STEWARD_KEY, {}), p.name, on));
  }
  /** Every day of the jakso is keskeytys (e.g. full loma) → LM +0:00. */
  function wholeKeskeytys(p) {
    return state.dayCount > 0 && sickCount(p) >= state.dayCount;
  }
  function stewardMinutes(p) {
    return ShiftCalc.stewardBonus(isSteward(p), null, wholeKeskeytys(p));
  }
  function activeLm(p) {
    return sumPerson(p, "me") + stewardMinutes(p);
  }
  function activeCompany(p) {
    return sumPerson(p, "company") + stewardMinutes(p);
  }
  function stewardLabel() {
    const m = ShiftCalc.stewardBonus(true);
    return "LM (+" + (m % 60 === 0 ? String(m / 60) : fmt(m)) + " h / jakso)";
  }
  function stewardLine(p) {
    return wholeKeskeytys(p) ? "LM-tunnit +0:00 (koko jakso keskeytyksellä)" : "LM-tunnit +" + fmt(ShiftCalc.stewardBonus(true));
  }
  function personOt(p) {
    if (!state.otEffective) return null;
    const lm = activeLm(p);
    if (state.dayCount !== 21) return ShiftCalc.overtimeSplit(lm, state.otEffective);
    const n = personN(p);
    if (n < 1) return { lisa: 0, yli50: 0, yli100: 0, n: 0, unsupported: null, cap50Hit: false, capLisaHit: false, norms: null };
    return ShiftCalc.interruptedSplit(lm, state.otEffective, n);
  }
  function otNote(ot) {
    if (!ot) return "";
    if (ot.unsupported === "override") return "muutetut rajat: tarkista käsin";
    if (ot.unsupported === "table") return "ei taulukossa";
    const bits = [];
    if (ot.capLisaHit && ot.norms) bits.push("Lisätyö katto " + ot.norms.lisaCapH + " h");
    if (ot.cap50Hit && ot.norms) bits.push("50 % katto " + ot.norms.cap50H + " h");
    return bits.join(" · ");
  }
  function laskentaLine(p) {
    const n = personN(p);
    const th = state.otEffective;
    if (!th) return "";
    if (n < 1) return "0 laskentapäivää";
    if (th.overridden && n < 21) return "muutetut rajat: tarkista käsin";
    if (th.overridden) return "21 pv · muutetut rajat käytössä";
    const norms = ShiftCalc.interruptedNorms(th.lisaK, th.yliK, n);
    if (!norms) return "ei taulukossa";
    let t = "Taulukko " + norms.table + " · " + n + " pv: ";
    if (!norms.noLisa) t += "Lisätyö yli " + fmt(norms.lisaK) + " (enint. " + norms.lisaCapH + " h) · ";
    return t + "50 % yli " + fmt(norms.yliK) + " (enint. " + norms.cap50H + " h) · 100 % yli " + fmt(norms.k100);
  }
  function fmtOt(ot, key) {
    return !ot || ot[key] == null ? "—" : fmt(ot[key]);
  }
  function overtimeText(th) {
    const def = state.overtime;
    if (!th) return "Ylityörajat: ei oletusta " + state.dayCount + " pv jaksolle";
    let t = (th.lisaK < th.yliK ? "Lisätyö yli " + fmt(th.lisaK) + " · " : "") + "50 % yli " + fmt(th.yliK) + " · 100 % yli " + fmt(th.sataK);
    if (def && def.holidays.length) t += " · Arkipyhät: " + def.holidays.map((h) => h.name + " " + dateFi(h.date).replace(/\d{4}$/, "")).join(", ");
    return t;
  }
  function renderOtMeta() {
    const th = state.otEffective;
    const ov = loadOtOverride() || {};
    const def = state.overtime;
    const val = (k) => (ov[k] != null ? fmt(ov[k]) : "");
    const ph = (k) => (def ? fmt(def[k]) : "h:mm");
    $("otMeta").innerHTML =
      '<span class="ot-text">' + overtimeText(th) + "</span>" +
      (th && th.overridden ? ' <span class="ot-mod">muutettu</span>' : "") +
      ' <button type="button" class="ot-edit-btn" id="otEditBtn">Muuta</button>' +
      '<div class="ot-editor" id="otEditor" hidden>' +
      '<label>Lisätyö alkaa<input id="otLisa" inputmode="decimal" autocomplete="off" placeholder="' + ph("lisaK") + '" value="' + val("lisaK") + '"></label>' +
      '<label>50 % alkaa<input id="ot50" inputmode="decimal" autocomplete="off" placeholder="' + ph("yliK") + '" value="' + val("yliK") + '"></label>' +
      '<label>100 % alkaa<input id="ot100" inputmode="decimal" autocomplete="off" placeholder="' + ph("sataK") + '" value="' + val("sataK") + '"></label>' +
      '<div class="ot-actions"><button type="button" class="ot-save" id="otSave">Tallenna</button><button type="button" class="ot-reset" id="otReset">Palauta oletus</button></div>' +
      '<p class="ot-err" id="otErr" hidden></p>' +
      '<p class="ot-note">' + (def ? "Tyhjä kenttä = oletus. " : "Anna 50 % ja 100 % rajat (h:mm); tyhjä lisätyö = sama kuin 50 %. ") + "Tallentuu vain tähän laitteeseen.</p></div>";
    $("otEditBtn").addEventListener("click", () => {
      const ed = $("otEditor");
      ed.hidden = !ed.hidden;
      if (!ed.hidden) $("otLisa").focus();
    });
    $("otReset").addEventListener("click", () => {
      saveOtOverride(null);
      renderOverview(true);
    });
    $("otSave").addEventListener("click", () => {
      const err = (msg) => {
        $("otErr").textContent = msg;
        $("otErr").hidden = false;
      };
      const read = (id) => {
        const raw = $(id).value.trim();
        return { raw: raw, v: raw ? ShiftCalc.parseHM(raw) : null };
      };
      const fL = read("otLisa"), f50 = read("ot50"), f100 = read("ot100");
      if ([fL, f50, f100].some((f) => f.raw && f.v == null)) return err("Anna aika muodossa h:mm, esim. 106:45.");
      const same = (v, k) => (def && v === def[k] ? null : v);
      const o = { lisaK: same(fL.v, "lisaK"), yliK: same(f50.v, "yliK"), sataK: same(f100.v, "sataK") };
      const dk = (k) => (def ? def[k] : null);
      const e50 = f50.v != null ? f50.v : dk("yliK");
      const e100 = f100.v != null ? f100.v : dk("sataK");
      const eL = fL.v != null ? fL.v : def ? def.lisaK : e50;
      if (!def && (e50 == null || e100 == null) && (fL.v != null || f50.v != null || f100.v != null)) return err("Anna ainakin 50 % ja 100 % rajat.");
      if (eL != null && e50 != null && eL > e50) return err("Lisätyö ei voi alkaa 50 % rajan jälkeen.");
      if (e50 != null && e100 != null && e100 < e50) return err("100 % raja ei voi olla pienempi kuin 50 % raja.");
      saveOtOverride(o);
      renderOverview(true);
    });
  }

  // --- Checks ---
  function lab(s) {
    if (s === "Md,s Check") return "TARKISTA";
    if (s === "Not Allowed") return "EI SALLITTU";
    if (s === "ok") return "OK";
    return s || "";
  }
  function restClass(s) {
    return s === "ok" ? "cell-ok" : s === "Md,s Check" ? "cell-md" : s === "Not Allowed" ? "cell-bad" : "";
  }
  function personProblems(p, long) {
    const problems = [];
    const rest = personRest(p);
    const dshort = (iso) => (long ? dateFi(iso) : dateFi(iso).replace(/\.\d{4}$/, "."));
    ShiftCalc.sheetMismatches(p.days).forEach((d) => {
      problems.push({ kind: "md", text: "Tarkista erotus " + dshort(d.date) + " = " + fmt(d.check) + " (Tunnit yhteensä ≠ Yritys)" });
    });
    p.days.forEach((d) => {
      if (isSickDay(p, d)) return;
      const r = rest[d.date] || {};
      if (r.restBefore === "Not Allowed" || r.restBefore === "Md,s Check") {
        const lepo = r.restBeforeMin != null ? fmt(r.restBeforeMin) : "";
        problems.push({
          kind: r.restBefore === "Not Allowed" ? "bad" : "md",
          text: lab(r.restBefore) + " " + (WD_SHORT[d.weekday] || "") + " " + dshort(d.date) + (lepo ? " (vuorojen väli " + lepo + ")" : ""),
        });
      }
    });
    return problems;
  }
  function personStatus(problems) {
    if (problems.some((x) => x.kind === "bad")) return "bad";
    return problems.length ? "md" : "ok";
  }
  function renderAlerts(target, p) {
    const items = personProblems(p, true);
    if (!items.length) {
      target.hidden = true;
      target.innerHTML = "";
      return;
    }
    const hasBad = items.some((i) => i.kind === "bad");
    target.hidden = false;
    target.className = "alerts" + (hasBad ? " bad" : "");
    target.innerHTML = "<h3>" + (hasBad ? "Huomio: ongelmia löytyi" : "Huomioita") + "</h3><ul>" + items.map((i) => "<li>" + i.text + "</li>").join("") + "</ul>";
  }
  function breakMarker(p) {
    if (state.dayCount !== 21) return "";
    const n = personN(p);
    return n === 21 ? "" : '<p class="card-break">Keskeytynyt · ' + n + " pv</p>";
  }

  // --- Planned lisät + comparison metrics ---
  function personLisat(p) {
    return ShiftCalc.lisatSummary(p.days, personMarks(p), NONE, NONE, state.holidayMap);
  }
  function suVuorot(p) {
    return p.days.filter((d) => d.start != null && !isSickDay(p, d) && (d.weekday === "Sunnuntai" || !!(state.holidayMap || {})[d.date])).length;
  }
  /** One person's numbers for the Vertailu table / history (minutes; counts). */
  function personMetrics(p) {
    const L = personLisat(p).minutes;
    const ot = personOt(p);
    return {
      name: p.name,
      n: state.dayCount === 21 ? personN(p) : null,
      sick: sickCount(p),
      tunnit: activeLm(p),
      kovat: sumPerson(p, "hrs"),
      su: L.su,
      la: L.la,
      yo: L.yo,
      ilta: L.ilta,
      aatto: L.aatto,
      suVuorot: suVuorot(p),
      lisa: ot && ot.lisa != null ? ot.lisa : null,
      yli50: ot && ot.yli50 != null ? ot.yli50 : null,
      yli100: ot && ot.yli100 != null ? ot.yli100 : null,
    };
  }

  // --- Overview ---
  /** Which block (group) of a multi-group sheet is in use; other groups are ignored. */
  function renderGroupNote() {
    const el = $("groupNote");
    if (!el) return;
    const others = (state.groups || []).filter((g) => g.title !== state.group);
    if (!state.group && !state.groupNotice) {
      el.hidden = true;
      return;
    }
    el.hidden = false;
    el.className = "group-note" + (state.groupNotice ? " warn" : "");
    el.textContent = state.groupNotice
      ? state.groupNotice
      : "Ryhmä: " + state.group + (others.length ? " · muut ryhmät ohitettu (" + others.map((g) => g.title).join(", ") + ")" : "");
  }

  function renderOverview(keepScroll) {
    const people = state.people;
    state.otEffective = effectiveOt();
    $("periodMeta").textContent = (state.periodLabel ? state.periodLabel + " · " : "") + "Alku " + dateFi(state.startDate);
    renderOtMeta();
    $("overviewTitle").textContent = "Kaikki " + people.length + " henkilöä";
    let nBad = 0, nMd = 0, nOk = 0;
    const cards = people.map((p) => {
      const problems = personProblems(p);
      const status = personStatus(problems);
      if (status === "bad") nBad++;
      else if (status === "md") nMd++;
      else nOk++;
      const ot = personOt(p);
      const badge = status === "bad" ? '<span class="pill bad">EI SALLITTU</span>' : status === "md" ? '<span class="pill md">TARKISTA</span>' : '<span class="pill ok">OK</span>';
      const cell = (l, v) => "<div><span>" + l + "</span><b>" + v + "</b></div>";
      return (
        '<button type="button" class="person-card status-' + status + '" data-id="' + p.id + '">' +
        '<div class="person-card-top"><div><h3>' + esc(p.name) + (isSteward(p) ? ' <span class="lm-tag">LM</span>' : "") + "</h3>" +
        '<p class="card-meta">' + p.shiftCount + " vuoroa</p></div>" + badge + "</div>" +
        '<div class="card-hours">' +
        cell("Kovat tunnit", fmt(sumPerson(p, "hrs"))) +
        cell("Tunnit yhteensä", fmt(activeLm(p))) +
        cell("Yritys", fmt(activeCompany(p))) +
        cell("Yö h", fmt(sumPerson(p, "night"))) +
        cell("Lauantai h", fmt(sumPerson(p, "b25"))) +
        cell("Pyhä h", fmt(sumPerson(p, "b100"))) +
        cell("Lisätyö h", fmtOt(ot, "lisa")) +
        cell("Ylityö 50 %", fmtOt(ot, "yli50")) +
        cell("Ylityö 100 %", fmtOt(ot, "yli100")) +
        "</div>" +
        breakMarker(p) +
        (otNote(ot) ? '<p class="card-ot-note">' + otNote(ot) + "</p>" : "") +
        (problems.length ? '<ul class="card-problems">' + problems.map((pr) => '<li class="prob-' + pr.kind + '">' + pr.text + "</li>").join("") + "</ul>" : '<p class="card-ok-line">Ei huomautuksia</p>') +
        '<span class="card-open">Avaa vuorotaulu →</span></button>'
      );
    });
    $("personCards").innerHTML = cards.join("");
    $("personCards").querySelectorAll(".person-card").forEach((btn) => btn.addEventListener("click", () => openDetail(+btn.dataset.id)));
    const bits = [];
    if (nBad) bits.push(nBad + " EI SALLITTU");
    if (nMd) bits.push(nMd + " TARKISTA");
    if (nOk) bits.push(nOk + " OK");
    $("overviewSummary").textContent = people.reduce((a, p) => a + p.shiftCount, 0) + " vuoroa · " + bits.join(" · ");
    renderGroupNote();
    const y = window.scrollY;
    show("overview");
    if (keepScroll) window.scrollTo(0, y);
  }

  // --- Detail ---
  function renderLisat(p) {
    const L = personLisat(p);
    const C = ShiftCalc.LISA_CODES;
    const rows = [
      ["ilta", "Iltatyölisä", "klo 18–21 · §17"],
      ["yo", "Yötyölisä", "klo 21–06, jatko klo 12 asti jos alkanut viim. klo 4 · §18"],
      ["la", "Lauantaityökorvaus", "arkilauantai klo 06–18 · §19"],
      ["su", "Sunnuntaityökorvaus", "su ja pyhät klo 0–24 + edellinen päivä klo 18–24 · §20"],
      ["aatto", "Aattopäivänlisä", "pääsiäislauantai, juhannus- ja jouluaatto klo 0–18 · §21"],
    ];
    $("detailLisat").innerHTML =
      "<h3>Suunnitellut lisät (TES)</h3>" +
      '<p class="ot-meta">Suunniteltujen vuorojen mukaan, keskeytyspäivät pois. Tunnit lasketaan yhteen koko jaksolta ja pyöristetään kerran (§24): alle 30 min alas, 30 min tai yli ylös.</p>' +
      '<table class="lisat-table"><thead><tr><th>Tuntilisä</th><th>Koodi</th><th>Suunniteltu</th><th>Pyöristetty</th></tr></thead><tbody>' +
      rows.map((r) => "<tr><td><b>" + r[1] + "</b><small>" + r[2] + "</small></td><td>" + (C[r[0]] || "–") + "</td><td>" + fmt(L.minutes[r[0]]) + '</td><td class="lisat-pay">' + L.hours[r[0]] + " h</td></tr>").join("") +
      "</tbody></table>";
  }

  function openDetail(personId) {
    const p = state.people.find((x) => x.id === personId);
    if (!p) return;
    state.otEffective = effectiveOt();
    $("detailTitle").textContent = p.name + " · " + p.shiftCount + " vuoroa";
    const rests = personRest(p);
    const sickBtn = (d, sick) => '<button type="button" class="sick-btn' + (sick ? " on" : "") + '" data-date="' + d.date + '">' + (sick ? "Keskeytyspäivä ✓" : "Keskeytyspäivä") + "</button>";
    $("detailTable").querySelector("tbody").innerHTML = p.days
      .map((d) => {
        const sick = isSickDay(p, d);
        const cell = (v) => '<td class="' + (sick ? "struck" : "") + '">' + v + "</td>";
        const r = rests[d.date] || {};
        return (
          '<tr class="' + (sick ? "sick" : "") + '"><td>' + sickBtn(d, sick) + " " + dateFi(d.date).slice(0, 5) + "</td><td>" + (WD_SHORT[d.weekday] || "") + "</td><td>" + (d.special || "") + "</td>" +
          "<td>" + fmt(d.start) + "</td><td>" + fmt(d.end) + '</td><td class="' + (sick ? "" : restClass(r.restAfter)) + '">' + (sick ? "—" : lab(r.restAfter)) + "</td>" +
          cell(fmt(d.hrs)) + cell(fmt(d.night)) + cell(fmt(d.b25)) + cell(fmt(d.b100)) + cell(fmt(d.me)) + cell(fmt(d.company)) +
          '<td class="' + (sick ? "struck" : d.check ? "cell-md" : "") + '">' + fmt(d.check) + "</td></tr>"
        );
      })
      .join("");
    $("detailCards").innerHTML = p.days
      .map((d) => {
        const has = d.start != null;
        const sick = isSickDay(p, d);
        const rd = rests[d.date] || {};
        const restLabel = !sick && rd.restAfter ? lab(rd.restAfter) + (rd.restAfterMin != null ? " · " + fmt(rd.restAfterMin) : "") : "";
        const g = (l, v, cls) => '<div class="' + (cls || "") + '"><span>' + l + "</span><b>" + v + "</b></div>";
        return (
          '<article class="day-card' + (has ? "" : " empty") + (sick ? " sick" : "") + '"><header><strong>' + (WD_SHORT[d.weekday] || "") + " " + dateFi(d.date) + "</strong>" +
          (d.special ? '<span class="tag">' + d.special + "</span>" : "") + sickBtn(d, sick) + "</header>" +
          (has
            ? '<div class="day-grid' + (sick ? " struck" : "") + '">' +
              g("Alku", fmt(d.start)) + g("Loppu", fmt(d.end)) + g("Kovat tunnit", fmt(d.hrs)) + g("Yö h", fmt(d.night)) +
              g("Tunnit yhteensä", fmt(d.me)) + g("Yritys", fmt(d.company)) + g("Lauantai h", fmt(d.b25)) + g("Pyhä h", fmt(d.b100)) +
              g("Erotus", fmt(d.check), d.check ? "cell-md" : "") + g("Vuorojen väli →", restLabel || "—", sick ? "" : restClass(rd.restAfter)) +
              "</div>"
            : '<p class="muted">Ei vuoroa</p>') +
          "</article>"
        );
      })
      .join("");
    const ot = personOt(p);
    const stat = (v, l) => '<div class="stat"><b>' + v + "</b><span>" + l + "</span></div>";
    $("detailTotals").innerHTML =
      stat(fmt(sumPerson(p, "hrs")), "Kovat tunnit") + stat(fmt(sumPerson(p, "night")), "Yö h") + stat(fmt(sumPerson(p, "b25")), "Lauantai h") +
      stat(fmt(sumPerson(p, "b100")), "Pyhä h") + stat(fmt(activeLm(p)), "Tunnit yhteensä") + stat(fmt(activeCompany(p)), "Yritys yht.") +
      stat(fmtOt(ot, "lisa"), "Lisätyö h") + stat(fmtOt(ot, "yli50"), "Ylityö 50 %") + stat(fmtOt(ot, "yli100"), "Ylityö 100 %") +
      '<p class="ot-meta totals-note">' + overtimeText(state.otEffective) + (state.otEffective && state.otEffective.overridden ? ' <span class="ot-mod">muutettu</span>' : "") + "</p>" +
      (isSteward(p) ? '<p class="ot-meta totals-note">' + stewardLine(p) + "</p>" : "") +
      (otNote(ot) ? '<p class="ot-meta totals-note">' + otNote(ot) + "</p>" : "");
    const canLask = state.dayCount === 21 && !!state.otEffective;
    $("laskentaBox").hidden = !canLask;
    const paintLaskenta = (keepInput) => {
      if (!canLask) return;
      const info = personNInfo(p);
      const sick = sickCount(p);
      $("laskentaLabel").textContent = "Laskentapäivät " + info.n + (sick ? " (" + sick + " " + (sick === 1 ? "keskeytyspäivä" : "keskeytyspäivää") + ")" : "");
      if (!keepInput) $("laskentaInput").value = info.n == null ? "" : info.n;
      $("laskentaReset").hidden = !info.manual;
      $("laskentaLine").textContent = laskentaLine(p);
    };
    paintLaskenta(false);
    const redraw = () => {
      const y = window.scrollY;
      openDetail(p.id);
      window.scrollTo(0, y);
    };
    $("lmCheckLabel").textContent = stewardLabel() + " · vain tunnit";
    $("lmCheck").checked = isSteward(p);
    $("lmCheck").onchange = () => {
      setSteward(p, $("lmCheck").checked);
      redraw();
    };
    $("laskentaInput").oninput = () => {
      const v = parseInt($("laskentaInput").value, 10);
      if (!(v >= 1 && v <= 21)) return;
      savePersonN(p, v);
      redraw();
    };
    $("laskentaReset").onclick = () => {
      clearPersonN(p);
      redraw();
    };
    document.querySelectorAll("#view-detail .sick-btn").forEach((btn) =>
      btn.addEventListener("click", () => {
        toggleSick(p, btn.dataset.date);
        redraw();
      })
    );
    renderLisat(p);
    renderAlerts($("detailAlerts"), p);
    show("detail");
  }

  // --- Vertailu ---
  const COLS = [
    { k: "tunnit", l: "Tunnit yht.", t: "h" },
    { k: "kovat", l: "Kovat", t: "h" },
    { k: "su", l: "Su h", t: "h" },
    { k: "la", l: "La h", t: "h" },
    { k: "yo", l: "Yö h", t: "h" },
    { k: "ilta", l: "Ilta h", t: "h" },
    { k: "aatto", l: "Aatto h", t: "h" },
    { k: "suVuorot", l: "Su-vuorot", t: "n" },
    { k: "lisa", l: "Lisätyö", t: "h", ot: true },
    { k: "yli50", l: "50 %", t: "h", ot: true },
    { k: "yli100", l: "100 %", t: "h", ot: true },
  ];
  let sort = { k: "name", dir: 1 };
  function settings() {
    return Object.assign({ thr: 240, thrN: 1, scale: false, mode: "jakso", year: null }, getJSON(SETTINGS_KEY, {}));
  }
  function saveSettings(s) {
    setJSON(SETTINGS_KEY, s);
  }
  const PUB = "julkaistu", DRAFT = "luonnos";
  /** Old entries (saved before the Luonnos/Julkaistu flag) count as published: they were saved for the year on purpose,
   *  so this keeps year totals unchanged. They're tagged legacy so the list can offer "Merkitse luonnokseksi". */
  function loadHistory() {
    const h = getJSON(HISTORY_KEY, {}) || {};
    Object.keys(h).forEach((k) => {
      if (h[k] && h[k].status !== PUB && h[k].status !== DRAFT) {
        h[k].status = PUB;
        h[k].legacy = true;
      }
    });
    return h;
  }
  const isPub = (e) => !!e && e.status === PUB;
  function statusBadge(st, legacy) {
    return st === PUB
      ? '<span class="st-badge st-pub">Julkaistu' + (legacy ? " (vanha)" : "") + "</span>"
      : '<span class="st-badge st-draft">Luonnos</span>';
  }
  const pubCount = (hist, year) => Object.keys(hist).filter((k) => isPub(hist[k]) && (year == null || k.slice(0, 4) === String(year))).length;
  const draftCount = (hist, year) => Object.keys(hist).filter((k) => !isPub(hist[k]) && (year == null || k.slice(0, 4) === String(year))).length;
  function saveHistory(h) {
    setJSON(HISTORY_KEY, h);
  }
  /** Rows for the current mode. Scaling (× 21 / laskentapäivät) applies to hours and Su-vuorot, not to overtime (already per laskentapäivät). */
  function vertRows(s) {
    let rows = [];
    let info = { text: "" };
    if (s.mode === "vuosi") {
      const hist = loadHistory();
      const keys = Object.keys(hist).filter((k) => k.slice(0, 4) === String(s.year) && isPub(hist[k])).sort();
      const drafts = draftCount(hist, s.year);
      const by = {};
      keys.forEach((k) =>
        (hist[k].people || []).forEach((m) => {
          const a = by[m.name] || (by[m.name] = { name: m.name, jaksot: 0, nSum: 0, sick: 0 });
          a.jaksot++;
          a.nSum += m.n == null ? 21 : m.n;
          a.sick += m.sick || 0;
          COLS.forEach((c) => {
            if (m[c.k] == null) return;
            a[c.k] = (a[c.k] || 0) + m[c.k];
          });
        })
      );
      rows = Object.keys(by).map((n) => {
        const a = by[n];
        a.factor = s.scale && a.nSum > 0 ? (21 * a.jaksot) / a.nSum : 1;
        return a;
      });
      const dn = drafts ? " " + drafts + (drafts === 1 ? " luonnos ei mukana." : " luonnosta ei mukana.") : "";
      info = keys.length
        ? "Vuosi " + s.year + ": " + keys.length + " julkaistua jaksoa (" + dateFi(keys[0]) + " – " + dateFi(hist[keys[keys.length - 1]].endDate || keys[keys.length - 1]) + "). Vain julkaistut jaksot lasketaan." + dn
        : "Vuodelle " + s.year + " ei ole julkaistuja jaksoja." + dn;
      info = { text: info, pub: keys.length, drafts: drafts };
    } else if (state) {
      state.otEffective = effectiveOt();
      rows = state.people.map((p) => {
        const m = personMetrics(p);
        m.jaksot = 1;
        m.factor = s.scale && m.n && m.n < 21 ? 21 / m.n : 1;
        return m;
      });
      info = {
        text: (state.group ? state.group + " · " : "") + (state.periodLabel ? state.periodLabel + " · " : "") + "Jakso " + dateFi(state.startDate) + " – " + dateFi(state.dates[state.dates.length - 1]) + " · " + rows.length + " kuljettajaa" +
          (state.published ? "" : " · luonnos: ei lasketa vuoteen"),
      };
    }
    rows.forEach((r) => {
      r.v = {};
      COLS.forEach((c) => {
        const raw = r[c.k];
        r.v[c.k] = raw == null ? null : c.ot ? raw : raw * r.factor;
      });
    });
    const avg = {};
    COLS.forEach((c) => {
      const vals = rows.map((r) => r.v[c.k]).filter((x) => x != null);
      avg[c.k] = vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
    });
    return { rows: rows, avg: avg, info: info.text, pub: info.pub, drafts: info.drafts };
  }
  const fmtV = (c, v) => (v == null ? "—" : c.t === "n" ? (Math.round(v * 10) / 10).toString().replace(".", ",") : fmt(Math.round(v)));
  function fmtDiff(c, d) {
    if (d == null) return "";
    if (c.t === "n") {
      const r = Math.round(d * 10) / 10;
      return (r > 0 ? "+" : r < 0 ? "−" : "±") + Math.abs(r).toString().replace(".", ",");
    }
    const m = Math.round(d);
    return (m > 0 ? "+" : m < 0 ? "−" : "±") + fmt(Math.abs(m));
  }
  function renderVertailu() {
    const s = settings();
    const hist = loadHistory();
    const years = Array.from(new Set(Object.keys(hist).map((k) => k.slice(0, 4)))).sort().reverse();
    if (!s.year) s.year = state ? +state.startDate.slice(0, 4) : years.length ? +years[0] : new Date().getFullYear();
    if (!state && s.mode === "jakso") s.mode = "vuosi";
    $("modeJakso").checked = s.mode === "jakso";
    $("modeVuosi").checked = s.mode === "vuosi";
    $("modeJakso").disabled = !state;
    const ys = Array.from(new Set(years.concat([String(s.year)]))).sort().reverse();
    $("vertYear").innerHTML = ys.map((y) => '<option value="' + y + '"' + (+y === +s.year ? " selected" : "") + ">" + y + "</option>").join("");
    $("vertYear").hidden = s.mode !== "vuosi";
    $("vertScale").checked = !!s.scale;
    if (document.activeElement !== $("vertThr")) $("vertThr").value = fmt(s.thr);
    if (document.activeElement !== $("vertThrN")) $("vertThrN").value = s.thrN;
    const V = vertRows(s);
    $("vertMeta").innerHTML =
      s.mode === "vuosi"
        ? "Vuosi " + s.year + ' <span class="st-badge st-pub">' + V.pub + " julkaistua</span>"
        : state ? "Jakso " + dateFi(state.startDate) + " " + statusBadge(state.published ? PUB : DRAFT) : "";
    $("vertInfo").textContent = V.info + (s.scale ? " · Skaalattu 21 päivään." : "");
    const rows = V.rows.slice().sort((a, b) => {
      if (sort.k === "name") return sort.dir * a.name.localeCompare(b.name, "fi", { numeric: true });
      const x = a.v[sort.k], y = b.v[sort.k];
      if (x == null && y == null) return 0;
      if (x == null) return 1;
      if (y == null) return -1;
      return sort.dir * (x - y);
    });
    const arrow = (k) => (sort.k === k ? (sort.dir > 0 ? " ▲" : " ▼") : "");
    const head =
      '<thead><tr><th data-k="name" class="sticky">Kuljettaja' + arrow("name") + "</th>" +
      (s.mode === "vuosi" ? '<th data-k="jaksot">Jaksot</th>' : '<th data-k="n">Pv</th>') +
      COLS.map((c) => '<th data-k="' + c.k + '">' + c.l + arrow(c.k) + "</th>").join("") + "</tr></thead>";
    const body = rows
      .map((r) => {
        return (
          '<tr><td class="sticky"><b>' + esc(r.name) + "</b></td>" +
          (s.mode === "vuosi" ? "<td>" + r.jaksot + "</td>" : "<td>" + (r.n == null ? "—" : r.n) + (r.sick ? '<small class="dev-note">' + r.sick + " kesk.</small>" : "") + "</td>") +
          COLS.map((c) => {
            const v = r.v[c.k], a = V.avg[c.k];
            const d = v == null || a == null ? null : v - a;
            const thr = c.t === "n" ? s.thrN : s.thr;
            const cls = d == null || Math.abs(d) < thr || thr <= 0 ? "" : d > 0 ? "dev-hi" : "dev-lo";
            return '<td class="' + cls + '">' + fmtV(c, v) + '<small class="dev">' + fmtDiff(c, d) + "</small></td>";
          }).join("") +
          "</tr>"
        );
      })
      .join("");
    const avgRow =
      '<tr class="vert-avg"><td class="sticky"><b>Keskiarvo</b></td><td></td>' + COLS.map((c) => "<td>" + fmtV(c, V.avg[c.k]) + "</td>").join("") + "</tr>";
    $("vertTable").innerHTML = head + "<tbody>" + (rows.length ? body + avgRow : '<tr><td colspan="13">Ei tietoja.</td></tr>') + "</tbody>";
    $("vertTable").querySelectorAll("th[data-k]").forEach((th) =>
      th.addEventListener("click", () => {
        const k = th.dataset.k;
        if (k === "n" || k === "jaksot") return;
        sort = sort.k === k ? { k: k, dir: -sort.dir } : { k: k, dir: k === "name" ? 1 : -1 };
        renderVertailu();
      })
    );
    $("btnSaveHist").hidden = !state;
    $("pubBox").hidden = !state;
    if (state) {
      $("pubCheck").checked = !!state.published;
      const saved = hist[state.startDate];
      $("btnSaveHist").textContent = state.published ? "Tallenna julkaistuna" : "Tallenna luonnoksena";
      $("pubHint").innerHTML = saved
        ? "Tallennettu: " + statusBadge(saved.status, saved.legacy) + " " + (saved.savedAt ? dateFi(saved.savedAt.slice(0, 10)) : "") +
          (isPub(saved) && !state.published ? " · luonnosta ei tallenneta julkaistun päälle" : "")
        : "Ei vielä tallennettu.";
    }
    renderHistList(hist);
    saveSettings(s);
  }
  function renderHistList(hist) {
    const keys = Object.keys(hist).sort().reverse();
    $("histList").innerHTML = keys.length
      ? "<h3>Tallennetut jaksot</h3><p class=\"ot-meta\">" + pubCount(hist) + " julkaistua · " + draftCount(hist) + " luonnosta. Vuosi-näkymä laskee vain julkaistut.</p><ul>" +
        keys
          .map((k) => {
            const h = hist[k];
            return (
              "<li>" + statusBadge(h.status, h.legacy) + " " + (h.periodLabel ? esc(h.periodLabel) + " · " : "") + dateFi(k) + " · " + (h.people || []).length + " kuljettajaa " +
              '<button type="button" class="linkish hist-st" data-k="' + k + '">' + (isPub(h) ? "Merkitse luonnokseksi" : "Merkitse julkaistuksi") + "</button> · " +
              '<button type="button" class="linkish hist-del" data-k="' + k + '">Poista</button></li>'
            );
          })
          .join("") +
        "</ul>"
      : "";
    $("histList").querySelectorAll(".hist-st").forEach((b) =>
      b.addEventListener("click", () => {
        const h = loadHistory();
        const e = h[b.dataset.k];
        if (!e) return;
        const toPub = !isPub(e);
        if (!confirm((toPub ? "Merkitäänkö jakso " + dateFi(b.dataset.k) + " julkaistuksi? Se lasketaan vuoteen." : "Merkitäänkö jakso " + dateFi(b.dataset.k) + " luonnokseksi? Sitä ei lasketa vuoteen."))) return;
        e.status = toPub ? PUB : DRAFT;
        delete e.legacy;
        saveHistory(h);
        renderVertailu();
        renderHomeHistory();
      })
    );
    $("histList").querySelectorAll(".hist-del").forEach((b) =>
      b.addEventListener("click", () => {
        if (!confirm("Poistetaanko jakso " + dateFi(b.dataset.k) + " vertailusta?")) return;
        const h = loadHistory();
        delete h[b.dataset.k];
        saveHistory(h);
        renderVertailu();
        renderHomeHistory();
      })
    );
  }
  const DIFF_COLS = [
    { k: "tunnit", l: "Tunnit yht.", t: "h" },
    { k: "kovat", l: "Kovat", t: "h" },
    { k: "su", l: "Su h", t: "h" },
    { k: "suVuorot", l: "Su-vuorot", t: "n" },
  ];
  /** Per-driver changes between two saved summaries (only drivers with a change). */
  function diffPeople(oldP, newP) {
    const byName = (arr) => {
      const m = {};
      (arr || []).forEach((x) => (m[x.name] = x));
      return m;
    };
    const o = byName(oldP), n = byName(newP);
    const names = Array.from(new Set(Object.keys(o).concat(Object.keys(n)))).sort((a, b) => a.localeCompare(b, "fi", { numeric: true }));
    const out = [];
    names.forEach((name) => {
      if (!o[name]) return out.push({ name: name, note: "uusi" });
      if (!n[name]) return out.push({ name: name, note: "poistunut" });
      const d = {};
      let any = false;
      DIFF_COLS.forEach((c) => {
        const x = (n[name][c.k] || 0) - (o[name][c.k] || 0);
        d[c.k] = x;
        if (x) any = true;
      });
      if (any) out.push({ name: name, d: d });
    });
    return out;
  }
  function diffText(diffs) {
    if (!diffs.length) return "Ei muutoksia tunteihin tai sunnuntaihin.";
    return diffs
      .slice(0, 12)
      .map((r) => r.name + ": " + (r.note || DIFF_COLS.filter((c) => r.d[c.k]).map((c) => c.l + " " + fmtDiff(c, r.d[c.k])).join(", ")))
      .join("\n") + (diffs.length > 12 ? "\n… ja " + (diffs.length - 12) + " muuta" : "");
  }
  function diffHtml(diffs, title) {
    if (!diffs.length) return '<p class="ot-meta">' + title + ": ei muutoksia tunteihin tai sunnuntaihin.</p>";
    return (
      '<h3 class="diff-title">' + title + '</h3><div class="table-wrap"><table class="vert-table diff-table"><thead><tr><th class="sticky">Kuljettaja</th>' +
      DIFF_COLS.map((c) => "<th>" + c.l + "</th>").join("") + "</tr></thead><tbody>" +
      diffs
        .map((r) =>
          '<tr><td class="sticky"><b>' + esc(r.name) + "</b></td>" +
          (r.note ? '<td colspan="' + DIFF_COLS.length + '">' + r.note + "</td>" : DIFF_COLS.map((c) => '<td class="' + (r.d[c.k] > 0 ? "dev-hi" : r.d[c.k] < 0 ? "dev-lo" : "") + '">' + (r.d[c.k] ? fmtDiff(c, r.d[c.k]) : "") + "</td>").join("")) +
          "</tr>"
        )
        .join("") +
      "</tbody></table></div>"
    );
  }
  function saveCurrentToHistory() {
    if (!state) return;
    state.otEffective = effectiveOt();
    const h = loadHistory();
    const prev = h[state.startDate];
    const status = state.published ? PUB : DRAFT;
    const people = state.people.map(personMetrics);
    const diffs = prev ? diffPeople(prev.people, people) : null;
    if (prev && isPub(prev) && status === DRAFT) {
      $("histMsg").textContent = "Jaksolle " + dateFi(state.startDate) + " on jo tallennettu julkaistu versio. Luonnosta ei tallenneta sen päälle. Rastita Julkaistu vuorotaulu, jos tämä on uusi julkaistu versio.";
      $("histDiff").innerHTML = diffHtml(diffs, "Ero tallennettuun julkaistuun");
      return;
    }
    if (prev && isPub(prev) && status === PUB && !confirm("Korvataanko aiempi julkaistu versio?\n\nMuutokset edelliseen:\n" + diffText(diffs))) {
      $("histMsg").textContent = "Ei tallennettu.";
      return;
    }
    h[state.startDate] = {
      startDate: state.startDate,
      endDate: state.dates[state.dates.length - 1],
      periodLabel: state.periodLabel || "",
      dayCount: state.dayCount,
      status: status,
      group: state.group || null,
      savedAt: new Date().toISOString(),
      people: people,
    };
    saveHistory(h);
    $("histMsg").textContent =
      "Jakso " + dateFi(state.startDate) + " tallennettu: " + (status === PUB ? "julkaistu, lasketaan vuoteen" : "luonnos, ei lasketa vuoteen") +
      (prev ? " (korvasi aiemman " + (isPub(prev) ? "julkaistun version" : "luonnoksen") + ")." : ".");
    $("histDiff").innerHTML = prev ? diffHtml(diffs, "Muutokset edelliseen tallennettuun (" + (isPub(prev) ? "julkaistu" : "luonnos") + ")") : "";
    renderVertailu();
    renderHomeHistory();
  }
  function openVertailu(from) {
    prevView = from || "overview";
    sort = { k: "name", dir: 1 };
    $("histMsg").textContent = "";
    $("histDiff").innerHTML = "";
    renderVertailu();
    show("vertailu");
  }

  // --- Lataa PDF: print view (window.print → "Tallenna PDF:nä"), works offline ---
  function printVertailu() {
    const s = settings();
    const V = vertRows(s);
    const rows = V.rows.slice().sort((a, b) => a.name.localeCompare(b.name, "fi", { numeric: true }));
    const head = "<tr><th>Kuljettaja</th><th>" + (s.mode === "vuosi" ? "Jaksot" : "Pv") + "</th>" + COLS.map((c) => "<th>" + c.l + "</th>").join("") + "</tr>";
    const body = rows
      .map((r) =>
        "<tr><td><b>" + esc(r.name) + "</b></td><td>" + (s.mode === "vuosi" ? r.jaksot : r.n == null ? "—" : r.n + (r.sick ? "<small>" + r.sick + " kesk.</small>" : "")) + "</td>" +
        COLS.map((c) => {
          const v = r.v[c.k], a = V.avg[c.k];
          const d = v == null || a == null ? null : v - a;
          const thr = c.t === "n" ? s.thrN : s.thr;
          const cls = d == null || Math.abs(d) < thr || thr <= 0 ? "" : d > 0 ? "pv-hi" : "pv-lo";
          return '<td class="' + cls + '">' + fmtV(c, v) + "<small>" + fmtDiff(c, d) + "</small></td>";
        }).join("") + "</tr>"
      )
      .join("");
    const avg = '<tr class="pv-avg"><td>Keskiarvo</td><td></td>' + COLS.map((c) => "<td>" + fmtV(c, V.avg[c.k]) + "</td>").join("") + "</tr>";
    let summary = "";
    if (state) {
      state.otEffective = effectiveOt();
      summary =
        "<h2>Kuljettajien tunnit</h2>" +
        '<table class="pv-table pv-vert"><thead><tr><th>Kuljettaja</th><th>Vuorot</th><th>Kesk.</th><th>Kovat</th><th>Tunnit yht.</th><th>Yritys</th><th>Yö h</th><th>La h</th><th>Pyhä h</th><th>Lisätyö</th><th>50 %</th><th>100 %</th><th class="pv-note">Huomautukset</th></tr></thead><tbody>' +
        state.people
          .map((p) => {
            const ot = personOt(p);
            const pr = personProblems(p);
            return (
              "<tr><td><b>" + esc(p.name) + "</b>" + (isSteward(p) ? " <small>LM</small>" : "") + "</td><td>" + p.shiftCount + "</td><td>" + (sickCount(p) || "") + "</td><td>" + fmt(sumPerson(p, "hrs")) + "</td><td>" + fmt(activeLm(p)) + "</td><td>" + fmt(activeCompany(p)) +
              "</td><td>" + fmt(sumPerson(p, "night")) + "</td><td>" + fmt(sumPerson(p, "b25")) + "</td><td>" + fmt(sumPerson(p, "b100")) + "</td><td>" + fmtOt(ot, "lisa") + "</td><td>" + fmtOt(ot, "yli50") + "</td><td>" + fmtOt(ot, "yli100") +
              '</td><td class="pv-note">' + (pr.length ? "<small>" + pr.map((x) => x.text).join("<br/>") + "</small>" : "OK") + "</td></tr>"
            );
          })
          .join("") +
        "</tbody></table>" +
        '<p class="pv-meta">' + overtimeText(state.otEffective) + "</p>";
    }
    const range = s.mode === "vuosi" ? "Vuosi " + s.year : dateFi(state.startDate) + "–" + dateFi(state.dates[state.dates.length - 1]);
    const html =
      "<style>@page { size: A4 landscape; }</style>" +
      '<header class="pv-head"><h1>Vuorotaulun tarkastus · Vertailu <span class="pv-badge">' +
      (s.mode === "vuosi" ? "Vain julkaistut (" + V.pub + ")" : state.published ? "Julkaistu" : "Luonnos") + "</span></h1><p>" + esc(V.info) + "</p><p class=\"pv-meta\">" +
      (s.scale ? "Skaalattu 21 päivään (× 21 / laskentapäivät, ei ylitöihin). " : "") +
      "Värirajat ± " + fmt(s.thr) + " h / ± " + s.thrN + " vuoroa. Pienet luvut = ero keskiarvoon. Tulostettu " + dateFi(new Date().toISOString().slice(0, 10)) + "</p></header>" +
      '<table class="pv-table pv-vert"><thead>' + head + "</thead><tbody>" + body + avg + "</tbody></table>" +
      '<p class="pv-meta">Su/La/Yö/Ilta/Aatto h = suunnitellut TES-lisätunnit ennen pyöristystä. Su-vuorot = vuorot, jotka alkavat sunnuntaina tai pyhänä.</p>' +
      summary +
      '<p class="pv-foot">© 2026 Lämpöpumppu Mafia · Kaikki oikeudet pidätetään</p>';
    let box = $("printView");
    if (!box) {
      box = document.createElement("div");
      box.id = "printView";
      document.body.appendChild(box);
    }
    box.innerHTML = html;
    const oldTitle = document.title;
    document.title = "Vuorotaulun tarkastus – Vertailu – " + range + (s.mode === "vuosi" ? "" : state.published ? " – Julkaistu" : " – Luonnos");
    const restore = () => {
      document.title = oldTitle;
      window.removeEventListener("afterprint", restore);
    };
    window.addEventListener("afterprint", restore);
    window.print();
  }

  // --- CSV (history backup) ---
  const CSV_COLS = ["jakso_alku", "jakso_loppu", "jakso", "tila", "nimi", "laskentapaivat", "keskeytyspaivat", "tunnit_yhteensa", "kovat_tunnit", "su_h", "la_h", "yo_h", "ilta_h", "aatto_h", "su_vuorot", "lisatyo", "ylityo_50", "ylityo_100"];
  const CSV_KEYS = { tunnit_yhteensa: "tunnit", kovat_tunnit: "kovat", su_h: "su", la_h: "la", yo_h: "yo", ilta_h: "ilta", aatto_h: "aatto", lisatyo: "lisa", ylityo_50: "yli50", ylityo_100: "yli100" };
  function csvCell(v) {
    const s = String(v == null ? "" : v);
    return /[;"\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }
  function exportCsv() {
    const h = loadHistory();
    const lines = [CSV_COLS.join(";")];
    Object.keys(h).sort().forEach((k) =>
      (h[k].people || []).forEach((m) => {
        lines.push(
          CSV_COLS.map((c) => {
            if (c === "jakso_alku") return k;
            if (c === "jakso_loppu") return h[k].endDate || "";
            if (c === "jakso") return csvCell(h[k].periodLabel || "");
            if (c === "tila") return h[k].status;
            if (c === "nimi") return csvCell(m.name);
            if (c === "laskentapaivat") return m.n == null ? "" : m.n;
            if (c === "keskeytyspaivat") return m.sick || 0;
            if (c === "su_vuorot") return m.suVuorot;
            const v = m[CSV_KEYS[c]];
            return v == null ? "" : fmt(v);
          }).join(";")
        );
      })
    );
    const blob = new Blob(["\ufeff" + lines.join("\r\n")], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "tarkastus-vertailu-" + new Date().toISOString().slice(0, 10) + ".csv";
    document.body.appendChild(a);
    a.click();
    a.remove();
    $("histMsg").textContent = "CSV ladattu (" + (lines.length - 1) + " riviä).";
  }
  function parseCsvLine(line) {
    const out = [];
    let cur = "", q = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (q) {
        if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
        else if (ch === '"') q = false;
        else cur += ch;
      } else if (ch === '"') q = true;
      else if (ch === ";" || ch === ",") { out.push(cur); cur = ""; }
      else cur += ch;
    }
    out.push(cur);
    return out;
  }
  function importCsv(text) {
    const lines = text.replace(/^\ufeff/, "").split(/\r?\n/).filter((l) => l.trim());
    if (!lines.length) throw new Error("Tyhjä tiedosto.");
    const head = parseCsvLine(lines[0]).map((x) => x.trim().toLowerCase());
    const idx = (c) => head.indexOf(c);
    if (idx("jakso_alku") < 0 || idx("nimi") < 0) throw new Error("Tiedosto ei ole Tarkastuksen vertailu-CSV.");
    const h = loadHistory();
    const touched = {};
    lines.slice(1).forEach((l) => {
      const v = parseCsvLine(l);
      const g = (c) => (idx(c) >= 0 ? (v[idx(c)] || "").trim() : "");
      const k = g("jakso_alku");
      if (!/^\d{4}-\d{2}-\d{2}$/.test(k)) return;
      if (!touched[k]) {
        touched[k] = true;
        // No "tila" column (older CSV) → published, same rule as old saved entries.
        const t = g("tila").toLowerCase();
        h[k] = { startDate: k, endDate: g("jakso_loppu") || null, periodLabel: g("jakso"), status: t === DRAFT ? DRAFT : PUB, savedAt: new Date().toISOString(), people: [] };
      }
      const m = { name: g("nimi"), n: g("laskentapaivat") ? +g("laskentapaivat") : null, sick: +g("keskeytyspaivat") || 0, suVuorot: +g("su_vuorot") || 0 };
      Object.keys(CSV_KEYS).forEach((c) => {
        const t = g(c);
        m[CSV_KEYS[c]] = t ? ShiftCalc.parseHM(t) : null;
      });
      h[k].people.push(m);
    });
    saveHistory(h);
    return Object.keys(touched).length;
  }
  function renderHomeHistory() {
    const h = loadHistory();
    const n = Object.keys(h).length;
    $("homeHistory").textContent = n ? pubCount(h) + " julkaistua jaksoa ja " + draftCount(h) + " luonnosta vertailussa." : "Ei tallennettuja jaksoja.";
    $("btnHomeVertailu").hidden = !n;
  }

  // --- File handling ---
  let pendingBuf = null;
  let pendingName = "";
  function showStartPrompt(result, fileName) {
    $("startPrompt").hidden = false;
    $("startPromptText").textContent = "Aloituspäivää ei löytynyt — anna jakson alkupäivä" + (result.periodLabel ? " (" + result.periodLabel + ")" : "");
    $("startDateErr").hidden = true;
    $("fileName").hidden = false;
    $("fileName").textContent = (fileName || "Tiedosto") + " — aloituspäivä puuttuu";
    show("home");
    $("startDateInput").focus();
  }
  function useResult(result, fileName) {
    if (result && result.needsStartDate) {
      showStartPrompt(result, fileName);
      return result;
    }
    $("startPrompt").hidden = true;
    state = result;
    if (state) state.published = false; // each load starts as a draft; tick Julkaistu vuorotaulu for the published version
    if (fileName) {
      $("fileName").hidden = false;
      $("fileName").textContent = fileName;
    }
    renderOverview();
    return state;
  }
  async function handleFile(file) {
    if (!file) return;
    $("startPrompt").hidden = true;
    $("fileName").hidden = false;
    $("fileName").textContent = "Luetaan: " + file.name + " …";
    try {
      const buf = await file.arrayBuffer();
      pendingBuf = buf;
      pendingName = file.name;
      useResult(ShiftCalc.parseArrayBuffer(buf), file.name);
    } catch (err) {
      $("fileName").textContent = "Virhe: " + (err && err.message ? err.message : String(err));
      alert("Tiedoston lukeminen epäonnistui: " + (err && err.message ? err.message : err));
    }
  }
  $("fileInput").addEventListener("change", (e) => handleFile(e.target.files && e.target.files[0]));
  $("btnNewFile").addEventListener("click", () => {
    state = null;
    $("fileInput").value = "";
    $("fileName").hidden = true;
    $("startPrompt").hidden = true;
    pendingBuf = null;
    renderHomeHistory();
    show("home");
  });
  $("btnBack").addEventListener("click", () => renderOverview());
  $("btnVertailu").addEventListener("click", () => openVertailu("overview"));
  $("btnHomeVertailu").addEventListener("click", () => {
    const s = settings();
    s.mode = "vuosi";
    saveSettings(s);
    openVertailu("home");
  });
  $("btnVertBack").addEventListener("click", () => {
    if (prevView === "overview" && state) renderOverview();
    else {
      renderHomeHistory();
      show("home");
    }
  });
  document.querySelectorAll('input[name="vertMode"]').forEach((r) =>
    r.addEventListener("change", () => {
      const s = settings();
      s.mode = r.value;
      saveSettings(s);
      renderVertailu();
    })
  );
  $("vertYear").addEventListener("change", () => {
    const s = settings();
    s.year = +$("vertYear").value;
    saveSettings(s);
    renderVertailu();
  });
  $("vertScale").addEventListener("change", () => {
    const s = settings();
    s.scale = $("vertScale").checked;
    saveSettings(s);
    renderVertailu();
  });
  $("vertThr").addEventListener("change", () => {
    const v = ShiftCalc.parseHM($("vertThr").value.trim());
    const s = settings();
    if (v != null) s.thr = v;
    saveSettings(s);
    renderVertailu();
  });
  $("vertThrN").addEventListener("change", () => {
    const v = parseFloat($("vertThrN").value.replace(",", "."));
    const s = settings();
    if (isFinite(v) && v >= 0) s.thrN = v;
    saveSettings(s);
    renderVertailu();
  });
  $("btnSaveHist").addEventListener("click", saveCurrentToHistory);
  $("pubCheck").addEventListener("change", (e) => {
    if (!state) return;
    state.published = e.target.checked;
    $("histDiff").innerHTML = "";
    $("histMsg").textContent = "";
    renderVertailu();
  });
  $("btnExport").addEventListener("click", exportCsv);
  $("btnPdf").addEventListener("click", printVertailu);
  $("csvInput").addEventListener("change", async (e) => {
    const f = e.target.files && e.target.files[0];
    if (!f) return;
    try {
      const n = importCsv(await f.text());
      $("histMsg").textContent = "Tuotu " + n + " jaksoa.";
    } catch (err) {
      $("histMsg").textContent = "Tuonti epäonnistui: " + (err && err.message ? err.message : err);
    }
    e.target.value = "";
    renderVertailu();
    renderHomeHistory();
  });
  $("startDateOk").addEventListener("click", () => {
    const v = $("startDateInput").value;
    const err = $("startDateErr");
    if (!v) {
      err.textContent = "Anna alkupäivä.";
      err.hidden = false;
      return;
    }
    if (!pendingBuf) return;
    try {
      useResult(ShiftCalc.parseArrayBuffer(pendingBuf, { startDate: v }), pendingName);
    } catch (e) {
      err.textContent = e && e.message ? e.message : String(e);
      err.hidden = false;
    }
  });
  renderHomeHistory();

  window.__ShiftApp = {
    loadArrayBuffer: (buf, opts) => {
      pendingBuf = buf;
      return useResult(ShiftCalc.parseArrayBuffer(buf, opts), "");
    },
    openDetail,
    openVertailu,
    getState: () => state,
    saveCurrentToHistory,
    importCsv,
    printVertailu,
  };
})();
