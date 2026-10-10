/*! © 2026 Lämpöpumppu Mafia. Kaikki oikeudet pidätetään. All rights reserved.
 * Tämän ohjelmiston kopioiminen, muokkaaminen tai jakaminen ilman lupaa on kielletty. */
/* global ShiftCalc, XLSX */
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const views = {
    home: $("view-home"),
    overview: $("view-overview"),
    detail: $("view-detail"),
  };

  let state = null;
  let summary = null;

  const WD_SHORT = {
    Maanantai: "Ma",
    Tiistai: "Ti",
    Keskiviikko: "Ke",
    Torstai: "To",
    Perjantai: "Pe",
    Lauantai: "La",
    Sunnuntai: "Su",
  };

  function show(view) {
    Object.keys(views).forEach((k) => {
      views[k].hidden = k !== view;
    });
    window.scrollTo(0, 0);
  }

  function fmt(mins) {
    return ShiftCalc.formatHM(mins);
  }

  /** ISO yyyy-mm-dd → d.m.yyyy */
  function dateFi(iso) {
    if (!iso) return "";
    const p = iso.split("-");
    if (p.length !== 3) return iso;
    return p[2] + "." + p[1] + "." + p[0];
  }

  // --- Overtime thresholds: defaults from ShiftCalc.OT_CONFIG, optional per-jakso override ---
  // Overrides live only in this device's localStorage, keyed by jakso start date.
  const OT_STORE_PREFIX = "nrtyokalut.otOverride.";

  function loadOtOverride() {
    if (!state) return null;
    try {
      const key = OT_STORE_PREFIX + state.startDate;
      const o = JSON.parse(localStorage.getItem(key) || "null");
      if (!o) return null;
      if (o.v !== 2) {
        // Old two-field format (before lisätyö was split out) — clear it
        localStorage.removeItem(key);
        return null;
      }
      if (o.lisaK == null && o.yliK == null && o.sataK == null) return null;
      return o;
    } catch (e) {
      return null;
    }
  }

  function saveOtOverride(o) {
    try {
      const key = OT_STORE_PREFIX + state.startDate;
      if (!o || (o.lisaK == null && o.yliK == null && o.sataK == null)) localStorage.removeItem(key);
      else localStorage.setItem(key, JSON.stringify(Object.assign({ v: 2 }, o)));
    } catch (e) {
      /* storage unavailable: override just won't persist */
    }
  }

  /** Effective thresholds (defaults + override); null when unknown (e.g. non-21-day jakso) */
  function effectiveOt() {
    return ShiftCalc.applyOvertimeOverride(state.overtime, loadOtOverride());
  }

  // Keskeytynyt jakso: laskentapäivät per person, on this device only.
  const LASK_PREFIX = "nrtyokalut.laskenta.";
  function laskKey(p) {
    return p.col + ":" + p.name;
  }
  function loadLaskMap() {
    if (!state) return {};
    try {
      return JSON.parse(localStorage.getItem(LASK_PREFIX + state.startDate) || "{}") || {};
    } catch (e) {
      return {};
    }
  }
  function saveLaskMap(map) {
    try {
      const key = LASK_PREFIX + state.startDate;
      if (!map || !Object.keys(map).length) localStorage.removeItem(key);
      else localStorage.setItem(key, JSON.stringify(map));
    } catch (e) {}
  }
  // Day marks: { [personKey]: { [yyyy-mm-dd]: "keskeytys" | "sairas" | "loma" } }.
  // The button stores "keskeytys" (no reason). "sairas"/"loma" are optional reasons for later.
  const MARK_PREFIX = "nrtyokalut.marks.";
  function loadMarkMap() {
    if (!state) return {};
    try {
      return JSON.parse(localStorage.getItem(MARK_PREFIX + state.startDate) || "{}") || {};
    } catch (e) {
      return {};
    }
  }
  function personMarks(p) {
    return loadMarkMap()[laskKey(p)] || {};
  }
  function saveMarkMap(map) {
    try {
      const key = MARK_PREFIX + state.startDate;
      if (!map || !Object.keys(map).length) localStorage.removeItem(key);
      else localStorage.setItem(key, JSON.stringify(map));
    } catch (e) {}
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
  /** Manual edit. Kept even when it equals the automatic value, until cleared. */
  function savePersonN(p, n) {
    const map = loadLaskMap();
    map[laskKey(p)] = n;
    saveLaskMap(map);
  }
  function clearPersonN(p) {
    const map = loadLaskMap();
    delete map[laskKey(p)];
    saveLaskMap(map);
  }
  function toggleSick(p, date) {
    if (personExtras(p)[date]) {
      dayFlash = { id: p.id, date: date, msg: CONFLICT_MSG };
      return false;
    }
    if (personTots(p)[date]) {
      dayFlash = { id: p.id, date: date, msg: TOT_CONFLICT_MSG, tot: true };
      return false;
    }
    dayFlash = null;
    const map = loadMarkMap();
    const key = laskKey(p);
    const cur = Object.assign({}, map[key] || {});
    if (ShiftCalc.isKeskeytys(cur[date])) delete cur[date];
    else cur[date] = "keskeytys";
    if (Object.keys(cur).length) map[key] = cur;
    else delete map[key];
    saveMarkMap(map);
    return true;
  }
  // Lisävuorot: { [personKey]: { [yyyy-mm-dd]: { start, end } } } minutes. On-device only.
  const EXTRA_PREFIX = "nrtyokalut.extra.";
  const CONFLICT_MSG = "Keskeytyspäivää ja lisävuoroa ei voi merkitä samalle päivälle.";
  let dayFlash = null;
  let openExtra = null;
  function loadExtraMap() {
    if (!state) return {};
    try {
      return JSON.parse(localStorage.getItem(EXTRA_PREFIX + state.startDate) || "{}") || {};
    } catch (e) {
      return {};
    }
  }
  function personExtras(p) {
    return loadExtraMap()[laskKey(p)] || {};
  }
  function saveExtraMap(map) {
    try {
      const key = EXTRA_PREFIX + state.startDate;
      if (!map || !Object.keys(map).length) localStorage.removeItem(key);
      else localStorage.setItem(key, JSON.stringify(map));
    } catch (e) {}
  }
  function saveExtra(p, date, start, end) {
    const map = loadExtraMap();
    const key = laskKey(p);
    const cur = Object.assign({}, map[key] || {});
    cur[date] = { start: start, end: end };
    map[key] = cur;
    saveExtraMap(map);
  }
  function deleteExtra(p, date) {
    const map = loadExtraMap();
    const key = laskKey(p);
    const cur = Object.assign({}, map[key] || {});
    delete cur[date];
    if (Object.keys(cur).length) map[key] = cur;
    else delete map[key];
    saveExtraMap(map);
  }
  function extraCount(p) {
    return Object.keys(personExtras(p)).length;
  }
  function extraMarker(p) {
    const n = extraCount(p);
    if (!n) return "";
    return '<p class="card-extra">' + (n === 1 ? "+1 lisävuoro" : "+" + n + " lisävuoroa") + "</p>";
  }
  function personRest(p) {
    return ShiftCalc.restWithExtras(p.days, personExtras(p), personMarks(p));
  }
  function sumPerson(p, key, tots) {
    return ShiftCalc.sumActive(p.days, personMarks(p), key, personExtras(p), state.holidayMap, tots || personTots(p));
  }
  function stewardN(p) {
    const days = ShiftCalc.OT_CONFIG.periodDays;
    return state.dayCount === days ? personN(p) : days;
  }
  function stewardMinutes(p) {
    return ShiftCalc.stewardBonus(isSteward(p), null, stewardN(p));
  }
  function activeLm(p, tots) {
    return sumPerson(p, "me", tots) + stewardMinutes(p);
  }
  function activeCompany(p) {
    return sumPerson(p, "company") + stewardMinutes(p);
  }
  function loadStewardMap() {
    try {
      return JSON.parse(localStorage.getItem(ShiftCalc.STEWARD_STORAGE_KEY) || "{}") || {};
    } catch (e) {
      return {};
    }
  }
  function isSteward(p) {
    return !!loadStewardMap()[p.name];
  }
  function setSteward(p, on) {
    const next = ShiftCalc.stewardMapSet(loadStewardMap(), p.name, on);
    try {
      const key = ShiftCalc.STEWARD_STORAGE_KEY;
      if (!Object.keys(next).length) localStorage.removeItem(key);
      else localStorage.setItem(key, JSON.stringify(next));
    } catch (e) {}
  }
  // TSV (työsuojeluvaltuutettu): remembered by name like LM, but adds NO hours; palkkio only.
  function loadTsvMap() {
    try {
      return JSON.parse(localStorage.getItem(ShiftCalc.TSV_STORAGE_KEY) || "{}") || {};
    } catch (e) {
      return {};
    }
  }
  function isTsv(p) {
    return !!loadTsvMap()[p.name];
  }
  function setTsv(p, on) {
    const next = ShiftCalc.stewardMapSet(loadTsvMap(), p.name, on);
    try {
      const key = ShiftCalc.TSV_STORAGE_KEY;
      if (!Object.keys(next).length) localStorage.removeItem(key);
      else localStorage.setItem(key, JSON.stringify(next));
    } catch (e) {}
  }
  function stewardLabel() {
    const m = ShiftCalc.stewardBonus(true);
    const h = m % 60 === 0 ? String(m / 60) : ShiftCalc.formatHM(m);
    return "LM (+" + h + " h / jakso)";
  }
  function stewardLine(p) {
    const days = ShiftCalc.OT_CONFIG.periodDays;
    const n = stewardN(p);
    const mins = ShiftCalc.stewardBonus(true, null, n);
    return "LM-tunnit +" + fmt(mins) + " (" + n + "/" + days + " pv)";
  }
  function personOt(p, tots) {
    if (!state.otEffective) return null;
    const lm = activeLm(p, tots);
    if (state.dayCount !== 21) return ShiftCalc.overtimeSplit(lm, state.otEffective);
    const n = personN(p);
    if (n < 1) {
      return { lisa: 0, yli50: 0, yli100: 0, n: 0, unsupported: null, cap50Hit: false, capLisaHit: false, norms: null };
    }
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
  /** Threshold line for this person's N, e.g. 'Taulukko 1 · 14 pv: 50 % yli 76:30 (enint. 12 h) · …' */
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
    t += "50 % yli " + fmt(norms.yliK) + " (enint. " + norms.cap50H + " h) · 100 % yli " + fmt(norms.k100);
    return t;
  }

  /** Overtime value or "—" when unknown / not in the table */
  function fmtOt(ot, key) {
    if (!ot || ot[key] == null) return "—";
    return fmt(ot[key]);
  }

  /** Short threshold text, e.g. '50 % yli 114:45 · 100 % yli 132:45 · Arkipyhät: …' */
  function overtimeText(th) {
    const def = state.overtime;
    if (!th) {
      return "Ylityörajat: ei oletusta " + state.dayCount + " pv jaksolle";
    }
    let t =
      (th.lisaK < th.yliK ? "Lisätyö yli " + fmt(th.lisaK) + " · " : "") +
      "50 % yli " +
      fmt(th.yliK) +
      " · 100 % yli " +
      fmt(th.sataK);
    if (def && def.holidays.length) {
      t +=
        " · Arkipyhät: " +
        def.holidays
          .map((h) => h.name + " " + dateFi(h.date).replace(/\d{4}$/, ""))
          .join(", ");
    }
    return t;
  }

  function renderOtMeta() {
    const th = state.otEffective;
    const ov = loadOtOverride() || {};
    const def = state.overtime;
    const val = (k) => (ov[k] != null ? fmt(ov[k]) : "");
    const ph = (k) => (def ? fmt(def[k]) : "h:mm");
    $("otMeta").innerHTML =
      '<span class="ot-text">' +
      overtimeText(th) +
      "</span>" +
      (th && th.overridden ? ' <span class="ot-mod">muutettu</span>' : "") +
      ' <button type="button" class="ot-edit-btn" id="otEditBtn">Muuta</button>' +
      '<div class="ot-editor" id="otEditor" hidden>' +
      '<label>Lisätyö alkaa<input id="otLisa" inputmode="decimal" autocomplete="off" placeholder="' +
      ph("lisaK") +
      '" value="' +
      val("lisaK") +
      '"></label>' +
      '<label>50 % alkaa<input id="ot50" inputmode="decimal" autocomplete="off" placeholder="' +
      ph("yliK") +
      '" value="' +
      val("yliK") +
      '"></label>' +
      '<label>100 % alkaa<input id="ot100" inputmode="decimal" autocomplete="off" placeholder="' +
      ph("sataK") +
      '" value="' +
      val("sataK") +
      '"></label>' +
      '<div class="ot-actions">' +
      '<button type="button" class="ot-save" id="otSave">Tallenna</button>' +
      '<button type="button" class="ot-reset" id="otReset">Palauta oletus</button>' +
      "</div>" +
      '<p class="ot-err" id="otErr" hidden></p>' +
      '<p class="ot-note">' +
      (def
        ? "Tyhjä kenttä = oletus. "
        : "Anna 50 % ja 100 % rajat (h:mm); tyhjä lisätyö = sama kuin 50 %. ") +
      "Tallentuu vain tähän laitteeseen.</p>" +
      "</div>";

    $("otEditBtn").addEventListener("click", () => {
      const ed = $("otEditor");
      ed.hidden = !ed.hidden;
      if (!ed.hidden) $("otLisa").focus();
    });
    $("otReset").addEventListener("click", () => {
      saveOtOverride(null);
      recalcOvertime();
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
      const fL = read("otLisa"),
        f50 = read("ot50"),
        f100 = read("ot100");
      if ([fL, f50, f100].some((f) => f.raw && f.v == null)) return err("Anna aika muodossa h:mm, esim. 106:45.");
      const same = (v, k) => (def && v === def[k] ? null : v);
      const o = { lisaK: same(fL.v, "lisaK"), yliK: same(f50.v, "yliK"), sataK: same(f100.v, "sataK") };
      const dk = (k) => (def ? def[k] : null);
      const e50 = f50.v != null ? f50.v : dk("yliK");
      const e100 = f100.v != null ? f100.v : dk("sataK");
      const eL = fL.v != null ? fL.v : def ? def.lisaK : e50;
      if (!def && (e50 == null || e100 == null) && (fL.v != null || f50.v != null || f100.v != null))
        return err("Anna ainakin 50 % ja 100 % rajat.");
      if (eL != null && e50 != null && eL > e50) return err("Lisätyö ei voi alkaa 50 % rajan jälkeen.");
      if (e50 != null && e100 != null && e100 < e50) return err("100 % raja ei voi olla pienempi kuin 50 % raja.");
      saveOtOverride(o);
      recalcOvertime();
    });
  }

  /** Recompute thresholds + all cards in place (keeps scroll position) */
  function recalcOvertime() {
    renderOverview(true);
  }

  /** e.g. To 29.10. */
  function dateFiShort(iso, weekday) {
    if (!iso) return "";
    const p = iso.split("-");
    const short = WD_SHORT[weekday] || "";
    return (short ? short + " " : "") + p[2] + "." + p[1] + ".";
  }

  function lab(s) {
    if (s === "Md,s Check") return "TARKISTA";
    if (s === "Not Allowed") return "EI SALLITTU";
    if (s === "ok") return "OK";
    return s || "";
  }

  function restClass(s) {
    if (s === "ok") return "cell-ok";
    if (s === "Md,s Check") return "cell-md";
    if (s === "Not Allowed") return "cell-bad";
    return "";
  }

  /**
   * Alerts from the SAME source as the overview grid: restBefore. (LM≠Yritys check moved to the Tarkastus app.)
   * One alert per flagged arrival day; Finnish date + actual rest duration.
   */
  function renderAlerts(target, result, personId) {
    const items = [];
    result.people.forEach((p) => {
      if (personId && p.id !== personId) return;
      const rest = personRest(p);
      p.days.forEach((d) => {
        if (isSickDay(p, d)) return;
        const r = rest[d.date] || {};
        if (r.restBefore === "Not Allowed" || r.restBefore === "Md,s Check") {
          const label = (WD_SHORT[d.weekday] || "") + " " + dateFi(d.date);
          const lepo = r.restBeforeMin != null ? fmt(r.restBeforeMin) : "";
          items.push({
            bad: r.restBefore === "Not Allowed",
            text:
              p.name +
              ": " +
              lab(r.restBefore) +
              " " +
              label +
              (lepo ? " (vuorojen väli " + lepo + ")" : ""),
          });
        }
        if (r.between === "Not Allowed" || r.between === "Md,s Check") {
          items.push({
            bad: r.between === "Not Allowed",
            text: p.name + ": " + lab(r.between) + " " + dateFi(d.date) + " (lisävuoron väli " + fmt(r.betweenMin) + ")",
          });
        }
      });
    });
    if (!items.length) {
      target.hidden = true;
      target.innerHTML = "";
      return;
    }
    const hasBad = items.some((i) => i.bad);
    target.hidden = false;
    target.className = "alerts" + (hasBad ? " bad" : "");
    target.innerHTML =
      "<h3>" +
      (hasBad ? "Huomio: ongelmia löytyi" : "Huomioita") +
      "</h3><ul>" +
      items.map((i) => "<li>" + i.text + "</li>").join("") +
      "</ul>";
  }

  function personProblems(p) {
    const problems = [];
    const rest = personRest(p);
    p.days.forEach((d) => {
      if (isSickDay(p, d)) return;
      const r = rest[d.date] || {};
      if (r.restBefore === "Not Allowed" || r.restBefore === "Md,s Check") {
        const label = (WD_SHORT[d.weekday] || "") + " " + dateFi(d.date).replace(/\.\d{4}$/, ".");
        const lepo = r.restBeforeMin != null ? fmt(r.restBeforeMin) : "";
        problems.push({
          kind: r.restBefore === "Not Allowed" ? "bad" : "md",
          text:
            lab(r.restBefore) +
            " " +
            label +
            (lepo ? " (" + lepo + ")" : ""),
        });
      }
      if (r.between === "Not Allowed" || r.between === "Md,s Check") {
        problems.push({
          kind: r.between === "Not Allowed" ? "bad" : "md",
          text: lab(r.between) + " " + dateFi(d.date).replace(/\.\d{4}$/, ".") + " (lisävuoron väli " + fmt(r.betweenMin) + ")",
        });
      }
    });
    return problems;
  }

  function personStatus(problems) {
    if (problems.some((x) => x.kind === "bad")) return "bad";
    if (problems.length) return "md";
    return "ok";
  }

  function breakMarker(p) {
    if (state.dayCount !== 21) return "";
    const n = personN(p);
    if (n === 21) return "";
    return '<p class="card-break">Keskeytynyt · ' + n + " pv</p>";
  }

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
    $("periodMeta").textContent =
      (state.periodLabel ? state.periodLabel + " · " : "") +
      "Alku " +
      dateFi(state.startDate);

    renderOtMeta();

    const titleEl = $("overviewTitle");
    if (titleEl) titleEl.textContent = "Kaikki " + people.length + " henkilöä";

    let nBad = 0,
      nMd = 0,
      nOk = 0;
    const cards = people.map((p) => {
      const problems = personProblems(p);
      const status = personStatus(problems);
      if (status === "bad") nBad++;
      else if (status === "md") nMd++;
      else nOk++;

      const sum = (key) => sumPerson(p, key);
      const hrs = sum("hrs");
      const me = sum("me");
      const night = sum("night");
      const b25 = sum("b25");
      const b100 = sum("b100");
      const badge =
        status === "bad"
          ? '<span class="pill bad">EI SALLITTU</span>'
          : status === "md"
            ? '<span class="pill md">TARKISTA</span>'
            : '<span class="pill ok">OK</span>';

      const problemList = problems.length
        ? '<ul class="card-problems">' +
          problems
            .map(
              (pr) =>
                '<li class="prob-' + pr.kind + '">' + pr.text + "</li>"
            )
            .join("") +
          "</ul>"
        : '<p class="card-ok-line">Ei huomautuksia</p>';

      return (
        '<button type="button" class="person-card status-' +
        status +
        '" data-id="' +
        p.id +
        '">' +
        '<div class="person-card-top">' +
        "<div><h3>" +
        p.name +
        (isSteward(p) ? ' <span class="lm-tag">LM</span>' : "") +
        "</h3>" +
        '<p class="card-meta">' +
        p.shiftCount +
        " vuoroa</p></div>" +
        badge +
        "</div>" +
        '<div class="card-hours">' +
        "<div><span>Kovat tunnit</span><b>" +
        fmt(hrs) +
        "</b></div>" +
        "<div><span>Tunnit yhteensä</span><b>" +
        fmt(activeLm(p)) +
        "</b></div>" +
        "<div><span>Yö h</span><b>" +
        fmt(night) +
        "</b></div>" +
        "<div><span>Lauantai h</span><b>" +
        fmt(b25) +
        "</b></div>" +
        "<div><span>Pyhä h</span><b>" +
        fmt(b100) +
        "</b></div>" +
        "<div><span>Lisätyö h</span><b>" +
        fmtOt(personOt(p), "lisa") +
        "</b></div>" +
        "<div><span>Ylityö 50 %</span><b>" +
        fmtOt(personOt(p), "yli50") +
        "</b></div>" +
        "<div><span>Ylityö 100 %</span><b>" +
        fmtOt(personOt(p), "yli100") +
        "</b></div>" +
        "</div>" +
        breakMarker(p) +
        extraMarker(p) +
        (otNote(personOt(p)) ? '<p class="card-ot-note">' + otNote(personOt(p)) + "</p>" : "") +
        problemList +
        '<span class="card-open">Avaa vuorotaulu →</span>' +
        "</button>"
      );
    });

    $("personCards").innerHTML = cards.join("");
    $("personCards").querySelectorAll(".person-card").forEach((btn) => {
      btn.addEventListener("click", () => openDetail(+btn.dataset.id));
    });

    const bits = [];
    if (nBad) bits.push(nBad + " EI SALLITTU");
    if (nMd) bits.push(nMd + " TARKISTA");
    if (nOk) bits.push(nOk + " OK");
    $("overviewSummary").textContent =
      people.reduce((a, p) => a + p.shiftCount, 0) +
      " vuoroa · " +
      bits.join(" · ");
    renderGroupNote();

    if (keepScroll) {
      const y = window.scrollY;
      show("overview");
      window.scrollTo(0, y);
    } else {
      show("overview");
    }
  }


  function extraBits(p, date) {
    const ex = personExtras(p)[date];
    const fig = ex ? ShiftCalc.computeShiftFigures(date, ex.start, ex.end, state.holidayMap) : null;
    const open = openExtra && openExtra.id === p.id && openExtra.date === date;
    const flash = dayFlash && !dayFlash.tot && dayFlash.id === p.id && dayFlash.date === date ? dayFlash.msg : "";
    const btn =
      '<button type="button" class="extra-toggle" data-date="' +
      date +
      '">' +
      (fig ? "Lisävuoro ✓" : "+ Lisävuoro") +
      "</button>";
    let block = "";
    const dd = p.days.find((x) => x.date === date);
    const ignored = fig && dd && ShiftCalc.extraOverlaps(dd, ex, personTots(p));
    if (fig) {
      block +=
        '<div class="extra-shift"><div class="extra-head"><span class="tag extra">Lisävuoro</span> <b>' +
        fmt(fig.start) +
        "–" +
        fmt(fig.end) +
        '</b></div><div class="day-grid">' +
        "<div><span>Kovat tunnit</span><b>" + fmt(fig.hrs) + "</b></div>" +
        "<div><span>Yö h</span><b>" + fmt(fig.night) + "</b></div>" +
        "<div><span>Tunnit yhteensä</span><b>" + fmt(fig.me) + "</b></div>" +
        "<div><span>Lauantai h</span><b>" + fmt(fig.b25) + "</b></div>" +
        "<div><span>Pyhä h</span><b>" + fmt(fig.b100) + "</b></div>" +
        "</div>" +
        (ignored ? '<p class="extra-note">Lisävuoro on päällekkäin toteutuneen vuoron kanssa, joten sitä ei lasketa. Tunnit lasketaan vain kerran.</p>' : "") +
        "</div>";
    }
    const editor =
      '<div class="extra-editor"' +
      (open ? "" : " hidden") +
      ' data-date="' +
      date +
      '">' +
      '<label>Alku <input class="ex-start" inputmode="decimal" value="' +
      (fig ? fmt(fig.start) : "") +
      '" placeholder="8:00"></label>' +
      '<label>Loppu <input class="ex-end" inputmode="decimal" value="' +
      (fig ? fmt(fig.end) : "") +
      '" placeholder="16:00"></label>' +
      '<div class="ot-actions"><button type="button" class="ot-save ex-save" data-date="' +
      date +
      '">Tallenna</button>' +
      (fig ? '<button type="button" class="ot-reset ex-del" data-date="' + date + '">Poista</button>' : "") +
      "</div>" +
      (flash && open ? '<p class="extra-err">' + flash + "</p>" : "") +
      "</div>" +
      (flash && !open ? '<p class="extra-err">' + flash + "</p>" : "");
    return { btn: btn, body: block + editor, fig: fig };
  }

  // --- Toteuma (actual times) + Lisät (TES) card ---
  // { [personKey]: { [yyyy-mm-dd]: toteuma } } — see ShiftCalc.lisatSummary. On-device only.
  const TOT_PREFIX = "nrtyokalut.toteuma.";
  const TOT_CONFLICT_MSG = "Keskeytyspäivää ja toteumaa ei voi merkitä samalle päivälle.";
  let openTot = null;
  function loadTotMap() {
    if (!state) return {};
    try {
      return JSON.parse(localStorage.getItem(TOT_PREFIX + state.startDate) || "{}") || {};
    } catch (e) {
      return {};
    }
  }
  function personTots(p) {
    return loadTotMap()[laskKey(p)] || {};
  }
  function saveTot(p, date, value) {
    const map = loadTotMap();
    const key = laskKey(p);
    const cur = Object.assign({}, map[key] || {});
    if (value) cur[date] = value;
    else delete cur[date];
    if (Object.keys(cur).length) map[key] = cur;
    else delete map[key];
    try {
      const k = TOT_PREFIX + state.startDate;
      if (!Object.keys(map).length) localStorage.removeItem(k);
      else localStorage.setItem(k, JSON.stringify(map));
    } catch (e) {}
  }
  function personLisat(p) {
    return ShiftCalc.lisatSummary(p.days, personMarks(p), personExtras(p), personTots(p), state.holidayMap);
  }
  function eur(x) {
    return x.toFixed(2).replace(".", ",") + " €";
  }
  const FIXED_LABEL = { poikkeama: "Poikkeamakorvaus", vapaa: "Vapaa-ajan poikkeamakorvaus", peruutus: "Työvuoron peruutuskorvaus", peruutusTunnit: "Peruttu vuoro" };

  /** Editor choice ↔ stored toteuma */
  function totChoice(t) {
    if (!t) return "";
    if (t.type === "muutos") return t.oma ? "oma" : "muutos";
    if (t.type === "peruttu") return t.late ? "peruttu_late" : "peruttu_ok";
    return t.type;
  }
  function totSummaryText(t) {
    if (!t) return "";
    if (t.type === "muutos") return "Toteutui " + fmt(t.start) + "–" + fmt(t.end) + (t.oma ? " (omasta pyynnöstä)" : "");
    if (t.type === "korvattu") return "Peruttu, tilalle " + fmt(t.start) + "–" + fmt(t.end);
    if (t.type === "vaihto") return "Vaihdettu vuoro työkaverin kanssa, ajettu " + fmt(t.start) + "–" + fmt(t.end);
    if (t.type === "peruttu")
      return t.late
        ? "Peruttu klo 17 jälkeen, ei uutta vuoroa · " + (t.korvaus ? "valittu peruutuskorvaus, tunnit eivät lasketa" : "valittu tunnit (paikallinen sopimus)")
        : "Peruttu ajoissa";
    if (t.type === "kutsu") {
      const a = t.aStart != null ? t.aStart : t.start;
      const b = t.aEnd != null ? t.aEnd : t.end;
      return "Kutsuttu vapaapäivänä " + fmt(t.start) + "–" + fmt(t.end) + (a !== t.start || b !== t.end ? " · toteutui " + fmt(a) + "–" + fmt(b) : "");
    }
    return "";
  }

  function totBits(p, d, lisat) {
    const t = ShiftCalc.normToteuma(personTots(p)[d.date]);
    const planned = d.start != null;
    const open = openTot && openTot.id === p.id && openTot.date === d.date;
    const flash = dayFlash && dayFlash.id === p.id && dayFlash.date === d.date && dayFlash.tot ? dayFlash.msg : "";
    const btn =
      '<button type="button" class="tot-toggle' + (t ? " on" : "") + '" data-date="' + d.date + '">' +
      (t ? "Toteuma ✓" : "Toteuma") + "</button>";
    let line = "";
    if (t && !(t.type === "kutsu" && planned)) {
      const ev = lisat.events.filter((e) => e.date === d.date);
      const c = ShiftCalc.countedShift(d, personTots(p));
      const fig = c && c.kind !== "plan" ? ShiftCalc.computeShiftFigures(d.date, c.start, c.end, state.holidayMap) : null;
      const korv = c && c.kind === "korvattu" && fig;
      const hrsTxt = korv
        ? "<br/>Lasketaan vähintään suunnitellun mukaan: kovat " + fmt(Math.max(d.hrs, fig.hrs)) + " · tunnit yht. " + fmt(Math.max(d.me, fig.me)) + " (suunniteltu " + fmt(d.me) + ", uusi vuoro " + fmt(fig.me) + "). Lisät lajeittain suuremman mukaan, km uuden vuoron mukaan."
        : fig
        ? "<br/>Lasketaan: kovat " + fmt(fig.hrs) + " · tunnit yht. " + fmt(fig.me) + (planned ? " (suunniteltu " + fmt(d.me) + ")" : "")
        : t.type === "vaihto" ? "<br/>Tunnit oman suunnitellun vuoron mukaan (kovat " + fmt(d.hrs) + " · tunnit yht. " + fmt(d.me) + "), lisät ajetun vuoron mukaan. Ei poikkeamakorvausta."
        : t.type === "peruttu" ? (t.late && t.korvaus ? "<br/>Tunnit eivät lasketa (tunnit yht. 0:00, suunniteltu " + fmt(d.me) + ")" : "<br/>Suunnitellut tunnit lasketaan (tunnit yht. " + fmt(d.me) + ")") : "";
      line =
        '<div class="tot-line"><span class="tag tot">Toteuma</span> ' + totSummaryText(t) + hrsTxt +
        (ev.filter((e) => e.eur).length ? "<br/>" + ev.filter((e) => e.eur).map((e) => "+ " + FIXED_LABEL[e.key] + " " + eur(e.eur)).join("<br/>") : "") +
        "</div>";
    }
    const c = totChoice(t);
    const opt = (v, label) => '<option value="' + v + '"' + (c === v ? " selected" : "") + ">" + label + "</option>";
    const val = (x) => (x != null ? fmt(x) : "");
    const timed = t && (t.type === "muutos" || t.type === "korvattu" || t.type === "vaihto");
    const editor =
      '<div class="tot-editor"' + (open ? "" : " hidden") + ' data-date="' + d.date + '">' +
      '<label class="tot-wide">Mitä tapahtui?<select class="tot-type">' +
      (planned
        ? opt("", "Toteutui suunnitellusti (" + fmt(d.start) + "–" + fmt(d.end) + ")") +
          opt("muutos", "Aika muuttui (työnjohdon määräys)") +
          opt("oma", "Oman vuoron aika muuttui omasta pyynnöstä") +
          opt("vaihto", "Vaihdettu vuoro (työkaverin kanssa)") +
          opt("korvattu", "Peruttu, tilalle annettiin toinen vuoro") +
          opt("peruttu_late", "Peruttu klo 17 jälkeen edellisenä päivänä, ei uutta vuoroa") +
          opt("peruttu_ok", "Peruttu ajoissa (ennen klo 17)")
        : opt("", "Vapaapäivä") + opt("kutsu", "Kutsuttu vapaapäivänä")) +
      "</select></label>" +
      (planned
        ? '<p class="tot-help" data-for="oma">Oma vuorosi alkoi tai päättyi eri aikaan pyynnöstäsi. Toteutuneet tunnit lasketaan kovina tunteina, tunteina yhteensä ja ylitöinä. Poikkeamakorvausta ei makseta.</p>' +
          '<p class="tot-help" data-for="korvattu">Vuoro peruttiin ja tilalle annettiin toinen vuoro. Anna uuden vuoron alku ja loppu. Tunnit ja lisät vähintään suunnitellun vuoron mukaan, km uuden vuoron mukaan. Kovat tunnit, Yö h ja Tunnit yhteensä ovat suunnitellun (' + fmt(d.start) + "–" + fmt(d.end) + ') ja uuden vuoron suurempi. Ilta-, yö-, la-, su- ja aattolisät lasketaan lajeittain suuremman mukaan. Poikkeamakorvaus, jos uusi vuoro alkaa aiemmin tai päättyy vähintään 30 min myöhemmin. Ei vapaa-ajan korvausta.</p>' +
          '<p class="tot-help" data-for="vaihto">Ajoit työkaverin vuoron (pidemmän tai lyhyemmän) ja hän ajoi sinun. Anna ajetun vuoron alku ja loppu. Kovat tunnit, Yö h ja Tunnit yhteensä lasketaan oman suunnitellun vuorosi mukaan (' + fmt(d.start) + "–" + fmt(d.end) + ', tunnit yhteensä ' + fmt(d.me) + '), joten vaihto ei lisää eikä vähennä ylitöitä (Lisälehti 11). Ilta-, yö-, la-, su- ja aattolisät lasketaan ajetun vuoron mukaan. Ei poikkeama- eikä vapaa-ajan korvausta.</p>'
        : '<p class="tot-help" data-for="">Vaihtoa vapaapäivälle ei voi merkitä. Merkitse vaihto sille päivälle, jolla oma vuorosi oli.</p>') +
      (planned
        ? '<label data-for="muutos oma korvattu vaihto">Toteutunut alku <input class="tot-start" inputmode="decimal" value="' + (timed ? val(t.start) : "") + '" placeholder="' + fmt(d.start) + '"></label>' +
          '<label data-for="muutos oma korvattu vaihto">Toteutunut loppu <input class="tot-end" inputmode="decimal" value="' + (timed ? val(t.end) : "") + '" placeholder="' + fmt(d.end) + '"></label>' +
          peruutusChoice(p, d, t)
        : '<label data-for="kutsu">Ilmoitettu alku <input class="tot-start" inputmode="decimal" value="' + (t && t.type === "kutsu" ? val(t.start) : "") + '" placeholder="8:00"></label>' +
          '<label data-for="kutsu">Ilmoitettu loppu <input class="tot-end" inputmode="decimal" value="' + (t && t.type === "kutsu" ? val(t.end) : "") + '" placeholder="16:00"></label>' +
          '<label data-for="kutsu">Toteutunut alku (jos eri) <input class="tot-astart" inputmode="decimal" value="' + (t && t.type === "kutsu" ? val(t.aStart) : "") + '"></label>' +
          '<label data-for="kutsu">Toteutunut loppu (jos eri) <input class="tot-aend" inputmode="decimal" value="' + (t && t.type === "kutsu" ? val(t.aEnd) : "") + '"></label>') +
      '<div class="ot-actions"><button type="button" class="ot-save tot-save" data-date="' + d.date + '">Tallenna</button>' +
      (t ? '<button type="button" class="ot-reset tot-del" data-date="' + d.date + '">Poista</button>' : "") +
      "</div>" +
      (flash && open ? '<p class="extra-err">' + flash + "</p>" : "") +
      "</div>" +
      (flash && !open ? '<p class="extra-err">' + flash + "</p>" : "");
    return { btn: btn, body: line + editor, t: t };
  }

  /** Late cancellation: hours (default) OR peruutuskorvaus, with the overtime effect of each. */
  function peruutusChoice(p, d, t) {
    const korv = !!(t && t.type === "peruttu" && t.late && t.korvaus);
    const rate = ShiftCalc.fixedRatesFor(d.date).peruutus;
    const otFor = (korvaus) => {
      const tots = Object.assign({}, personTots(p));
      tots[d.date] = { type: "peruttu", late: true, korvaus: korvaus };
      return personOt(p, tots);
    };
    let hint = "";
    const a = otFor(false), b = otFor(true);
    if (a && b && !a.unsupported && !b.unsupported) {
      const txt = (o) => "lisätyö " + fmt(o.lisa) + " · 50 % " + fmt(o.yli50) + " · 100 % " + fmt(o.yli100);
      const same = a.lisa === b.lisa && a.yli50 === b.yli50 && a.yli100 === b.yli100;
      hint = same
        ? "Jakson ylityöt eivät muutu (" + txt(a) + ")."
        : "Jakson ylityöt: tunnit → " + txt(a) + "; korvaus → " + txt(b) + ".";
    }
    return (
      '<fieldset class="tot-choice" data-for="peruttu_late"><legend>Valitse jompikumpi</legend>' +
      '<label><input type="radio" name="pk-' + d.date + '" class="tot-korv" value="0"' + (korv ? "" : " checked") + "> Tunnit lasketaan (paikallinen sopimus)<small>Suunnitellut tunnit (yhteensä " + fmt(d.me) + ") lasketaan, ei peruutuskorvausta eikä ilta-/yö-/la-/su-lisiä.</small></label>" +
      '<label><input type="radio" name="pk-' + d.date + '" class="tot-korv" value="1"' + (korv ? " checked" : "") + "> Työvuoron peruutuskorvaus " + eur(rate) + " (tunnit eivät lasketa)<small>Koodi 1313. Vuoron tunnit 0:00 kovissa tunneissa, tunneissa yhteensä ja ylitöissä.</small></label>" +
      (hint ? '<p class="tot-hint">' + hint + "</p>" : "") +
      "</fieldset>"
    );
  }

  function syncTotEditor(box) {
    const v = box.querySelector(".tot-type").value;
    box.querySelectorAll("label[data-for], fieldset[data-for], p[data-for]").forEach((l) => {
      l.hidden = l.dataset.for.split(" ").indexOf(v) < 0;
    });
  }

  function readTotEditor(box, d) {
    const v = box.querySelector(".tot-type").value;
    const get = (cls) => {
      const el = box.querySelector(cls);
      const s = el ? el.value.trim() : "";
      return s === "" ? null : ShiftCalc.parseHM(s);
    };
    const bad = (cls) => {
      const el = box.querySelector(cls);
      return el && el.value.trim() !== "" && ShiftCalc.parseHM(el.value.trim()) == null;
    };
    if (v === "") return { value: null };
    if (v === "peruttu_ok") return { value: { type: "peruttu", late: false } };
    if (v === "peruttu_late") {
      const r = box.querySelector(".tot-korv:checked");
      const o = { type: "peruttu", late: true };
      if (r && r.value === "1") o.korvaus = true;
      return { value: o };
    }
    if ([".tot-start", ".tot-end", ".tot-astart", ".tot-aend"].some(bad)) return { err: "Anna ajat muodossa h:mm." };
    const a = get(".tot-start");
    const b = get(".tot-end");
    if (a == null || b == null) return { err: v === "kutsu" ? "Anna ilmoitettu alku ja loppu." : "Anna toteutunut alku ja loppu." };
    if (a === b) return { err: "Alku ja loppu eivät voi olla samat." };
    if (v === "muutos" || v === "oma") {
      if (a === d.start && b === d.end) return { value: null };
      return { value: { type: "muutos", start: a, end: b, oma: v === "oma" } };
    }
    if (v === "korvattu") return { value: { type: "korvattu", start: a, end: b } };
    if (v === "vaihto") return { value: { type: "vaihto", start: a, end: b } };
    const o = { type: "kutsu", start: a, end: b };
    const as = get(".tot-astart");
    const ae = get(".tot-aend");
    if (as != null) o.aStart = as;
    if (ae != null) o.aEnd = ae;
    if ((as != null ? as : a) === (ae != null ? ae : b)) return { err: "Alku ja loppu eivät voi olla samat." };
    return { value: o };
  }

  function palkkioHtml(p) {
    const lm = isSteward(p);
    const tsv = isTsv(p);
    if (!lm && !tsv) return "";
    const endKey = state.dates && state.dates.length ? state.dates[state.dates.length - 1] : state.startDate;
    const rates = ShiftCalc.palkkioRatesInRange(state.startDate, endKey);
    const rateTxt = rates
      .map((r, i) => eur(r.eur) + "/kk" + (i ? " " + dateFi(r.from).slice(0, 6) + " alkaen" : ""))
      .join(", ");
    const row = (name, code, small) =>
      "<tr><td><b>" + name + "</b><small>" + small + '</small></td><td>–</td><td></td><td class="lisat-pay">' + rateTxt + "</td></tr>";
    return (
      '<table class="lisat-table palkkio-table"><thead><tr><th>Palkkiot (€/kk)</th><th>Koodi</th><th></th><th>€/kk</th></tr></thead><tbody>' +
      (lm ? row("Luottamusmiespalkkio", "", "§18.17 · maksetaan 10 kk/vuosi (varamiehelle 2 kk)") : "") +
      (tsv ? row("Työsuojeluvaltuutetun palkkio", "", "§21 · maksetaan 12 kk/vuosi (varahenkilölle 2 kk: huhti- ja lokakuu) · ei lisätunteja") : "") +
      "</tbody></table>" +
      '<p class="ot-meta lisat-note">Palkkiot ovat kuukausikohtaisia, eivät jaksokohtaisia: jakso ja kalenterikuukausi eivät osu yhteen, joten vertaa niitä kuukauden palkkalaskelmaan.</p>' +
      (lm && tsv
        ? '<p class="palkkio-warn">Huom: TES §21:n mukaan luottamusmiespalkkiota ja työsuojeluvaltuutetun palkkiota ei makseta samalle henkilölle samanaikaisesti. Jos olet esimerkiksi LM ja vara-TSV, TSV-palkkio maksetaan vain niiltä kuukausilta, joina toimit TSV:n sijaisena.</p>'
        : "")
    );
  }

  // --- Junat (veturiraha) + autolla-ajo ---
  // Junat: { [personKey]: { [yyyy-mm-dd]: [{ junanumero, paino, km, yksin?, hidas?, ivyvak?, veturina? }] } }
  // yksin is stored only as a manual override; missing = auto from the sheet (same planned shift as someone → kaksinajo).
  // veturina: true = veturina ajo (veturi ilman junaa), km only, no paino.
  // Same shape a future Railcube import can write: one array of trains per day. paino = todellinen
  // jarrupainojärjestelmän kokonaisjunapaino (t), km = lähtö- ja tulopaikan välinen matka.
  // Autolla-ajo is kept in its own key so a Railcube train import can't wipe it:
  // { [personKey]: { [yyyy-mm-dd]: [{ tyyppi: "paikallis" | "ulko", km }] } }, one entry per one-way trip.
  // All on-device only.
  const JUNA_PREFIX = "nrtyokalut.junat.";
  const AUTOAJO_PREFIX = "nrtyokalut.autoajo.";
  let openJuna = null;
  let junaUid = 0;
  function escHtml(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
  }
  function loadDayListMap(prefix) {
    if (!state) return {};
    try {
      return JSON.parse(localStorage.getItem(prefix + state.startDate) || "{}") || {};
    } catch (e) {
      return {};
    }
  }
  function saveDayList(prefix, p, date, list) {
    const map = loadDayListMap(prefix);
    const key = laskKey(p);
    const cur = Object.assign({}, map[key] || {});
    if (list && list.length) cur[date] = list;
    else delete cur[date];
    if (Object.keys(cur).length) map[key] = cur;
    else delete map[key];
    try {
      const k = prefix + state.startDate;
      if (!Object.keys(map).length) localStorage.removeItem(k);
      else localStorage.setItem(k, JSON.stringify(map));
    } catch (e) {}
  }
  function personJunat(p) {
    return loadDayListMap(JUNA_PREFIX)[laskKey(p)] || {};
  }
  function saveJunat(p, date, list) {
    saveDayList(JUNA_PREFIX, p, date, list);
  }
  function personAutoajot(p) {
    return loadDayListMap(AUTOAJO_PREFIX)[laskKey(p)] || {};
  }
  function saveAutoajot(p, date, list) {
    saveDayList(AUTOAJO_PREFIX, p, date, list);
  }
  function pairInfo(p, date) {
    return ShiftCalc.pairPartners(state.people, p, date);
  }
  function pairHint(info) {
    if (!info.planned) return "Yksinajo: ei suunniteltua vuoroa tänä päivänä";
    return info.yksin ? "Yksinajo: ei samaa vuoroa" : "Kaksinajo: sama vuoro kuin " + info.partners.join(", ");
  }
  function personVeturiraha(p) {
    return ShiftCalc.veturirahaSummary(
      personJunat(p),
      (date) => isSickDay(p, { date: date }),
      (date) => pairInfo(p, date).yksin
    );
  }
  function personAutoajo(p) {
    return ShiftCalc.autoajoSummary(personAutoajot(p), (date) => isSickDay(p, { date: date }));
  }
  function numFi(x) {
    return x == null ? "" : String(Math.round(x * 1000) / 1000).replace(".", ",");
  }
  function rateFi(x) {
    return String(x).replace(".", ",");
  }
  function tonnit(x) {
    return x == null ? "?" : Math.round(x).toLocaleString("fi-FI").replace(/\u00a0/g, " ");
  }
  /** info = pairInfo(): the auto default; a saved j.yksin (true/false) is a manual choice and wins. */
  function junaRowHtml(j, info) {
    const v = j || {};
    const manual = v.yksin === true || v.yksin === false;
    const on = manual ? v.yksin : info.yksin;
    const uid = ++junaUid;
    return (
      '<div class="juna-row' + (v.veturina ? " is-veturina" : "") + '" data-auto="' + (info.yksin ? 1 : 0) + '">' +
      '<label>Junanumero <input class="j-num" value="' + escHtml(v.junanumero || "") + '" placeholder="esim. 3345"></label>' +
      '<label class="j-paino-l">Paino (t) <input class="j-paino" inputmode="decimal" value="' + numFi(v.paino) + '" placeholder="todellinen"></label>' +
      '<label>Km <input class="j-km" inputmode="decimal" value="' + numFi(v.km) + '" placeholder="lähtö–tulo"></label>' +
      '<div class="juna-flags">' +
      '<span class="ajo-toggle" role="radiogroup" aria-label="Yksin- vai kaksinajo">' +
      '<label><input type="radio" class="j-ajo" name="ajo-' + uid + '" value="yksin"' + (on ? " checked" : "") + "><span>Yksinajo</span></label>" +
      '<label><input type="radio" class="j-ajo" name="ajo-' + uid + '" value="kaksin"' + (on ? "" : " checked") + "><span>Kaksinajo</span></label>" +
      "</span>" +
      '<label><input type="checkbox" class="j-veturina"' + (v.veturina ? " checked" : "") + "> Veturina ajo (ilman junaa)</label>" +
      '<label class="j-train-only"><input type="checkbox" class="j-hidas"' + (v.hidas ? " checked" : "") + "> Hidas (enint. 40 km/h)</label>" +
      '<label class="j-train-only"><input type="checkbox" class="j-ivyvak"' + (v.ivyvak ? " checked" : "") + "> IVY-VAK</label>" +
      "</div>" +
      '<small class="j-hint">' + escHtml(junaHintText(info, manual ? v.yksin : null)) + "</small>" +
      '<button type="button" class="j-remove" aria-label="Poista rivi">✕</button>' +
      "</div>"
    );
  }
  function autoRowHtml(a) {
    const v = a || { tyyppi: "paikallis" };
    const opt = (k, label) => '<option value="' + k + '"' + (v.tyyppi === k ? " selected" : "") + ">" + label + "</option>";
    return (
      '<div class="juna-row auto-row">' +
      '<label class="a-type-l">Ajo <select class="a-type">' +
      opt("paikallis", "Paikallisajo") +
      opt("ulko", "Ulkopuolinen, 1 suunta") +
      "</select></label>" +
      '<label>Km <input class="a-km" inputmode="decimal" value="' + numFi(v.km) + '" placeholder="todellinen"></label>' +
      '<button type="button" class="j-remove" aria-label="Poista rivi">✕</button>' +
      "</div>"
    );
  }
  function junaHintText(info, manualYksin) {
    if (manualYksin == null || manualYksin === info.yksin) return pairHint(info);
    return "Valittu käsin: " + (manualYksin ? "Yksinajo" : "Kaksinajo") + " (vuorotaulun mukaan " + (info.yksin ? "Yksinajo" : "Kaksinajo") + ")";
  }
  function junaBits(p, d, vr, aa) {
    const info = pairInfo(p, d.date);
    const list = (personJunat(p)[d.date] || []).map(ShiftCalc.normJuna).filter(Boolean);
    const autos = (personAutoajot(p)[d.date] || []).map(ShiftCalc.normAutoajo).filter(Boolean);
    const n = list.length + autos.length;
    const open = openJuna && openJuna.id === p.id && openJuna.date === d.date;
    const btn =
      '<button type="button" class="juna-toggle' + (n ? " on" : "") + '" data-date="' + d.date + '">' +
      (n ? "Ajot ✓ " + n : "+ Juna") + "</button>";
    let line = "";
    if (n) {
      const counted = vr.trains.filter((t) => t.date === d.date);
      const cAuto = aa.trips.filter((t) => t.date === d.date);
      const dayEur = counted.reduce((s, t) => s + t.eur, 0) + cAuto.reduce((s, t) => s + t.eur, 0);
      line =
        '<div class="juna-line"><span class="tag juna">Ajot</span> ' +
        list
          .map((j, i) => {
            const t = counted.find((x) => x.i === i);
            const yk = j.yksin != null ? j.yksin : info.yksin;
            const who = "<b>" + (yk ? "Yksinajo" : "Kaksinajo") + "</b>" + (j.yksin != null && j.yksin !== info.yksin ? " (käsin)" : "");
            const flags = [j.hidas && !j.veturina ? "hidas" : "", j.ivyvak && !j.veturina ? "IVY-VAK" : ""].filter(Boolean).join(", ");
            const what = j.veturina
              ? "veturina ajo" + (j.junanumero ? " " + escHtml(j.junanumero) : "")
              : escHtml(j.junanumero || "Juna " + (i + 1)) + " · " + tonnit(j.paino) + " t";
            return (
              "<br/>" + who + " · " + what + " · " + (j.km != null ? numFi(j.km) : "?") + " km" + (flags ? " · " + flags : "") +
              (t ? " → " + t.code + " · " + numFi(j.km) + " × " + rateFi(t.rate) + " = <b>" + eur(t.eur) + "</b>" : " → ei lasketa")
            );
          })
          .join("") +
        autos
          .map((a, i) => {
            const t = cAuto.find((x) => x.i === i);
            return (
              "<br/>Autolla: " + ShiftCalc.AUTOAJO_LABELS[a.tyyppi].toLowerCase() + " · " + numFi(a.km) + " km" +
              (t ? " → " + (t.paidKm !== a.km ? "väh. " + ShiftCalc.AUTOAJO_MIN_KM + " km: " : "") + numFi(t.paidKm) + " × " + rateFi(t.rate) + " = <b>" + eur(t.eur) + "</b>" : " → ei lasketa")
            );
          })
          .join("") +
        (counted.length + cAuto.length > 1 ? "<br/>Päivä yhteensä <b>" + eur(dayEur) + "</b>" : "") +
        "</div>";
    }
    const editor =
      '<div class="juna-editor"' + (open ? "" : " hidden") + ' data-date="' + d.date + '">' +
      '<p class="tot-help">Lisää päivän junat. Paino on junan todellinen paino jarrupainojärjestelmästä. Km on lähtö- ja tulopaikan välinen matka. Veturiraha = km × TES:n hinta. Veturina ajo: vain km.</p>' +
      '<p class="juna-pair' + (info.yksin ? "" : " kaksin") + '">' + escHtml(pairHint(info)) + " (suunniteltu vuoro)</p>" +
      '<div class="juna-rows">' + (list.length ? list.map((j) => junaRowHtml(j, info)).join("") : autos.length ? "" : junaRowHtml(null, info)) + "</div>" +
      '<button type="button" class="j-add">+ Lisää juna</button>' +
      '<h4 class="auto-head">Autolla-ajo <small>(vain auton kuljettaja, miehistönvaihto)</small></h4>' +
      '<div class="auto-rows">' + autos.map(autoRowHtml).join("") + "</div>" +
      '<button type="button" class="a-add">+ Lisää autolla-ajo</button>' +
      '<div class="ot-actions"><button type="button" class="ot-save j-save" data-date="' + d.date + '">Tallenna</button>' +
      (n ? '<button type="button" class="ot-reset j-del" data-date="' + d.date + '">Poista kaikki</button>' : "") +
      "</div>" +
      "</div>";
    return { btn: btn, body: line + editor, n: n };
  }
  function readJunaEditor(box) {
    const out = [];
    const autos = [];
    const num = (s) => {
      const t = s.trim().replace(/\s/g, "").replace(",", ".");
      if (t === "") return null;
      const n = Number(t);
      return isFinite(n) && n > 0 ? n : NaN;
    };
    const rows = box.querySelectorAll(".juna-rows .juna-row");
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const vet = r.querySelector(".j-veturina").checked;
      const nm = (vet ? "Veturina ajo " : "Juna ") + (i + 1);
      const nro = r.querySelector(".j-num").value.trim();
      const paino = vet ? null : num(r.querySelector(".j-paino").value);
      const km = num(r.querySelector(".j-km").value);
      if (!nro && paino == null && km == null) continue; // empty row
      if (Number.isNaN(paino) || Number.isNaN(km)) return { err: nm + ": anna " + (vet ? "km" : "paino ja km") + " numeroina." };
      if (!vet && paino == null) return { err: nm + ": anna paino (t)." };
      if (km == null) return { err: nm + ": anna km." };
      const o = vet ? { junanumero: nro, km: km, veturina: true } : { junanumero: nro, paino: paino, km: km };
      // Store yksin only as a manual override (differs from the sheet's auto value); otherwise auto.
      const sel = r.querySelector(".j-ajo:checked");
      const yk = sel ? sel.value === "yksin" : r.dataset.auto === "1";
      if (yk !== (r.dataset.auto === "1")) o.yksin = yk;
      if (!vet && r.querySelector(".j-hidas").checked) o.hidas = true;
      if (!vet && r.querySelector(".j-ivyvak").checked) o.ivyvak = true;
      out.push(o);
    }
    const arows = box.querySelectorAll(".auto-rows .auto-row");
    for (let i = 0; i < arows.length; i++) {
      const km = num(arows[i].querySelector(".a-km").value);
      if (km == null) continue;
      if (Number.isNaN(km)) return { err: "Autolla-ajo " + (i + 1) + ": anna km numerona." };
      autos.push({ tyyppi: arows[i].querySelector(".a-type").value === "ulko" ? "ulko" : "paikallis", km: km });
    }
    return { value: out, autos: autos };
  }
  const WD_JS = ["Su", "Ma", "Ti", "Ke", "To", "Pe", "La"];
  function dayLabel(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
    if (!m) return iso;
    const d = new Date(+m[1], +m[2] - 1, +m[3]);
    return WD_JS[d.getDay()] + " " + +m[3] + "." + +m[2] + ".";
  }
  function notesHtml(notes) {
    return notes.length ? '<ul class="lisat-events warn">' + notes.map((n) => "<li>" + dayLabel(n.date) + " · " + escHtml(n.msg) + "</li>").join("") + "</ul>" : "";
  }
  /** Per-pay-code summary rows (rounded per code, as on the payslip). */
  function codeRows(rows, labelOf, codeOf, unit) {
    const multi = {};
    rows.forEach((r) => (multi[codeOf(r)] = (multi[codeOf(r)] || 0) + 1));
    return rows
      .map(
        (r) =>
          "<tr><td><b>" + labelOf(r) + "</b><small>" + r.n + " " + (typeof unit === "function" ? unit(r) : unit)[r.n === 1 ? 0 : 1] + " · " + rateFi(r.rate) + " €/km" +
          (multi[codeOf(r)] > 1 ? " · " + dateFi(r.from).slice(0, 6) + " alk." : "") +
          "</small></td><td>" + (r.code || "–") + "</td><td>" + numFi(r.km) + '</td><td class="lisat-pay">' + eur(r.eur) + "</td></tr>"
      )
      .join("");
  }
  function veturirahaHtml(vr) {
    const vrNotes = notesHtml(vr.notes);
    if (!vr.rows.length)
      return (
        '<table class="lisat-table"><thead><tr><th>Veturiraha</th><th>Koodi</th><th>Km</th><th>€</th></tr></thead><tbody>' +
        '<tr><td colspan="4"><small>Ei laskettavia junia. Lisää päivän junat <b>+ Juna</b>-napista (junanumero, paino, km).</small></td></tr></tbody></table>' +
        vrNotes
      );
    const list = vr.trains
      .map((t) => {
        const what = t.veturina ? (t.junanumero ? escHtml(t.junanumero) : "") : (t.junanumero ? escHtml(t.junanumero) + " · " : "") + tonnit(t.paino) + " t";
        return (
          "<tr><td><b>" + ShiftCalc.veturirahaTitle(t) + "</b><small>" + dayLabel(t.date) + (what ? " · " + what : "") + (t.auto ? "" : " · valittu käsin") + " · " + rateFi(t.rate) + " €/km</small></td>" +
          "<td>" + t.code + "</td><td>" + numFi(t.km) + '</td><td class="lisat-pay">' + eur(t.eur) + "</td></tr>"
        );
      })
      .join("");
    return (
      '<table class="lisat-table vr-list"><thead><tr><th>Veturiraha · junat ja ajot</th><th>Koodi</th><th>Km</th><th>€</th></tr></thead><tbody>' +
      list +
      "</tbody></table>" +
      '<table class="lisat-table vr-codes"><thead><tr><th>Veturiraha palkkalajeittain</th><th>Koodi</th><th>Km</th><th>€</th></tr></thead><tbody>' +
      codeRows(vr.rows, (r) => r.label, (r) => r.code, (r) => (r.code === 1456 || r.code === 1463 ? ["ajo", "ajoa"] : ["juna", "junaa"])) +
      '<tr class="lisat-sum"><td colspan="2"><b>Veturiraha yhteensä</b></td><td>' + numFi(vr.km) + '</td><td class="lisat-pay">' + eur(vr.eur) + "</td></tr>" +
      "</tbody></table>" +
      vrNotes +
      '<p class="ot-meta lisat-note">Veturiraha = km × hinta junapainon mukaan (TES Lisäpalkkiot, veturiraha). Yksinajon hinnat ovat taulukossa jo kaksinkertaisia. Veturina ajo maksetaan henkilöjunan hinnalla (vakiintunut käytäntö). Hinta ajopäivän mukaan (1.9.2026 alkaen uudet). Yhteissumma pyöristetään palkkalajeittain koko jaksolta, joten se voi poiketa rivien summasta sentillä.</p>'
    );
  }
  function autoajoHtml(aa) {
    if (!aa.rows.length && !aa.notes.length) return "";
    const list = aa.trips
      .map(
        (t) =>
          "<tr><td><b>" + ShiftCalc.AUTOAJO_LABELS[t.tyyppi] + "</b><small>" + dayLabel(t.date) + " · " + numFi(t.km) + " km ajettu" +
          (t.paidKm !== t.km ? " → väh. " + ShiftCalc.AUTOAJO_MIN_KM + " km" : "") + " · " + rateFi(t.rate) + " €/km</small></td>" +
          '<td>–</td><td>' + numFi(t.paidKm) + '</td><td class="lisat-pay">' + eur(t.eur) + "</td></tr>"
      )
      .join("");
    return (
      '<table class="lisat-table aa-list"><thead><tr><th>Autolla-ajo · ajot</th><th>Koodi</th><th>Km</th><th>€</th></tr></thead><tbody>' +
      list +
      "</tbody></table>" +
      '<table class="lisat-table aa-codes"><thead><tr><th>Autolla-ajokorvaus lajeittain</th><th>Koodi</th><th>Km</th><th>€</th></tr></thead><tbody>' +
      codeRows(aa.rows, (r) => r.label + (r.tyyppi === "ulko" ? " (väh. " + ShiftCalc.AUTOAJO_MIN_KM + " km/ajo)" : ""), (r) => r.tyyppi, ["ajo", "ajoa"]) +
      '<tr class="lisat-sum"><td colspan="2"><b>Autolla-ajo yhteensä</b></td><td>' + numFi(aa.km) + '</td><td class="lisat-pay">' + eur(aa.eur) + "</td></tr>" +
      "</tbody></table>" +
      notesHtml(aa.notes) +
      '<p class="ot-meta lisat-note">TES Lisäpalkkiot, autolla-ajokorvaus: paikallisajo hitaan yksinajon (1501–4800 t) hinnalla, paikkakunnan ulkopuolinen ajo kaksinajon (enint. 1500 t) hinnalla, vähintään 30 km yhteen suuntaan. Maksetaan vain auton kuljettajalle. Palkkalaji ei tiedossa.</p>'
    );
  }

  function renderLisat(p, lisat) {
    const L = lisat;
    const C = ShiftCalc.LISA_CODES;
    const rows = [
      ["ilta", "Iltatyölisä", "klo 18–21 · §17"],
      ["yo", "Yötyölisä", "klo 21–06, jatko klo 12 asti jos alkanut viim. klo 4 · §18"],
      ["la", "Lauantaityökorvaus", "arkilauantai klo 06–18 · §19"],
      ["su", "Sunnuntaityökorvaus", "su ja pyhät klo 0–24 + edellinen päivä klo 18–24 · §20"],
      ["aatto", "Aattopäivänlisä", "pääsiäislauantai, juhannus- ja jouluaatto klo 0–18 · §21"],
    ];
    const hourRows = rows
      .map((r) =>
        "<tr><td><b>" + r[1] + "</b><small>" + r[2] + "</small></td><td>" + (C[r[0]] || "–") + "</td><td>" +
        fmt(L.minutes[r[0]]) + '</td><td class="lisat-pay">' + L.hours[r[0]] + " h</td></tr>"
      )
      .join("");
    const rateText = (f) => {
      const r = Object.keys(f.rates);
      return r.length === 1 ? f.n + " × " + r[0].replace(".", ",") + " €" : f.n + " kpl";
    };
    const fixedRows = ["poikkeama", "vapaa", "peruutus"]
      .map((k) => {
        const f = L.fixed[k];
        const unit = ShiftCalc.fixedRatesFor(state.startDate)[k];
        return (
          "<tr><td><b>" + FIXED_LABEL[k] + "</b><small>" + (f.n ? rateText(f) : eur(unit) + " / kpl") +
          (k === "peruutus" ? " · §6" : " · §22") + "</small></td><td>" + C[k] + "</td><td>" + f.n + ' kpl</td><td class="lisat-pay">' +
          eur(f.eur) + "</td></tr>"
        );
      })
      .join("");
    const evs = L.events.length
      ? '<ul class="lisat-events">' +
        L.events
          .map((e) => {
            const d = p.days.find((x) => x.date === e.date) || {};
            return "<li>" + (WD_SHORT[d.weekday] || "") + " " + dateFi(e.date).slice(0, 6) + " · " + FIXED_LABEL[e.key] + (e.eur ? " " + eur(e.eur) : "") + " – " + e.why + "</li>";
          })
          .join("") +
        "</ul>"
      : "";
    const notes = L.notes.length
      ? '<ul class="lisat-events warn">' + L.notes.map((n) => "<li>" + dateFi(n.date).slice(0, 6) + " · " + n.msg + "</li>").join("") + "</ul>"
      : "";
    $("detailLisat").innerHTML =
      "<h3>Lisät (TES)</h3>" +
      '<p class="ot-meta">Vertaa palkkalaskelman määriin. Tunnit lasketaan yhteen koko jaksolta ja pyöristetään kerran (§24): alle 30 min alas, 30 min tai yli ylös. Toteuma-merkinnät korvaavat suunnitellun ajan.</p>' +
      '<table class="lisat-table"><thead><tr><th>Tuntilisä</th><th>Koodi</th><th>Tehty</th><th>Maksetaan</th></tr></thead><tbody>' +
      hourRows +
      "</tbody></table>" +
      '<p class="ot-meta lisat-note">Tuntilisien euromäärät lisätään, kun tuntipalkka ja täydennysosa on vahvistettu.</p>' +
      '<table class="lisat-table"><thead><tr><th>Korvaus</th><th>Koodi</th><th>Määrä</th><th>€</th></tr></thead><tbody>' +
      fixedRows +
      '<tr class="lisat-sum"><td colspan="3"><b>Yhteensä</b></td><td class="lisat-pay">' + eur(L.fixedTotal) + "</td></tr>" +
      "</tbody></table>" +
      evs +
      notes +
      '<div class="vr-block">' + veturirahaHtml(personVeturiraha(p)) + autoajoHtml(personAutoajo(p)) + "</div>" +
      palkkioHtml(p) +
      '<p class="ot-meta lisat-note">Merkitse muutokset päivän <b>Toteuma</b>-napista. Toteuman tunnit lasketaan myös kovien tuntien, Tunnit yhteensä -luvun ja ylitöiden (lisätyö, 50 %, 100 %) yhteismääriin: muuttunut vuoro toteutuneen ajan mukaan ja kutsu vapaapäivänä kokonaan, ilman erillistä lisävuoroa. Ajoissa peruttu vuoro pitää suunnitellut tunnit. Klo 17 jälkeen perutusta vuorosta valitset joko tunnit tai peruutuskorvauksen, et molempia. Lisävuoroa, joka on päällekkäin saman päivän toteutuneen vuoron kanssa, ei lasketa.</p>';
  }

  // --- Lataa PDF: print view (window.print → "Tallenna PDF:nä"), works offline ---
  function printDoc(title, html) {
    let box = $("printView");
    if (!box) {
      box = document.createElement("div");
      box.id = "printView";
      document.body.appendChild(box);
    }
    box.innerHTML = html + '<p class="pv-foot">© 2026 Lämpöpumppu Mafia · Kaikki oikeudet pidätetään</p>';
    const oldTitle = document.title;
    document.title = title;
    const restore = () => {
      document.title = oldTitle;
      window.removeEventListener("afterprint", restore);
    };
    window.addEventListener("afterprint", restore);
    window.print();
  }
  function jaksoRange() {
    const end = state.dates && state.dates.length ? state.dates[state.dates.length - 1] : state.startDate;
    return dateFi(state.startDate) + "–" + dateFi(end);
  }
  function pvTable(head, rows, cls) {
    return (
      '<table class="pv-table ' + (cls || "") + '"><thead><tr>' + head.map((h) => "<th>" + h + "</th>").join("") + "</tr></thead><tbody>" +
      rows.map((r) => "<tr" + (r.cls ? ' class="' + r.cls + '"' : "") + ">" + (r.cells || r).map((c) => "<td>" + (c == null ? "" : c) + "</td>").join("") + "</tr>").join("") +
      "</tbody></table>"
    );
  }
  function printPerson(p) {
    const marks = personMarks(p), ex = personExtras(p), tots = personTots(p), hm = state.holidayMap;
    const L = personLisat(p);
    const vr = personVeturiraha(p);
    const aa = personAutoajo(p);
    const ot = personOt(p);
    const fm = (m) => (m ? fmt(m) : "");
    const dayRows = p.days.map((d) => {
      const sick = isSickDay(p, d);
      const t = ShiftCalc.normToteuma(tots[d.date]);
      const one = [d];
      const me = sick ? 0 : ShiftCalc.sumActive(one, marks, "me", ex, hm, tots);
      const hrs = sick ? 0 : ShiftCalc.sumActive(one, marks, "hrs", ex, hm, tots);
      const m = sick ? null : ShiftCalc.lisatSummary(one, marks, ex, tots, hm).minutes;
      const extra = ex[d.date];
      const plan = d.start != null ? fmt(d.start) + "–" + fmt(d.end) : "–";
      const act = [t ? totSummaryText(t) : "", extra ? "Lisävuoro " + fmt(extra.start) + "–" + fmt(extra.end) : ""].filter(Boolean).join("; ");
      return {
        cls: sick ? "pv-sick" : "",
        cells: [
          (WD_SHORT[d.weekday] || "") + " " + dateFi(d.date).slice(0, 6) + (d.special ? "<small>" + d.special + "</small>" : ""),
          plan, act ? "<small>" + escHtml(act) + "</small>" : "", fm(me), fm(hrs),
          m ? fm(m.ilta) : "", m ? fm(m.yo) : "", m ? fm(m.la) : "", m ? fm(m.su) : "", m ? fm(m.aatto) : "",
          sick ? "Kyllä" : "",
        ],
      };
    });
    const lmMin = stewardMinutes(p);
    if (lmMin) dayRows.push(["Luottamusmiestunnit", "", "<small>" + escHtml(stewardLine(p)) + "</small>", "+" + fmt(lmMin), "", "", "", "", "", "", ""]);
    dayRows.push({
      cls: "pv-sum",
      cells: ["<b>Yhteensä</b>", "", "", fmt(activeLm(p)), fmt(sumPerson(p, "hrs")), fmt(L.minutes.ilta), fmt(L.minutes.yo), fmt(L.minutes.la), fmt(L.minutes.su), fmt(L.minutes.aatto), sickCount(p) || ""],
    });
    const trainRows = vr.trains.map((t) => [dayLabel(t.date), escHtml(t.junanumero || "–"), t.veturina ? "–" : tonnit(t.paino) + " t", "<b>" + (t.yksin ? "Yksinajo" : "Kaksinajo") + "</b>" + (t.veturina ? " · veturina ajo" : "") + (t.hidas && !t.veturina ? " · hidas" : "") + (t.ivyvak && !t.veturina ? " · IVY-VAK" : ""), t.code, numFi(t.km), eur(t.eur)]);
    const autoRows = aa.trips.map((t) => [dayLabel(t.date), ShiftCalc.AUTOAJO_LABELS[t.tyyppi], numFi(t.km) + (t.paidKm !== t.km ? " → " + numFi(t.paidKm) : ""), "–", eur(t.eur)]);
    const C = ShiftCalc.LISA_CODES;
    const lisaRows = [["ilta", "Iltatyölisä"], ["yo", "Yötyölisä"], ["la", "Lauantaityökorvaus"], ["su", "Sunnuntaityökorvaus"], ["aatto", "Aattopäivänlisä"]].map((r) => [r[1], C[r[0]] || "–", fmt(L.minutes[r[0]]), L.hours[r[0]] + " h"]);
    const fixedRows = ["poikkeama", "vapaa", "peruutus"].map((k) => [FIXED_LABEL[k], C[k], L.fixed[k].n + " kpl", eur(L.fixed[k].eur)]);
    const endKey = state.dates[state.dates.length - 1];
    const rates = ShiftCalc.palkkioRatesInRange(state.startDate, endKey).map((r, i) => eur(r.eur) + "/kk" + (i ? " " + dateFi(r.from).slice(0, 6) + " alkaen" : "")).join(", ");
    const pk = [];
    if (isSteward(p)) pk.push(["Luottamusmiespalkkio", rates, "maksetaan 10 kk/vuosi (varamiehelle 2 kk)"]);
    if (isTsv(p)) pk.push(["Työsuojeluvaltuutetun palkkio", rates, "12 kk/vuosi (varahenkilölle 2 kk)"]);
    const flags = [isSteward(p) ? "LM (+8 h)" : "", isTsv(p) ? "TSV" : ""].filter(Boolean).join(" · ");
    const html =
      '<header class="pv-head"><h1>Jakso päiväkirja</h1><p><b>' + escHtml(p.name) + "</b> · " + (state.group ? escHtml(state.group) + " · " : "") + (state.periodLabel ? escHtml(state.periodLabel) + " · " : "") + jaksoRange() +
      (flags ? " · " + flags : "") + "</p><p class=\"pv-meta\">Tulostettu " + dateFi(new Date().toISOString().slice(0, 10)) + "</p></header>" +
      "<h2>Päivät</h2>" +
      pvTable(["Päivä", "Suunniteltu", "Toteuma", "Tunnit yht.", "Kovat", "Ilta", "Yö", "La", "Su", "Aatto", "Kesk."], dayRows, "pv-days") +
      "<h2>Junat ja ajot</h2>" +
      (trainRows.length ? pvTable(["Päivä", "Juna", "Paino", "Ajo", "Koodi", "Km", "€"], trainRows) : '<p class="pv-meta">Ei junia.</p>') +
      (autoRows.length ? "<h2>Autolla-ajo</h2>" + pvTable(["Päivä", "Ajo", "Km", "Koodi", "€"], autoRows) : "") +
      "<h2>Lisät (TES)</h2>" +
      pvTable(["Tuntilisä", "Koodi", "Tehty", "Maksetaan"], lisaRows) +
      pvTable(["Korvaus", "Koodi", "Määrä", "€"], fixedRows.concat([{ cls: "pv-sum", cells: ["<b>Yhteensä</b>", "", "", eur(L.fixedTotal)] }])) +
      (vr.rows.length
        ? pvTable(["Veturiraha palkkalajeittain", "Koodi", "Km", "€"], vr.rows.map((r) => [r.label + " <small>" + String(r.rate).replace(".", ",") + " €/km</small>", r.code, numFi(r.km), eur(r.eur)]).concat([{ cls: "pv-sum", cells: ["<b>Veturiraha yhteensä</b>", "", numFi(vr.km), eur(vr.eur)] }]))
        : "") +
      (aa.rows.length
        ? pvTable(["Autolla-ajokorvaus", "Koodi", "Km", "€"], aa.rows.map((r) => [r.label + " <small>" + String(r.rate).replace(".", ",") + " €/km</small>", "–", numFi(r.km), eur(r.eur)]).concat([{ cls: "pv-sum", cells: ["<b>Autolla-ajo yhteensä</b>", "", numFi(aa.km), eur(aa.eur)] }]))
        : "") +
      "<h2>Tunnit ja ylityöt</h2>" +
      pvTable(["", "Tunnit"], [
        ["Kovat tunnit", fmt(sumPerson(p, "hrs"))],
        ["Tunnit yhteensä", fmt(activeLm(p))],
        ["Lisätyö", fmtOt(ot, "lisa")],
        ["Ylityö 50 %", fmtOt(ot, "yli50")],
        ["Ylityö 100 %", fmtOt(ot, "yli100")],
      ], "pv-narrow") +
      '<p class="pv-meta">' + overtimeText(state.otEffective) + (isSteward(p) ? " · " + stewardLine(p) : "") + "</p>" +
      (pk.length ? "<h2>Palkkiot</h2>" + pvTable(["Palkkio", "€/kk", ""], pk) + '<p class="pv-meta">Palkkiot ovat kuukausikohtaisia, eivät jaksokohtaisia.</p>' : "");
    printDoc("Jakso päiväkirja – " + p.name + " – " + jaksoRange(), html);
  }

  function openDetail(personId) {
    const p = state.people.find((x) => x.id === personId);
    if (!p) return;
    $("detailTitle").textContent = p.name + " · " + p.shiftCount + " vuoroa";
    const rests = personRest(p);
    const lisat = personLisat(p);
    const vr = personVeturiraha(p);
    const aa = personAutoajo(p);

    // Desktop/wide: classic table; mobile: compact day cards (see CSS)
    const tbody = $("detailTable").querySelector("tbody");
    tbody.innerHTML = p.days
      .map((d) => {
        const sick = isSickDay(p, d);
        const cell = (v) => '<td class="' + (sick ? "struck" : "") + '">' + v + "</td>";
        return (
          '<tr class="' + (sick ? "sick" : "") + '">' +
          "<td>" +
          '<button type="button" class="sick-btn' +
          (sick ? " on" : "") +
          '" data-date="' +
          d.date +
          '">' +
          (sick ? "Keskeytyspäivä ✓" : "Keskeytyspäivä") +
          "</button> " +
          extraBits(p, d.date).btn +
          totBits(p, d, lisat).btn +
          junaBits(p, d, vr, aa).btn +
          " " +
          dateFi(d.date).slice(0, 5) +
          extraBits(p, d.date).body +
          totBits(p, d, lisat).body +
          junaBits(p, d, vr, aa).body +
          "</td>" +
          "<td>" +
          (WD_SHORT[d.weekday] || "") +
          "</td>" +
          "<td>" +
          (d.special || "") +
          "</td>" +
          "<td>" +
          fmt(d.start) +
          "</td>" +
          "<td>" +
          fmt(d.end) +
          "</td>" +
          '<td class="' +
          (sick ? "" : restClass((rests[d.date] || {}).restAfter)) +
          '">' +
          (sick ? "—" : lab((rests[d.date] || {}).restAfter)) +
          "</td>" +
          cell(fmt(d.hrs)) +
          cell(fmt(d.night)) +
          cell(fmt(d.b25)) +
          cell(fmt(d.b100)) +
          cell(fmt(d.me)) +
          "</tr>"
        );
      })
      .join("");

    // Mobile card list (always in DOM; CSS shows/hides)
    const cards = $("detailCards");
    cards.innerHTML = p.days
      .map((d) => {
        const hasShift = d.start != null;
        const sick = isSickDay(p, d);
        const rd = rests[d.date] || {};
        const restLabel = !sick && rd.restAfter
          ? lab(rd.restAfter) +
            (rd.restAfterMin != null ? " · " + fmt(rd.restAfterMin) : "")
          : "";
        const extra = extraBits(p, d.date);
        const tot = totBits(p, d, lisat);
        const juna = junaBits(p, d, vr, aa);
        return (
          '<article class="day-card' +
          (hasShift ? "" : " empty") +
          (extra.fig ? " has-extra" : "") +
          (sick ? " sick" : "") +
          (tot.t ? " has-tot" : "") +
          (juna.n ? " has-juna" : "") +
          '">' +
          '<header><strong>' +
          (WD_SHORT[d.weekday] || "") +
          " " +
          dateFi(d.date) +
          "</strong>" +
          (d.special ? '<span class="tag">' + d.special + "</span>" : "") +
          '<button type="button" class="sick-btn' +
          (sick ? " on" : "") +
          '" data-date="' +
          d.date +
          '">' +
          (sick ? "Keskeytyspäivä ✓" : "Keskeytyspäivä") +
          "</button>" +
          extra.btn +
          tot.btn +
          juna.btn +
          "</header>" +
          (hasShift
            ? '<div class="day-grid' +
              (sick ? " struck" : "") +
              '">' +
              "<div><span>Alku</span><b>" +
              fmt(d.start) +
              "</b></div>" +
              "<div><span>Loppu</span><b>" +
              fmt(d.end) +
              "</b></div>" +
              "<div><span>Kovat tunnit</span><b>" +
              fmt(d.hrs) +
              "</b></div>" +
              "<div><span>Yö h</span><b>" +
              fmt(d.night) +
              "</b></div>" +
              "<div><span>Tunnit yhteensä</span><b>" +
              fmt(d.me) +
              "</b></div>" +

              "<div><span>Lauantai h</span><b>" +
              fmt(d.b25) +
              "</b></div>" +
              "<div><span>Pyhä h</span><b>" +
              fmt(d.b100) +
              "</b></div>" +

              '<div class="' +
              (sick ? "" : restClass(rd.restAfter)) +
              '"><span>Vuorojen väli →</span><b>' +
              (restLabel || "—") +
              "</b></div>" +
              "</div>"
            : extra.fig
              ? ""
              : '<p class="muted">Ei vuoroa</p>') +
          extra.body +
          tot.body +
          juna.body +
          "</article>"
        );
      })
      .join("");

    const sum = (key) => sumPerson(p, key);
    const paintTotals = () => {
    $("detailTotals").innerHTML =
      '<div class="stat"><b>' +
      fmt(sum("hrs")) +
      "</b><span>Kovat tunnit</span></div>" +
      '<div class="stat"><b>' +
      fmt(sum("night")) +
      "</b><span>Yö h</span></div>" +
      '<div class="stat"><b>' +
      fmt(sum("b25")) +
      "</b><span>Lauantai h</span></div>" +
      '<div class="stat"><b>' +
      fmt(sum("b100")) +
      "</b><span>Pyhä h</span></div>" +
      '<div class="stat"><b>' +
      fmt(activeLm(p)) +
      "</b><span>Tunnit yhteensä</span></div>" +

      '<div class="stat"><b>' +
      fmtOt(personOt(p), "lisa") +
      "</b><span>Lisätyö h</span></div>" +
      '<div class="stat"><b>' +
      fmtOt(personOt(p), "yli50") +
      "</b><span>Ylityö 50 %</span></div>" +
      '<div class="stat"><b>' +
      fmtOt(personOt(p), "yli100") +
      "</b><span>Ylityö 100 %</span></div>" +
      '<p class="ot-meta totals-note">' +
      overtimeText(state.otEffective) +
      (state.otEffective && state.otEffective.overridden ? ' <span class="ot-mod">muutettu</span>' : "") +
      "</p>" +
      (isSteward(p) ? '<p class="ot-meta totals-note">' + stewardLine(p) + "</p>" : "") +
      (otNote(personOt(p)) ? '<p class="ot-meta totals-note">' + otNote(personOt(p)) + "</p>" : "");
    };
    paintTotals();

    function paintLaskenta(keepInput) {
      const canLask = state.dayCount === 21 && !!state.otEffective;
      $("laskentaBox").hidden = !canLask;
      if (!canLask) return;
      const info = personNInfo(p);
      const sick = sickCount(p);
      const word = sick === 1 ? "keskeytyspäivä" : "keskeytyspäivää";
      $("laskentaLabel").textContent =
        "Laskentapäivät " + info.n + (sick ? " (" + sick + " " + word + ")" : "");
      if (!keepInput) $("laskentaInput").value = info.n == null ? "" : info.n;
      $("laskentaReset").hidden = !info.manual;
      $("laskentaLine").textContent = laskentaLine(p);
    }
    paintLaskenta(false);
    $("lmCheckLabel").textContent = stewardLabel();
    $("lmCheck").checked = isSteward(p);
    $("lmCheck").onchange = () => {
      setSteward(p, $("lmCheck").checked);
      paintTotals();
      renderLisat(p, personLisat(p));
    };
    $("tsvCheck").checked = isTsv(p);
    $("tsvCheck").onchange = () => {
      setTsv(p, $("tsvCheck").checked);
      renderLisat(p, personLisat(p));
    };
    $("laskentaInput").oninput = () => {
      const v = parseInt($("laskentaInput").value, 10);
      if (!(v >= 1 && v <= 21)) return;
      savePersonN(p, v);
      paintLaskenta(true);
      paintTotals();
    };
    $("laskentaReset").onclick = () => {
      clearPersonN(p);
      paintLaskenta(false);
      paintTotals();
    };

    const redraw = () => {
      const y = window.scrollY;
      openDetail(p.id);
      window.scrollTo(0, y);
    };
    document.querySelectorAll("#view-detail .sick-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        toggleSick(p, btn.dataset.date);
        redraw();
      });
    });
    document.querySelectorAll("#view-detail .extra-toggle").forEach((btn) => {
      btn.addEventListener("click", () => {
        const date = btn.dataset.date;
        if (isSickDay(p, { date: date })) {
          dayFlash = { id: p.id, date: date, msg: CONFLICT_MSG };
          openExtra = null;
          redraw();
          return;
        }
        dayFlash = null;
        if (openExtra && openExtra.id === p.id && openExtra.date === date) openExtra = null;
        else openExtra = { id: p.id, date: date };
        redraw();
      });
    });
    document.querySelectorAll("#view-detail .ex-save").forEach((btn) => {
      btn.addEventListener("click", () => {
        const date = btn.dataset.date;
        const box = btn.closest(".extra-editor");
        const a = ShiftCalc.parseHM(box.querySelector(".ex-start").value);
        const b = ShiftCalc.parseHM(box.querySelector(".ex-end").value);
        openExtra = { id: p.id, date: date };
        if (a == null || b == null) {
          dayFlash = { id: p.id, date: date, msg: "Anna alku ja loppu muodossa h:mm." };
          redraw();
          return;
        }
        if (a === b) {
          dayFlash = { id: p.id, date: date, msg: "Alku ja loppu eivät voi olla samat." };
          redraw();
          return;
        }
        if (isSickDay(p, { date: date })) {
          dayFlash = { id: p.id, date: date, msg: CONFLICT_MSG };
          redraw();
          return;
        }
        saveExtra(p, date, a, b);
        dayFlash = null;
        openExtra = null;
        redraw();
      });
    });
    document.querySelectorAll("#view-detail .ex-del").forEach((btn) => {
      btn.addEventListener("click", () => {
        deleteExtra(p, btn.dataset.date);
        dayFlash = null;
        openExtra = null;
        redraw();
      });
    });

    renderLisat(p, lisat);
    document.querySelectorAll("#view-detail .tot-editor").forEach((box) => {
      syncTotEditor(box);
      box.querySelector(".tot-type").addEventListener("change", () => syncTotEditor(box));
    });
    document.querySelectorAll("#view-detail .tot-toggle").forEach((btn) => {
      btn.addEventListener("click", () => {
        const date = btn.dataset.date;
        if (isSickDay(p, { date: date })) {
          dayFlash = { id: p.id, date: date, msg: TOT_CONFLICT_MSG, tot: true };
          openTot = { id: p.id, date: date };
          redraw();
          return;
        }
        dayFlash = null;
        if (openTot && openTot.id === p.id && openTot.date === date) openTot = null;
        else openTot = { id: p.id, date: date };
        redraw();
      });
    });
    document.querySelectorAll("#view-detail .tot-save").forEach((btn) => {
      btn.addEventListener("click", () => {
        const date = btn.dataset.date;
        const d = p.days.find((x) => x.date === date);
        const box = btn.closest(".tot-editor");
        openTot = { id: p.id, date: date };
        if (isSickDay(p, { date: date })) {
          dayFlash = { id: p.id, date: date, msg: TOT_CONFLICT_MSG, tot: true };
          redraw();
          return;
        }
        const r = readTotEditor(box, d);
        if (r.err) {
          dayFlash = { id: p.id, date: date, msg: r.err, tot: true };
          redraw();
          return;
        }
        saveTot(p, date, r.value);
        dayFlash = null;
        openTot = null;
        redraw();
      });
    });
    document.querySelectorAll("#view-detail .tot-del").forEach((btn) => {
      btn.addEventListener("click", () => {
        saveTot(p, btn.dataset.date, null);
        dayFlash = null;
        openTot = null;
        redraw();
      });
    });

    document.querySelectorAll("#view-detail .juna-toggle").forEach((btn) => {
      btn.addEventListener("click", () => {
        const date = btn.dataset.date;
        dayFlash = null;
        if (openJuna && openJuna.id === p.id && openJuna.date === date) openJuna = null;
        else openJuna = { id: p.id, date: date };
        redraw();
      });
    });
    document.querySelectorAll("#view-detail .juna-editor").forEach((box) => {
      box.querySelector(".j-add").addEventListener("click", () => {
        box.querySelector(".juna-rows").insertAdjacentHTML("beforeend", junaRowHtml(null, pairInfo(p, box.dataset.date)));
      });
      box.querySelector(".a-add").addEventListener("click", () => {
        box.querySelector(".auto-rows").insertAdjacentHTML("beforeend", autoRowHtml(null));
      });
      box.addEventListener("change", (e) => {
        const vt = e.target.closest(".j-veturina");
        if (vt) vt.closest(".juna-row").classList.toggle("is-veturina", vt.checked);
        const cb = e.target.closest(".j-ajo");
        if (!cb) return;
        const info = pairInfo(p, box.dataset.date);
        const hint = cb.closest(".juna-row").querySelector(".j-hint");
        hint.textContent = junaHintText(info, cb.value === "yksin");
      });
      box.addEventListener("click", (e) => {
        const rm = e.target.closest(".j-remove");
        if (!rm) return;
        const row = rm.closest(".juna-row");
        if (row.classList.contains("auto-row") || box.querySelectorAll(".juna-rows .juna-row").length > 1) row.remove();
        else row.querySelectorAll("input:not([type=checkbox])").forEach((i) => (i.value = ""));
      });
      box.querySelector(".j-save").addEventListener("click", () => {
        const date = box.dataset.date;
        openJuna = { id: p.id, date: date };
        const r = readJunaEditor(box);
        if (r.err) {
          // keep typed values: show the message without redrawing
          let el = box.querySelector(".extra-err");
          if (!el) {
            el = document.createElement("p");
            el.className = "extra-err";
            box.appendChild(el);
          }
          el.textContent = r.err;
          return;
        }
        saveJunat(p, date, r.value);
        saveAutoajot(p, date, r.autos);
        dayFlash = null;
        openJuna = null;
        redraw();
      });
      const del = box.querySelector(".j-del");
      if (del)
        del.addEventListener("click", () => {
          saveJunat(p, box.dataset.date, null);
          saveAutoajot(p, box.dataset.date, null);
          dayFlash = null;
          openJuna = null;
          redraw();
        });
    });

    $("btnPdf").onclick = () => printPerson(p);
    renderAlerts($("detailAlerts"), state, personId);
    show("detail");
  }

  let pendingBuf = null;
  let pendingName = "";

  function showStartPrompt(result, fileName) {
    $("startPrompt").hidden = false;
    $("startPromptText").textContent =
      "Aloituspäivää ei löytynyt — anna jakson alkupäivä" +
      (result.periodLabel ? " (" + result.periodLabel + ")" : "");
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
    summary = ShiftCalc.summarize(state);
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
      console.error(err);
      $("fileName").textContent =
        "Virhe: " + (err && err.message ? err.message : String(err));
      alert(
        "Tiedoston lukeminen epäonnistui: " +
          (err && err.message ? err.message : err)
      );
    }
  }

  $("fileInput").addEventListener("change", (e) => {
    handleFile(e.target.files && e.target.files[0]);
  });
  $("btnNewFile").addEventListener("click", () => {
    state = null;
    $("fileInput").value = "";
    $("fileName").hidden = true;
    $("startPrompt").hidden = true;
    pendingBuf = null;
    show("home");
  });
  $("btnBack").addEventListener("click", () => renderOverview());

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

  window.__ShiftApp = {
    loadArrayBuffer: (buf, opts) => {
      pendingBuf = buf;
      return useResult(ShiftCalc.parseArrayBuffer(buf, opts), "");
    },
    openDetail,
    getState: () => state,
  };
})();
