/*! © 2026 Pekka Rautiainen. Kaikki oikeudet pidätetään. All rights reserved.
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
  function sumPerson(p, key) {
    return ShiftCalc.sumActive(p.days, personMarks(p), key, personExtras(p), state.holidayMap);
  }
  function stewardN(p) {
    const days = ShiftCalc.OT_CONFIG.periodDays;
    return state.dayCount === days ? personN(p) : days;
  }
  function stewardMinutes(p) {
    return ShiftCalc.stewardBonus(isSteward(p), null, stewardN(p));
  }
  function activeLm(p) {
    return sumPerson(p, "me") + stewardMinutes(p);
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
  function personOt(p) {
    if (!state.otEffective) return null;
    const lm = activeLm(p);
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
   * Alerts from the SAME source as the overview grid: restBefore (+ check mismatches).
   * One alert per flagged arrival day; Finnish date + actual rest duration.
   */
  function renderAlerts(target, result, personId) {
    const items = [];
    result.people.forEach((p) => {
      if (personId && p.id !== personId) return;
      const rest = personRest(p);
      ShiftCalc.sheetMismatches(p.days).forEach((d) => {
        items.push({
          bad: false,
          text: p.name + ": Tarkista erotus " + dateFi(d.date) + " = " + fmt(d.check),
        });
      });
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
    ShiftCalc.sheetMismatches(p.days).forEach((d) => {
      problems.push({
        kind: "md",
        text: "Tarkista erotus " + dateFi(d.date).replace(/\.\d{4}$/, ".") + " = " + fmt(d.check),
      });
    });
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
      const company = sum("company");
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
        "<div><span>LM</span><b>" +
        fmt(activeLm(p)) +
        "</b></div>" +
        "<div><span>Yritys</span><b>" +
        fmt(activeCompany(p)) +
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
    const flash = dayFlash && dayFlash.id === p.id && dayFlash.date === date ? dayFlash.msg : "";
    const btn =
      '<button type="button" class="extra-toggle" data-date="' +
      date +
      '">' +
      (fig ? "Lisävuoro ✓" : "+ Lisävuoro") +
      "</button>";
    let block = "";
    if (fig) {
      block +=
        '<div class="extra-shift"><div class="extra-head"><span class="tag extra">Lisävuoro</span> <b>' +
        fmt(fig.start) +
        "–" +
        fmt(fig.end) +
        '</b></div><div class="day-grid">' +
        "<div><span>Kovat tunnit</span><b>" + fmt(fig.hrs) + "</b></div>" +
        "<div><span>Yö h</span><b>" + fmt(fig.night) + "</b></div>" +
        "<div><span>LM</span><b>" + fmt(fig.me) + "</b></div>" +
        "<div><span>Yritys</span><b>" + fmt(fig.company) + "</b></div>" +
        "<div><span>Lauantai h</span><b>" + fmt(fig.b25) + "</b></div>" +
        "<div><span>Pyhä h</span><b>" + fmt(fig.b100) + "</b></div>" +
        "</div></div>";
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

  function openDetail(personId) {
    const p = state.people.find((x) => x.id === personId);
    if (!p) return;
    $("detailTitle").textContent = p.name + " · " + p.shiftCount + " vuoroa";
    const rests = personRest(p);

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
          " " +
          dateFi(d.date).slice(0, 5) +
          extraBits(p, d.date).body +
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
          cell(fmt(d.company)) +
          cell(fmt(d.check)) +
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
        return (
          '<article class="day-card' +
          (hasShift ? "" : " empty") +
          (extra.fig ? " has-extra" : "") +
          (sick ? " sick" : "") +
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
              "<div><span>LM</span><b>" +
              fmt(d.me) +
              "</b></div>" +
              "<div><span>Yritys</span><b>" +
              fmt(d.company) +
              "</b></div>" +
              "<div><span>Lauantai h</span><b>" +
              fmt(d.b25) +
              "</b></div>" +
              "<div><span>Pyhä h</span><b>" +
              fmt(d.b100) +
              "</b></div>" +
              "<div><span>Erotus</span><b>" +
              fmt(d.check) +
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
      "</b><span>LM yht.</span></div>" +
      '<div class="stat"><b>' +
      fmt(activeCompany(p)) +
      "</b><span>Yritys yht.</span></div>" +
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
