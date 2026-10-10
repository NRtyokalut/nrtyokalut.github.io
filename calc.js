/*! © 2026 Lämpöpumppu Mafia. Kaikki oikeudet pidätetään. All rights reserved.
 * Tämän ohjelmiston kopioiminen, muokkaaminen tai jakaminen ilman lupaa on kielletty. */
/**
 * Vuorotaulun tarkastus — calculation engine (ported from v6 Excel workbook).
 * Works in browser and Node. Times are integer minutes from midnight.
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.ShiftCalc = api;
  return api;
})(typeof globalThis !== "undefined" ? globalThis : typeof window !== "undefined" ? window : undefined, function () {
  "use strict";

  const WEEKDAYS_FI = [
    "Maanantai",
    "Tiistai",
    "Keskiviikko",
    "Torstai",
    "Perjantai",
    "Lauantai",
    "Sunnuntai",
  ];

  const REST_NOT_ALLOWED_MIN = 7 * 60; // Final Sheet I1
  const REST_MD_CHECK_MIN = 10 * 60; // Final Sheet I2
  const NIGHT_FACTOR = 1 / 3;

  // Segment bounds (Matching R8:AA9) as minutes; 0:00:01 ≈ 0 for practical overlap
  const SEG_CUR = [
    [0, 4 * 60], // R 00–04
    [4 * 60, 6 * 60], // S 04–06
    [6 * 60, 12 * 60], // T 06–12
    [12 * 60, 18 * 60], // U 12–18
    [18 * 60, 21 * 60], // V 18–21
    [21 * 60, 24 * 60], // W 21–24 (23:59:59 ≈ 24:00)
  ];
  const SEG_NEXT = [
    [0, 4 * 60], // X
    [4 * 60, 6 * 60], // Y
    [6 * 60, 12 * 60], // Z
    [12 * 60, 15 * 60], // AA
  ];
  const EARLY_FULL = 2 * 60; // AA7 = Y9-Y8 = 2h

  function pad2(n) {
    return String(n).padStart(2, "0");
  }

  function formatHM(mins) {
    if (mins == null || mins === "") return "";
    const s = mins < 0 ? "-" : "";
    const a = Math.abs(Math.round(mins));
    return s + Math.floor(a / 60) + ":" + pad2(a % 60);
  }

  function parseTimeToken(tok) {
    tok = String(tok).trim().replace(".", ":");
    if (tok === "24:00") return 24 * 60;
    const m = tok.match(/^(\d{1,2}):(\d{2})$/);
    if (!m) return null;
    return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
  }

  /** Parse 'hh:mm-hh:mm' → {start, end} minutes, or null */
  function parseShift(text) {
    if (text == null || text === "" || text === 0) return null;
    if (typeof text !== "string") {
      // Excel time serial or Date — treat as invalid for shift text
      return null;
    }
    const s = text.trim().replace(/\u2013|\u2014/g, "-");
    if (s.length < 9 || s.indexOf("-") < 0) return null;
    // Prefer exact 11-char hh:mm-hh:mm; also allow slight variants
    const parts = s.split("-");
    if (parts.length !== 2) return null;
    const start = parseTimeToken(parts[0]);
    const end = parseTimeToken(parts[1]);
    if (start == null || end == null) return null;
    return { start, end };
  }

  /** Company hours from form cell (Excel fraction, Date, or "h:mm") → minutes */
  function parseCompanyHours(v) {
    if (v == null || v === "" || v === 0) return null;
    if (typeof v === "number" && isFinite(v)) {
      // Excel day fraction (hours) or <1 time serial
      const frac = v >= 1 ? v % 1 : v;
      return Math.round(frac * 24 * 60);
    }
    if (v instanceof Date) {
      return v.getUTCHours() * 60 + v.getUTCMinutes() + Math.round(v.getUTCSeconds() / 60);
    }
    if (typeof v === "string") {
      const t = parseTimeToken(v);
      return t;
    }
    return null;
  }

  function excelSerialToDate(serial) {
    // Excel epoch 1899-12-30
    const utc = Date.UTC(1899, 11, 30) + Math.round(serial * 86400000);
    return new Date(utc);
  }

  function toDate(v) {
    if (v == null || v === "") return null;
    if (v instanceof Date) {
      // Prefer UTC parts — SheetJS date cells are UTC-based
      return new Date(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate());
    }
    if (typeof v === "number") {
      // Excel serial: integer = days since 1899-12-30
      if (v >= 1) {
        const utc = Date.UTC(1899, 11, 30) + Math.round(v) * 86400000;
        const d = new Date(utc);
        return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
      }
      // time-only fraction handled elsewhere
      return null;
    }
    if (typeof v === "string") {
      const m = v.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
      const m2 = v.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
      if (m2) return new Date(+m2[3], +m2[2] - 1, +m2[1]);
    }
    return null;
  }


  const WD_PREFIX_RE = /^(maanantai|tiistai|keskiviikko|torstai|perjantai|lauantai|sunnuntai|ma|ti|ke|to|pe|la|su)\.?\s+/i;

  function stripWeekdayPrefix(text) {
    return String(text).trim().replace(WD_PREFIX_RE, "");
  }

  /** Excel serial → local Date, only for real date serials (not hours-fractions or small numbers). */
  function serialToDateSafe(n) {
    if (typeof n !== "number" || !isFinite(n)) return null;
    if (n < 20000 || n >= 80000) return null;
    if (Math.abs(n - Math.round(n)) > 0.001) return null;
    const utc = Date.UTC(1899, 11, 30) + Math.round(n) * 86400000;
    const d = new Date(utc);
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }

  /**
   * A cell value that is clearly one date: JS Date, Excel date serial, or text
   * '12.10.2026' / '12.10.26' / '2026-10-12' / 'Ma 12.10.2026'. No year → null
   * (see parsePartialDM).
   */
  function parseDateValue(v) {
    if (v == null || v === "") return null;
    if (v instanceof Date && !isNaN(v.getTime())) {
      return new Date(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate());
    }
    if (typeof v === "number") return serialToDateSafe(v);
    if (typeof v === "string") {
      const s = stripWeekdayPrefix(v).trim();
      let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
      if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
      m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})\.?$/);
      if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
      m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2})\.?$/);
      if (m) {
        const yy = +m[3];
        return new Date(yy >= 70 ? 1900 + yy : 2000 + yy, +m[2] - 1, +m[1]);
      }
    }
    return null;
  }

  /** '12.10.' / '12.10' with no year. */
  function parsePartialDM(v) {
    if (typeof v !== "string") return null;
    const s = stripWeekdayPrefix(v).trim();
    const m = s.match(/^(\d{1,2})\.(\d{1,2})\.?$/);
    if (!m) return null;
    const day = +m[1], month = +m[2];
    if (day < 1 || day > 31 || month < 1 || month > 12) return null;
    return { day: day, month: month };
  }

  function isWeekdayText(v) {
    if (v == null || typeof v === "number") return false;
    const s = String(v).trim().toLowerCase();
    for (let i = 0; i < WEEKDAYS_FI.length; i++) {
      if (WEEKDAYS_FI[i].toLowerCase() === s) return true;
    }
    return false;
  }

  function isTunnitText(v) {
    if (v == null || typeof v === "number") return false;
    const s = String(v).trim().toLowerCase();
    return s === "tunnit" || s.indexOf("tunnit ") === 0 || s.indexOf("tunnit\n") === 0;
  }

  function dayDiff(a, b) {
    const ua = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
    const ub = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
    return Math.round((ub - ua) / 86400000);
  }

  function dateKey(d) {
    return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
  }

  function addDays(d, n) {
    const x = new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
    return x;
  }

  function weekdayFi(d) {
    // JS: 0=Sun..6=Sat → Mon-first index
    const i = (d.getDay() + 6) % 7;
    return WEEKDAYS_FI[i];
  }

  // --- Finnish holidays (same 13 as v6) ---
  function easterSunday(y) {
    const a = y % 19;
    const b = Math.floor(y / 100);
    const c = y % 100;
    const d = Math.floor(b / 4);
    const e = b % 4;
    const f = Math.floor((b + 8) / 25);
    const g = Math.floor((b - f + 1) / 3);
    const h = (19 * a + b - d - g + 15) % 30;
    const i = Math.floor(c / 4);
    const k = c % 4;
    const l = (32 + 2 * e + 2 * i - h - k) % 7;
    const m = Math.floor((a + 11 * h + 22 * l) / 451);
    const month = Math.floor((h + l - 7 * m + 114) / 31);
    const day = ((h + l - 7 * m + 114) % 31) + 1;
    return new Date(y, month - 1, day);
  }

  function saturdayOnOrAfter(y, monthIndex0, day) {
    const start = new Date(y, monthIndex0, day);
    const wd = (start.getDay() + 6) % 7; // Mon=0..Sun=6; Sat=5
    const delta = (5 - wd + 7) % 7;
    return addDays(start, delta);
  }

  function holidaysForYear(y) {
    const e = easterSunday(y);
    const map = {};
    function put(d, name) {
      map[dateKey(d)] = name;
    }
    put(new Date(y, 0, 1), "Uudenvuodenpäivä");
    put(new Date(y, 0, 6), "Loppiainen");
    put(addDays(e, -2), "Pitkäperjantai");
    put(e, "Pääsiäispäivä");
    put(addDays(e, 1), "2. pääsiäispäivä");
    put(new Date(y, 4, 1), "Vappu");
    put(addDays(e, 39), "Helatorstai");
    put(addDays(e, 49), "Helluntaipäivä");
    put(saturdayOnOrAfter(y, 5, 20), "Juhannuspäivä");
    put(saturdayOnOrAfter(y, 9, 31), "Pyhäinpäivä");
    put(new Date(y, 11, 6), "Itsenäisyyspäivä");
    put(new Date(y, 11, 25), "Joulupäivä");
    put(new Date(y, 11, 26), "Tapaninpäivä");
    return map;
  }

  /** TES §21 aattopäivät: pääsiäislauantai, juhannusaatto, jouluaatto (else ""). */
  function aattoName(d) {
    if (!d) return "";
    const y = d.getFullYear();
    const k = dateKey(d);
    if (k === dateKey(addDays(easterSunday(y), -1))) return "Pääsiäislauantai";
    if (k === dateKey(addDays(saturdayOnOrAfter(y, 5, 20), -1))) return "Juhannusaatto";
    if (d.getMonth() === 11 && d.getDate() === 24) return "Jouluaatto";
    return "";
  }

  function holidayName(d) {
    const y = d.getFullYear();
    const map = Object.assign({}, holidaysForYear(y), holidaysForYear(y + 1), holidaysForYear(y - 1));
    return map[dateKey(d)] || "";
  }

  // --- Lisätyö / ylityö (overtime) ---
  /**
   * Default overtime rules. All base numbers live here so they can be changed for another
   * year / TES. Source: employer's 2026 'Työajan laskenta' table (TES VML, 3-week jakso).
   * Users can still override a single jakso's thresholds in the app (stored on-device).
   *   periodDays           rules apply only to jaksos of this length (others: no default)
   *   baseMin              normal target 114:45 → 50 % starts above this when no arkipyhät
   *   holidayReductionMin  each listed holiday on Mon–Fri inside the jakso lowers by 8 h
   *   limit100Min          100 % starts above 132:45 (never lowered)
   *   lisatyoHolidays      lower the lisätyö threshold
   *   ylityoHolidays       also lower the ylityö (50 %) threshold
   */
  const OT_CONFIG = {
    periodDays: 21,
    baseMin: 114 * 60 + 45,
    holidayReductionMin: 8 * 60,
    limit100Min: 132 * 60 + 45,
    lisatyoHolidays: [
      "jouluaatto",
      "joulupaiva",
      "tapaninpaiva",
      "uudenvuodenpaiva",
      "loppiainen",
      "pitkaperjantai",
      "toinenPaasiaispaiva",
      "vappu",
      "juhannusaatto",
      "itsenaisyyspaiva",
    ],
    ylityoHolidays: ["uudenvuodenpaiva", "vappu", "itsenaisyyspaiva"],
    // Luottamusmies: fixed hours added to LM once per jakso. Not scaled for a short jakso.
    stewardMin: 8 * 60,
  };

  /** Holiday id → { name, date(year) }. Easter computed algorithmically (works for any year). */
  const OT_HOLIDAY_DEFS = {
    uudenvuodenpaiva: { name: "Uudenvuodenpäivä", date: (y) => new Date(y, 0, 1) },
    loppiainen: { name: "Loppiainen", date: (y) => new Date(y, 0, 6) },
    pitkaperjantai: { name: "Pitkäperjantai", date: (y) => addDays(easterSunday(y), -2) },
    toinenPaasiaispaiva: { name: "2. pääsiäispäivä", date: (y) => addDays(easterSunday(y), 1) },
    vappu: { name: "Vappu", date: (y) => new Date(y, 4, 1) },
    helatorstai: { name: "Helatorstai", date: (y) => addDays(easterSunday(y), 39) },
    juhannusaatto: {
      name: "Juhannusaatto",
      // Friday between 19–25 June
      date: (y) => {
        const jun19 = new Date(y, 5, 19);
        return addDays(jun19, (5 - jun19.getDay() + 7) % 7);
      },
    },
    itsenaisyyspaiva: { name: "Itsenäisyyspäivä", date: (y) => new Date(y, 11, 6) },
    jouluaatto: { name: "Jouluaatto", date: (y) => new Date(y, 11, 24) },
    joulupaiva: { name: "Joulupäivä", date: (y) => new Date(y, 11, 25) },
    tapaninpaiva: { name: "Tapaninpäivä", date: (y) => new Date(y, 11, 26) },
  };

  /** Listed holidays for one year with flags from the config. */
  function overtimeHolidaysForYear(y, config) {
    const cfg = config || OT_CONFIG;
    return cfg.lisatyoHolidays
      .filter((id) => OT_HOLIDAY_DEFS[id])
      .map((id) => ({
        id: id,
        d: OT_HOLIDAY_DEFS[id].date(y),
        name: OT_HOLIDAY_DEFS[id].name,
        lowersYlityo: cfg.ylityoHolidays.indexOf(id) >= 0,
      }));
  }

  /**
   * Default thresholds for a jakso. Returns null unless the jakso length equals
   * config.periodDays (rules are only defined for the 3-week jakso).
   * start: Date or 'yyyy-mm-dd' / 'd.m.yyyy'.
   *   lisaK = lisätyö starts, yliK = ylityö 50 % starts, sataK = ylityö 100 % starts.
   */
  /**
   * Luottamusmies hours for one jakso. Full stewardMin at n = periodDays.
   * Scales with laskentapäivät: round(stewardMin * n / 21). n omitted = full amount.
   * Same minutes are added to LM and to Yritys.
   */
  function stewardBonus(on, config, n) {
    if (!on) return 0;
    const cfg = config || OT_CONFIG;
    const full = cfg.stewardMin || 0;
    const days = cfg.periodDays || 21;
    if (n == null) return full;
    return Math.round((full * n) / days);
  }

  /**
   * Steward flags are stored by person name, not by jakso. map is { [name]: 1 }.
   * Returns a new map.
   */
  function stewardMapSet(map, name, on) {
    const next = Object.assign({}, map || {});
    const key = String(name || "");
    if (!key) return next;
    if (on) next[key] = 1;
    else delete next[key];
    return next;
  }

  /**
   * Erillinen palkkio €/kk (not per jakso): luottamusmies §18.17 and työsuojeluvaltuutettu §21
   * have the same amounts. TSV gets no extra hours. TES §21: not paid to the same person at
   * the same time.
   */
  const PALKKIO_RATES = [
    { from: "0000-00-00", eur: 161 },
    { from: "2026-09-01", eur: 165 },
    { from: "2027-08-01", eur: 169 },
  ];
  function palkkioFor(key) {
    let r = PALKKIO_RATES[0];
    PALKKIO_RATES.forEach(function (x) {
      if (key >= x.from) r = x;
    });
    return r.eur;
  }
  /** Distinct monthly rates over a jakso [startKey, endKey] (a jakso can cross a rate change). */
  function palkkioRatesInRange(startKey, endKey) {
    const out = [{ from: startKey, eur: palkkioFor(startKey) }];
    PALKKIO_RATES.forEach(function (x) {
      if (x.from > startKey && x.from <= endKey) out.push({ from: x.from, eur: x.eur });
    });
    return out;
  }

  function overtimeThresholds(start, dayCount, config) {
    const cfg = config || OT_CONFIG;
    const s = toDate(start);
    if (!s || dayCount !== cfg.periodDays) return null;
    const end = addDays(s, dayCount - 1);
    const startKey = dateKey(s),
      endKey = dateKey(end);
    const holidays = [];
    for (let y = s.getFullYear(); y <= end.getFullYear(); y++) {
      overtimeHolidaysForYear(y, cfg).forEach(function (h) {
        const k = dateKey(h.d);
        const wd = h.d.getDay(); // 0=Sun, 6=Sat
        if (k >= startKey && k <= endKey && wd >= 1 && wd <= 5) {
          holidays.push({ date: k, name: h.name, lowersYlityo: h.lowersYlityo });
        }
      });
    }
    holidays.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
    const lisaK = cfg.baseMin - cfg.holidayReductionMin * holidays.length;
    const yliK = cfg.baseMin - cfg.holidayReductionMin * holidays.filter((h) => h.lowersYlityo).length;
    return { lisaK: lisaK, yliK: yliK, sataK: cfg.limit100Min, holidays: holidays, start: startKey, end: endKey };
  }

  /**
   * Apply a user override {lisaK?, yliK?, sataK?} (minutes) on top of defaults (may be null).
   * Missing lisätyö start falls back to the default, or to the 50 % start when there is no default.
   * Returns effective thresholds (lisaK ≤ yliK ≤ sataK) or null if 50 %/100 % are still unknown.
   */
  function applyOvertimeOverride(defaults, override) {
    const o = override || {};
    const pick = (k) => (o[k] != null ? o[k] : defaults ? defaults[k] : null);
    const yliK = pick("yliK");
    const sataK = pick("sataK");
    if (yliK == null || sataK == null) return null;
    let lisaK = pick("lisaK");
    if (lisaK == null) lisaK = yliK;
    return Object.assign({}, defaults || { holidays: [] }, {
      lisaK: Math.min(lisaK, yliK),
      yliK: yliK,
      sataK: Math.max(sataK, yliK),
      overridden: o.lisaK != null || o.yliK != null || o.sataK != null,
    });
  }

  /** Parse 'h:mm' / 'h.mm' / 'h' (hours may exceed 24) → minutes, or null */
  function parseHM(text) {
    if (text == null) return null;
    const t = String(text).trim().replace(",", ".");
    let m = t.match(/^(\d{1,3})[:.](\d{2})$/);
    if (m) return +m[2] < 60 ? +m[1] * 60 + +m[2] : null;
    m = t.match(/^(\d{1,3})$/);
    return m ? +m[1] * 60 : null;
  }

  function clampMin(x, lo, hi) {
    return Math.min(Math.max(x, lo), hi);
  }

  /**
   * Split LM total (minutes) into lisätyö (lisaK → yliK), ylityö 50 % (yliK → 132:45)
   * and ylityö 100 % (above 132:45). null if no thresholds (non-21-day jakso).
   */
  function overtimeSplit(lmMin, th) {
    if (!th) return null;
    const lm = lmMin || 0;
    return {
      lisa: clampMin(lm, th.lisaK, th.yliK) - th.lisaK,
      yli50: clampMin(lm, th.yliK, th.sataK) - th.yliK,
      yli100: Math.max(0, lm - th.sataK),
    };
  }

  /**
   * Keskeytynyt työjakso: 'Suhteellisten työtuntinormien taulukot' 1 and 2,
   * embedded exactly as printed (N = 1..21). Not computed by formula.
   * Columns: II 50 % threshold, III 100 % threshold, IV max whole hours at 50 %,
   * then (lisätyö threshold, max whole lisätyö hours) pairs.
   *   Table 1 lisätyö columns: 106:45, 98:45, 90:45, 82:45 (ylityö base 114:45)
   *   Table 2 lisätyö columns: 98:45, 90:45, 82:45, 74:45 (ylityö base 106:45)
   */
  const INT_TABLE_SRC = {
    1: [
      "1 | 5 t 28 min | 6 t 19 min | 1 | 5 t 05 min | 0 | 4 t 42 min | 1 | 4 t 19 min | 1 | 3 t 56 min | 2",
      "2 | 10 t 56 min | 12 t 39 min | 2 | 10 t 10 min | 1 | 9 t 24 min | 2 | 8 t 39 min | 2 | 7 t 53 min | 3",
      "3 | 16 t 24 min | 18 t 58 min | 3 | 15 t 15 min | 1 | 14 t 06 min | 2 | 12 t 58 min | 3 | 11 t 49 min | 5",
      "4 | 21 t 51 min | 25 t 17 min | 3 | 20 t 20 min | 2 | 18 t 49 min | 3 | 17 t 17 min | 5 | 15 t 46 min | 6",
      "5 | 27 t 19 min | 31 t 36 min | 4 | 25 t 25 min | 2 | 23 t 31 min | 4 | 21 t 36 min | 6 | 19 t 42 min | 8",
      "6 | 32 t 47 min | 37 t 56 min | 5 | 30 t 30 min | 2 | 28 t 13 min | 5 | 25 t 56 min | 7 | 23 t 39 min | 9",
      "7 | 38 t 15 min | 44 t 15 min | 6 | 35 t 35 min | 3 | 32 t 55 min | 5 | 30 t 15 min | 8 | 27 t 35 min | 11",
      "8 | 43 t 43 min | 50 t 34 min | 7 | 40 t 40 min | 3 | 37 t 37 min | 6 | 34 t 34 min | 9 | 31 t 31 min | 12",
      "9 | 49 t 11 min | 56 t 54 min | 8 | 45 t 45 min | 3 | 42 t 19 min | 7 | 38 t 54 min | 10 | 35 t 28 min | 14",
      "10 | 54 t 39 min | 63 t 13 min | 9 | 50 t 50 min | 4 | 47 t 01 min | 8 | 43 t 13 min | 11 | 39 t 24 min | 15",
      "11 | 60 t 06 min | 69 t 32 min | 9 | 55 t 55 min | 4 | 51 t 44 min | 8 | 47 t 32 min | 13 | 43 t 21 min | 17",
      "12 | 65 t 34 min | 75 t 51 min | 10 | 61 t 00 min | 5 | 56 t 26 min | 9 | 51 t 51 min | 14 | 47 t 17 min | 18",
      "13 | 71 t 02 min | 82 t 11 min | 11 | 66 t 05 min | 5 | 61 t 08 min | 10 | 56 t 11 min | 15 | 51 t 14 min | 20",
      "14 | 76 t 30 min | 88 t 30 min | 12 | 71 t 10 min | 5 | 65 t 50 min | 11 | 60 t 30 min | 16 | 55 t 10 min | 21",
      "15 | 81 t 58 min | 94 t 49 min | 13 | 76 t 15 min | 6 | 70 t 32 min | 11 | 64 t 49 min | 17 | 59 t 06 min | 23",
      "16 | 87 t 26 min | 101 t 09 min | 14 | 81 t 20 min | 6 | 75 t 14 min | 12 | 69 t 09 min | 18 | 63 t 03 min | 24",
      "17 | 92 t 54 min | 107 t 28 min | 15 | 86 t 25 min | 6 | 79 t 56 min | 13 | 73 t 28 min | 19 | 66 t 59 min | 26",
      "18 | 98 t 21 min | 113 t 47 min | 15 | 91 t 30 min | 7 | 84 t 39 min | 14 | 77 t 47 min | 21 | 70 t 56 min | 27",
      "19 | 103 t 49 min | 120 t 06 min | 16 | 96 t 35 min | 7 | 89 t 21 min | 14 | 82 t 06 min | 22 | 74 t 52 min | 29",
      "20 | 109 t 17 min | 126 t 26 min | 17 | 101 t 40 min | 8 | 94 t 03 min | 15 | 86 t 26 min | 23 | 78 t 49 min | 30",
      "21 | 114 t 45 min | 132 t 45 min | 18 | 106 t 45 min | 8 | 98 t 45 min | 16 | 90 t 45 min | 24 | 82 t 45 min | 32",
    ],
    2: [
      "1 | 5 t 05 min | 6 t 19 min | 1 | 4 t 42 min | 0 | 4 t 19 min | 1 | 3 t 56 min | 1 | 3 t 34 min | 2",
      "2 | 10 t 10 min | 12 t 39 min | 2 | 9 t 24 min | 1 | 8 t 39 min | 2 | 7 t 53 min | 2 | 7 t 07 min | 3",
      "3 | 15 t 15 min | 18 t 58 min | 4 | 14 t 06 min | 1 | 12 t 58 min | 2 | 11 t 49 min | 3 | 10 t 41 min | 5",
      "4 | 20 t 20 min | 25 t 17 min | 5 | 18 t 49 min | 2 | 17 t 17 min | 3 | 15 t 46 min | 5 | 14 t 14 min | 6",
      "5 | 25 t 25 min | 31 t 36 min | 6 | 23 t 31 min | 2 | 21 t 36 min | 4 | 19 t 42 min | 6 | 17 t 48 min | 8",
      "6 | 30 t 30 min | 37 t 56 min | 7 | 28 t 13 min | 2 | 25 t 56 min | 5 | 23 t 39 min | 7 | 21 t 21 min | 9",
      "7 | 35 t 35 min | 44 t 15 min | 9 | 32 t 55 min | 3 | 30 t 15 min | 5 | 27 t 35 min | 8 | 24 t 55 min | 11",
      "8 | 40 t 40 min | 50 t 34 min | 10 | 37 t 37 min | 3 | 34 t 34 min | 6 | 31 t 31 min | 9 | 28 t 29 min | 12",
      "9 | 45 t 45 min | 56 t 54 min | 11 | 42 t 19 min | 3 | 38 t 54 min | 7 | 35 t 28 min | 10 | 32 t 02 min | 14",
      "10 | 50 t 50 min | 63 t 13 min | 12 | 47 t 01 min | 4 | 43 t 13 min | 8 | 39 t 24 min | 11 | 35 t 36 min | 15",
      "11 | 55 t 55 min | 69 t 32 min | 14 | 51 t 44 min | 4 | 47 t 32 min | 8 | 43 t 21 min | 13 | 39 t 09 min | 17",
      "12 | 61 t 00 min | 75 t 51 min | 15 | 56 t 26 min | 5 | 51 t 51 min | 9 | 47 t 17 min | 14 | 42 t 43 min | 18",
      "13 | 66 t 05 min | 82 t 11 min | 16 | 61 t 08 min | 5 | 56 t 11 min | 10 | 51 t 14 min | 15 | 46 t 16 min | 20",
      "14 | 71 t 10 min | 88 t 30 min | 17 | 65 t 50 min | 5 | 60 t 30 min | 11 | 55 t 10 min | 16 | 49 t 50 min | 21",
      "15 | 76 t 15 min | 94 t 49 min | 19 | 70 t 32 min | 6 | 64 t 49 min | 11 | 59 t 06 min | 17 | 53 t 24 min | 23",
      "16 | 81 t 20 min | 101 t 09 min | 20 | 75 t 14 min | 6 | 69 t 09 min | 12 | 63 t 03 min | 18 | 56 t 57 min | 24",
      "17 | 86 t 25 min | 107 t 28 min | 21 | 79 t 56 min | 6 | 73 t 28 min | 13 | 66 t 59 min | 19 | 60 t 31 min | 26",
      "18 | 91 t 30 min | 113 t 47 min | 22 | 84 t 39 min | 7 | 77 t 47 min | 14 | 70 t 56 min | 21 | 64 t 04 min | 27",
      "19 | 96 t 35 min | 120 t 06 min | 24 | 89 t 21 min | 7 | 82 t 06 min | 14 | 74 t 52 min | 22 | 67 t 38 min | 29",
      "20 | 101 t 40 min | 126 t 26 min | 25 | 94 t 03 min | 8 | 86 t 26 min | 15 | 78 t 49 min | 23 | 71 t 11 min | 30",
      "21 | 106 t 45 min | 132 t 45 min | 26 | 98 t 45 min | 8 | 90 t 45 min | 16 | 82 t 45 min | 24 | 74 t 45 min | 32",
    ],
  };
  const INT_LISA_KEYS = {
    1: [106 * 60 + 45, 98 * 60 + 45, 90 * 60 + 45, 82 * 60 + 45],
    2: [98 * 60 + 45, 90 * 60 + 45, 82 * 60 + 45, 74 * 60 + 45],
  };

  function parsePrintedCell(cell) {
    const t = cell.trim();
    if (/^\d+$/.test(t)) return { capH: +t };
    const m = t.match(/^(\d+)\s*t\s*(\d+)\s*min$/);
    if (!m) throw new Error("Bad table cell: " + cell);
    return { min: +m[1] * 60 + +m[2] };
  }

  function buildIntTables() {
    const out = { 1: [], 2: [] };
    [1, 2].forEach(function (table) {
      INT_TABLE_SRC[table].forEach(function (line) {
        const c = line.split("|").map(parsePrintedCell);
        if (c.length !== 12) throw new Error("row width " + c.length);
        const lisa = [];
        for (let i = 0; i < 4; i++) {
          lisa.push({ k: c[4 + i * 2].min, capH: c[5 + i * 2].capH });
        }
        out[table][+line.split("|")[0].trim()] = {
          yli50: c[1].min,
          k100: c[2].min,
          cap50H: c[3].capH,
          lisa: lisa,
        };
      });
    });
    return out;
  }
  const INT_TABLES = buildIntTables();

  /**
   * Norms for N laskentapäiviä given the full jakso's lisäK and yliK (minutes).
   * null when the pair is not in the printed tables (caller shows 'ei taulukossa').
   * When lisäK === yliK there is no lisätyö band (lisaK/lisaCap null).
   */
  function interruptedNorms(lisaK, yliK, n) {
    const table = yliK === 114 * 60 + 45 ? 1 : yliK === 106 * 60 + 45 ? 2 : 0;
    if (!table || n < 1 || n > 21) return null;
    const row = INT_TABLES[table][n];
    if (!row) return null;
    const noLisa = lisaK === yliK;
    let lisa = null;
    if (!noLisa) {
      const idx = INT_LISA_KEYS[table].indexOf(lisaK);
      if (idx < 0) return null;
      lisa = row.lisa[idx];
    }
    return {
      table: table,
      n: n,
      yliK: row.yli50,
      k100: row.k100,
      cap50: row.cap50H * 60,
      cap50H: row.cap50H,
      lisaK: lisa ? lisa.k : null,
      lisaCap: lisa ? lisa.capH * 60 : null,
      lisaCapH: lisa ? lisa.capH : null,
      noLisa: noLisa,
    };
  }

  /**
   * Lisätyö / 50 % / 100 % for a person with N laskentapäiviä.
   * N = 21 uses the ordinary split (caps equal the band widths, overrides included).
   * An active Muuta override is not scaled: N < 21 → unsupported 'override'.
   * Unlisted thresholds → unsupported 'table'.
   * Caps cut the paid hours; time between a cap and the next threshold is paid as neither.
   */
  function interruptedSplit(lmMin, th, n) {
    if (!th) return null;
    n = n || 21;
    if (n === 21) {
      const s = overtimeSplit(lmMin, th);
      return {
        lisa: s.lisa,
        yli50: s.yli50,
        yli100: s.yli100,
        n: 21,
        unsupported: null,
        cap50Hit: false,
        capLisaHit: false,
        norms: th.overridden ? null : interruptedNorms(th.lisaK, th.yliK, 21),
      };
    }
    if (th.overridden) {
      return { lisa: null, yli50: null, yli100: null, n: n, unsupported: "override", cap50Hit: false, capLisaHit: false, norms: null };
    }
    const norms = interruptedNorms(th.lisaK, th.yliK, n);
    if (!norms) {
      return { lisa: null, yli50: null, yli100: null, n: n, unsupported: "table", cap50Hit: false, capLisaHit: false, norms: null };
    }
    const lm = lmMin || 0;
    let lisa = 0;
    let capLisaHit = false;
    if (!norms.noLisa) {
      const raw = clampMin(lm, norms.lisaK, norms.yliK) - norms.lisaK;
      capLisaHit = raw > norms.lisaCap;
      lisa = Math.min(norms.lisaCap, raw);
    }
    const raw50 = clampMin(lm, norms.yliK, norms.k100) - norms.yliK;
    const cap50Hit = raw50 > norms.cap50;
    return {
      lisa: lisa,
      yli50: Math.min(norms.cap50, raw50),
      yli100: Math.max(0, lm - norms.k100),
      n: n,
      unsupported: null,
      cap50Hit: cap50Hit,
      capLisaHit: capLisaHit,
      norms: norms,
    };
  }

    /** Interval overlap length (same logic as Matching R15..) */
  function segOverlap(start, end, segStart, segEnd) {
    // Mirror nested IFs exactly
    if (start < segStart && start < segEnd && end > segStart && end < segEnd) return end - segStart;
    if (start >= segStart && start < segEnd && end > segStart && end < segEnd) return end - start;
    if (start >= segStart && start < segEnd && end > segStart && end >= segEnd) return segEnd - start;
    if (start < segStart && start < segEnd && end > segStart && end >= segEnd) return segEnd - segStart;
    return 0;
  }

  function allSegOverlaps(dayStart, dayEnd, segs) {
    return segs.map(([a, b]) => segOverlap(dayStart, dayEnd, a, b));
  }

  /**
   * Night hours AB = SUM(R:S)+IF(early)T + SUM(W:Y)+IF(early next)Z
   * Early: if S (or Y) equals full 2h AND R+S (or X+Y) > 2h, include T (or Z).
   */
  function nightMinutes(curSegs, nextSegs) {
    const R = curSegs[0],
      S = curSegs[1],
      T = curSegs[2];
    const W = curSegs[5];
    const X = nextSegs[0],
      Y = nextSegs[1],
      Z = nextSegs[2];
    // TES §18: work started at the latest 04:00 that runs past 06:00 is night work
    // until 12:00 (no rest info in the form). S === 2 h means the shift covers 04–06.
    let n = R + S;
    if (S === EARLY_FULL) n += T;
    n += W + X + Y;
    if (Y === EARLY_FULL) n += Z;
    return n;
  }

  /** 25% bonus AC */
  function bonus25(weekday, holidayToday, holidayTomorrow, curSegs, nextSegs, aattoToday, aattoTomorrow) {
    // AB10=Perjantai, AC10=Lauantai. TES §19: not on pääsiäislauantai, not when aattopäivänlisä is paid.
    if (weekday === "Perjantai" && !holidayTomorrow && !aattoTomorrow) {
      return nextSegs[2] + nextSegs[3]; // Z+AA
    }
    if (weekday === "Lauantai" && !holidayToday && !aattoToday) {
      return curSegs[2] + curSegs[3]; // T+U
    }
    return 0;
  }

  /** 100% bonus AD */
  function bonus100(weekday, holidayToday, holidayTomorrow, curSegs, nextSegs) {
    let n = 0;
    const isSun = weekday === "Sunnuntai";
    if (holidayToday || isSun) {
      n += curSegs[0] + curSegs[1] + curSegs[2] + curSegs[3] + curSegs[4] + curSegs[5]; // R:W
    } else if ((holidayTomorrow || false) && !holidayToday && weekday !== "Sunnuntai") {
      // next holiday or… AD uses C16=$AD$10 (Sunday) OR D16<>""
      // handled below via holidayTomorrow || tomorrowIsSunday
    }
    // Re-read AD:
    // IF(OR(D15<>"", C15=Sunday), SUM(R:W),
    //    IF(AND(OR(D16<>"", C16=Sunday), D15="", C15<>Sunday), SUM(V:W), 0))
    // + IF(OR(D16<>"", C16=Sunday), SUM(X:AA), 0)
    // So day-before uses V:W only when tomorrow is holiday OR Sunday.
    if (!(holidayToday || isSun)) {
      // second branch needs tomorrow info — caller passes holidayTomorrow and tomorrow weekday
    }
    return n; // completed in computeDay with tomorrow weekday
  }

  function bonus100Full(weekday, tomorrowWeekday, holidayToday, holidayTomorrow, curSegs, nextSegs) {
    let n = 0;
    const sun = "Sunnuntai";
    if (holidayToday || weekday === sun) {
      n += curSegs.reduce((a, b) => a + b, 0);
    } else if ((holidayTomorrow || tomorrowWeekday === sun) && !holidayToday && weekday !== sun) {
      n += curSegs[4] + curSegs[5]; // V:W
    }
    if (holidayTomorrow || tomorrowWeekday === sun) {
      n += nextSegs.reduce((a, b) => a + b, 0); // X:AA
    }
    return n;
  }

  function restStatus(restMin) {
    if (restMin == null || restMin === "") return "";
    if (restMin < REST_NOT_ALLOWED_MIN) return "Not Allowed";
    if (restMin < REST_MD_CHECK_MIN) return "Md,s Check";
    return "ok";
  }

  function mroundMin(x) {
    if (x == null) return null;
    return Math.round(x);
  }

  /**
   * Compute one person's days (any length).
   * shifts[i] = {start,end}|null, company[i]=minutes|null, dates[i]=Date
   */
  function computePerson(dates, shifts, companyHours, holidayMap) {
    const days = [];
    const n = dates.length;

    // Precompute end datetime (absolute minutes from epoch-ish) for rest
    const endAbs = new Array(n).fill(null);
    const startAbs = new Array(n).fill(null);

    for (let i = 0; i < n; i++) {
      const d = dates[i];
      const sh = shifts[i];
      const hol = holidayMap[dateKey(d)] || "";
      const wd = weekdayFi(d);
      const tomorrow = i + 1 < n ? dates[i + 1] : addDays(d, 1);
      const holTom = holidayMap[dateKey(tomorrow)] || "";
      const wdTom = weekdayFi(tomorrow);

      let start = null,
        end = null,
        hrs = null,
        night = null,
        b25 = null,
        b100 = null,
        me = null,
        company = null,
        check = null;

      if (sh) {
        start = sh.start;
        end = sh.end;
        const overnight = end <= start;
        hrs = overnight ? end - start + 24 * 60 : end - start;

        // H=start, I=overnight? 24*60 : end, J=0, K=overnight? end : 0
        const H = start;
        const I = overnight ? 24 * 60 : end;
        const J = 0;
        const K = overnight ? end : 0;

        const curSegs = allSegOverlaps(H, I, SEG_CUR);
        const nextSegs = K > 0 ? allSegOverlaps(J, K, SEG_NEXT) : [0, 0, 0, 0];

        night = nightMinutes(curSegs, nextSegs);
        b25 = bonus25(wd, !!hol, !!holTom, curSegs, nextSegs, !!aattoName(d), !!aattoName(tomorrow));
        b100 = bonus100Full(wd, wdTom, !!hol, !!holTom, curSegs, nextSegs);

        me = mroundMin(hrs + night * NIGHT_FACTOR);
        company = companyHours[i] != null ? mroundMin(companyHours[i]) : null;
        check = company != null && me != null ? mroundMin(Math.abs(me - company)) : null;

        // absolute times for rest: date + minutes
        const day0 = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 60000;
        startAbs[i] = day0 + start;
        endAbs[i] = overnight ? day0 + 24 * 60 + end : day0 + end;
      }

      days.push({
        date: dateKey(d),
        dateObj: d,
        weekday: wd,
        special: hol,
        start,
        end,
        hrs,
        night: sh ? night : null,
        b25: sh ? b25 : null,
        b100: sh ? b100 : null,
        me,
        company,
        check,
        restAfter: "",
        restBefore: "",
        restAfterMin: null,
        restBeforeMin: null,
      });
    }

    // Rest AFTER (Final/Matching): only when the NEXT calendar day also has a shift.
    for (let i = 0; i < n - 1; i++) {
      if (endAbs[i] == null || startAbs[i + 1] == null) continue;
      const rest = startAbs[i + 1] - endAbs[i];
      days[i].restAfterMin = rest;
      days[i].restAfter = restStatus(rest);
    }

    // Rest BEFORE (Kaikki/_calc): carry last end across days off; no status without a real previous end.
    let lastEnd = null;
    for (let i = 0; i < n; i++) {
      if (startAbs[i] != null && lastEnd != null) {
        const rest = startAbs[i] - lastEnd;
        days[i].restBeforeMin = rest;
        days[i].restBefore = restStatus(rest);
      }
      if (endAbs[i] != null) lastEnd = endAbs[i];
    }

    return days;
  }

  /**
   * One shift's hours from start/end, same segment rules as an Excel shift.
   * Yritys has no company cell for an app-entered shift, so it follows LM
   * and the difference (Erotus) is 0.
   */
  function computeShiftFigures(date, startMin, endMin, holidayMap) {
    const d = date instanceof Date ? new Date(date.getFullYear(), date.getMonth(), date.getDate()) : toDate(date);
    if (!d || startMin == null || endMin == null) return null;
    const tomorrow = addDays(d, 1);
    const hol = (holidayMap && holidayMap[dateKey(d)]) || "";
    const holTom = (holidayMap && holidayMap[dateKey(tomorrow)]) || "";
    const wd = weekdayFi(d);
    const wdTom = weekdayFi(tomorrow);
    const overnight = endMin <= startMin;
    const hrs = overnight ? endMin - startMin + 24 * 60 : endMin - startMin;
    const I = overnight ? 24 * 60 : endMin;
    const K = overnight ? endMin : 0;
    const curSegs = allSegOverlaps(startMin, I, SEG_CUR);
    const nextSegs = K > 0 ? allSegOverlaps(0, K, SEG_NEXT) : [0, 0, 0, 0];
    const night = nightMinutes(curSegs, nextSegs);
    const b25 = bonus25(wd, !!hol, !!holTom, curSegs, nextSegs, !!aattoName(d), !!aattoName(tomorrow));
    const b100 = bonus100Full(wd, wdTom, !!hol, !!holTom, curSegs, nextSegs);
    const me = mroundMin(hrs + night * NIGHT_FACTOR);
    return {
      start: startMin,
      end: endMin,
      hrs: hrs,
      night: night,
      b25: b25,
      b100: b100,
      me: me,
      company: me,
      check: 0,
    };
  }

  function shiftAbs(date, start, end) {
    const d = date instanceof Date ? date : toDate(date);
    const day0 = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 60000;
    const overnight = end <= start;
    return { startAbs: day0 + start, endAbs: overnight ? day0 + 24 * 60 + end : day0 + end };
  }

  /**
   * Rest gaps including lisävuorot. extras: { date: {start, end} }.
   * A keskeytyspäivä drops the planned shift. Same rules as the Excel rest:
   * restAfter only to the next calendar day, restBefore carries across days off.
   * between = the gap when a planned shift and a lisävuoro share a day.
   */
  function restWithExtras(days, extras, marks) {
    const byDate = {};
    function add(date, dateObj, start, end, kind) {
      const a = shiftAbs(dateObj || date, start, end);
      (byDate[date] = byDate[date] || []).push({ date: date, kind: kind, startAbs: a.startAbs, endAbs: a.endAbs });
    }
    (days || []).forEach(function (d) {
      const off = isKeskeytys(marks && marks[d.date]);
      if (!off && d.start != null) add(d.date, d.dateObj, d.start, d.end, "plan");
      const ex = !off && extras && extras[d.date];
      if (ex && ex.start != null && ex.end != null) add(d.date, d.dateObj, ex.start, ex.end, "extra");
    });
    Object.keys(byDate).forEach(function (date) {
      byDate[date].sort(function (a, b) { return a.startAbs - b.startAbs; });
    });
    const out = {};
    (days || []).forEach(function (d) {
      out[d.date] = { restAfter: "", restAfterMin: null, restBefore: "", restBeforeMin: null, between: "", betweenMin: null };
      const list = byDate[d.date];
      if (list && list.length === 2) {
        const gap = list[1].startAbs - list[0].endAbs;
        out[d.date].betweenMin = gap;
        out[d.date].between = restStatus(gap);
      }
    });
    for (let i = 0; i < days.length - 1; i++) {
      const a = byDate[days[i].date];
      const b = byDate[days[i + 1].date];
      if (!a || !b) continue;
      const rest = b[0].startAbs - a[a.length - 1].endAbs;
      out[days[i].date].restAfterMin = rest;
      out[days[i].date].restAfter = restStatus(rest);
    }
    let lastEnd = null;
    for (let i = 0; i < days.length; i++) {
      const list = byDate[days[i].date];
      if (list && lastEnd != null) {
        const rest = list[0].startAbs - lastEnd;
        out[days[i].date].restBeforeMin = rest;
        out[days[i].date].restBefore = restStatus(rest);
      }
      if (list) lastEnd = list[list.length - 1].endAbs;
    }
    return out;
  }

  /**
   * Parse company Jakso form sheet (Taul1 / first sheet) into structured data.
   * Accepts a SheetJS workbook or a 2D array sheet.
   */
  function parseJaksoForm(workbook, opts) {
    opts = opts || {};
    const sheetName = workbook.SheetNames.includes("Taul1")
      ? "Taul1"
      : workbook.SheetNames.includes("Jakso_form")
        ? "Jakso_form"
        : workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    function raw(addr) {
      const c = sheet[addr];
      if (!c) return null;
      return c.v !== undefined ? c.v : null;
    }

    function hasFormula(addr) {
      const c = sheet[addr];
      return !!(c && c.f);
    }

    /** 0-based column index → Excel letter (0=A) */
    function colLetter(idx0) {
      let n = idx0 + 1;
      let s = "";
      while (n > 0) {
        const m = (n - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        n = Math.floor((n - 1) / 26);
      }
      return s;
    }

    function sheetMaxCol0() {
      if (sheet["!ref"] && typeof XLSX !== "undefined" && XLSX.utils && XLSX.utils.decode_range) {
        return XLSX.utils.decode_range(sheet["!ref"]).e.c;
      }
      let max = 2;
      Object.keys(sheet).forEach(function (k) {
        if (k.charAt(0) === "!") return;
        const m = k.match(/^([A-Z]+)/);
        if (!m) return;
        let n = 0;
        const letters = m[1];
        for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
        if (n - 1 > max) max = n - 1;
      });
      return max;
    }

    const weekdaySet = {};
    WEEKDAYS_FI.forEach(function (w) {
      weekdaySet[w] = true;
    });

    // Layout is found by searching, not by fixed cells.
    // Strategy 1: a cell "Jakso 16" — the start date is directly under it,
    // or 1–2 cells further down (the row in between is usually the weekday).
    // Then: a label (Alku / Aloitus / …) with a date beside it;
    // then a run of 7+ consecutive daily dates; then any date in the top-left.
    // If nothing yields a start date, return { needsStartDate: true } so the UI can ask.
    const forcedStart = opts.startDate ? parseDateValue(opts.startDate) || toDate(opts.startDate) : null;
    const MAX_SCAN_ROW = 200;

    function forEachValue(limitR, limitC, fn) {
      if (sheet["!ref"] && typeof XLSX !== "undefined" && XLSX.utils && XLSX.utils.decode_range) {
        const rg = XLSX.utils.decode_range(sheet["!ref"]);
        const r2 = Math.min(limitR, rg.e.r);
        const c2 = Math.min(limitC, rg.e.c);
        for (let r = 0; r <= r2; r++) {
          for (let c = 0; c <= c2; c++) {
            const cell = sheet[XLSX.utils.encode_cell({ r: r, c: c })];
            if (cell && cell.v != null && cell.v !== "") fn(r, c, cell.v);
          }
        }
        return;
      }
      Object.keys(sheet).forEach(function (k) {
        if (k.charAt(0) === "!") return;
        const cell = sheet[k];
        if (!cell || cell.v == null || cell.v === "") return;
        const m = k.match(/^([A-Z]+)(\d+)$/);
        if (!m) return;
        let c = 0;
        for (let i = 0; i < m[1].length; i++) c = c * 26 + (m[1].charCodeAt(i) - 64);
        const r = +m[2] - 1;
        if (r <= limitR && c - 1 <= limitC) fn(r, c - 1, cell.v);
      });
    }

    let hintYear = null;
    forEachValue(120, 25, function (_r, _c, v) {
      if (hintYear != null) return;
      const d = parseDateValue(v);
      if (d) hintYear = d.getFullYear();
    });

    /** row(1-based) → Date, with 'd.m.' years filled in from neighbours */
    function resolvedDatesInCol(col0, maxRow) {
      const items = [];
      for (let r = 1; r <= maxRow; r++) {
        const v = raw(colLetter(col0) + r);
        const full = parseDateValue(v);
        if (full) {
          items.push({ row: r, date: full, day: full.getDate(), month: full.getMonth() + 1, year: full.getFullYear() });
          continue;
        }
        const part = parsePartialDM(v);
        if (part) items.push({ row: r, date: null, day: part.day, month: part.month, year: null });
      }
      function pick(day, month, neighbour, dir) {
        let best = null;
        for (let y = neighbour.getFullYear() - 1; y <= neighbour.getFullYear() + 1; y++) {
          const dt = new Date(y, month - 1, day);
          if (dt.getMonth() !== month - 1 || dt.getDate() !== day) continue;
          const diff = dir === 1 ? dayDiff(neighbour, dt) : dayDiff(dt, neighbour);
          if (diff <= 0 || diff > 40) continue;
          if (!best || diff < best.diff) best = { dt: dt, diff: diff };
        }
        return best ? best.dt : null;
      }
      for (let i = 0; i < items.length; i++) {
        if (items[i].date) continue;
        let prev = null;
        for (let j = i - 1; j >= 0; j--) if (items[j].date) { prev = items[j]; break; }
        if (!prev) continue;
        const dt = pick(items[i].day, items[i].month, prev.date, 1);
        if (dt) items[i].date = dt;
      }
      for (let i = items.length - 1; i >= 0; i--) {
        if (items[i].date) continue;
        let next = null;
        for (let j = i + 1; j < items.length; j++) if (items[j].date) { next = items[j]; break; }
        if (!next) continue;
        const dt = pick(items[i].day, items[i].month, next.date, -1);
        if (dt) items[i].date = dt;
      }
      if (hintYear != null) {
        items.forEach(function (it, i) {
          if (it.date) return;
          let dt = new Date(hintYear, it.month - 1, it.day);
          if (i > 0 && items[i - 1].date && dayDiff(items[i - 1].date, dt) < 0) {
            dt = new Date(hintYear + 1, it.month - 1, it.day);
          }
          it.date = dt;
        });
      }
      const map = {};
      items.forEach(function (it) {
        if (it.date && !isNaN(it.date.getTime())) map[it.row] = it.date;
      });
      return map;
    }

    function findDateRun() {
      let best = null;
      for (let c = 0; c <= 15; c++) {
        const map = resolvedDatesInCol(c, 120);
        const rows = Object.keys(map).map(Number).sort(function (a, b) { return a - b; });
        for (let i = 0; i < rows.length - 1; i++) {
          const stride = rows[i + 1] - rows[i];
          if (stride < 1 || stride > 6) continue;
          let len = 1;
          for (let k = i + 1; k < rows.length; k++) {
            if (rows[k] - rows[k - 1] !== stride) break;
            if (dayDiff(map[rows[k - 1]], map[rows[k]]) !== 1) break;
            len++;
          }
          if (len >= 7 && (!best || len > best.len)) {
            best = { col: c, startRow: rows[i], startDate: map[rows[i]], stride: stride, len: len };
          }
        }
      }
      return best;
    }

    function findWeekdayRun() {
      let best = null;
      for (let c = 0; c <= 15; c++) {
        const rows = [];
        for (let r = 1; r <= MAX_SCAN_ROW; r++) {
          if (isWeekdayText(raw(colLetter(c) + r))) rows.push(r);
        }
        for (let i = 0; i < rows.length - 1; i++) {
          const stride = rows[i + 1] - rows[i];
          if (stride < 2 || stride > 6) continue;
          let len = 1;
          for (let k = i + 1; k < rows.length; k++) {
            if (rows[k] - rows[k - 1] !== stride) break;
            len++;
          }
          if (len >= 7 && (!best || len > best.len)) {
            best = { col: c, firstRow: rows[i], stride: stride, len: len };
          }
        }
      }
      return best;
    }

    const LABEL_RE = /^\s*(alkupäivä|alkupaiva|jakson\s*alku|aloituspäivä|aloituspaiva|aloitus|alkaa|alku|pvm)\b\s*[:.]?\s*(.*)$/i;
    function findLabelDate() {
      let found = null;
      forEachValue(80, 30, function (r, c, v) {
        if (found || typeof v !== "string") return;
        const m = v.trim().match(LABEL_RE);
        if (!m) return;
        let dt = m[2] ? parseDateValue(m[2]) || parseDateValue(stripWeekdayPrefix(m[2])) : null;
        if (!dt) dt = parseDateValue(raw(colLetter(c + 1) + (r + 1)));
        if (!dt) dt = parseDateValue(raw(colLetter(c) + (r + 2)));
        if (dt) found = dt;
      });
      return found;
    }

    let jakso = null;
    forEachValue(MAX_SCAN_ROW, 40, function (r, c, v) {
      if (jakso || typeof v !== "string") return;
      const m = v.match(/^\s*jakso\s*(\d+)/i);
      if (m) jakso = { row: r + 1, col: c, label: String(v).trim(), number: +m[1] };
    });

    let nameRow = null, weekRow0 = null, dateRow0 = null, hoursRow0 = null;
    let stride = 3, metaCol = 1, startDate = null, periodLabel = "";

    if (jakso) {
      periodLabel = jakso.label;
      metaCol = jakso.col;
      nameRow = jakso.row;
      const map = resolvedDatesInCol(jakso.col, MAX_SCAN_ROW);
      let hitRow = null;
      for (let off = 1; off <= 3 && !hitRow; off++) {
        if (map[jakso.row + off]) hitRow = jakso.row + off;
      }
      if (hitRow) {
        dateRow0 = hitRow;
        startDate = map[hitRow];
        const later = Object.keys(map).map(Number).filter(function (r) { return r > hitRow; }).sort(function (a, b) { return a - b; });
        if (later.length && later[0] - hitRow >= 2 && later[0] - hitRow <= 6) stride = later[0] - hitRow;
        weekRow0 = hitRow - 1 > jakso.row ? hitRow - 1 : hitRow;
        hoursRow0 = hitRow + 1 < hitRow + stride ? hitRow + 1 : hitRow;
      } else {
        // "Jakso" found but the date cell is empty: keep the usual block
        // (weekday, date, hours) and let a label or the user supply the date.
        // Do NOT steal a later day's date from the column.
        weekRow0 = jakso.row + 1;
        dateRow0 = jakso.row + 2;
        hoursRow0 = jakso.row + 3;
        stride = 3;
      }
    }

    if (!startDate) {
      const labelled = findLabelDate();
      if (labelled) startDate = labelled;
    }
    if (!startDate && !jakso) {
      const run = findDateRun();
      if (run) {
        metaCol = run.col;
        dateRow0 = run.startRow;
        startDate = run.startDate;
        stride = run.stride;
        weekRow0 = isWeekdayText(raw(colLetter(metaCol) + (dateRow0 - 1))) ? dateRow0 - 1 : dateRow0;
        hoursRow0 = stride > 1 ? dateRow0 + 1 : dateRow0;
        nameRow = weekRow0 > 1 ? weekRow0 - 1 : dateRow0;
      }
    }
    if (!startDate && !jakso && dateRow0 == null) {
      const wd = findWeekdayRun();
      if (wd) {
        metaCol = wd.col;
        stride = wd.stride;
        weekRow0 = wd.firstRow;
        dateRow0 = wd.firstRow + 1;
        hoursRow0 = wd.firstRow + 2 < wd.firstRow + stride ? wd.firstRow + 2 : wd.firstRow + 1;
        nameRow = wd.firstRow - 1;
      }
    }
    if (!startDate && dateRow0 == null) {
      let any = null;
      forEachValue(20, 8, function (_r, _c, v) {
        if (any) return;
        const d = parseDateValue(v);
        if (d) any = d;
      });
      if (any) startDate = any;
    }
    if (!startDate && forcedStart) startDate = forcedStart;

    if (!startDate || dateRow0 == null || nameRow == null) {
      if (jakso || findWeekdayRun()) {
        return { needsStartDate: true, periodLabel: periodLabel, periodNumber: jakso ? jakso.number : null };
      }
      throw new Error("Jakso-lomaketta ei voitu lukea (ei jaksoa eikä päiviä).");
    }
    if (weekRow0 == null) weekRow0 = dateRow0;
    if (hoursRow0 == null) hoursRow0 = dateRow0 + 1;

    const metaLetter = colLetter(metaCol);
    const MAX_DAYS = 62;
    let dayCount = 0;
    for (let i = 0; i < MAX_DAYS; i++) {
      const dateRow = dateRow0 + stride * i;
      const weekRow = weekRow0 + stride * i;
      const bDate = raw(metaLetter + dateRow);
      const bWeek = raw(metaLetter + weekRow);
      if (isTunnitText(bWeek) || isTunnitText(bDate)) break;
      if (parseDateValue(bDate) || parsePartialDM(bDate)) {
        dayCount++;
        continue;
      }
      if (isWeekdayText(bWeek) || (weekRow !== dateRow && isWeekdayText(bDate))) {
        dayCount++;
        continue;
      }
      // First day may have an empty date cell when the user typed the start date.
      if (i === 0 && isWeekdayText(raw(metaLetter + (dateRow0 - 1)))) {
        dayCount++;
        continue;
      }
      break;
    }
    if (dayCount < 1) dayCount = 1;

    const dates = [];
    for (let i = 0; i < dayCount; i++) dates.push(addDays(startDate, i));

    function isHelperName(n) {
      if (n == null) return false;
      const s = String(n).trim().toLowerCase();
      if (!s) return false;
      if (/käytössä|ukkoja|countif|yhteensä|tunti/.test(s)) return true;
      if (/^jakso\b/i.test(s)) return true;
      return false;
    }

    /** Person = column with a real name in the name row. Nameless shift columns are ignored. */
    function colHasPersonData(letter) {
      const n = raw(letter + nameRow);
      if (n == null || String(n).trim() === "") return false;
      if (isHelperName(n)) return false;
      return true;
    }

    const maxCol0 = Math.min(sheetMaxCol0(), metaCol + 1 + 45);
    const personCols = [];
    for (let c = metaCol + 1; c <= maxCol0; c++) {
      const letter = colLetter(c);
      if (colHasPersonData(letter)) personCols.push(letter);
    }
    if (!personCols.length) {
      throw new Error("Jakso-lomakkeelta ei löytynyt yhtään henkilöä (sarakkeet C…).");
    }

    const y0 = startDate.getFullYear();
    const holidayMap = Object.assign(
      {},
      holidaysForYear(y0 - 1),
      holidaysForYear(y0),
      holidaysForYear(y0 + 1)
    );

    const otThresholds = overtimeThresholds(startDate, dayCount);

    const people = [];
    for (let p = 0; p < personCols.length; p++) {
      const letter = personCols[p];
      const rawName = raw(letter + nameRow);
      const name = String(rawName).trim();
      const shifts = [];
      const company = [];
      for (let i = 0; i < dayCount; i++) {
        const formShiftRow = dateRow0 + stride * i;
        const formHoursRow = hoursRow0 + stride * i;
        let sh = null;
        if (!hasFormula(letter + formShiftRow)) {
          sh = parseShift(raw(letter + formShiftRow));
        }
        shifts.push(sh);
        let hrs = null;
        if (sh) hrs = parseCompanyHours(raw(letter + formHoursRow));
        company.push(hrs);
      }
      const days = computePerson(dates, shifts, company, holidayMap);
      const lmTotal = days.reduce(function (a, d) {
        return a + (d.me != null ? d.me : 0);
      }, 0);
      people.push({
        id: p + 1,
        name: name,
        col: letter,
        days,
        shiftCount: days.filter(function (d) {
          return d.start != null;
        }).length,
        lmTotal: lmTotal,
        overtime: overtimeSplit(lmTotal, otThresholds),
      });
    }

    return {
      periodLabel: periodLabel,
      periodNumber: jakso ? jakso.number : null,
      startDate: dateKey(startDate),
      dates: dates.map(dateKey),
      dayCount: dayCount,
      layout: { nameRow: nameRow, weekRow0: weekRow0, dateRow0: dateRow0, hoursRow0: hoursRow0, stride: stride, metaCol: metaCol },
      people,
      holidayMap,
      overtime: otThresholds,
    };
  }

  /** SheetJS sheet → array-of-arrays (1-based mentally; 0-based array) */
  function XLSXorAOA(sheet, ref) {
    // Prefer global XLSX when in browser/node with sheetjs
    if (typeof XLSX !== "undefined") {
      return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
    }
    throw new Error("SheetJS (XLSX) required to parse workbook");
  }

  /** Parse from ArrayBuffer */
  function parseArrayBuffer(buf, opts) {
    if (typeof XLSX === "undefined") throw new Error("SheetJS ei ole ladattu");
    const wb = XLSX.read(buf, { type: "array", cellDates: false, raw: true });
    return parseJaksoForm(wb, opts);
  }

  /**
   * Keskeytyspäivä on any day of the jakso (shift or day off): that day's hours
   * are left out and laskentapäivät drops by one. One generic mark covers sick
   * leave and holiday. Stored value:
   *   "keskeytys" — marked in the app, no reason
   *   "sairas" | "loma" — optional reason (future Excel colours: pink = sairas, green = loma)
   *   { reason: "sairas"|"loma" } — same, object form
   * The UI does not ask for a reason.
   */
  const DAY_MARK_KINDS = ["keskeytys", "sairas", "loma"];

  /** Optional reason, or "" if the day is marked with no reason. null if not a mark. */
  function markReason(v) {
    if (v == null || v === false || v === "") return null;
    if (typeof v === "string") {
      if (v === "keskeytys") return "";
      if (v === "sairas" || v === "loma") return v;
      return null;
    }
    if (typeof v === "object") {
      if (v.reason === "sairas" || v.reason === "loma") return v.reason;
      if (v.on) return "";
    }
    return null;
  }

  function isKeskeytys(v) {
    return markReason(v) !== null;
  }

  function sickDateSet(days, marks) {
    const set = {};
    if (!marks) return set;
    (days || []).forEach(function (d) {
      if (isKeskeytys(marks[d.date])) set[d.date] = true;
    });
    return set;
  }

  function sickDates(days, marks) {
    return Object.keys(sickDateSet(days, marks)).sort();
  }

  /**
   * The shift that counts for a day (Toteuma applied), or null.
   *   no toteuma / peruttu → planned shift (kind "plan"; peruttu keeps the planned hours, local agreement)
   *   peruttu late + korvaus → null: peruutuskorvaus chosen instead, hours count 0
   *   vaihto               → planned shift (hours as planned); lisät use the driven times
   *   muutos               → actual times (kind "actual")
   *   korvattu             → new shift times (kind "korvattu"); figures = MAX(planned, new) per field
   *   kutsu on a day off   → actual times, or the told times when no actual given (kind "kutsu")
   */
  function countedShift(d, toteumat) {
    const t = normToteuma(toteumat && toteumat[d.date]);
    if (d.start != null) {
      if (t && t.type === "muutos") return { kind: "actual", start: t.start, end: t.end, t: t };
      // cancelled, new shift given: each figure = MAX(planned, new shift); km follow the new shift
      if (t && t.type === "korvattu") return { kind: "korvattu", start: t.start, end: t.end, t: t };
      // late cancellation: Pekka chooses hours (default) OR peruutuskorvaus — never both
      if (t && t.type === "peruttu" && t.late && t.korvaus) return null;
      return { kind: "plan", start: d.start, end: d.end, t: t };
    }
    if (t && t.type === "kutsu") {
      return { kind: "kutsu", start: t.aStart != null ? t.aStart : t.start, end: t.aEnd != null ? t.aEnd : t.end, t: t };
    }
    return null;
  }

  /**
   * Double-count guard: a lisävuoro that overlaps the day's counted shift
   * (planned, actual or kutsu) is not counted anywhere. Returns true when ignored.
   */
  function extraOverlaps(d, ex, toteumat) {
    if (!ex || ex.start == null || ex.end == null) return false;
    const c = countedShift(d, toteumat);
    if (!c) return false;
    const b = shiftSpan(ex.start, ex.end);
    const hit = function (st, en) {
      const a = shiftSpan(st, en);
      return Math.min(a.e, b.e) - Math.max(a.s, b.s) > 0;
    };
    if (hit(c.start, c.end)) return true;
    if (c.kind === "korvattu" && hit(d.start, d.end)) return true;
    return !!(c.t && c.t.type === "vaihto" && hit(c.t.start, c.t.end));
  }

  /**
   * Sum a day field, skipping keskeytyspäivät. Toteuma (optional) replaces planned times
   * (muutos), takes MAX(planned, new) per field (korvattu) or adds a kutsu shift;
   * a cancelled shift keeps its planned hours.
   * "company" for a changed planned shift = sheet figure + (actual LM − planned LM),
   * so the sheet's Erotus is unchanged. App-only shifts: Yritys follows LM.
   */
  function sumActive(days, marks, key, extras, holidayMap, toteumat) {
    const sick = sickDateSet(days, marks);
    let total = 0;
    (days || []).forEach(function (d) {
      if (sick[d.date]) return;
      const c = countedShift(d, toteumat);
      if (c && c.kind === "plan") {
        if (d[key] != null) total += d[key];
      } else if (c) {
        const fig = computeShiftFigures(d.dateObj || d.date, c.start, c.end, holidayMap || {});
        if (fig && c.kind === "korvattu") {
          const meMax = Math.max(d.me || 0, fig.me || 0);
          if (key === "company") {
            if (d.company != null) total += d.company + (meMax - d.me);
          } else if (key === "check") {
            if (d.check != null) total += d.check;
          } else if (key === "me") total += meMax;
          else total += Math.max(d[key] || 0, fig[key] || 0);
        } else if (fig) {
          if (c.kind === "actual" && key === "company") {
            if (d.company != null) total += d.company + (fig.me - d.me);
          } else if (c.kind === "actual" && key === "check") {
            if (d.check != null) total += d.check;
          } else if (fig[key] != null) total += fig[key];
        }
      }
      const ex = extras && extras[d.date];
      if (!ex || ex.start == null || ex.end == null || extraOverlaps(d, ex, toteumat)) return;
      const fx = computeShiftFigures(d.dateObj || d.date, ex.start, ex.end, holidayMap || {});
      if (fx && fx[key] != null) total += fx[key];
    });
    return total;
  }

  /**
   * Laskentapäivät. Auto = 21 − sick days, only for a 21-day jakso.
   * A manual value (1–21) wins until cleared.
   */
  function resolveLaskenta(periodDays, sickCount, manual) {
    const auto = periodDays === 21 ? Math.max(0, 21 - (sickCount || 0)) : null;
    if (manual != null && manual >= 1 && manual <= 21) {
      return { n: manual, auto: auto, manual: true };
    }
    return { n: auto == null ? null : auto, auto: auto, manual: false };
  }

  /**
   * LM≠Yritys warnings use only the imported sheet difference (day.check).
   * Marks, lisävuorot and steward hours are not arguments, so they cannot
   * create or clear a mismatch.
   */
  function sheetMismatches(days) {
    return (days || []).filter(function (d) { return d.check != null && d.check > 0; });
  }

  // --- Lisät (TES III luku §17–§22, §24) and §6 peruutuskorvaus ---
  /**
   * Time windows (minutes from the shift's own date 00:00, k = day offset):
   *   ilta  §17  every day 18–21
   *   yö    §18  21–06, plus work started at the latest 04:00 that runs past 06:00
   *              counts as night until 12:00 (no rest info → no 2 h rest break)
   *   la    §19  arkilauantai 06–18; not pääsiäislauantai, not a holiday Saturday,
   *              not when aattopäivänlisä is paid (jouluaatto on a Saturday)
   *   su    §20  Sunday / kirkollinen juhlapäivä / itsenäisyyspäivä / vappu 00–24
   *              and the day before 18–24
   *   aatto §21  pääsiäislauantai, juhannusaatto, jouluaatto 00–18
   */
  const LISA_KEYS = ["ilta", "yo", "la", "su", "aatto"];
  const LISA_CODES = { ilta: "1420", yo: "1421", la: "1430", su: "1400", aatto: "", poikkeama: "1312", vapaa: "1311", peruutus: "1313" };
  /** Fixed-euro items by date (TES lisäpalkkiotaulukko; peruutus = 2 × työhönsidonnaisuuslisä, §6). */
  const FIXED_RATES = [
    { from: "0000-00-00", poikkeama: 36.12, vapaa: 22.85, peruutus: 56.08 },
    { from: "2026-09-01", poikkeama: 37.17, vapaa: 23.51, peruutus: 57.7 },
    { from: "2027-08-01", poikkeama: 38.06, vapaa: 24.07, peruutus: 59.08 },
  ];
  const LATE_END_MIN = 30;

  function fixedRatesFor(key) {
    let r = FIXED_RATES[0];
    FIXED_RATES.forEach(function (x) {
      if (key >= x.from) r = x;
    });
    return r;
  }

  function ov(s, e, a, b) {
    return Math.max(0, Math.min(e, b) - Math.max(s, a));
  }

  function asDate(date) {
    return date instanceof Date ? new Date(date.getFullYear(), date.getMonth(), date.getDate()) : toDate(date);
  }

  /** start/end minutes of one shift → absolute minutes from its date 00:00 (end after start). */
  function shiftSpan(start, end) {
    return { s: start, e: end <= start ? end + 24 * 60 : end };
  }

  function isSundayLike(d, holidayMap) {
    return d.getDay() === 0 || !!(holidayMap && holidayMap[dateKey(d)]);
  }

  /** Lisä minutes of one shift (any length up to 48 h). */
  function lisaShiftMinutes(date, startMin, endMin, holidayMap) {
    const d = asDate(date);
    const out = { ilta: 0, yo: 0, la: 0, su: 0, aatto: 0 };
    if (!d || startMin == null || endMin == null) return out;
    const sp = shiftSpan(startMin, endMin);
    const s = sp.s,
      e = sp.e;
    const H = 60,
      DAY = 24 * 60;
    for (let k = -1; k <= 2; k++) {
      const o = k * DAY;
      const day = addDays(d, k);
      out.ilta += ov(s, e, o + 18 * H, o + 21 * H);
      // night 21:00 (day k) → 06:00 (day k+1)
      out.yo += ov(s, e, o + 21 * H, o + DAY + 6 * H);
      // continuation on morning k: started ≤ 04:00 and still working after 06:00
      if (s <= o + 4 * H && e > o + 6 * H) out.yo += ov(s, e, o + 6 * H, o + 12 * H);
      const aatto = aattoName(day);
      if (day.getDay() === 6 && !(holidayMap && holidayMap[dateKey(day)]) && !aatto) {
        out.la += ov(s, e, o + 6 * H, o + 18 * H);
      }
      if (isSundayLike(day, holidayMap)) out.su += ov(s, e, o, o + DAY);
      else if (isSundayLike(addDays(day, 1), holidayMap)) out.su += ov(s, e, o + 18 * H, o + DAY);
      if (aatto) out.aatto += ov(s, e, o, o + 18 * H);
    }
    return out;
  }

  /** §24: jakso total to whole hours; under 30 min down, 30 min or more up. */
  function roundTes24(min) {
    if (min == null) return null;
    const h = Math.floor(min / 60);
    return min - h * 60 >= 30 ? h + 1 : h;
  }

  /**
   * Toteuma per day (on-device). Stored value:
   *   { type: "muutos", start, end, oma }   planned shift ran at other times (oma = own request)
   *   { type: "peruttu", late, korvaus }   cancelled, no new shift (late = notice after 17:00 the day before;
   *                                       korvaus = peruutuskorvaus chosen instead of the planned hours)
   *   { type: "korvattu", start, end }     cancelled and another shift given instead: hours, LM and each
   *                                       lisä = MAX(planned, new shift); km follow the new shift
   *   { type: "vaihto", start, end }       shift swapped with a colleague (Lisälehti 11): start/end = shift
   *                                       actually driven. Hours/LM stay as the OWN planned shift (longer
   *                                       or shorter swap alike), lisät follow the driven shift, no korvaus.
   *   { type: "kutsu", start, end, aStart, aEnd }  called in on a planned day off
   *                                       (start/end as told, aStart/aEnd actual, optional)
   */
  const TOTEUMA_TYPES = ["muutos", "peruttu", "korvattu", "kutsu", "vaihto"];

  function normToteuma(v) {
    if (!v || typeof v !== "object" || TOTEUMA_TYPES.indexOf(v.type) < 0) return null;
    return v;
  }

  /** Poikkeama check between a reference (planned/told) shift and the shift actually worked. */
  function deviation(ref, act) {
    const a = shiftSpan(ref.start, ref.end);
    const b = shiftSpan(act.start, act.end);
    const early = b.s < a.s;
    const late = b.e - a.e;
    return { early: early, lateMin: late, hit: early || late >= LATE_END_MIN };
  }

  /**
   * Lisät for one person over the jakso.
   * days: computePerson rows; marks: keskeytys map; extras: lisävuorot; toteumat: { date: toteuma }.
   * Keskeytyspäivä drops the day. A lisävuoro overlapping the day's counted shift is not
   * counted (extraOverlaps), same rule as the hour totals.
   */
  function lisatSummary(days, marks, extras, toteumat, holidayMap) {
    const mins = { ilta: 0, yo: 0, la: 0, su: 0, aatto: 0 };
    const fixed = {
      poikkeama: { n: 0, eur: 0, rates: {} },
      vapaa: { n: 0, eur: 0, rates: {} },
      peruutus: { n: 0, eur: 0, rates: {} },
    };
    const events = [];
    const notes = [];
    function addMin(date, start, end) {
      const m = lisaShiftMinutes(date, start, end, holidayMap);
      LISA_KEYS.forEach(function (k) {
        mins[k] += m[k];
      });
    }
    function addFixed(key, date, why) {
      const r = fixedRatesFor(date)[key];
      fixed[key].n += 1;
      fixed[key].eur = Math.round((fixed[key].eur + r) * 100) / 100;
      fixed[key].rates[r.toFixed(2)] = true;
      events.push({ date: date, key: key, eur: r, why: why });
    }
    function devWhy(dv) {
      const bits = [];
      if (dv.early) bits.push("alkoi suunniteltua aiemmin");
      if (dv.lateMin >= LATE_END_MIN) bits.push("päättyi " + formatHM(dv.lateMin) + " myöhemmin");
      return bits.join(", ");
    }
    (days || []).forEach(function (d) {
      if (isKeskeytys(marks && marks[d.date])) return;
      const t = normToteuma(toteumat && toteumat[d.date]);
      const planned = d.start != null ? { start: d.start, end: d.end } : null;
      if (planned) {
        if (!t || t.type === "kutsu") {
          addMin(d.date, planned.start, planned.end);
        } else if (t.type === "vaihto") {
          addMin(d.date, t.start, t.end); // Lisälehti 11: lisät from the driven shift, no korvaus
        } else if (t.type === "muutos") {
          addMin(d.date, t.start, t.end);
          const dv = deviation(planned, t);
          if (dv.hit && !t.oma) addFixed("poikkeama", d.date, devWhy(dv));
        } else if (t.type === "korvattu") {
          // each lisä type = MAX(planned shift's minutes, new shift's minutes)
          const mp = lisaShiftMinutes(d.date, planned.start, planned.end, holidayMap);
          const mn = lisaShiftMinutes(d.date, t.start, t.end, holidayMap);
          LISA_KEYS.forEach(function (k) {
            mins[k] += Math.max(mp[k], mn[k]);
          });
          const dv = deviation(planned, t);
          if (dv.hit) addFixed("poikkeama", d.date, "peruttu, tilalle vuoro joka " + devWhy(dv));
        } else if (t.type === "peruttu") {
          if (t.late && t.korvaus) addFixed("peruutus", d.date, "peruttu edellisenä päivänä klo 17 jälkeen, valittu korvaus (tunnit eivät lasketa)");
          else if (t.late) events.push({ date: d.date, key: "peruutusTunnit", eur: 0, why: "peruttu klo 17 jälkeen, valittu tunnit (suunnitellut tunnit lasketaan, ei peruutuskorvausta)" });
        }
      } else if (t && t.type === "kutsu") {
        const act = {
          start: t.aStart != null ? t.aStart : t.start,
          end: t.aEnd != null ? t.aEnd : t.end,
        };
        addMin(d.date, act.start, act.end);
        addFixed("vapaa", d.date, "kutsuttu vapaapäivänä");
        const dv = deviation(t, act);
        if (dv.hit) addFixed("poikkeama", d.date, "kutsuvuoro " + devWhy(dv));
      }
      const ex = extras && extras[d.date];
      if (ex && ex.start != null && ex.end != null) {
        if (extraOverlaps(d, ex, toteumat)) notes.push({ date: d.date, msg: "lisävuoro on päällekkäin toteutuneen vuoron kanssa – sitä ei lasketa kahteen kertaan" });
        else addMin(d.date, ex.start, ex.end);
      }
    });
    const hours = {};
    LISA_KEYS.forEach(function (k) {
      hours[k] = roundTes24(mins[k]);
    });
    const fixedTotal = Math.round((fixed.poikkeama.eur + fixed.vapaa.eur + fixed.peruutus.eur) * 100) / 100;
    return { minutes: mins, hours: hours, fixed: fixed, fixedTotal: fixedTotal, events: events, notes: notes };
  }

  /** Summaries for overview */

  function summarize(result) {
    const flags = { notAllowed: 0, mdCheck: 0, checkMismatch: 0 };
    const perPerson = result.people.map((p) => {
      let na = 0,
        md = 0;
      p.days.forEach((d) => {
        if (d.restBefore === "Not Allowed") na++;
        if (d.restBefore === "Md,s Check") md++;
      });
      const mism = sheetMismatches(p.days).length;
      flags.notAllowed += na;
      flags.mdCheck += md;
      flags.checkMismatch += mism;
      return { id: p.id, name: p.name, shiftCount: p.shiftCount, notAllowed: na, mdCheck: md, checkMismatch: mism };
    });
    return { flags, perPerson };
  }

  // --- Veturiraha (TES Lisäpalkkiot – Veturiraha, PDF s. 61–62; hinnat lisäpalkkiotaulukko s. 59) ---
  /**
   * TES: "Veturirahaa maksetaan kultakin veturikilometriltä." → € = km × hinta.
   * "Veturikilometrit määräytyvät kultakin matkalta lähtö- ja tulopaikan välisen matkan mukaan."
   * Hinta valitaan tavarajunien taulukosta junapainon mukaan:
   *   junapaino 3: 1500 tn tai alle · 4: yli 1500 enint. 4800 tn · 5: yli 4800 enint. 5600 tn tai IVY-VAK-juna
   *   "hitaat, 40 km/t tai alle" ja "yksinajo" ovat omat rivinsä. Yksinajorivit ovat taulukossa jo
   *   kaksinkertaisia ("Yhden miehen ajossa on veturiraha kaksinkertainen"), joten niitä ei kerrota uudelleen.
   * Huomautus 3: junapaino = jarrupainojärjestelmän kokonaisjunapaino; Pekan työpaikalla todellinen paino.
   * Huomautus 2: itäisen liikenteen VAK-juna (nopeus vähintään 40 km/h) → kuten yli 4800 tn, painosta riippumatta.
   * Palkkalajit 1450–1462 palkkataulukosta (1456/1463 "Veturina ajo" ei kuulu tähän).
   * OLETUKSET: (a) yli 5600 tn: TES ei anna riviä → käytetään junapainoa 5 ja näytetään huomautus.
   *   (b) hinta valitaan junan päivämäärän (vuoron päivän) mukaan.
   *   (c) € pyöristetään senteiksi palkkalajeittain koko jaksolta (km yhteensä × hinta), kuten palkkalaskelmassa.
   *   (d) Huomautus 1 (Vartius–Kontiomäki–Oulu–Raahe, suunniteltu paino) ei koske Pekan junia, ei toteutettu.
   * Veturina ajo (veturi ilman junaa, 1456 / yksin 1463): TES:ssä ei omaa riviä; vakiintuneen käytännön
   *   (palkkataulukko) mukaan km × henkilöjunan hinta (Henkilöjuna / Henkilöjuna yksinajo). Paino ei vaikuta.
   */
  const VETURIRAHA_RATES = [
    { from: "0000-00-00", r: { 1450: 0.0833, 1451: 0.1677, 1452: 0.2515, 1453: 0.1677, 1454: 0.3357, 1455: 0.504, 1457: 0.1677, 1458: 0.3357, 1459: 0.504, 1460: 0.3357, 1461: 0.6718, 1462: 1.0081, 1456: 0.0576, 1463: 0.1156 } },
    { from: "2026-09-01", r: { 1450: 0.0857, 1451: 0.1725, 1452: 0.2587, 1453: 0.1725, 1454: 0.3454, 1455: 0.5186, 1457: 0.1725, 1458: 0.3454, 1459: 0.5186, 1460: 0.3454, 1461: 0.6912, 1462: 1.0373, 1456: 0.0592, 1463: 0.1189 } },
    { from: "2027-08-01", r: { 1450: 0.0877, 1451: 0.1766, 1452: 0.2649, 1453: 0.1766, 1454: 0.3536, 1455: 0.531, 1457: 0.1766, 1458: 0.3536, 1459: 0.531, 1460: 0.3536, 1461: 0.7077, 1462: 1.0621, 1456: 0.0606, 1463: 0.1217 } },
  ];
  const VETURIRAHA_LABELS = {
    1450: "Kaksinajo · tavarajuna enint. 1500 t", 1451: "Kaksinajo · tavarajuna 1501–4800 t", 1452: "Kaksinajo · tavarajuna yli 4800 t tai IVY-VAK",
    1453: "Kaksinajo · hidas tavarajuna enint. 1500 t", 1454: "Kaksinajo · hidas tavarajuna 1501–4800 t", 1455: "Kaksinajo · hidas tavarajuna yli 4800 t",
    1456: "Kaksinajo · veturina ajo", 1463: "Yksinajo · veturina ajo",
    1457: "Yksinajo · tavarajuna enint. 1500 t", 1458: "Yksinajo · tavarajuna 1501–4800 t", 1459: "Yksinajo · tavarajuna yli 4800 t tai IVY-VAK",
    1460: "Yksinajo · hidas tavarajuna enint. 1500 t", 1461: "Yksinajo · hidas tavarajuna 1501–4800 t", 1462: "Yksinajo · hidas tavarajuna yli 4800 t",
  };
  /** Row title for one counted train: always starts with Yksinajo / Kaksinajo. */
  function veturirahaTitle(t) {
    const who = t.yksin ? "Yksinajo" : "Kaksinajo";
    if (t.veturina) return who + " · veturina ajo";
    const lk = junapainoLuokka(t);
    if (t.ivyvak && !t.hidas) return who + " · tavarajuna, IVY-VAK (kuten yli 4800 t)";
    const w = lk === 3 ? "enint. 1500 t" : lk === 4 ? "1501–4800 t" : "yli 4800 t";
    return who + " · " + (t.hidas ? "hidas tavarajuna (enint. 40 km/h) " : "tavarajuna ") + w;
  }

  function veturirahaRatesFor(key) {
    let r = VETURIRAHA_RATES[0];
    VETURIRAHA_RATES.forEach(function (x) {
      if (key >= x.from) r = x;
    });
    return r.r;
  }

  /** Normalise one stored train → { junanumero, paino, km, yksin, hidas, ivyvak } or null. */
  function normJuna(j) {
    if (!j || typeof j !== "object") return null;
    const num = function (v) {
      const n = typeof v === "number" ? v : parseFloat(String(v == null ? "" : v).replace(/\s/g, "").replace(",", "."));
      return isFinite(n) && n > 0 ? n : null;
    };
    return {
      junanumero: String(j.junanumero == null ? "" : j.junanumero).trim(),
      paino: num(j.paino),
      km: num(j.km),
      // true/false = saved manual choice; null = not chosen → auto-detected from the sheet (see pairPartners)
      yksin: j.yksin === true ? true : j.yksin === false ? false : null,
      hidas: !!j.hidas,
      ivyvak: !!j.ivyvak,
      veturina: !!j.veturina,
    };
  }

  /**
   * Autolla-ajokorvaus (TES Lisäpalkkiot, PDF s. 63): veturimies ajaa työnantajan autoa miehistönvaihto-
   * tai miehitystarkoituksessa; maksetaan vain autoa kuljettavalle. € = todelliset km × hinta:
   *   1) Paikallisajo (paikkakunnan sisällä): "samansuuruisena kuin … hitaissa yksinajettavissa junapainoltaan
   *      yli 1500 tonnin mutta alle 4800 tonnin tavarajunissa" → koodin 1461 hinta.
   *   2) Paikkakunnan ulkopuolinen ajo: "kuin … kaksinajettavissa junapainoltaan alle 1500 tonnin tavarajunissa,
   *      kuitenkin vähintään 30 kilometriltä/yhdensuuntainen ajomatka" → koodin 1450 hinta, jokainen
   *      yhdensuuntainen ajo erikseen vähintään 30 km.
   * Palkkalaji ei tiedossa (palkkataulukossa ei riviä) → näytetään "–".
   * trips: { [date]: [{ tyyppi: "paikallis" | "ulko", km, kuvaus? }] }
   */
  const AUTOAJO_LABELS = { paikallis: "Paikallisajo", ulko: "Paikkakunnan ulkopuolinen ajo" };
  const AUTOAJO_MIN_KM = 30;
  function normAutoajo(t) {
    if (!t || typeof t !== "object") return null;
    const km = typeof t.km === "number" ? t.km : parseFloat(String(t.km == null ? "" : t.km).replace(/\s/g, "").replace(",", "."));
    return { tyyppi: t.tyyppi === "ulko" ? "ulko" : "paikallis", km: isFinite(km) && km > 0 ? km : null, kuvaus: String(t.kuvaus || "").trim() };
  }
  function autoajoSummary(trips, skip) {
    const out = [];
    const groups = {};
    const notes = [];
    Object.keys(trips || {})
      .sort()
      .forEach(function (date) {
        (Array.isArray(trips[date]) ? trips[date] : []).forEach(function (raw, i) {
          const t = normAutoajo(raw);
          if (!t || t.km == null) return;
          if (skip && skip(date)) {
            notes.push({ date: date, msg: "Autolla-ajo " + (i + 1) + ": keskeytyspäivä, ei lasketa." });
            return;
          }
          const rates = veturirahaRatesFor(date);
          const rate = t.tyyppi === "ulko" ? rates[1450] : rates[1461];
          const paid = t.tyyppi === "ulko" ? Math.max(AUTOAJO_MIN_KM, t.km) : t.km;
          const row = Object.assign({ date: date, i: i, rate: rate, paidKm: paid, eur: round2(paid * rate) }, t);
          out.push(row);
          const gk = t.tyyppi + "@" + rate;
          if (!groups[gk]) groups[gk] = { tyyppi: t.tyyppi, label: AUTOAJO_LABELS[t.tyyppi], rate: rate, km: 0, n: 0, from: date };
          groups[gk].km += paid;
          groups[gk].n += 1;
        });
      });
    const rows = Object.keys(groups)
      .map(function (k) {
        const g = groups[k];
        g.km = Math.round(g.km * 1000) / 1000;
        g.eur = round2(g.km * g.rate);
        return g;
      })
      .sort(function (a, b) {
        return a.tyyppi < b.tyyppi ? -1 : a.tyyppi > b.tyyppi ? 1 : a.from < b.from ? -1 : 1;
      });
    const eur = round2(rows.reduce(function (s, r) { return s + r.eur; }, 0));
    const km = Math.round(rows.reduce(function (s, r) { return s + r.km; }, 0) * 1000) / 1000;
    return { trips: out, rows: rows, km: km, eur: eur, notes: notes };
  }

  /**
   * Yksin- vai kaksinajo vuorotaulusta (Pekan työpaikan käytäntö, ei TES-sääntö): jos jollakin muulla
   * lomakkeen henkilöllä on samana päivänä täsmälleen sama suunniteltu vuoro (sama alku- ja loppuaika
   * minuutilleen), juna on ajettu kahdestaan; muuten yksin. Käytetään suunniteltua vuoroa (d.start/d.end),
   * ei toteumaa eikä keskeytyspäiviä. Nimettömät sarakkeet eivät ole henkilöitä (parseJaksoForm), joten
   * niitä ei verrata.
   * → { yksin: bool, partners: [name], planned: bool }
   */
  function pairPartners(people, person, date) {
    const own = (person.days || []).find(function (d) { return d.date === date; });
    if (!own || own.start == null || own.end == null) return { yksin: true, partners: [], planned: false };
    const partners = [];
    (people || []).forEach(function (q) {
      if (q === person || q.id === person.id) return;
      const d = (q.days || []).find(function (x) { return x.date === date; });
      if (d && d.start === own.start && d.end === own.end) partners.push(q.name);
    });
    return { yksin: partners.length === 0, partners: partners, planned: true };
  }

  /** Junapaino class 3/4/5 (TES rows 3.–5.) */
  function junapainoLuokka(j) {
    if (j.ivyvak && !j.hidas) return 5; // Huomautus 2 / "tai IVY-VAK-juna"
    if (j.paino == null) return null;
    if (j.paino <= 1500) return 3;
    if (j.paino <= 4800) return 4;
    return 5;
  }

  /** Pay code for one train (1450–1462), or null if weight is missing. */
  function veturirahaCode(j) {
    const yksin = j.yksin == null ? true : j.yksin;
    if (j.veturina) return yksin ? 1463 : 1456;
    const lk = junapainoLuokka(j);
    if (lk == null) return null;
    const base = yksin ? (j.hidas ? 1460 : 1457) : j.hidas ? 1453 : 1450;
    return base + (lk - 3);
  }

  function round2(x) {
    return Math.round(x * 100 + 1e-9) / 100;
  }

  /**
   * junat: { [yyyy-mm-dd]: [train, ...] }; skip(date) → true for days not counted (keskeytyspäivä);
   * autoYksin(date) → true/false for trains without a saved choice (default yksinajo).
   * → { trains:[{date,i,...train,code,rate,eur}], rows:[{code,label,rate,km,eur,from}], km, eur, notes }
   */
  function veturirahaSummary(junat, skip, autoYksin) {
    const trains = [];
    const groups = {};
    const notes = [];
    Object.keys(junat || {})
      .sort()
      .forEach(function (date) {
        const list = Array.isArray(junat[date]) ? junat[date] : [];
        list.forEach(function (raw, i) {
          const j = normJuna(raw);
          if (!j) return;
          j.auto = j.yksin == null;
          if (j.auto) {
            const a = autoYksin ? autoYksin(date) : null;
            j.yksin = a == null ? true : !!a;
          }
          const nm = (j.veturina ? "Veturina ajo " : "Juna ") + (j.junanumero || i + 1);
          if (skip && skip(date)) {
            notes.push({ date: date, msg: nm + ": keskeytyspäivä, ei lasketa." });
            return;
          }
          const code = veturirahaCode(j);
          if (code == null || j.km == null) {
            notes.push({ date: date, msg: nm + ": puuttuu " + (code == null ? "paino" : "km") + ", ei lasketa." });
            return;
          }
          if (!j.veturina && j.paino != null && j.paino > 5600 && !(j.ivyvak && !j.hidas))
            notes.push({ date: date, msg: nm + ": paino yli 5600 t. TES:n taulukko päättyy 5600 t:iin, laskettu junapainolla 5 (tarkista palkanlaskennasta)." });
          if (!j.veturina && j.ivyvak && j.hidas)
            notes.push({ date: date, msg: nm + ": IVY-VAK-sääntö koskee vain junia, joiden nopeus on vähintään 40 km/h. Laskettu hitaana painon mukaan." });
          const rate = veturirahaRatesFor(date)[code];
          const t = Object.assign({ date: date, i: i, code: code, rate: rate, eur: round2(j.km * rate) }, j);
          trains.push(t);
          const gk = code + "@" + rate;
          if (!groups[gk]) groups[gk] = { code: code, label: VETURIRAHA_LABELS[code], rate: rate, km: 0, n: 0, from: date };
          groups[gk].km += j.km;
          groups[gk].n += 1;
        });
      });
    const rows = Object.keys(groups)
      .map(function (k) {
        const g = groups[k];
        g.km = Math.round(g.km * 1000) / 1000;
        g.eur = round2(g.km * g.rate);
        return g;
      })
      .sort(function (a, b) {
        return a.code - b.code || (a.from < b.from ? -1 : 1);
      });
    const eur = round2(rows.reduce(function (s, r) { return s + r.eur; }, 0));
    const km = Math.round(rows.reduce(function (s, r) { return s + r.km; }, 0) * 1000) / 1000;
    return { trains: trains, rows: rows, km: km, eur: eur, notes: notes };
  }


  return {
    WEEKDAYS_FI,
    parseShift,
    parseCompanyHours,
    parseJaksoForm,
    parseArrayBuffer,
    computePerson,
    computeShiftFigures,
    restWithExtras,
    holidaysForYear,
    holidayName,
    easterSunday,
    OT_CONFIG,
    STEWARD_STORAGE_KEY: "nrtyokalut.lm",
    TSV_STORAGE_KEY: "nrtyokalut.tsv",
    PALKKIO_RATES,
    palkkioFor,
    palkkioRatesInRange,
    stewardBonus,
    stewardMapSet,
    overtimeHolidaysForYear,
    overtimeThresholds,
    applyOvertimeOverride,
    overtimeSplit,
    INT_TABLES,
    interruptedNorms,
    interruptedSplit,
    DAY_MARK_KINDS,
    markReason,
    isKeskeytys,
    sickDates,
    sumActive,
    sheetMismatches,
    resolveLaskenta,
    parseHM,
    formatHM,
    summarize,
    dateKey,
    weekdayFi,
    REST_NOT_ALLOWED_MIN,
    REST_MD_CHECK_MIN,
    aattoName,
    LISA_KEYS,
    LISA_CODES,
    FIXED_RATES,
    fixedRatesFor,
    lisaShiftMinutes,
    roundTes24,
    TOTEUMA_TYPES,
    normToteuma,
    lisatSummary,
    countedShift,
    extraOverlaps,
    VETURIRAHA_RATES,
    VETURIRAHA_LABELS,
    veturirahaRatesFor,
    normJuna,
    junapainoLuokka,
    veturirahaCode,
    veturirahaSummary,
    veturirahaTitle,
    pairPartners,
    AUTOAJO_LABELS,
    AUTOAJO_MIN_KM,
    normAutoajo,
    autoajoSummary,
  };
});
