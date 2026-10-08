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
    let n = R + S;
    if (S + R > EARLY_FULL && S === EARLY_FULL) n += T;
    n += W + X + Y;
    if (Y + X > EARLY_FULL && Y === EARLY_FULL) n += Z;
    return n;
  }

  /** 25% bonus AC */
  function bonus25(weekday, holidayToday, holidayTomorrow, curSegs, nextSegs) {
    // AB10=Perjantai, AC10=Lauantai
    if (weekday === "Perjantai" && !holidayTomorrow) {
      return nextSegs[2] + nextSegs[3]; // Z+AA
    }
    if (weekday === "Lauantai" && !holidayToday) {
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
        b25 = bonus25(wd, !!hol, !!holTom, curSegs, nextSegs);
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

  /** Summaries for overview */
  function summarize(result) {
    const flags = { notAllowed: 0, mdCheck: 0, checkMismatch: 0 };
    const perPerson = result.people.map((p) => {
      let na = 0,
        md = 0,
        mism = 0;
      p.days.forEach((d) => {
        if (d.restBefore === "Not Allowed") na++;
        if (d.restBefore === "Md,s Check") md++;
        if (d.check != null && d.check > 0) mism++;
      });
      flags.notAllowed += na;
      flags.mdCheck += md;
      flags.checkMismatch += mism;
      return { id: p.id, name: p.name, shiftCount: p.shiftCount, notAllowed: na, mdCheck: md, checkMismatch: mism };
    });
    return { flags, perPerson };
  }

  return {
    WEEKDAYS_FI,
    parseShift,
    parseCompanyHours,
    parseJaksoForm,
    parseArrayBuffer,
    computePerson,
    holidaysForYear,
    holidayName,
    easterSunday,
    OT_CONFIG,
    overtimeHolidaysForYear,
    overtimeThresholds,
    applyOvertimeOverride,
    overtimeSplit,
    parseHM,
    formatHM,
    summarize,
    dateKey,
    weekdayFi,
    REST_NOT_ALLOWED_MIN,
    REST_MD_CHECK_MIN,
  };
});
