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
   *   lisatyoHolidays      lower the lisätyö threshold (= where 50 % starts; lisätyö is paid at 50 %)
   *   ylityoHolidays       also lower the statutory ylityö threshold (shown for information)
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
   *   lisaK = where 50 % starts (lowered lisätyö threshold), yliK = statutory ylityö threshold
   *   (informational), sataK = where 100 % starts.
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
   * Apply a user override {lisaK?, sataK?} (minutes) on top of defaults (may be null).
   * Returns effective thresholds or null if either limit is still unknown.
   */
  function applyOvertimeOverride(defaults, override) {
    const o = override || {};
    const lisaK = o.lisaK != null ? o.lisaK : defaults ? defaults.lisaK : null;
    const sataK = o.sataK != null ? o.sataK : defaults ? defaults.sataK : null;
    if (lisaK == null || sataK == null) return null;
    return Object.assign({}, defaults || { holidays: [] }, {
      lisaK: lisaK,
      sataK: Math.max(sataK, lisaK),
      overridden: o.lisaK != null || o.sataK != null,
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
   * Split LM total (minutes). Lisätyö is paid at the same 50 % as ylityö, so the 50 %
   * band runs from the (possibly lowered) lisätyö threshold up to 132:45.
   * Ylityö 100 % is everything above 132:45. null if no thresholds (non-21-day jakso).
   */
  function overtimeSplit(lmMin, th) {
    if (!th) return null;
    const lm = lmMin || 0;
    return {
      yli50: clampMin(lm, th.lisaK, th.sataK) - th.lisaK,
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
  function parseJaksoForm(workbook) {
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

    // Layout detection:
    //  Modern: B2=Jakso, names row2, weekday row3, date row4, hours row5
    //  Older:  B3=Jakso, names row3, weekday row4, date row5, hours row6
    let nameRow, weekRow0, dateRow0, hoursRow0, periodAddr;
    if (toDate(raw("B4"))) {
      nameRow = 2;
      weekRow0 = 3;
      dateRow0 = 4;
      hoursRow0 = 5;
      periodAddr = "B2";
    } else if (toDate(raw("B5"))) {
      nameRow = 3;
      weekRow0 = 4;
      dateRow0 = 5;
      hoursRow0 = 6;
      periodAddr = "B3";
    } else if (toDate(raw("B4")) || toDate(raw("B3"))) {
      // fallback attempt
      nameRow = 2;
      weekRow0 = 3;
      dateRow0 = 4;
      hoursRow0 = 5;
      periodAddr = "B2";
    } else {
      throw new Error("Jakso-lomakkeelta ei löytynyt aloituspäivää (B4/B5).");
    }

    const startDate = toDate(raw("B" + dateRow0));
    if (!startDate) throw new Error("Jakso-lomakkeelta ei löytynyt aloituspäivää.");

    const MAX_DAYS = 62;
    let dayCount = 0;
    for (let i = 0; i < MAX_DAYS; i++) {
      const dateRow = dateRow0 + 3 * i;
      const weekRow = weekRow0 + 3 * i;
      const bDate = raw("B" + dateRow);
      const bWeek = raw("B" + weekRow);
      if (bWeek != null && String(bWeek).trim() === "Tunnit") break;
      if (bDate != null && String(bDate).trim() === "Tunnit") break;
      if (toDate(bDate)) {
        dayCount++;
        continue;
      }
      if (bWeek != null && weekdaySet[String(bWeek).trim()]) {
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

    const maxCol0 = Math.min(sheetMaxCol0(), 2 + 45);
    const personCols = [];
    for (let c = 2; c <= maxCol0; c++) {
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
        const formShiftRow = dateRow0 + 3 * i;
        const formHoursRow = hoursRow0 + 3 * i;
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

    const periodCell = raw(periodAddr);
    return {
      periodLabel: periodCell ? String(periodCell) : "",
      startDate: dateKey(startDate),
      dates: dates.map(dateKey),
      dayCount: dayCount,
      layout: { nameRow: nameRow, weekRow0: weekRow0, dateRow0: dateRow0, hoursRow0: hoursRow0 },
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
  function parseArrayBuffer(buf) {
    if (typeof XLSX === "undefined") throw new Error("SheetJS ei ole ladattu");
    const wb = XLSX.read(buf, { type: "array", cellDates: false, raw: true });
    return parseJaksoForm(wb);
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
