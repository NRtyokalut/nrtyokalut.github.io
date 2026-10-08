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
      p.days.forEach((d) => {
        if (d.restBefore === "Not Allowed" || d.restBefore === "Md,s Check") {
          const label = (WD_SHORT[d.weekday] || "") + " " + dateFi(d.date);
          const lepo = d.restBeforeMin != null ? fmt(d.restBeforeMin) : "";
          items.push({
            bad: d.restBefore === "Not Allowed",
            text:
              p.name +
              ": " +
              lab(d.restBefore) +
              " " +
              label +
              (lepo ? " (vuorojen väli " + lepo + ")" : ""),
          });
        }
        if (d.check != null && d.check > 0) {
          items.push({
            bad: false,
            text: p.name + ": Tarkista erotus " + dateFi(d.date) + " = " + fmt(d.check),
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
    p.days.forEach((d) => {
      if (d.restBefore === "Not Allowed" || d.restBefore === "Md,s Check") {
        const label = (WD_SHORT[d.weekday] || "") + " " + dateFi(d.date).replace(/\.\d{4}$/, ".");
        const lepo = d.restBeforeMin != null ? fmt(d.restBeforeMin) : "";
        problems.push({
          kind: d.restBefore === "Not Allowed" ? "bad" : "md",
          text:
            lab(d.restBefore) +
            " " +
            label +
            (lepo ? " (" + lepo + ")" : ""),
        });
      }
      if (d.check != null && d.check > 0) {
        problems.push({
          kind: "md",
          text: "Tarkista erotus " + dateFi(d.date).replace(/\.\d{4}$/, ".") + " = " + fmt(d.check),
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

  function renderOverview() {
    const people = state.people;
    $("periodMeta").textContent =
      (state.periodLabel ? state.periodLabel + " · " : "") +
      "Alku " +
      dateFi(state.startDate);

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

      const sum = (key) =>
        p.days.reduce((a, d) => a + (d[key] != null ? d[key] : 0), 0);
      const hrs = sum("hrs");
      const me = sum("me");
      const company = sum("company");
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
        "</h3>" +
        '<p class="card-meta">' +
        p.shiftCount +
        " vuoroa</p></div>" +
        badge +
        "</div>" +
        '<div class="card-hours">' +
        "<div><span>Tunnit</span><b>" +
        fmt(hrs) +
        "</b></div>" +
        "<div><span>LM</span><b>" +
        fmt(me) +
        "</b></div>" +
        "<div><span>Yritys</span><b>" +
        fmt(company) +
        "</b></div>" +
        "</div>" +
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

    show("overview");
  }

  function openDetail(personId) {
    const p = state.people.find((x) => x.id === personId);
    if (!p) return;
    $("detailTitle").textContent = p.name + " · " + p.shiftCount + " vuoroa";

    // Desktop/wide: classic table; mobile: compact day cards (see CSS)
    const tbody = $("detailTable").querySelector("tbody");
    tbody.innerHTML = p.days
      .map((d) => {
        return (
          "<tr>" +
          "<td>" +
          dateFi(d.date).slice(0, 5) +
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
          restClass(d.restAfter) +
          '">' +
          lab(d.restAfter) +
          "</td>" +
          "<td>" +
          fmt(d.hrs) +
          "</td>" +
          "<td>" +
          fmt(d.night) +
          "</td>" +
          "<td>" +
          fmt(d.b25) +
          "</td>" +
          "<td>" +
          fmt(d.b100) +
          "</td>" +
          "<td>" +
          fmt(d.me) +
          "</td>" +
          "<td>" +
          fmt(d.company) +
          "</td>" +
          "<td>" +
          fmt(d.check) +
          "</td>" +
          "</tr>"
        );
      })
      .join("");

    // Mobile card list (always in DOM; CSS shows/hides)
    const cards = $("detailCards");
    cards.innerHTML = p.days
      .map((d) => {
        const hasShift = d.start != null;
        const restLabel = d.restAfter
          ? lab(d.restAfter) +
            (d.restAfterMin != null ? " · " + fmt(d.restAfterMin) : "")
          : "";
        return (
          '<article class="day-card' +
          (hasShift ? "" : " empty") +
          '">' +
          '<header><strong>' +
          (WD_SHORT[d.weekday] || "") +
          " " +
          dateFi(d.date) +
          "</strong>" +
          (d.special ? '<span class="tag">' + d.special + "</span>" : "") +
          "</header>" +
          (hasShift
            ? '<div class="day-grid">' +
              "<div><span>Alku</span><b>" +
              fmt(d.start) +
              "</b></div>" +
              "<div><span>Loppu</span><b>" +
              fmt(d.end) +
              "</b></div>" +
              "<div><span>Tunnit</span><b>" +
              fmt(d.hrs) +
              "</b></div>" +
              "<div><span>Yö</span><b>" +
              fmt(d.night) +
              "</b></div>" +
              "<div><span>LM</span><b>" +
              fmt(d.me) +
              "</b></div>" +
              "<div><span>Yritys</span><b>" +
              fmt(d.company) +
              "</b></div>" +
              "<div><span>25%</span><b>" +
              fmt(d.b25) +
              "</b></div>" +
              "<div><span>100%</span><b>" +
              fmt(d.b100) +
              "</b></div>" +
              "<div><span>Erotus</span><b>" +
              fmt(d.check) +
              "</b></div>" +
              '<div class="' +
              restClass(d.restAfter) +
              '"><span>Vuorojen väli →</span><b>' +
              (restLabel || "—") +
              "</b></div>" +
              "</div>"
            : '<p class="muted">Ei vuoroa</p>') +
          "</article>"
        );
      })
      .join("");

    const sum = (key) => p.days.reduce((a, d) => a + (d[key] != null ? d[key] : 0), 0);
    $("detailTotals").innerHTML =
      '<div class="stat"><b>' +
      fmt(sum("hrs")) +
      "</b><span>Tunnit</span></div>" +
      '<div class="stat"><b>' +
      fmt(sum("night")) +
      "</b><span>Yö</span></div>" +
      '<div class="stat"><b>' +
      fmt(sum("b25")) +
      "</b><span>25 %</span></div>" +
      '<div class="stat"><b>' +
      fmt(sum("b100")) +
      "</b><span>100 %</span></div>" +
      '<div class="stat"><b>' +
      fmt(sum("me")) +
      "</b><span>LM yht.</span></div>" +
      '<div class="stat"><b>' +
      fmt(sum("company")) +
      "</b><span>Yritys yht.</span></div>";

    renderAlerts($("detailAlerts"), state, personId);
    show("detail");
  }

  async function handleFile(file) {
    if (!file) return;
    $("fileName").hidden = false;
    $("fileName").textContent = "Luetaan: " + file.name + " …";
    try {
      const buf = await file.arrayBuffer();
      state = ShiftCalc.parseArrayBuffer(buf);
      summary = ShiftCalc.summarize(state);
      $("fileName").textContent = file.name;
      renderOverview();
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
    show("home");
  });
  $("btnBack").addEventListener("click", () => show("overview"));

  window.__ShiftApp = {
    loadArrayBuffer: (buf) => {
      state = ShiftCalc.parseArrayBuffer(buf);
      summary = ShiftCalc.summarize(state);
      renderOverview();
      return state;
    },
    openDetail,
    getState: () => state,
  };
})();
