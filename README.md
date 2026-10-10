# Jakso päiväkirja (PWA)

Sovellus: https://nrtyokalut.github.io/

Paikallinen, asennettava web-sovellus junakuljettajan jakso-lomakkeen tarkistukseen.

## Käyttö

1. Avaa `index.html` paikallisella HTTP-palvelimella (PWA / file picker vaatii http/https).
2. Valitse yrityksen `.xlsx` (Taul1 / Jakso-lomake).
3. Katso kaikkien henkilöiden yhteenveto ja avaa henkilön detalji.

**Data ei lähde laitteelta.** SheetJS + laskentalogiikka ajetaan selaimessa. Service worker mahdollistaa offline-käytön.

## Asennus laitteelle

- **iPhone (Safari):** Jaa → *Lisää Koti-valikkoon*
- **Windows (Chrome/Edge):** valikkopalkin asennuskuvake / *Asenna sovellus*

## Testi

```bash
node test.mjs
```

Vertaa laskentaa LibreOffice-uudelleenlaskettuun v6-työkirjaan (`golden.json`).

## Tiedostot

- `calc.js` — Excel-v6-logiikan porttaus
- `app.js` / `index.html` / `styles.css` — UI
- `lib/xlsx.full.min.js` — SheetJS (paikallinen)
- `sw.js` / `manifest.webmanifest` / `icons/` — PWA

---
© 2026 Lämpöpumppu Mafia. Kaikki oikeudet pidätetään. All rights reserved. Katso LICENSE.
