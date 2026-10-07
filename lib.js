/* lib.js - reine Hilfsfunktionen (kein Browser/Excel noetig, daher testbar) */
(function (root) {
  'use strict';

  // Google/ORS "encoded polyline" -> [[lat, lng], ...]
  function decodePolyline(str, precision) {
    var factor = Math.pow(10, precision == null ? 5 : precision);
    var index = 0, lat = 0, lng = 0, out = [];
    while (index < str.length) {
      var b, shift = 0, result = 0;
      do { b = str.charCodeAt(index++) - 63; result |= (b & 0x1f) << shift; shift += 5; } while (b >= 0x20);
      lat += (result & 1) ? ~(result >> 1) : (result >> 1);
      shift = 0; result = 0;
      do { b = str.charCodeAt(index++) - 63; result |= (b & 0x1f) << shift; shift += 5; } while (b >= 0x20);
      lng += (result & 1) ? ~(result >> 1) : (result >> 1);
      out.push([lat / factor, lng / factor]);
    }
    return out;
  }

  // "AT-1230" -> {country:'AT', zip:'1230'}; "83313" -> DE; "1230" -> AT
  function parsePlz(plz) {
    var s = String(plz == null ? '' : plz).trim();
    var m = /^([A-Za-z]{1,3})\s*[-\s]\s*(\d{4,5})$/.exec(s);
    if (m) return { country: m[1].toUpperCase(), zip: m[2] };
    var d = /^(\d{4,5})$/.exec(s);
    if (d) return { country: d[1].length === 5 ? 'DE' : 'AT', zip: d[1] };
    return { country: 'AT', zip: s.replace(/\D/g, '') };
  }

  // Entfernt Zusaetze wie "Cineplex Donau Plex TOP 601A", "Tuer 604", "; Donauzentrum", "/7-12", Hausnummern-Bereiche.
  // "Wagramer Straße 79 Tür 604" -> "Wagramer Straße 79" ; "Mariahilferstrasse 42-48" -> "Mariahilferstrasse 42"
  function cleanStrasse(str) {
    var s = String(str == null ? '' : str).trim().split(/[;,]/)[0].trim();
    var m = /^(.*?[^\d\s])\s*(\d+[a-zA-Z]?)(?![a-zA-Z])/.exec(s);
    return m ? (m[1].trim() + ' ' + m[2]) : s;
  }
  function nurStrasse(str) {
    var s = cleanStrasse(str);
    return s.replace(/\s*\d+[a-zA-Z]?$/, '').trim() || s;
  }
  // Suchvarianten: Original, bereinigt, nur Strasse (ohne Duplikate)
  function addressVariants(str) {
    var out = [], o = String(str == null ? '' : str).trim();
    [o, cleanStrasse(o), nurStrasse(o)].forEach(function (v) { if (v && out.indexOf(v) < 0) out.push(v); });
    return out;
  }

  function norm(s) { return String(s == null ? '' : s).trim().replace(/\s+/g, ' ').toLowerCase(); }

  function addressKey(strasse, plz, ort) {
    var p = parsePlz(plz);
    return [norm(strasse), p.country + '-' + p.zip, norm(ort)].join('|');
  }

  function addressText(strasse, plz, ort) {
    var p = parsePlz(plz);
    return String(strasse).trim() + ', ' + p.zip + ' ' + String(ort).trim() + ', ' + (p.country === 'DE' ? 'Deutschland' : 'Oesterreich');
  }

  function num(v) { var n = Number(String(v == null ? '' : v).replace(',', '.')); return isFinite(n) ? n : 0; }

  /**
   * Gruppiert die Zeilen von Tabelle1 nach Fahrer ("Tour final") und Kunde (= ein Stopp).
   * headers: Spaltenueberschriften, values: 2D-Array der Datenzeilen.
   * Rueckgabe: { tours: { Fahrer: {name, kfz, kfzAlle:[], stops:[...], zeilen, gewicht} }, fehlend:[Spalten] }
   */
  function groupTours(headers, values) {
    var col = {};
    headers.forEach(function (h, i) { col[String(h).trim().toLowerCase()] = i; });
    var need = ['kd-nr.', 'kunde', 'strasse', 'ort', 'plz', 'gew', 'reihung final', 'kfz', 'tour final'];
    var fehlend = need.filter(function (n) { return !(n in col); });
    if (fehlend.length) return { tours: {}, fehlend: fehlend };
    var tours = {};
    values.forEach(function (row, bodyIdx) {
      var fahrer = String(row[col['tour final']] == null ? '' : row[col['tour final']]).trim();
      var strasse = String(row[col['strasse']] == null ? '' : row[col['strasse']]).trim();
      if (!fahrer || !strasse) return;
      var t = tours[fahrer] || (tours[fahrer] = { name: fahrer, kfzZaehler: {}, stops: [], index: {}, zeilen: 0, gewicht: 0 });
      var kfz = String(row[col['kfz']] == null ? '' : row[col['kfz']]).trim();
      if (kfz) t.kfzZaehler[kfz] = (t.kfzZaehler[kfz] || 0) + 1;
      var key = String(row[col['kd-nr.']]).trim() + '#' + addressKey(strasse, row[col['plz']], row[col['ort']]);
      var cleanKey = addressKey(cleanStrasse(strasse), row[col['plz']], row[col['ort']]);
      var st = t.index[key];
      var reihung = row[col['reihung final']];
      reihung = (reihung === '' || reihung == null) ? 1e6 : num(reihung);
      if (!st) {
        st = t.index[key] = {
          key: key, addrKey: cleanKey,
          kdnr: row[col['kd-nr.']], name: String(row[col['kunde']] == null ? '' : row[col['kunde']]).trim(),
          strasse: strasse, plz: row[col['plz']], ort: String(row[col['ort']] == null ? '' : row[col['ort']]).trim(),
          reihung: reihung, gew: 0, zeilen: [], erst: bodyIdx
        };
        t.stops.push(st);
      }
      st.reihung = Math.min(st.reihung, reihung);
      st.gew += num(row[col['gew']]);
      st.zeilen.push(bodyIdx);
      t.zeilen++; t.gewicht += num(row[col['gew']]);
    });
    Object.keys(tours).forEach(function (n) {
      var t = tours[n];
      t.stops.sort(function (a, b) { return a.reihung - b.reihung || a.erst - b.erst; });
      var ks = Object.keys(t.kfzZaehler).sort(function (a, b) { return t.kfzZaehler[b] - t.kfzZaehler[a]; });
      t.kfz = ks[0] || '';
      t.kfzAlle = ks;
      delete t.index;
    });
    return { tours: tours, fehlend: [] };
  }

  // Teilt eine Punkteliste in Stuecke mit max. n Punkten; benachbarte Stuecke teilen sich den Uebergabepunkt.
  function chunkPoints(points, n) {
    if (points.length <= n) return [points.slice()];
    var out = [], i = 0;
    while (i < points.length - 1) {
      var end = Math.min(i + n - 1, points.length - 1);
      out.push(points.slice(i, end + 1));
      i = end;
    }
    return out;
  }

  function parseGoogleDuration(s) { return s ? Number(String(s).replace('s', '')) || 0 : 0; }

  function fmtDur(sec) {
    sec = Math.max(0, Math.round(sec / 60));
    var h = Math.floor(sec / 60), m = sec % 60;
    return h ? h + ' h ' + (m < 10 ? '0' : '') + m + ' min' : m + ' min';
  }
  function fmtKm(m) { return (m / 1000).toFixed(1).replace('.', ',') + ' km'; }
  function fmtClock(date) {
    var h = date.getHours(), m = date.getMinutes();
    return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
  }

  // ORS-Optimierung: Reihenfolge der Job-IDs (1-basiert) aus der Antwort
  function orderFromOptimization(resp) {
    var r = resp && resp.routes && resp.routes[0];
    if (!r) return [];
    return r.steps.filter(function (s) { return s.type === 'job'; }).map(function (s) { return s.job; });
  }

  // Fahrzeugdaten je Typ -> ORS-Optionen
  function orsOptions(fz, cfg) {
    var typ = (cfg.typen && cfg.typen[fz && fz.typ]) || { vehicle_type: 'hgv' };
    var opt = { vehicle_type: typ.vehicle_type || 'hgv' };
    if (fz && fz.gewicht) {
      var r = {};
      if (fz.gewicht) r.weight = fz.gewicht;
      if (fz.hoehe) r.height = fz.hoehe;
      if (fz.breite) r.width = fz.breite;
      if (fz.laenge) r.length = fz.laenge;
      opt.profile_params = { restrictions: r };
    }
    return opt;
  }

  var api = {
    decodePolyline: decodePolyline, parsePlz: parsePlz, addressKey: addressKey, addressText: addressText,
    cleanStrasse: cleanStrasse, addressVariants: addressVariants, groupTours: groupTours, chunkPoints: chunkPoints, parseGoogleDuration: parseGoogleDuration,
    fmtDur: fmtDur, fmtKm: fmtKm, fmtClock: fmtClock, orderFromOptimization: orderFromOptimization,
    orsOptions: orsOptions, num: num
  };
  root.TP = root.TP || {};
  Object.keys(api).forEach(function (k) { root.TP[k] = api[k]; });
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
