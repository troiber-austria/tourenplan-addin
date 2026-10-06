/* demo.js - nur fuer ?demo=1: erfundene Beispieldaten und ein Netzwerk-Ersatz, damit die Oberflaeche ohne Excel/Schluessel ausprobiert werden kann. */
(function () {
  'use strict';
  var ORTE = {   // erfundene Kunden an ungefähren Wiener Koordinaten
    'Teststraße 1': [48.2082, 16.3738], 'Beispielgasse 5': [48.1850, 16.3200], 'Musterweg 12': [48.2400, 16.4200],
    'Demoplatz 3': [48.1500, 16.3600], 'Probering 8': [48.2700, 16.3300], 'Versuchsallee 20': [48.1200, 16.2800], 'Kontrollgasse 9': [48.2200, 16.5000]
  };
  var DEPOT = [47.84, 16.19];
  var kopf = ['Kd-Nr.', 'Kunde', 'Identnummer', 'Strasse', 'Ort', 'Plz', 'Rb', 'Gew', 'TK RB', 'RB frisch', 'Tour orig', 'Reihung orig', 'Hilfsspalte Kunden', 'Reihung final', 'KFZ', 'Tour final'];
  var zeile = function (kd, name, str, plz, gew, rf, kfz, tour) { return [kd, name, 1, str, 'Wien', plz, 0, gew, 0, 0, 'Wien', 1, 1, rf, kfz, tour]; };
  var werte = [
    zeile(1, 'Demo Gastro Eins', 'Teststraße 1', 'AT-1010', 120, 1, 'WB 764EW', 'Anna'),
    zeile(2, 'Demo Gastro Zwei', 'Beispielgasse 5', 'AT-1060', 80, 2, 'WB 764EW', 'Anna'),
    zeile(2, 'Demo Gastro Zwei', 'Beispielgasse 5', 'AT-1060', 40, 2, 'WB 764EW', 'Anna'),
    zeile(3, 'Demo Gastro Drei', 'Musterweg 12', 'AT-1220', 200, 3, 'WB 764EW', 'Anna'),
    zeile(4, 'Demo Gastro Vier', 'Demoplatz 3', 'AT-1100', 90, 4, 'WB 764EW', 'Anna'),
    zeile(5, 'Demo Gastro Fünf', 'Probering 8', 'AT-1190', 60, 1, 'WB 804FJ', 'Ben'),
    zeile(6, 'Demo Gastro Sechs', 'Versuchsallee 20', 'AT-1230', 75, 2, 'WB 804FJ', 'Ben'),
    zeile(7, 'Demo Gastro Sieben', 'Kontrollgasse 9', 'AT-1220', 55, 3, 'WB 804FJ', 'Ben')
  ];
  function hav(a, b) {
    var R = 6371000, r = Math.PI / 180, dl = (b[0] - a[0]) * r, dn = (b[1] - a[1]) * r;
    var x = Math.sin(dl / 2) * Math.sin(dl / 2) + Math.cos(a[0] * r) * Math.cos(b[0] * r) * Math.sin(dn / 2) * Math.sin(dn / 2);
    return 2 * R * Math.asin(Math.sqrt(x));
  }
  function encNum(v) { v = v < 0 ? ~(v << 1) : (v << 1); var s = ''; while (v >= 0x20) { s += String.fromCharCode((0x20 | (v & 0x1f)) + 63); v >>= 5; } return s + String.fromCharCode(v + 63); }
  function encode(pts) { var la = 0, lo = 0, s = ''; pts.forEach(function (p) { var a = Math.round(p[0] * 1e5), b = Math.round(p[1] * 1e5); s += encNum(a - la) + encNum(b - lo); la = a; lo = b; }); return s; }
  function resp(b) { return Promise.resolve({ ok: true, status: 200, json: function () { return Promise.resolve(b); } }); }

  window.TP_DEMO = {
    lesen: function () {
      return Promise.resolve({ headers: kopf, values: werte.map(function (r) { return r.slice(); }), fzBlatt: true, fahrzeuge: {
        'WB764EW': { kfz: 'WB 764EW', typ: '18-Tonner', gewicht: 18, hoehe: 4, breite: 2.55, laenge: 12 },
        'WB804FJ': { kfz: 'WB 804FJ', typ: '3,5-Tonner', gewicht: 3.5, hoehe: 2.8, breite: 2.1, laenge: 7 } } });
    },
    reihungSchreiben: function (ups) { var alt = []; ups.forEach(function (u) { alt.push({ idx: u.idx, wert: werte[u.idx][13], kd: werte[u.idx][0] }); werte[u.idx][13] = u.wert; }); return Promise.resolve(alt); },
    fetch: function (url, opts) {
      opts = opts || {};
      if (url.indexOf('/geocode/search/structured') > -1) {
        var a = new URL(url).searchParams.get('address'); var c = ORTE[a];
        return resp({ features: c ? [{ geometry: { coordinates: [c[1], c[0]] }, properties: { confidence: 1, label: a } }] : [] });
      }
      if (url.indexOf('/geocode/search') > -1) return resp({ features: [{ geometry: { coordinates: [DEPOT[1], DEPOT[0]] }, properties: { label: 'Depot (Demo)' } }] });
      if (url.indexOf('/v2/directions/') > -1) {
        var co = JSON.parse(opts.body).coordinates.map(function (p) { return [p[1], p[0]]; }), segs = [], d = 0, t = 0;
        for (var i = 1; i < co.length; i++) { var m = hav(co[i - 1], co[i]) * 1.3; segs.push({ distance: m, duration: m / 12 }); d += m; t += m / 12; }
        return resp({ routes: [{ summary: { distance: d, duration: t }, geometry: encode(co), segments: segs }] });
      }
      if (url.indexOf('routes.googleapis.com') > -1) {
        var b = JSON.parse(opts.body), pts = [b.origin].concat(b.intermediates || [], [b.destination]).map(function (w) { return [w.location.latLng.latitude, w.location.latLng.longitude]; });
        var legs = []; for (var j = 1; j < pts.length; j++) { var mm = hav(pts[j - 1], pts[j]) * 1.3; legs.push({ duration: Math.round(mm / 10.5) + 's', staticDuration: Math.round(mm / 12) + 's', distanceMeters: Math.round(mm) }); }
        return resp({ routes: [{ legs: legs }] });
      }
      if (url.indexOf('/optimization') > -1) {
        var jobs = JSON.parse(opts.body).jobs.slice(), cur = DEPOT, order = [];
        while (jobs.length) { jobs.sort(function (x, y) { return hav(cur, [x.location[1], x.location[0]]) - hav(cur, [y.location[1], y.location[0]]); }); var n = jobs.shift(); order.push(n.id); cur = [n.location[1], n.location[0]]; }
        return resp({ routes: [{ steps: [{ type: 'start' }].concat(order.map(function (id) { return { type: 'job', job: id }; }), [{ type: 'end' }]) }] });
      }
      return Promise.reject(new Error('Demo: unbekannte Adresse ' + url));
    }
  };
})();
