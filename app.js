/* app.js - Oberflaeche, Excel-Anbindung, Google-Karte */
(function () {
  'use strict';
  var T = window.TP, cfg = window.TP_CONFIG;
  var $ = function (id) { return document.getElementById(id); };
  var FARBEN = ['#305496', '#c2410c', '#15803d', '#7e22ce', '#be123c', '#0e7490', '#a16207', '#4d7c0f', '#9d174d', '#1d4ed8'];

  var S = {
    keys: { ors: '', google: '' }, tours: {}, fahrzeuge: {}, geoCache: {}, geoNeu: [], results: {},
    map: null, overlays: [], info: null, aktuell: null, vorschlag: null, busy: false, demo: false, daten: null
  };

  // ---------- kleine Helfer ----------
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function lsGet(k) { try { return window.localStorage.getItem(k); } catch (e) { return null; } }
  function lsSet(k, v) { try { window.localStorage.setItem(k, v); } catch (e) { /* Speicher gesperrt */ } }
  function status(text, art) { var e = $('status'); e.textContent = text || ''; e.className = 'status' + (art ? ' ' + art : ''); }
  function normKfz(s) { return String(s || '').replace(/\s+/g, '').toUpperCase(); }
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function setBusy(b) {
    S.busy = b;
    ['btnBerechnen', 'btnOptimieren', 'btnAlle', 'btnNeu', 'fahrer'].forEach(function (id) { $(id).disabled = b; });
  }

  // ---------- Schluessel ----------
  function ladeKeys() {
    var k = {}; try { k = JSON.parse(lsGet('tp_keys') || '{}'); } catch (e) { k = {}; }
    S.keys.ors = k.ors || (window.TP_KEYS && window.TP_KEYS.ors) || '';
    S.keys.google = k.google || (window.TP_KEYS && window.TP_KEYS.google) || '';
    var d = lsGet('tp_depot'); if (d) cfg.depot.adresse = d;
  }

  // ---------- Datenzugriff (Excel) ----------
  var Daten = {
    async lesen() {
      if (S.demo) return window.TP_DEMO.lesen();
      return Excel.run(async function (ctx) {
        var t = ctx.workbook.tables.getItemOrNullObject(cfg.tabelle);
        t.load('isNullObject');
        var ws = ctx.workbook.worksheets.getItemOrNullObject(cfg.fahrzeugBlatt);
        ws.load('isNullObject');
        await ctx.sync();
        if (t.isNullObject) throw new Error('Die Tabelle "' + cfg.tabelle + '" wurde in dieser Mappe nicht gefunden.');
        var h = t.getHeaderRowRange().load('values');
        var b = t.getDataBodyRange().load('values');
        var f = ws.isNullObject ? null : ws.getRange('A2:F20').load('values');
        await ctx.sync();
        var fz = {};
        if (f) f.values.forEach(function (r) {
          if (!r[0]) return;
          fz[normKfz(r[0])] = { kfz: String(r[0]), typ: String(r[1] || ''), gewicht: T.num(r[2]), hoehe: T.num(r[3]), breite: T.num(r[4]), laenge: T.num(r[5]) };
        });
        return { headers: h.values[0], values: b.values, fahrzeuge: fz, fzBlatt: !ws.isNullObject };
      });
    },
    async cacheLesen() {
      if (S.demo) return {};
      return Excel.run(async function (ctx) {
        var ws = ctx.workbook.worksheets.getItemOrNullObject(cfg.cacheBlatt);
        ws.load('isNullObject'); await ctx.sync();
        if (ws.isNullObject) return {};
        var ur = ws.getUsedRangeOrNullObject(true); ur.load('values,isNullObject'); await ctx.sync();
        var out = {};
        if (!ur.isNullObject) ur.values.slice(1).forEach(function (r) {
          if (r[0] === '' || r[1] === '' || r[2] === '') return;
          var lat = Number(r[1]), lon = Number(r[2]), label = String(r[4] || ''), ebene = String(r[5] || '');
          // alte Fehltreffer (Land/Landesmitte) nicht mehr verwenden
          var grob = ['country', 'macroregion', 'region', 'macrocounty', 'county'].indexOf(ebene) > -1 ||
            /^(austria|österreich|germany|deutschland)$/i.test(label.trim()) ||
            (Math.abs(lat - 47.3333) < 0.01 && Math.abs(lon - 13.3333) < 0.01);
          if (!grob) out[String(r[0])] = { lat: lat, lon: lon, weak: r[3] === 1 || r[3] === '1', label: label, layer: ebene };
        });
        return out;
      });
    },
    async cacheSchreiben(eintraege) {
      if (S.demo || !eintraege.length) return;
      return Excel.run(async function (ctx) {
        var ws = ctx.workbook.worksheets.getItemOrNullObject(cfg.cacheBlatt);
        ws.load('isNullObject'); await ctx.sync();
        var start = 1;
        if (ws.isNullObject) {
          ws = ctx.workbook.worksheets.add(cfg.cacheBlatt);
          ws.getRange('A1:F1').values = [['Adresse (Schlüssel)', 'Breite', 'Länge', 'unsicher', 'Treffer', 'Ebene']];
          ws.visibility = Excel.SheetVisibility.hidden;
        } else {
          var ur = ws.getUsedRangeOrNullObject(true); ur.load('rowCount,isNullObject'); await ctx.sync();
          start = ur.isNullObject ? 1 : ur.rowCount;
        }
        var rows = eintraege.map(function (e) { return [e.key, e.v.lat, e.v.lon, e.v.weak ? 1 : 0, e.v.label || '', e.v.layer || '']; });
        ws.getRangeByIndexes(start, 0, rows.length, 6).values = rows;
        await ctx.sync();
      });
    },
    async cacheLeeren() {
      if (S.demo) return;
      return Excel.run(async function (ctx) {
        var ws = ctx.workbook.worksheets.getItemOrNullObject(cfg.cacheBlatt);
        ws.load('isNullObject'); await ctx.sync();
        if (!ws.isNullObject) { ws.getRange('A2:F100000').clear(Excel.ClearApplyTo.contents); await ctx.sync(); }
      });
    },
    async zeileZeigen(idx) {
      if (S.demo) return;
      return Excel.run(async function (ctx) {
        var t = ctx.workbook.tables.getItem(cfg.tabelle);
        t.worksheet.activate();
        t.getDataBodyRange().getRow(idx).select();
        await ctx.sync();
      });
    },
    // updates: [{idx, wert}] -> Spalte "Reihung final"; liefert die alten Werte zurueck
    async reihungSchreiben(updates) {
      if (S.demo) return window.TP_DEMO.reihungSchreiben(updates);
      return Excel.run(async function (ctx) {
        var t = ctx.workbook.tables.getItem(cfg.tabelle);
        var col = t.columns.getItem('Reihung final').getDataBodyRange().load('values,formulas');
        var kd = t.columns.getItem('Kd-Nr.').getDataBodyRange().load('values');
        await ctx.sync();
        var hatFormel = col.formulas.some(function (r) { return typeof r[0] === 'string' && r[0].charAt(0) === '='; });
        if (hatFormel) throw new Error('Die Spalte "Reihung final" enthält Formeln - es wird nichts überschrieben.');
        var alt = [];
        updates.forEach(function (u) {
          alt.push({ idx: u.idx, wert: col.values[u.idx][0], kd: kd.values[u.idx][0] });
          col.getCell(u.idx, 0).values = [[u.wert]];
        });
        await ctx.sync();
        return alt;
      });
    },
    async reihungPruefenUndSetzen(alt) {   // fuer "Rueckgaengig": nur wenn Kd-Nr. noch passt
      if (S.demo) return window.TP_DEMO.reihungSchreiben(alt.map(function (a) { return { idx: a.idx, wert: a.wert }; }));
      return Excel.run(async function (ctx) {
        var t = ctx.workbook.tables.getItem(cfg.tabelle);
        var col = t.columns.getItem('Reihung final').getDataBodyRange();
        var kd = t.columns.getItem('Kd-Nr.').getDataBodyRange().load('values');
        await ctx.sync();
        var ok = alt.every(function (a) { return kd.values[a.idx] && String(kd.values[a.idx][0]) === String(a.kd); });
        if (!ok) throw new Error('Die Tabelle hat sich seit der Übernahme verändert - Rückgängig nicht möglich.');
        alt.forEach(function (a) { col.getCell(a.idx, 0).values = [[a.wert]]; });
        await ctx.sync();
      });
    }
  };

  // ---------- Geocode-Merkliste ----------
  var cache = {
    get: function (k) { return S.geoCache[k]; },
    put: function (k, v) { S.geoCache[k] = v; S.geoNeu.push({ key: k, v: v }); },
    flush: async function () {
      var n = S.geoNeu; S.geoNeu = [];
      try { await Daten.cacheSchreiben(n); } catch (e) { status('Merkliste konnte nicht gespeichert werden: ' + e.message, 'fehler'); }
    }
  };

  // ---------- Engine ----------
  var engine = T.createEngine({
    cfg: cfg, fetch: function (u, o) { return (S.demo ? window.TP_DEMO.fetch : window.fetch.bind(window))(u, o); },
    cache: cache, getKeys: function () { return S.keys; }
  });

  // ---------- Karte: im Seitenbereich oder im eigenen Fenster ----------
  var mv = T.MapView.create($('karte'), function () { return S.demo ? '' : S.keys.google; });
  var Dlg = { d: null, bereit: false, wartend: null };

  function dialogMoeglich() {
    return !S.demo && window.Office && Office.context && Office.context.ui && typeof Office.context.ui.displayDialogAsync === 'function';
  }
  function fensterModus() { return dialogMoeglich() && lsGet('tp_fenster') !== '0'; }

  function einfach(l) {      // Schicht in ein schlankes, uebertragbares Objekt umwandeln
    var res = l.res; if (!res || !res.ors) return null;
    var g = res.google, last = null, path = [];
    res.ors.path.forEach(function (p, i) {      // Punkte ausduennen (ca. 20 m), Start und Ende bleiben
      if (!last || i === res.ors.path.length - 1 || Math.abs(p[0] - last[0]) + Math.abs(p[1] - last[1]) > 0.0002) { path.push(p); last = p; }
    });
    return {
      farbe: l.farbe, gestrichelt: !!l.gestrichelt, nummern: !!l.nummern,
      depot: { lat: res.depot.lat, lon: res.depot.lon, name: cfg.depot.name }, path: path,
      stops: res.stops.map(function (s, i) { return { name: s.name, strasse: s.strasse, ort: s.ort, gew: Math.round(s.gew), lat: s.lat, lon: s.lon, weak: !!s.weak, eta: g ? T.fmtClock(g.eta[i]) : '' }; })
    };
  }

  function paneKarteZeigen(an) {
    $('karte').style.display = an ? '' : 'none';
    $('karteHinweis').hidden = an;
  }

  function dialogSenden() {
    if (Dlg.d && Dlg.bereit && Dlg.wartend) { Dlg.d.messageChild(JSON.stringify(Dlg.wartend)); }
  }

  // Liefert true, wenn das Kartenfenster die Anzeige uebernommen hat
  function imFensterZeigen(paket) {
    return new Promise(function (resolve) {
      Dlg.wartend = paket;
      if (Dlg.d) { dialogSenden(); return resolve(true); }
      var url = location.href.replace(/[?#].*$/, '').replace(/[^\/]*$/, '') + 'karte.html';
      Office.context.ui.displayDialogAsync(url, { height: 88, width: 88, displayInIframe: false }, function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) { Dlg.wartend = null; return resolve(false); }
        Dlg.d = r.value; Dlg.bereit = false;
        Dlg.d.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          if (arg.message === 'bereit') { Dlg.bereit = true; dialogSenden(); }
        });
        Dlg.d.addEventHandler(Office.EventType.DialogEventReceived, function () {   // Fenster wurde geschlossen
          Dlg.d = null; Dlg.bereit = false; paneKarteZeigen(true);
          if (S.letzteLayers) mv.draw(S.letzteLayers);
        });
        paneKarteZeigen(false);
        resolve(true);
      });
    });
  }

  function dialogSchliessen() {
    if (Dlg.d) { try { Dlg.d.close(); } catch (e) { /* schon zu */ } Dlg.d = null; Dlg.bereit = false; }
    paneKarteZeigen(true);
  }

  // layers: [{res, farbe, gestrichelt, nummern}] ; bereiche: Elemente, deren Text im Kartenfenster links stehen soll
  async function zeichne(layers, bereiche) { return zeigePlain(layers.map(einfach).filter(Boolean), bereiche); }
  async function zeigePlain(plain, bereiche) {
    S.letzteLayers = plain; S.letzteBereiche = bereiche;
    if (fensterModus()) {
      var html = (bereiche || ['vorschlag', 'ergebnis']).map(function (id) { var e = $(id); return e && !e.hidden ? e.innerHTML : ''; }).join('');
      var ok = await imFensterZeigen({ googleKey: S.keys.google, titel: 'Tourenplan-Karte', html: html, layers: plain });
      if (ok) return;
    }
    paneKarteZeigen(true);
    await mv.draw(plain);
  }

  // ---------- Daten laden ----------
  async function datenLaden() {
    status('Lese Tourenplan ...');
    var d = await Daten.lesen();
    S.daten = d; S.fahrzeuge = d.fahrzeuge;
    var g = T.groupTours(d.headers, d.values);
    if (g.fehlend.length) throw new Error('In Tabelle1 fehlen die Spalten: ' + g.fehlend.join(', '));
    S.tours = g.tours; S.results = {}; S.vorschlag = null;
    var sel = $('fahrer'), alt = sel.value; sel.innerHTML = '';
    var namen = Object.keys(S.tours).sort(function (a, b) { return a.localeCompare(b, 'de'); });
    namen.forEach(function (n) {
      var t = S.tours[n], fz = S.fahrzeuge[normKfz(t.kfz)];
      var o = document.createElement('option'); o.value = n;
      o.textContent = n + ' · ' + (t.kfz || 'kein KFZ') + (fz && fz.typ ? ' (' + fz.typ + ')' : '') + ' · ' + t.stops.length + ' Kunden';
      sel.appendChild(o);
    });
    if (alt && S.tours[alt]) sel.value = alt;
    if (!namen.length) status('In Tabelle1 sind keine Touren mit Fahrer ("Tour final") und Adresse vorhanden.', 'fehler');
    else status(namen.length + ' Touren eingelesen.', 'ok');
    $('ergebnis').hidden = true; $('vorschlag').hidden = true;
    if (!d.fzBlatt) hinweisBlock('Das Blatt "' + cfg.fahrzeugBlatt + '" fehlt - die Route wird ohne Gewichts-/Maßbeschränkung berechnet.');
  }
  function hinweisBlock(text) { var e = $('ergebnis'); e.hidden = false; e.innerHTML = '<div class="hinweis">' + esc(text) + '</div>'; }

  function aktuelleTour() { return S.tours[$('fahrer').value]; }
  function startZeit() {
    var d = $('datum').value, z = $('zeit').value || '06:00';
    var dt = d ? new Date(d + 'T' + z + ':00') : new Date();
    return isNaN(dt) ? new Date() : dt;
  }
  function standSek() { return Math.max(0, T.num($('standzeit').value)) * 60; }
  function fzFuer(tour) { return S.fahrzeuge[normKfz(tour.kfz)] || null; }

  async function lokalisieren(tour) {
    var fehler = await engine.geocodeStops(tour.stops, function (i, n) { status('Adressen suchen ' + i + ' / ' + n + ' ...'); });
    return fehler;
  }

  // ---------- Berechnen ----------
  async function berechne(name, ohneAnzeige) {
    var tour = S.tours[name];
    await lokalisieren(tour);
    var fz = fzFuer(tour);
    status('Route für ' + name + ' wird berechnet ...');
    var res = await engine.evaluate(tour, tour.stops, fz, { start: startZeit(), serviceSec: standSek() });
    if (!fz) res.warnings.push('Für das Kennzeichen "' + (tour.kfz || '-') + '" ist kein Fahrzeug im Blatt "' + cfg.fahrzeugBlatt + '" hinterlegt - Route ohne Fahrzeugbeschränkung.');
    if (tour.kfzAlle.length > 1) res.warnings.push('In dieser Tour kommen mehrere Kennzeichen vor (' + tour.kfzAlle.join(', ') + ') - verwendet wird ' + tour.kfz + '.');
    S.results[name] = res;
    if (!ohneAnzeige) await zeigeErgebnis(name);
    return res;
  }

  function kennzahl(wert, text, gross) { return '<div class="kz' + (gross ? ' gross' : '') + '"><b>' + wert + '</b><span>' + text + '</span></div>'; }

  async function zeigeErgebnis(name, layersExtra) {
    var res = S.results[name], tour = S.tours[name]; S.aktuell = name;
    var e = $('ergebnis'); e.hidden = false;
    var h = '<h2>' + esc(name) + ' · ' + esc(tour.kfz || 'kein KFZ') + (res.fz && res.fz.typ ? ' · ' + esc(res.fz.typ) : '') + '</h2>';
    res.warnings.forEach(function (w) { h += '<div class="hinweis">' + esc(w) + '</div>'; });
    if (res.skipped.length) h += '<div class="hinweis">Nicht gefunden (in der Route ausgelassen): ' + res.skipped.map(function (s) { return esc(s.name) + ' (' + esc(s.strasse) + ', ' + esc(s.ort) + ')' + (s.geoFehler ? ' - ' + esc(s.geoFehler) : ''); }).join('; ') + '</div>';
    if (res.ors) {
      var svc = res.stops.length * res.serviceSec, g = res.google;
      h += '<div class="kennzahlen">' +
        kennzahl(T.fmtKm(res.ors.dist), 'Strecke (LKW-Route)') +
        kennzahl(res.stops.length + ' Kunden', Math.round(tour.gewicht) + ' kg gesamt') +
        kennzahl(T.fmtDur(res.ors.dur), 'Fahrzeit LKW (Standard)') +
        kennzahl(g ? T.fmtDur(g.dur) : '–', 'Fahrzeit live (Google)' + (g ? ', ohne Verkehr ' + T.fmtDur(g.staticDur) : '')) +
        kennzahl(T.fmtDur(svc), 'Standzeit gesamt') +
        kennzahl(T.fmtClock(g ? g.back : new Date(startZeit().getTime() + (res.ors.dur + svc) * 1000)), 'Rückkehr ca. ' + (g ? '(live)' : '(Standardzeit)')) +
        '</div>';
      if (g && g.adjusted) h += '<div class="hinweis">Die gewählte Abfahrt liegt in der Vergangenheit - die Live-Zeit gilt ab jetzt.</div>';
      if (!S.keys.google) h += '<div class="hinweis">Ohne Google-Schlüssel gibt es keine Live-Zeit und keine Karte.</div>';
      h += '<table><thead><tr><th class="l">#</th><th class="l">Kunde</th><th>kg</th><th>km</th><th>LKW</th><th>Live</th><th>an</th></tr></thead><tbody>';
      res.stops.forEach(function (s, i) {
        var l = res.ors.legs[i] || {}, gl = g && g.legs[i];
        h += '<tr class="klick" data-idx="' + s.zeilen[0] + '"><td class="l"><span class="nr">' + (i + 1) + '</span></td>' +
          '<td class="l">' + esc(s.name) + (s.weak ? ' ⚠' : '') + '<div class="sub">' + esc(s.strasse) + ', ' + esc(s.ort) + '</div></td>' +
          '<td>' + Math.round(s.gew) + '</td><td>' + ((l.dist || 0) / 1000).toFixed(1).replace('.', ',') + '</td>' +
          '<td>' + Math.round((l.dur || 0) / 60) + '′</td><td>' + (gl ? Math.round(gl.dur / 60) + '′' : '–') + '</td>' +
          '<td>' + (g ? T.fmtClock(g.eta[i]) : '–') + '</td></tr>';
      });
      var lr = res.ors.legs[res.stops.length] || {}, glr = g && g.legs[res.stops.length];
      h += '<tr><td class="l"></td><td class="l">Rückfahrt zum Depot</td><td></td><td>' + ((lr.dist || 0) / 1000).toFixed(1).replace('.', ',') + '</td><td>' + Math.round((lr.dur || 0) / 60) + '′</td><td>' + (glr ? Math.round(glr.dur / 60) + '′' : '–') + '</td><td>' + (g ? T.fmtClock(g.back) : '–') + '</td></tr>';
      h += '</tbody></table><p class="klein">⚠ = Adresse nur ungefähr gefunden, bitte prüfen. Zeile anklicken = in Excel anzeigen. LKW = Standardfahrzeit des LKW-Routings, Live = Google mit aktuellem/prognostiziertem Verkehr (Pkw-Basis).</p>';
    }
    e.innerHTML = h;
    Array.prototype.forEach.call(e.querySelectorAll('tr.klick'), function (tr) {
      tr.addEventListener('click', function () { Daten.zeileZeigen(Number(tr.getAttribute('data-idx'))).catch(function () {}); });
    });
    var layers = [{ res: res, farbe: FARBEN[Object.keys(S.tours).sort().indexOf(name) % FARBEN.length], nummern: true }];
    await zeichne(layersExtra ? layersExtra.concat(layers) : layers, ['ergebnis']);
  }

  // ---------- Optimieren ----------
  async function optimiere() {
    var name = $('fahrer').value, tour = S.tours[name];
    await lokalisieren(tour);
    var fz = fzFuer(tour);
    status('Aktuelle Reihenfolge wird berechnet ...');
    var aktuell = await engine.evaluate(tour, tour.stops, fz, { start: startZeit(), serviceSec: standSek() });
    status('Beste Reihenfolge für den LKW wird gesucht ...');
    var order = await engine.optimize(tour, aktuell.stops, fz, standSek());
    status('Vorschlag wird bewertet ...');
    var neu = await engine.evaluate(tour, order.concat(aktuell.skipped), fz, { start: startZeit(), serviceSec: standSek() });
    S.results[name] = aktuell; S.vorschlag = { name: name, res: neu, order: order, skipped: aktuell.skipped };
    await zeigeErgebnis(name, null);
    zeigeVorschlag(aktuell, neu);
    var layers = [{ res: aktuell, farbe: '#6b7280', gestrichelt: true, nummern: false }, { res: neu, farbe: '#c2410c', nummern: true }];
    await zeichne(layers, ['vorschlag', 'ergebnis']);
    status('Vorschlag fertig - grau gestrichelt = bisherige Route, orange = Vorschlag.', 'ok');
  }

  function diffZeile(text, alt, neu, fmt) {
    var d = neu - alt, cls = d > 0 ? 'diff-plus' : 'diff-minus', vz = d > 0 ? '+' : '−';
    return '<tr><td class="l">' + text + '</td><td>' + fmt(alt) + '</td><td>' + fmt(neu) + '</td><td class="' + cls + '">' + (Math.abs(d) < 1 ? '±0' : vz + fmt(Math.abs(d))) + '</td></tr>';
  }
  function zeigeVorschlag(a, n) {
    var v = $('vorschlag'); v.hidden = false;
    var h = '<h2>Optimierungsvorschlag für ' + esc(S.vorschlag.name) + '</h2><table><thead><tr><th class="l"></th><th>bisher</th><th>Vorschlag</th><th>Differenz</th></tr></thead><tbody>';
    h += diffZeile('Strecke', a.ors.dist, n.ors.dist, T.fmtKm);
    h += diffZeile('Fahrzeit LKW', a.ors.dur, n.ors.dur, T.fmtDur);
    if (a.google && n.google) h += diffZeile('Fahrzeit live', a.google.dur, n.google.dur, T.fmtDur);
    h += '</tbody></table><ol>';
    n.stops.forEach(function (s) { h += '<li>' + esc(s.name) + '<span class="sub"> · ' + esc(s.strasse) + ', ' + esc(s.ort) + '</span></li>'; });
    h += '</ol><p class="klein">Der Vorschlag berücksichtigt nur Fahrstrecke und Standzeit - keine Lieferzeitfenster, Kühlketten oder Ladereihenfolge. Bitte fachlich prüfen.</p>' +
      '<div class="zeile"><button id="btnUebernehmen" class="primaer">Reihenfolge in Tabelle übernehmen</button><button id="btnVerwerfen" class="sek">Verwerfen</button></div>';
    v.innerHTML = h;
    $('btnUebernehmen').addEventListener('click', uebernehmen);
    $('btnVerwerfen').addEventListener('click', function () { S.vorschlag = null; v.hidden = true; zeigeErgebnis(S.aktuell); status(''); });
  }

  async function uebernehmen() {
    var vs = S.vorschlag; if (!vs || S.busy) return;
    var tour = S.tours[vs.name];
    if (!window.confirm('Die Spalte "Reihung final" von ' + vs.name + ' wird neu nummeriert (' + vs.res.stops.length + ' Kunden, ' + tour.zeilen + ' Zeilen). Fortfahren?\n\nDas lässt sich nur über "Rückgängig" in diesem Add-in zurücknehmen, nicht mit Strg+Z.')) return;
    setBusy(true);
    try {
      var updates = [], pos = 1;
      vs.order.concat(vs.skipped).forEach(function (s) { s.zeilen.forEach(function (idx) { updates.push({ idx: idx, wert: pos }); }); pos++; });
      var alt = await Daten.reihungSchreiben(updates);
      lsSet('tp_undo', JSON.stringify({ tour: vs.name, alt: alt, zeit: Date.now() }));
      status('Reihenfolge übernommen. ' + updates.length + ' Zeilen neu nummeriert.', 'ok');
      $('vorschlag').hidden = true; S.vorschlag = null;
      await datenLaden(); $('fahrer').value = vs.name; await berechne(vs.name);
      undoAnzeigen();
    } catch (e) { status(e.message, 'fehler'); }
    setBusy(false);
  }
  function undoAnzeigen() {
    var u; try { u = JSON.parse(lsGet('tp_undo') || 'null'); } catch (e) { u = null; }
    if (!u) return;
    var e = $('vorschlag'); e.hidden = false;
    e.innerHTML = '<div class="zeile"><span class="klein">Letzte Übernahme: ' + esc(u.tour) + ' (' + new Date(u.zeit).toLocaleString('de-AT') + ')</span><button id="btnUndo" class="sek">Rückgängig</button></div>';
    $('btnUndo').addEventListener('click', async function () {
      setBusy(true);
      try { await Daten.reihungPruefenUndSetzen(u.alt); lsSet('tp_undo', ''); status('Rückgängig gemacht.', 'ok'); e.hidden = true; await datenLaden(); }
      catch (err) { status(err.message, 'fehler'); }
      setBusy(false);
    });
  }

  // ---------- Alle Fahrer ----------
  async function alleFahrer() {
    var namen = Object.keys(S.tours).sort(function (a, b) { return a.localeCompare(b, 'de'); });
    var fehlerGesamt = [];
    for (var i = 0; i < namen.length; i++) {
      status('(' + (i + 1) + '/' + namen.length + ') ' + namen[i] + ' ...');
      try { await berechne(namen[i], true); } catch (e) { fehlerGesamt.push(namen[i] + ': ' + e.message); }
      await sleep(S.demo ? 0 : 1500);
    }
    var u = $('uebersicht'); u.hidden = false;
    var h = '<h2>Übersicht aller Touren</h2><table><thead><tr><th class="l">Fahrer</th><th>Kd.</th><th>km</th><th>LKW</th><th>Live</th><th>zurück</th></tr></thead><tbody>';
    var tk = 0, tkd = 0;
    namen.forEach(function (n, i) {
      var r = S.results[n];
      if (!r || !r.ors) { h += '<tr><td class="l">' + esc(n) + '</td><td colspan="5">keine Berechnung</td></tr>'; return; }
      tk += r.ors.dist; tkd += r.stops.length;
      h += '<tr class="klick" data-n="' + esc(n) + '"><td class="l"><span class="nr" style="background:' + FARBEN[i % FARBEN.length] + '">&nbsp;</span> ' + esc(n) + '<div class="sub">' + esc(S.tours[n].kfz) + '</div></td><td>' + r.stops.length + '</td><td>' + (r.ors.dist / 1000).toFixed(0) + '</td><td>' + T.fmtDur(r.ors.dur) + '</td><td>' + (r.google ? T.fmtDur(r.google.dur) : '–') + '</td><td>' + (r.google ? T.fmtClock(r.google.back) : '–') + '</td></tr>';
    });
    h += '<tr><td class="l"><b>Summe</b></td><td><b>' + tkd + '</b></td><td><b>' + (tk / 1000).toFixed(0) + '</b></td><td></td><td></td><td></td></tr></tbody></table>';
    if (fehlerGesamt.length) h += '<div class="hinweis">' + fehlerGesamt.map(esc).join('<br>') + '</div>';
    h += '<div class="zeile"><button id="btnAlleKarte" class="sek">Alle Routen auf der Karte zeigen</button></div>';
    u.innerHTML = h;
    Array.prototype.forEach.call(u.querySelectorAll('tr.klick'), function (tr) {
      tr.addEventListener('click', function () { var n = tr.getAttribute('data-n'); $('fahrer').value = n; zeigeErgebnis(n); });
    });
    $('btnAlleKarte').addEventListener('click', function () {
      zeichne(namen.filter(function (n) { return S.results[n] && S.results[n].ors; }).map(function (n) { return { res: S.results[n], farbe: FARBEN[namen.indexOf(n) % FARBEN.length], nummern: false }; }), ['uebersicht']);
    });
    status(fehlerGesamt.length ? 'Fertig, aber mit Hinweisen.' : 'Alle Touren berechnet.', fehlerGesamt.length ? 'fehler' : 'ok');
  }

  // ---------- Einstellungen ----------
  async function einstellungenSpeichern() {
    var ors = $('keyOrs').value.trim(), g = $('keyGoogle').value.trim(), dep = $('depotAdresse').value.trim();
    lsSet('tp_keys', JSON.stringify({ ors: ors, google: g })); if (dep) lsSet('tp_depot', dep);
    S.keys.ors = ors; S.keys.google = g; if (dep) cfg.depot.adresse = dep;
    mv.reset();
    var out = $('einstStatus'); out.textContent = 'Teste ...';
    var teile = [];
    try { var d = await engine.getDepot(); teile.push('Depot gefunden (' + d.lat.toFixed(4) + ', ' + d.lon.toFixed(4) + ').'); }
    catch (e) { teile.push('OpenRouteService: ' + e.message); }
    if (g) { try { var okMap = await mv.ensure(); teile.push(okMap ? 'Google Maps geladen.' : 'Google Maps konnte nicht geladen werden (Schlüssel/Freigaben prüfen).'); } catch (e) { teile.push('Google: ' + e.message); } }
    else teile.push('Kein Google-Schlüssel: ohne Karte und Live-Zeit.');
    out.textContent = teile.join(' ');
  }

  // ---------- Start ----------
  function wrap(fn) {
    return async function () {
      if (S.busy) return;
      setBusy(true);
      try { await fn(); } catch (e) { status(e.message || String(e), 'fehler'); }
      setBusy(false);
    };
  }

  async function start() {
    ladeKeys();
    if (S.demo) { S.keys.ors = S.keys.ors || 'demo'; S.keys.google = S.keys.google || 'demo'; }
    $('keyOrs').value = S.keys.ors; $('keyGoogle').value = S.keys.google; $('depotAdresse').value = cfg.depot.adresse;
    $('standzeit').value = cfg.standzeitMin; $('zeit').value = cfg.abfahrt;
    var d = new Date(); if (d.getHours() >= 12) d.setDate(d.getDate() + 1);
    $('datum').value = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
    $('btnEinst').addEventListener('click', function () { $('einstellungen').hidden = !$('einstellungen').hidden; });
    $('btnSpeichern').addEventListener('click', einstellungenSpeichern);
    $('btnCacheLeeren').addEventListener('click', async function () {
      try { await Daten.cacheLeeren(); S.geoCache = {}; $('einstStatus').textContent = 'Merkliste geleert.'; } catch (e) { $('einstStatus').textContent = e.message; }
    });
    // Autostart: wird in der Mappe selbst gespeichert (Dokumenteinstellung), gilt dann fuer jeden, der sie oeffnet
    var auto = $('chkAutostart'), dokSet = null;
    try { dokSet = (!S.demo && Office.context.document.settings) || null; } catch (e) { dokSet = null; }
    if (!dokSet) { $('autostartZeile').hidden = true; }
    else {
      auto.checked = dokSet.get('Office.AutoShowTaskpaneWithDocument') === true;
      auto.addEventListener('change', function () {
        dokSet.set('Office.AutoShowTaskpaneWithDocument', auto.checked);
        dokSet.saveAsync(function (r) {
          $('einstStatus').textContent = r.status === Office.AsyncResultStatus.Succeeded
            ? 'Gespeichert. Bitte die Mappe speichern, damit die Einstellung dauerhaft in der Datei steht.' : 'Konnte nicht gespeichert werden.';
        });
      });
    }
    var chk = $('chkFenster');
    chk.checked = fensterModus(); chk.disabled = !dialogMoeglich();
    if (!dialogMoeglich()) chk.parentNode.hidden = true;
    chk.addEventListener('change', function () {
      lsSet('tp_fenster', chk.checked ? '1' : '0');
      if (chk.checked) { if (S.letzteLayers) zeigePlain(S.letzteLayers, S.letzteBereiche); }
      else { dialogSchliessen(); if (S.letzteLayers) mv.draw(S.letzteLayers); }
    });
    $('btnNeu').addEventListener('click', wrap(datenLaden));
    $('btnBerechnen').addEventListener('click', wrap(function () { S.vorschlag = null; $('vorschlag').hidden = true; return berechne($('fahrer').value); }));
    $('btnOptimieren').addEventListener('click', wrap(optimiere));
    $('btnAlle').addEventListener('click', wrap(alleFahrer));
    $('fahrer').addEventListener('change', function () {
      S.vorschlag = null; $('vorschlag').hidden = true;
      var n = $('fahrer').value; if (S.results[n]) zeigeErgebnis(n); else { $('ergebnis').hidden = true; }
    });
    if (!S.keys.ors) { $('einstellungen').hidden = false; $('einstStatus').textContent = 'Bitte zuerst die Schlüssel eintragen.'; }
    setBusy(true);
    try {
      S.geoCache = await Daten.cacheLesen();
      await datenLaden();
      if (S.keys.google && !fensterModus()) await mv.ensure();
    } catch (e) { status(e.message || String(e), 'fehler'); }
    setBusy(false);
    undoAnzeigen();
  }

  var demo = /[?&]demo=1/.test(location.search);
  function los() {
    var go = function () { S.demo = demo; start(); };
    if (demo) { var s = document.createElement('script'); s.src = 'demo.js'; s.onload = go; document.head.appendChild(s); return; }
    if (window.Office && Office.onReady) Office.onReady(function (info) {
      if (info && info.host === Office.HostType.Excel) go();
      else { status('Dieses Add-in läuft nur in Excel. Zum Ansehen: ?demo=1 an die Adresse anhängen.', 'fehler'); setBusy(true); }
    });
    else { status('Office.js konnte nicht geladen werden (Internetverbindung?).', 'fehler'); }
  }
  los();
})();
