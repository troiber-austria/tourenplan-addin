/* mapview.js - Google-Karte mit Routen und nummerierten Stopps.
   Wird im Seitenbereich (index.html) und im grossen Kartenfenster (karte.html) verwendet.
   Eine "Schicht" (layer) ist ein einfaches Objekt:
   { farbe, gestrichelt, nummern, depot:{lat,lon,name}, path:[[lat,lon],...], stops:[{name,strasse,ort,gew,lat,lon,weak,eta}] } */
(function (root) {
  'use strict';
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

  function create(el, getKey) {
    var map = null, overlays = [], info = null, laden = null;

    function ladeGoogle() {
      if (laden) return laden;
      laden = new Promise(function (resolve, reject) {
        var key = getKey();
        if (!key) return reject(new Error('kein Schlüssel'));
        root.gm_authFailure = function () { el.innerHTML = '<div class="karte-leer">Google lehnt den Schlüssel ab (Maps JavaScript API aktiviert? Abrechnung verknüpft? Einschränkungen des Schlüssels prüfen).</div>'; };
        root.__gmInit = function () { resolve(); };
        var s = document.createElement('script');
        s.src = 'https://maps.googleapis.com/maps/api/js?key=' + encodeURIComponent(key) + '&loading=async&v=weekly&language=de&region=AT&callback=__gmInit';
        s.onerror = function () { reject(new Error('Google Maps konnte nicht geladen werden.')); };
        document.head.appendChild(s);
      }).then(async function () {
        await google.maps.importLibrary('maps'); await google.maps.importLibrary('marker');
      });
      laden.catch(function () { laden = null; });
      return laden;
    }

    async function ensure() {
      if (map) return true;
      try { await ladeGoogle(); } catch (e) { return false; }
      el.innerHTML = '';
      map = new google.maps.Map(el, { center: { lat: 48.2, lng: 16.37 }, zoom: 9, mapId: 'DEMO_MAP_ID', mapTypeControl: false, streetViewControl: false, fullscreenControl: false });
      info = new google.maps.InfoWindow();
      return true;
    }

    function clear() { overlays.forEach(function (o) { if (o.setMap) o.setMap(null); else o.map = null; }); overlays = []; }
    function reset() { clear(); map = null; laden = null; }

    function marker(pos, glyph, farbe, html) {
      var pin = new google.maps.marker.PinElement({ glyph: String(glyph), background: farbe, borderColor: farbe, glyphColor: '#ffffff', scale: 0.9 });
      var m = new google.maps.marker.AdvancedMarkerElement({ map: map, position: pos, content: pin.element });
      m.addListener('click', function () { info.setContent(html); info.open({ map: map, anchor: m }); });
      overlays.push(m);
    }

    async function draw(layers) {
      if (!(await ensure())) return false;
      clear();
      var b = new google.maps.LatLngBounds(), erstes = true;
      layers.forEach(function (l) {
        if (!l || !l.path) return;
        var path = l.path.map(function (p) { return { lat: p[0], lng: p[1] }; });
        var o = { path: path, map: map, strokeColor: l.farbe, strokeOpacity: l.gestrichelt ? 0 : 0.85, strokeWeight: l.gestrichelt ? 3 : 5 };
        if (l.gestrichelt) o.icons = [{ icon: { path: 'M 0,-1 0,1', strokeOpacity: 1, scale: 3 }, offset: '0', repeat: '14px' }];
        overlays.push(new google.maps.Polyline(o));
        path.forEach(function (p) { b.extend(p); });
        if (erstes && l.depot) {
          marker({ lat: l.depot.lat, lng: l.depot.lon }, 'D', '#111827', '<b>' + esc(l.depot.name) + '</b><br>Depot (Start/Ziel)');
          erstes = false;
        }
        if (l.nummern) (l.stops || []).forEach(function (s, i) {
          marker({ lat: s.lat, lng: s.lon }, i + 1, l.farbe,
            '<b>' + (i + 1) + '. ' + esc(s.name) + '</b><br>' + esc(s.strasse) + ', ' + esc(s.ort) + '<br>' + s.gew + ' kg' +
            (s.eta ? ' · Ankunft ca. ' + esc(s.eta) : '') + (s.weak ? '<br>⚠ Adresse nur ungefähr lokalisiert' : ''));
        });
      });
      if (!b.isEmpty()) map.fitBounds(b, 40);
      return true;
    }

    return { ensure: ensure, draw: draw, clear: clear, reset: reset };
  }

  root.TP = root.TP || {};
  root.TP.MapView = { create: create };
})(window);
