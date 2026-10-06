/* engine.js - Berechnungs-Ablauf (Geocoding, LKW-Route, Live-Zeit, Optimierung).
   Alle Abhaengigkeiten werden uebergeben, damit der Ablauf ohne Browser getestet werden kann. */
(function (root) {
  'use strict';
  var T = (typeof module !== 'undefined' && module.exports) ? require('./lib.js') : root.TP;

  function createEngine(deps) {
    var cfg = deps.cfg, fetchFn = deps.fetch, sleep = deps.sleep || function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
    var cache = deps.cache;                   // { get(key), put(key, entry), flush() }
    var keys = function () { return deps.getKeys(); };

    function errText(body, status) {
      if (body && body.error) {
        if (typeof body.error === 'string') return body.error;
        if (body.error.message) return body.error.message;
      }
      return 'HTTP ' + status;
    }

    async function jfetch(url, opts, label) {
      for (var versuch = 0; versuch < 4; versuch++) {
        var res = await fetchFn(url, opts);
        var body = null;
        try { body = await res.json(); } catch (e) { body = null; }
        if (res.status === 429 && versuch < 3) { await sleep(4000 * (versuch + 1)); continue; }
        if (!res.ok) { var err = new Error(label + ': ' + errText(body, res.status)); err.status = res.status; throw err; }
        return body;
      }
    }

    function orsHeaders() {
      var k = keys().ors;
      if (!k) throw new Error('OpenRouteService-Schlüssel fehlt (Einstellungen).');
      return { 'Authorization': k, 'Content-Type': 'application/json', 'Accept': 'application/json, application/geo+json' };
    }

    // ---------- Geocoding ----------
    async function geocodeAddress(strasse, plz, ort) {
      var p = T.parsePlz(plz);
      var h = orsHeaders();
      var u = cfg.ors + '/geocode/search/structured?address=' + encodeURIComponent(strasse) +
        '&postalcode=' + encodeURIComponent(p.zip) + '&locality=' + encodeURIComponent(ort) +
        '&country=' + encodeURIComponent(p.country) + '&size=1';
      var body = await jfetch(u, { headers: h }, 'Geocoding');
      if (!body.features || !body.features.length) {
        u = cfg.ors + '/geocode/search?text=' + encodeURIComponent(T.addressText(strasse, plz, ort)) +
          '&boundary.country=' + encodeURIComponent(p.country) + '&size=1';
        body = await jfetch(u, { headers: h }, 'Geocoding');
      }
      var f = body.features && body.features[0];
      if (!f) return null;
      var pr = f.properties || {};
      var weak = (pr.confidence != null && pr.confidence < 0.6) || pr.accuracy === 'centroid' || pr.match_type === 'fallback';
      return { lat: f.geometry.coordinates[1], lon: f.geometry.coordinates[0], weak: !!weak, label: pr.label || '' };
    }

    async function geocodeStops(stops, onProgress) {
      var fehler = [], offen = 0;
      for (var i = 0; i < stops.length; i++) {
        var s = stops[i];
        var c = cache.get(s.addrKey);
        if (!c) {
          offen++;
          try {
            c = await geocodeAddress(s.strasse, s.plz, s.ort);
          } catch (e) { fehler.push(s); s.geoFehler = e.message; if (onProgress) onProgress(i + 1, stops.length); continue; }
          if (c) cache.put(s.addrKey, c);
          await sleep(cfg.geocodeWartezeitMs == null ? 700 : cfg.geocodeWartezeitMs);
        }
        if (c) { s.lat = c.lat; s.lon = c.lon; s.weak = !!c.weak; s.geoLabel = c.label; }
        else { fehler.push(s); }
        if (onProgress) onProgress(i + 1, stops.length);
      }
      if (offen) await cache.flush();
      return fehler;
    }

    async function getDepot() {
      var d = cfg.depot;
      if (d.lat != null && d.lon != null) return { lat: d.lat, lon: d.lon };
      var c = cache.get('depot|' + d.adresse);
      if (!c) {
        var body = await jfetch(cfg.ors + '/geocode/search?text=' + encodeURIComponent(d.adresse) + '&size=1', { headers: orsHeaders() }, 'Depot-Geocoding');
        var f = body.features && body.features[0];
        if (!f) throw new Error('Depot-Adresse nicht gefunden: ' + d.adresse);
        c = { lat: f.geometry.coordinates[1], lon: f.geometry.coordinates[0], weak: false, label: f.properties && f.properties.label || '' };
        cache.put('depot|' + d.adresse, c);
        await cache.flush();
      }
      return { lat: c.lat, lon: c.lon };
    }

    // ---------- LKW-Route (OpenRouteService, driving-hgv) ----------
    async function orsRoute(points, fz) {   // points: [[lat,lon],...]
      var legs = [], path = [], dist = 0, dur = 0;
      var chunks = T.chunkPoints(points, 50);
      for (var c = 0; c < chunks.length; c++) {
        var coords = chunks[c].map(function (p) { return [p[1], p[0]]; });
        var body = await jfetch(cfg.ors + '/v2/directions/driving-hgv', {
          method: 'POST', headers: orsHeaders(),
          body: JSON.stringify({ coordinates: coords, instructions: false, preference: 'recommended', options: T.orsOptions(fz, cfg) })
        }, 'LKW-Route');
        var r = body.routes && body.routes[0];
        if (!r) throw new Error('LKW-Route: keine Route gefunden.');
        dist += r.summary.distance; dur += r.summary.duration;
        (r.segments || []).forEach(function (sg) { legs.push({ dist: sg.distance, dur: sg.duration }); });
        var pts = T.decodePolyline(r.geometry, 5);
        path = path.concat(c ? pts.slice(1) : pts);
      }
      return { dist: dist, dur: dur, legs: legs, path: path };
    }

    // ---------- Live-Verkehr (Google Routes API) ----------
    async function googleTimes(points, start, serviceSec) {
      var k = keys().google;
      if (!k) return null;
      var chunks = T.chunkPoints(points, 27);      // Start + Ziel + max. 25 Zwischenstopps
      var t = start.getTime(), now = Date.now() + 30000, adjusted = false;
      var legs = [], eta = [], dur = 0, stat = 0, dist = 0, idx = 0;
      var last = points.length - 1;
      var wp = function (p) { return { location: { latLng: { latitude: p[0], longitude: p[1] } } }; };
      for (var c = 0; c < chunks.length; c++) {
        var ch = chunks[c];
        var dep = t;
        if (dep < now) { dep = now; adjusted = true; }
        var body = await jfetch(cfg.googleRoutes, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json', 'X-Goog-Api-Key': k,
            'X-Goog-FieldMask': 'routes.duration,routes.staticDuration,routes.distanceMeters,routes.legs.duration,routes.legs.staticDuration,routes.legs.distanceMeters'
          },
          body: JSON.stringify({
            origin: wp(ch[0]), destination: wp(ch[ch.length - 1]),
            intermediates: ch.slice(1, -1).map(wp),
            travelMode: 'DRIVE', routingPreference: 'TRAFFIC_AWARE',
            departureTime: new Date(dep).toISOString(), languageCode: 'de-AT', units: 'METRIC'
          })
        }, 'Google Routes');
        var r = body.routes && body.routes[0];
        if (!r || !r.legs) throw new Error('Google Routes: keine Route gefunden.');
        r.legs.forEach(function (lg) {
          var d = T.parseGoogleDuration(lg.duration), sd = T.parseGoogleDuration(lg.staticDuration || lg.duration);
          legs.push({ dur: d, staticDur: sd, dist: lg.distanceMeters || 0 });
          dur += d; stat += sd; dist += lg.distanceMeters || 0;
          t += d * 1000; idx++;
          eta.push(new Date(t));
          if (idx < last) t += serviceSec * 1000;      // Standzeit beim Kunden (nicht am End-Depot)
        });
      }
      return { dur: dur, staticDur: stat, dist: dist, legs: legs, eta: eta, back: new Date(t), adjusted: adjusted };
    }

    // ---------- Gesamtauswertung einer Tour ----------
    async function evaluate(tour, orderedStops, fz, opts) {
      var depot = await getDepot();
      var geo = orderedStops.filter(function (s) { return s.lat != null; });
      var skipped = orderedStops.filter(function (s) { return s.lat == null; });
      var res = { tour: tour.name, kfz: tour.kfz, fz: fz, depot: depot, stops: geo, skipped: skipped, ors: null, google: null, warnings: [], serviceSec: opts.serviceSec };
      if (!geo.length) { res.warnings.push('Keine Adresse konnte lokalisiert werden.'); return res; }
      var pts = [[depot.lat, depot.lon]].concat(geo.map(function (s) { return [s.lat, s.lon]; }), [[depot.lat, depot.lon]]);
      res.ors = await orsRoute(pts, fz);
      if (keys().google) {
        try { res.google = await googleTimes(pts, opts.start, opts.serviceSec); }
        catch (e) { res.warnings.push(e.message); }
      }
      return res;
    }

    // ---------- Optimierung (ORS / VROOM, LKW-Profil) ----------
    async function optimize(tour, stops, fz, serviceSec) {
      var depot = await getDepot();
      var geo = stops.filter(function (s) { return s.lat != null; });
      if (geo.length < 3) throw new Error('Für eine Optimierung braucht es mindestens 3 lokalisierte Kunden.');
      var body = await jfetch(cfg.ors + '/optimization', {
        method: 'POST', headers: orsHeaders(),
        body: JSON.stringify({
          jobs: geo.map(function (s, i) { return { id: i + 1, location: [s.lon, s.lat], service: serviceSec }; }),
          vehicles: [{ id: 1, profile: 'driving-hgv', start: [depot.lon, depot.lat], end: [depot.lon, depot.lat] }]
        })
      }, 'Optimierung');
      var order = T.orderFromOptimization(body);
      if (order.length !== geo.length) throw new Error('Optimierung: nicht alle Kunden konnten eingeplant werden.');
      return order.map(function (id) { return geo[id - 1]; });
    }

    return { geocodeStops: geocodeStops, getDepot: getDepot, orsRoute: orsRoute, googleTimes: googleTimes, evaluate: evaluate, optimize: optimize, geocodeAddress: geocodeAddress };
  }

  var api = { createEngine: createEngine };
  root.TP = root.TP || {};
  root.TP.createEngine = createEngine;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
