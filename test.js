const assert = require('assert');
const T = require('../lib.js');
const { createEngine } = require('../engine.js');
global.window = undefined;
const cfg = require('fs').readFileSync(__dirname + '/../config.js', 'utf8');
const TP_CONFIG = new Function('window', cfg + '; return window.TP_CONFIG;')({});

// --- lib ---
assert.deepStrictEqual(T.parsePlz('AT-1230'), { country: 'AT', zip: '1230' });
assert.deepStrictEqual(T.parsePlz('83313'), { country: 'DE', zip: '83313' });
assert.deepStrictEqual(T.parsePlz(1060), { country: 'AT', zip: '1060' });
// Google-Beispiel aus der Doku: _p~iF~ps|U_ulLnnqC_mqNvxq`@ -> (38.5,-120.2),(40.7,-120.95),(43.252,-126.453)
const dp = T.decodePolyline('_p~iF~ps|U_ulLnnqC_mqNvxq`@');
assert.deepStrictEqual(dp, [[38.5, -120.2], [40.7, -120.95], [43.252, -126.453]]);
assert.deepStrictEqual(T.chunkPoints([1,2,3,4,5,6,7], 3), [[1,2,3],[3,4,5],[5,6,7]]);
assert.deepStrictEqual(T.chunkPoints([1,2], 50), [[1,2]]);
assert.strictEqual(T.fmtDur(3725), '1 h 02 min'); assert.strictEqual(T.fmtKm(12345), '12,3 km');
assert.strictEqual(T.parseGoogleDuration('1234s'), 1234);

const headers = ['Kd-Nr.','Kunde','Identnummer','Strasse','Ort','Plz','Rb','Gew','TK RB','RB frisch','Tour orig','Reihung orig','Hilfsspalte Kunden','Reihung final','KFZ','Tour final'];
const row = (kd, name, str, ort, plz, gew, rf, kfz, tour) => [kd, name, 1, str, ort, plz, 0, gew, 0, 0, 'x', 1, 1, rf, kfz, tour];
const values = [
  row(1, 'A', 'Weg 1', 'Wien', 'AT-1010', 100, 2, 'WB 764EW', 'Max'),
  row(2, 'B', 'Weg 2', 'Wien', 'AT-1020', 50, 1, 'WB 764EW', 'Max'),
  row(2, 'B', 'Weg 2', 'Wien', 'AT-1020', 25, 1, 'WB 764EW', 'Max'),   // gleicher Kunde -> ein Stopp
  row(3, 'C', 'Weg 3', 'Wien', 'AT-1030', 10, 3, 'WB 764EW', 'Max'),
  row(4, 'D', 'Weg 4', 'Graz', 'AT-8010', 10, 1, 'WB 804FJ', 'Eva'),
  row(5, 'E', 'Weg 5', 'Graz', 'AT-8010', 10, '', '', ''),            // ohne Fahrer -> ignoriert
];
const g = T.groupTours(headers, values);
assert.deepStrictEqual(Object.keys(g.tours).sort(), ['Eva', 'Max']);
const max = g.tours.Max;
assert.strictEqual(max.stops.length, 3); assert.deepStrictEqual(max.stops.map(s => s.name), ['B', 'A', 'C']);
assert.strictEqual(max.stops[0].gew, 75); assert.deepStrictEqual(max.stops[0].zeilen, [1, 2]);
assert.strictEqual(max.kfz, 'WB 764EW'); assert.strictEqual(max.zeilen, 4);
assert.deepStrictEqual(T.groupTours(['a'], []).fehlend.length > 0, true);
assert.deepStrictEqual(T.orsOptions({ typ: '18-Tonner', gewicht: 18, hoehe: 4, breite: 2.55, laenge: 12 }, TP_CONFIG),
  { vehicle_type: 'hgv', profile_params: { restrictions: { weight: 18, height: 4, width: 2.55, length: 12 } } });
assert.deepStrictEqual(T.addressVariants('Wagramer Straße 79 Tür 604'), ['Wagramer Straße 79 Tür 604', 'Wagramer Straße 79', 'Wagramer Straße']);
assert.strictEqual(T.cleanStrasse('Wagramer Straße 94; Donauzentrum'), 'Wagramer Straße 94');
assert.strictEqual(T.cleanStrasse('Mariahilferstrasse 42-48'), 'Mariahilferstrasse 42');
console.log('lib ok');

// --- engine mit gefaelschtem Netzwerk ---
(async () => {
  const calls = [];
  const fakeFetch = async (url, opts = {}) => {
    calls.push({ url, opts });
    const ok = (b) => ({ ok: true, status: 200, json: async () => b });
    if (url.includes('/geocode/search/structured')) {
      const u = new URL(url); const n = Number(u.searchParams.get('address').replace(/\D/g, ''));
      if (n === 9 || n === 0) return ok({ features: [] });
      return ok({ features: [{ geometry: { coordinates: [16 + n / 100, 48 + n / 100] }, properties: { confidence: n === 3 ? 0.4 : 1, label: 'x', layer: 'address' } }] });
    }
    if (url.includes('/geocode/search') && url.includes('Industrie')) return ok({ features: [{ geometry: { coordinates: [16.2, 47.85] }, properties: { label: 'Depot', layer: 'address' } }] });
    if (url.includes('/geocode/search')) return ok({ features: [] });
    if (url.includes('/v2/directions/driving-hgv')) {
      const b = JSON.parse(opts.body); const n = b.coordinates.length;
      return ok({ routes: [{ summary: { distance: 1000 * (n - 1), duration: 600 * (n - 1) }, geometry: '_p~iF~ps|U_ulLnnqC_mqNvxq`@',
        segments: Array.from({ length: n - 1 }, () => ({ distance: 1000, duration: 600 })) }] });
    }
    if (url.includes('routes.googleapis.com')) {
      const b = JSON.parse(opts.body); const n = (b.intermediates || []).length + 1;
      return ok({ routes: [{ legs: Array.from({ length: n }, () => ({ duration: '700s', staticDuration: '600s', distanceMeters: 1100 })) }] });
    }
    if (url.includes('/optimization')) {
      const b = JSON.parse(opts.body);
      return ok({ routes: [{ steps: [{ type: 'start' }, ...b.jobs.map(j => j.id).reverse().map(id => ({ type: 'job', job: id })), { type: 'end' }] }] });
    }
    throw new Error('unerwartet ' + url);
  };
  const store = {};
  const eng = createEngine({
    cfg: Object.assign({}, TP_CONFIG, { geocodeWartezeitMs: 0 }), fetch: fakeFetch, sleep: async () => {},
    getKeys: () => ({ ors: 'KEY', google: 'GKEY' }),
    cache: { get: k => store[k], put: (k, v) => { store[k] = v; }, flush: async () => {} }
  });
  const stops = max.stops.concat([{ addrKey: 'z', strasse: 'Weg 9', plz: 'AT-1090', ort: 'Wien', name: 'Z', zeilen: [] }]);
  const fehler = await eng.geocodeStops(stops);
  assert.strictEqual(fehler.length, 1); assert.strictEqual(fehler[0].name, 'Z');
  assert.strictEqual(stops.find(s => s.name === 'C').weak, true);
  const bad = createEngine({ cfg: Object.assign({}, TP_CONFIG, { geocodeWartezeitMs: 0 }), sleep: async () => {}, getKeys: () => ({ ors: 'K' }),
    cache: { get: () => undefined, put() {}, flush: async () => {} },
    fetch: async () => ({ ok: true, status: 200, json: async () => ({ features: [{ geometry: { coordinates: [13.33333, 47.33333] }, properties: { layer: 'country', label: 'Austria', confidence: 1 } }] }) }) });
  assert.strictEqual(await bad.geocodeAddress('Nirgendwo 1', 'AT-9999', 'Ort'), null);
  const nCalls = calls.length; await eng.geocodeStops(max.stops);   // jetzt aus dem Cache
  assert.strictEqual(calls.length, nCalls);
  const fz = { typ: '18-Tonner', gewicht: 18, hoehe: 4, breite: 2.55, laenge: 12 };
  const start = new Date(Date.now() + 86400000);
  const ev = await eng.evaluate(max, stops, fz, { start, serviceSec: 900 });
  assert.strictEqual(ev.skipped.length, 1); assert.strictEqual(ev.stops.length, 3);
  assert.strictEqual(ev.ors.legs.length, 4); assert.strictEqual(ev.ors.dist, 4000);
  assert.strictEqual(ev.google.legs.length, 4); assert.strictEqual(ev.google.dur, 2800);
  // ETA Stopp 1 = Start + 700 s ; Stopp 2 = + 700 + 900 + 700
  assert.strictEqual(ev.google.eta[0] - start, 700000); assert.strictEqual(ev.google.eta[1] - start, 2300000);
  assert.strictEqual(ev.google.back - start, (2800 + 3 * 900) * 1000);
  const dirCall = calls.find(c => c.url.includes('driving-hgv'));
  assert.strictEqual(JSON.parse(dirCall.opts.body).options.profile_params.restrictions.weight, 18);
  assert.strictEqual(dirCall.opts.headers.Authorization, 'KEY');
  const opt = await eng.optimize(max, ev.stops, fz, 900);
  assert.deepStrictEqual(opt.map(s => s.name), ['C', 'A', 'B']);
  // Lange Tour: > 50 Punkte -> mehrere ORS-Aufrufe, Summen stimmen
  const viele = Array.from({ length: 60 }, (_, i) => [48 + i / 1000, 16]);
  const before = calls.length; const r = await eng.orsRoute(viele, fz);
  assert.strictEqual(calls.length - before, 2); assert.strictEqual(r.legs.length, 59);
  // Google: > 25 Zwischenstopps -> mehrere Aufrufe
  const before2 = calls.length; const gt = await eng.googleTimes(viele, start, 0);
  assert.ok(calls.length - before2 >= 3); assert.strictEqual(gt.legs.length, 59);
  console.log('engine ok');
})().catch(e => { console.error(e); process.exit(1); });
