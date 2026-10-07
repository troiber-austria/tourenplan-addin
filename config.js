/* config.js - nicht geheime Einstellungen. API-Schluessel werden NICHT hier gespeichert,
   sondern im Add-in unter "Einstellungen" eingegeben (bleiben nur auf dem jeweiligen PC). */
window.TP_CONFIG = {
  depot: {
    name: 'Troiber',
    adresse: 'Industriestraße 383, 2722 Weikersdorf am Steinfelde, Österreich',
    // Optional fest hinterlegen, sonst wird die Adresse beim ersten Start geocodiert und gemerkt:
    lat: null, lon: null
  },
  tabelle: 'Tabelle1',          // Name der Excel-Tabelle im Blatt Tourenplan
  fahrzeugBlatt: 'Fahrzeuge',   // Blatt mit Kennzeichen/Typ/Maßen
  cacheBlatt: 'GeoCache',       // wird automatisch (ausgeblendet) angelegt
  standzeitMin: 15,             // Standard-Standzeit je Kunde
  abfahrt: '06:00',             // Standard-Abfahrtszeit
  // Routingprofil je Fahrzeugtyp aus dem Blatt "Fahrzeuge" (OpenRouteService, Profil driving-hgv)
  typen: {
    '18-Tonner':  { vehicle_type: 'hgv' },
    '3,5-Tonner': { vehicle_type: 'delivery' }
  },
  snapRadius: 1500,             // Meter: so weit darf eine Adresse zur nächsten befahrbaren Straße entfernt liegen
  ors: 'https://api.openrouteservice.org',
  googleRoutes: 'https://routes.googleapis.com/directions/v2:computeRoutes'
};
