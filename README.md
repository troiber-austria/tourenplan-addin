# Tourenplan-Karte (Excel-Add-in)

Zeigt die geplanten Touren aus `Tabelle1` auf einer Karte – mit LKW-Route, Kilometern, Fahrzeit (LKW-Routing und Google-Live-Verkehr), Ankunftszeiten je Kunde und einem Optimierungsvorschlag für die Reihenfolge.

* **Karte & Live-Zeit:** Google (Maps JavaScript API + Routes API)
* **LKW-Route, Adress-Suche, Optimierung:** OpenRouteService (Profil `driving-hgv`, mit Gewicht/Höhe/Breite/Länge aus dem Blatt „Fahrzeuge“)

## Was das Add-in liest und schreibt

| Blatt | Zugriff |
|---|---|
| `Tourenplan` / Tabelle1 | liest Kunde, Straße, Ort, PLZ, Gew, Reihung final, KFZ, Tour final. **Schreibt nur** die Spalte „Reihung final“ – ausschließlich nach Klick auf „Reihenfolge in Tabelle übernehmen“ und Bestätigung (mit „Rückgängig“-Knopf). |
| `Fahrzeuge` | liest Kennzeichen, Typ, Gewicht, Höhe, Breite, Länge |
| `GeoCache` | wird automatisch (ausgeblendet) angelegt und merkt sich die Koordinaten je Adresse, damit jede Adresse nur einmal gesucht wird |

## Einrichtung – Schritt für Schritt

### 1. Dateien nach GitHub

1. Auf github.com in der Organisation **troiber-austria**: *New repository* → Name **`tourenplan-addin`** → **Public** → *Create repository*.
   (GitHub Pages ist im kostenlosen Tarif nur für öffentliche Repositories möglich. Im Repository liegen **keine Schlüssel und keine Kundendaten** – nur Programmcode und erfundene Demo-Daten.)
2. *Add file → Upload files* → den **Inhalt** dieses Ordners hineinziehen (inkl. Ordner `icons` und `test`) → *Commit changes*.
3. *Settings → Pages* → *Source:* **Deploy from a branch**, *Branch:* **main**, Ordner **/ (root)** → *Save*.
4. Nach 1–2 Minuten ist die Seite erreichbar unter  
   `https://troiber-austria.github.io/tourenplan-addin/`  
   Test: `.../index.html?demo=1` im Browser öffnen → die Oberfläche läuft mit erfundenen Beispieldaten (ohne Excel, ohne Schlüssel).

### 2. Schlüssel besorgen (haben Sie bereits)

* **OpenRouteService:** Schlüssel (Token) aus dem Dashboard. Im kostenlosen Tarif gelten Tageslimits (u. a. Routen, Adress-Suche, Optimierung) – für einige Touren pro Tag reicht das.
* **Google Cloud:** ein Schlüssel, bei dem die APIs **Maps JavaScript API** und **Routes API** aktiviert sind.
  * Unter *Schlüssel einschränken → API-Einschränkungen* nur diese beiden APIs zulassen.
  * Unter *Kontingente* ein Tageslimit für die Routes API setzen und unter *Abrechnung → Budgets & Benachrichtigungen* ein Budget mit E-Mail-Warnung anlegen.

Die Schlüssel stehen **nicht** im Code. Sie werden im Add-in unter ⚙ *Einstellungen* eingetragen und bleiben im Browser-Speicher des jeweiligen PCs (jeder Benutzer trägt sie einmal ein).

### 3. Add-in in Excel einbinden

**Zum Ausprobieren (nur für Sie):**
* *Excel im Browser (Microsoft 365):* Datei öffnen → *Einfügen → Add-Ins → Meine Add-Ins → Eigenes Add-In hochladen* → `manifest.xml` wählen.
* *Excel am PC:* `manifest.xml` in einen Netzwerkordner legen → *Datei → Optionen → Trust Center → Einstellungen für das Trust Center → Vertrauenswürdige Add-In-Kataloge* → Ordnerpfad hinzufügen, „Im Menü anzeigen“ anhaken → Excel neu starten → *Einfügen → Meine Add-Ins → Freigegebener Ordner*.

**Für alle Mitarbeiter (empfohlen):**
1. Als Administrator im *Microsoft 365 Admin Center* anmelden → *Einstellungen → Integrierte Apps → Benutzerdefinierte Apps hochladen* (je nach Version: *Add-Ins → Add-In bereitstellen*).
2. `manifest.xml` hochladen, Benutzer bzw. Gruppen auswählen, bereitstellen.
3. Es kann bis zu 24 Stunden dauern, bis die Schaltfläche **Start → Tourenplan → Karte** erscheint.

*Optional – Karte fest im Tabellenblatt:* `manifest-inhalt.xml` ist die Variante als „Inhalts-Add-in“. Es wird über *Einfügen → Add-Ins* in das Blatt gesetzt (z. B. in ein eigenes Blatt „Karte“), bleibt dort an seiner Position und wird mit der Datei gespeichert. Es nutzt dieselbe Seite.

### 4. Erster Start

1. Karte öffnen → ⚙ → beide Schlüssel eintragen → *Speichern & testen*. Das Depot (Industriestraße 383, Weikersdorf) wird gefunden, Google Maps lädt.
2. Fahrer wählen → *Route berechnen*. Beim ersten Mal werden die Adressen gesucht (ca. 0,7 s je neue Adresse), danach kommen sie aus der Merkliste.
3. *Optimieren* zeigt die bisherige (grau gestrichelt) gegen die vorgeschlagene Route (orange); *Reihenfolge in Tabelle übernehmen* schreibt die neue Nummerierung in „Reihung final“.
4. *Alle Fahrer* rechnet alle Touren nacheinander und zeigt eine Tagesübersicht.

## Wichtige Hinweise

* **Fahrzeugdaten:** Die Maße im Blatt „Fahrzeuge“ sind Schätzwerte (gelb) – bitte mit den Zulassungsscheinen abgleichen. Sie steuern die LKW-Beschränkungen (Gewicht, Höhe, Breite, Länge) der Route.
* **Live-Zeit ≠ LKW-Zeit:** Google rechnet mit Pkw-Fahrzeiten, aber aktuellem bzw. prognostiziertem Verkehr. Das LKW-Routing liefert die Strecke und die Standardfahrzeit des LKW. Beide Werte werden getrennt angezeigt.
* **Abfahrt in der Zukunft:** Google kennt Verkehrsprognosen nur für die Zukunft; liegt die gewählte Abfahrt in der Vergangenheit, gilt die Live-Zeit ab „jetzt“ (Hinweis im Ergebnis).
* **Optimierung** berücksichtigt nur Fahrstrecke und Standzeit – keine Lieferzeitfenster, Kühlketten oder Ladereihenfolge. Vorschlag immer fachlich prüfen.
* **Adressen:** Kunden, deren Adresse nicht gefunden wird, werden rot aufgelistet und aus der Route ausgelassen; ⚠ markiert nur ungefähr gefundene Adressen. Die Geschäfts-Adressen werden zur Suche/Routenberechnung an OpenRouteService und Google übermittelt.
* **Datenänderung:** Nach Änderungen in Tabelle1 den Knopf ↻ drücken.
* **Rückgängig:** Strg+Z funktioniert für Änderungen durch das Add-in nicht; dafür gibt es „Rückgängig“ im Add-in (letzte Übernahme).

## Dateien

`index.html`, `style.css` – Oberfläche · `app.js` – Excel-Anbindung, Karte · `engine.js` – Berechnungsablauf · `lib.js` – Hilfsfunktionen · `config.js` – Einstellungen (Depot, Fahrzeugtypen) · `demo.js` – Demo-Modus · `manifest.xml` – Aufgabenbereich · `manifest-inhalt.xml` – optional im Blatt · `test/test.js` – automatische Tests (`node test/test.js`)

## Fehlersuche

| Meldung | Ursache |
|---|---|
| „OpenRouteService-Schlüssel fehlt“ | ⚙ Einstellungen ausfüllen |
| „Google lehnt den Schlüssel ab“ | Maps JavaScript API nicht aktiviert, Abrechnung nicht verknüpft oder Schlüssel eingeschränkt |
| „Google Routes: …“ | Routes API nicht aktiviert / nicht im Schlüssel erlaubt |
| „LKW-Route: Route nicht gefunden“ | mit den hinterlegten Maßen/Gewichten gibt es keine Straße dorthin, oder die Adresse liegt falsch – Adresse bzw. Fahrzeugmaße prüfen |
| „Die Tabelle „Tabelle1“ wurde nicht gefunden“ | Tabelle umbenannt – Namen in `config.js` (`tabelle`) anpassen |
