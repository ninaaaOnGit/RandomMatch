# 🎲 Random Teams

Rein clientseitige Web-App zur zufälligen, aber fairen Gruppeneinteilung bei Teamevents. Mischt Abteilungen und verteilt Führungskräfte (FK) möglichst gleichmäßig über alle Gruppen.

**Live:** https://ninaaaongit.github.io/RandomMatch/

## Funktionen

- Excel-Import (`.xlsx`) mit den Spalten `Name` und `Abteilung`
- Erkennt Führungskräfte automatisch am Text "FK" im Namen (z. B. `Nina Faulstich FK`, `Max Mustermann (FK)`)
- Bildet automatisch sinnvoll große Gruppen (3–6 Personen, bevorzugt 4–5)
- Gruppenanzahl optional manuell festlegbar – dann entfällt die 3–6er-Empfehlung und es wird gleichmäßig auf genau diese Anzahl verteilt
- Mischt Abteilungen und FKs bei jedem "Neu würfeln" zufällig neu
- Zeigt eine Qualitätsbewertung der Verteilung an (Abteilungsmischung, FK-Verteilung, Gruppengrößen)
- "Ergebnis kopieren" kopiert alle Gruppen formatiert in die Zwischenablage
- Funktioniert komplett offline im Browser – kein Server, keine Datenbank, keine Datenübertragung

## Excel-Format

| Name              | Abteilung   |
|-------------------|-------------|
| Janine Doe FK     | Abteilung A |
| Max Mustermann    | Abteilung B |
| Sabine Müller FK  | Abteilung B |
| Peter Schneider   | Abteilung A |

## Nutzung

1. Seite öffnen (siehe Live-Link oben, ideal auf dem iPhone via Safari)
2. Excel-Datei auswählen und auf **"Gruppen erstellen"** klicken
3. Mit **"Neu würfeln"** bei Bedarf eine neue Verteilung erzeugen
4. Mit **"Ergebnis kopieren"** die Gruppen z. B. in eine Nachricht einfügen

Tipp: Auf dem iPhone in Safari über "Teilen → Zum Home-Bildschirm" als App-Icon speichern.

## Technik

Eine einzige statische Seite ohne Build-Schritt:

- `index.html` – Struktur der App
- `style.css` – Design (hell/dunkel, responsive für Desktop & iPhone)
- `app.js` – Excel-Import, Gruppen-Algorithmus, Rendering
- `xlsx.full.min.js` – [SheetJS](https://sheetjs.com/)-Bibliothek zum Lesen von Excel-Dateien, lokal eingebunden (kein CDN, funktioniert offline)

## Lokal testen

Da es sich um eine reine Frontend-App ohne Build-Schritt handelt, reicht ein einfacher statischer Server:

```bash
python3 -m http.server 8000
```

Danach `http://localhost:8000` im Browser öffnen.

## Deployment

Die App liegt auf GitHub Pages und aktualisiert sich automatisch bei jedem Push auf `main`:

```bash
git add -A
git commit -m "Beschreibung der Änderung"
git push
```
