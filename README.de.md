# Form Designer

[English](README.md) · **Deutsch**

WYSIWYG-Formulardesigner für SharePoint-Online-Listen. Formulare werden im Browser per Drag-and-drop gestaltet und als **SPFx List Form Customizer** bereitgestellt – ohne Windows-Desktop-Tool.

Autor: Matthias Hummler

## Bestandteile

| Komponente | Typ | Aufgabe |
|---|---|---|
| `FormDesignerWebPart` | Webpart | Designer: Layout, Eigenschaften, Kopfbereich, Formular-Skript, Vorschau, Bereitstellung |
| `FormDesignerRuntimeFormCustomizer` | Form Customizer | Rendert die Formulare «Neu», «Bearbeiten» und «Anzeigen» als eigene Seite |
| `FormDesignerPanelCommandSet` | ListView Command Set | Öffnet die Formulare aus der Listenansicht im Panel oder Dialog |

Alle drei liegen im selben Paket (`form-designer.sppkg`), damit sie immer dasselbe Definitionsschema verwenden.

## Funktionen

- Tabs und Abschnitte (ein- oder zweispaltig), Felder in halber oder voller Breite
- Drag-and-drop (dnd-kit, auch per Tastatur): Spalten ins Formular ziehen, Felder umsortieren, auf Tabs ablegen, zurück in die Spaltenliste ziehen zum Entfernen, Abschnitte verschieben
- Titel je Formulartyp mit Platzhaltern, z. B. `{Title} bearbeiten`
- Kopfbereich mit Stil, Untertitel, Icon, Logo und Akzentfarbe
- Pflichtfeld, Schreibschutz, Hilfetext, bedingte Sichtbarkeit
- Validierung: Länge, Format (E-Mail, Telefon CH, PLZ CH, eigener regulärer Ausdruck), Zahlenbereich, Datumsbereich, eigene Fehlermeldung
- Standardwerte im Formular «Neu»: fester Wert, heute, heute + n (Werk-)Tage, nächster Werktag, Monatsende, aktueller Benutzer, JavaScript-Formel
- Formular-Skript mit `form`-API (siehe unten)
- Vorschau mit ausgeführtem Skript, optional mit einem bestehenden Element
- Darstellung je Formulartyp: Panel, Dialog oder ganze Seite, mit einstellbarer Breite
- Bereitstellung mit Sicherung der bisherigen Konfiguration und «Zurücksetzen»

### Unterstützte Spaltentypen

Einzelne und mehrere Textzeilen, Zahl und Währung, Auswahl, Mehrfachauswahl, Ja/Nein, Datum, Datum mit Uhrzeit, Person (einzeln), Nachschlagen (einzeln). Andere Typen werden angezeigt, aber noch nicht bearbeitet.

## Kopfbereich

Unter «Titel und Tabs» › «Kopfbereich»: Stil (schlicht, mit Akzentlinie, farbiges Band), Untertitel mit Platzhaltern, Fluent-UI-Icon, Logo oder Bild (URL oder Upload nach «Websiteobjekte/FormDesigner/Bilder»), Position des Bilds und Akzentfarbe. Beim farbigen Band wird die Schriftfarbe automatisch nach Kontrast gewählt.

Die JSON-Formularkonfiguration von Microsoft (Header/Body/Footer) wirkt nur auf das Standardformular von SharePoint und nicht auf Formulare des Form Designers.

## Panel, Dialog oder ganze Seite

Pro Formulartyp (Neu, Bearbeiten, Anzeigen) lässt sich unter «Titel und Tabs» wählen, wie sich das Formular aus der Listenansicht öffnet: **Panel**, **Dialog** (verschiebbar, in der Grösse änderbar) oder **ganze Seite**. Dazu kommt die Breite von Panel und Dialog.

So funktioniert es:

- Beim Bereitstellen registriert der Designer `FormDesignerPanelCommandSet` als benutzerdefinierte Aktion auf der Liste. Stehen alle Formulartypen auf «Ganze Seite», wird sie wieder entfernt, ebenso bei «Zurücksetzen».
- Die Erweiterung lädt mit der Befehlsleiste der Liste und fängt das Öffnen von Formularen ab: Navigationen (Navigation API), Klicks auf Formular-Links und `window.open`. Erkannt werden `SPListForm.aspx`, `listform.aspx` sowie `NewForm/EditForm/DispForm.aspx` der Liste. Doppelt gemeldete Navigationen werden nur einmal behandelt, und es ist immer höchstens ein Formular offen.
- Statt auf die Formularseite zu wechseln, öffnet sie das Formular in einem eigenen Fluent-UI-Panel oder -Dialog. Nach dem Speichern wird die Listenansicht neu geladen.
- Greift das Abfangen nicht (direkter Link, Listen-Webpart ohne Befehlsleiste, Browser ohne Navigation API, Fehler beim Laden), öffnet sich das Formular als ganze Seite über den Form Customizer.

Das Abfangen ist nicht von Microsoft dokumentiert und muss nach SharePoint-Updates eventuell angepasst werden.

Zusätzlich setzt der Form Customizer `isPanelExperienceEnabled: true`. Sobald Microsoft das native Listen-Panel für Form Customizer freischaltet (derzeit Bug, sp-dev-docs #10996), öffnet SharePoint das Formular selbst im Panel. Für diesen Fall setzt die Runtime die eingestellte Breite experimentell direkt am Panel-Container.

### Diagnose

Parameter an die URL der Listenansicht anhängen (gilt für die Browser-Sitzung). Die Ausgabe erscheint in der Browser-Konsole mit dem Präfix `[Form Designer]`. Tipp: In den Konsoleneinstellungen «Preserve log» aktivieren und nach `Form Designer` filtern.

| Parameter | Wirkung |
|---|---|
| `?fdDebug=log` | Nur protokollieren: jeder Klick mit Automation-IDs und jedes erkannte Formular, ohne umzuleiten |
| `?fdDebug=on` | Protokollieren und umleiten |
| `?fdDebug=off` | Zurück zum Normalbetrieb |

Im Diagnosemodus erscheint in der Befehlsleiste «Form Designer: Diagnose». Ein Klick darauf gibt die geladene Konfiguration aus.

## Speicherort der Definitionen

In der Bibliothek «Websiteobjekte» (Site Assets) der Website, Ordner `FormDesigner`:

| Datei | Inhalt |
|---|---|
| `{ListId}_{InhaltstypId}.draft.json` | Entwurf aus dem Designer |
| `{ListId}_{InhaltstypId}.json` | Bereitgestellte Definition, gelesen von Runtime und Panel-Erweiterung |
| `{ListId}_{InhaltstypId}.backup.json` | Formularkonfiguration des Inhaltstyps vor der ersten Bereitstellung |
| `Bilder/…` | Hochgeladene Logos und Bilder für den Kopfbereich |

Der Inhaltstyp erhält `New/Edit/DisplayFormClientSideComponentId` (ID der Runtime) und in `…ClientSideComponentProperties` den Pfad zur Definition (`{"definitionUrl": "…"}`).

Website-Mitglieder brauchen Lesezugriff auf «Websiteobjekte» (Standard).

## Formular-Skript

```js
form.onReady(() => {
  if (form.type === 'new') form.field('FaelligAm').value = form.date.addWorkdays(form.date.today(), 5);
});

form.field('Status').onChange(value => {
  form.field('BemerkungIntern').visible = value === 'Abgeschlossen';
  form.field('Monteur').required = value === 'In Arbeit';
});

form.field('FaelligAm').validate(value => {
  if (form.field('Prioritaet').value === 'Hoch' && !value) return 'Bei hoher Priorität ist ein Fälligkeitsdatum nötig.';
});

form.onBeforeSave(async () => {
  // Ein zurückgegebener Text bricht das Speichern ab.
});
```

| API | Beschreibung |
|---|---|
| `form.field(name).value` | Wert lesen/setzen (Datum: `'JJJJ-MM-TT'` oder Date-Objekt) |
| `.visible` · `.required` · `.disabled` | Feld ein-/ausblenden, Pflicht, sperren |
| `.onChange(fn)` · `.validate(fn)` | Änderungen und eigene Prüfungen |
| `form.onReady(fn)` · `form.onBeforeSave(fn)` | Laden und Speichern |
| `form.showMessage(text, art)` | `'info'`, `'warning'`, `'error'`, `'success'` |
| `form.type` · `form.title` | Formulartyp und Titel |
| `form.date` | `today`, `addDays`, `addWorkdays`, `endOfMonth`, `format`, `iso` |
| `form.currentUser` · `form.sp.get(pfad)` / `.post(pfad, body)` | Benutzer und REST-Zugriff auf die Website |

Das Skript läuft mit den Rechten der angemeldeten Person. Den Designer daher nur auf einer Seite platzieren, die Website-Besitzern vorbehalten ist.

## Entwicklung

Voraussetzung: Node.js 22 (≥ 22.14), SPFx 1.23.2 (Heft).

```bash
npm install
npm run start   # lokale Entwicklung
npm run build   # Tests, Produktions-Build und sharepoint/solution/form-designer.sppkg
```

Für `npm run start` in `config/serve.json` den Tenant und die Liste eintragen.

## Bereitstellung

1. `form-designer.sppkg` in den App-Katalog laden und bereitstellen.
2. App «Form Designer» auf der Website hinzufügen.
3. Webpart «Form Designer» auf einer Seite für Website-Besitzer einfügen (ganze Breite empfohlen).
4. Liste und Inhaltstyp wählen, Formular gestalten, «Entwurf speichern», dann «Bereitstellen».
5. Ist auf der Liste ein Power-Apps-Formular aktiv: Listeneinstellungen › Formulareinstellungen › «Standard-SharePoint-Formular».

## Bekannte Grenzen

- `addWorkdays` berücksichtigt keine Feiertage.
- Datumswerte werden in der Zeitzone des Browsers dargestellt.
- Personen- und Nachschlagefelder mit Mehrfachauswahl, URL-, Taxonomie- und Anlagenfelder sind noch nicht editierbar.
- Die Benutzeroberfläche ist derzeit nur auf Deutsch.
- Formular-Skript und Formeln verwenden `new Function`. Sollte SharePoint Online `unsafe-eval` per Content Security Policy unterbinden, braucht es dafür eine Alternative.
