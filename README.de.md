# Form Designer

[English](README.md) · **Deutsch**

WYSIWYG-Formulardesigner für SharePoint-Online-Listen. Formulare werden im Browser per Drag-and-drop gestaltet und als **SPFx List Form Customizer** bereitgestellt – ohne Windows-Desktop-Tool.
<img width="2555" height="660" alt="Bildschirmfoto_20261103_194310" src="https://github.com/user-attachments/assets/573f689a-20a0-43d1-9c06-adb93c852926" />
<img width="1102" height="1122" alt="Bildschirmfoto_20261103_194232" src="https://github.com/user-attachments/assets/7cfd552e-5582-47bf-9e83-ef681c342074" />
<img width="1261" height="989" alt="Bildschirmfoto_20261103_194206" src="https://github.com/user-attachments/assets/bdd9f012-a717-42b2-a1fd-0fac8a9a0b72" />
<img width="1094" height="989" alt="Bildschirmfoto_20261103_194146" src="https://github.com/user-attachments/assets/8625e4de-eaae-44be-880c-e6a2138aab09" />


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
- Vorbelegung über URL-Parameter, mit Link-Generator
- Darstellung je Formulartyp: Panel, Dialog oder ganze Seite, mit einstellbarer Breite
- Mehrsprachig: Oberfläche auf Deutsch und Englisch, Formularinhalte übersetzbar
- Bereitstellung mit Sicherung der bisherigen Konfiguration und «Zurücksetzen»

### Unterstützte Spaltentypen

| Spaltentyp | Bearbeitung | Wert im Formular-Skript |
|---|---|---|
| Einzelne/mehrere Textzeilen | Textfeld | `'Text'` |
| Zahl, Währung | Zahlenfeld | `'42'` (Text der Eingabe) |
| Auswahl / Mehrfachauswahl | Auswahlliste | `'Offen'` / `['A', 'B']` |
| Ja/Nein | Kontrollkästchen | `true` / `false` |
| Datum / Datum und Uhrzeit | Datumsauswahl (+ Uhrzeit) | `'2026-11-20'` / `'2026-11-20T14:30'` |
| Person (einzeln/mehrere) | Personenauswahl | `{ id, title, email }` / Array davon |
| Nachschlagen (einzeln/mehrere) | Auswahlliste | `{ id, title }` / Array davon |
| Verwaltete Metadaten (einzeln/mehrere) | Begriffsauswahl mit Suche | `{ termGuid, label }` / Array davon |
| Link oder Bild | URL und Beschreibung | `{ url, description }` |
| Berechnet | nur Anzeige | `'Text'` (formatiert von SharePoint) |

Anlagen, Ort, Bild- und Taxonomie-Spalten mit eigener Benutzeroberfläche (z. B. «Bild» als Dateiupload) werden angezeigt, aber noch nicht bearbeitet.

Verwaltete Metadaten und Links werden über `ValidateUpdateListItem` gespeichert; beim Speichern entsteht dadurch je nach Versionierung eine zusätzliche Version.

## Mehrsprachigkeit

**Oberfläche:** Designer, Formulare, Panel, Dialog und alle Meldungen erscheinen auf Deutsch oder Englisch – je nach Sprache der SharePoint-Seite (Deutsch für alle `de-*`-Kulturen, sonst Englisch). Datumswerte werden im Format der Kultur angezeigt und eingegeben, z. B. `03.11.2026` (de-CH) oder `11/03/2026` (en-US).

**Formularinhalte:** In der Ansicht «Übersetzungen» des Designers lassen sich Titel, Untertitel, Tabs, Abschnitte, Feldbezeichnungen, Hilfetexte und eigene Fehlermeldungen übersetzen. Dort wird auch festgelegt, in welcher Sprache die Texte erfasst sind. Fehlt eine Übersetzung, erscheint der erfasste Text. In der Vorschau lässt sich die Sprache umschalten.

Technisch sind die deutschen Texte zugleich die Schlüssel (`t('Speichern')`), die englischen stehen in `src/common/i18n.en.ts`. Ein Unit-Test prüft, dass zu jedem Text eine Übersetzung existiert.

## Vorbelegung per URL

Felder lassen sich über Parameter in der URL vorbelegen: `fd_<InternerName>=Wert`. Gross-/Kleinschreibung des Feldnamens spielt keine Rolle.

```
…/_layouts/15/SPListForm.aspx?PageType=8&List={ListId}&ContentTypeId=…&fd_Title=Heizung&fd_Prioritaet=Hoch&fd_FaelligAm=today+10
…/Lists/Auftraege?fdOpen=new&fd_Title=Heizung        (Listenansicht, öffnet das Formular im Panel/Dialog)
…/Lists/Auftraege?fdOpen=edit&fdId=8                  (Element 8 bearbeiten)
```

| Spaltentyp | Format |
|---|---|
| Text, Zahl, Auswahl | Wert, z. B. `Offen` |
| Mehrfachauswahl | `A;B` |
| Ja/Nein | `ja`/`nein`, `1`/`0`, `true`/`false` |
| Datum | `JJJJ-MM-TT`, `today`, `today+7`, `today-1` |
| Datum und Uhrzeit | `JJJJ-MM-TTTHH:mm` oder wie Datum |
| Person (einzeln/mehrere) | E-Mail-Adresse oder Anmeldename; mehrere mit `;` |
| Nachschlagen (einzeln/mehrere) | ID oder Titel; mehrere mit `;` |
| Verwaltete Metadaten | Begriff oder Begriffs-ID; mehrere mit `;` |
| Link oder Bild | `URL` oder `URL|Beschreibung` |

- Einstellung unter «Titel und Tabs» › «Vorbelegung per URL»: aus, nur im Formular «Neu» (Standard) oder in «Neu» und «Bearbeiten».
- Reihenfolge: Standardwerte → URL-Parameter → Formular-Skript. Ungültige Werte erscheinen als Hinweis im Formular.
- Alle Parameter der URL stehen im Skript als `form.urlParams` zur Verfügung, z. B. `form.urlParams.quelle`.
- «Link erzeugen» im Designer baut den Link mit Vorbelegung (Formularseite oder Listenansicht mit Panel).
- `fdOpen` wird nach dem Öffnen aus der Adresszeile entfernt, damit das Formular beim Neuladen nicht erneut aufgeht.
- Die Designer-Vorschau hat ein Feld für URL-Parameter.

### Beispiele

Die Beispiele verwenden eine Liste «Auftraege» unter `https://contoso.sharepoint.com/sites/service` mit den Spalten `Title`, `Kunde` (Nachschlagen), `Prioritaet` (Auswahl), `FaelligAm` (Datum), `Verantwortlich` (Person, mehrere), `Standort` (verwaltete Metadaten) und `Kanal` (Auswahl).

**1. Einfacher Link mit Vorbelegung**

```
https://contoso.sharepoint.com/sites/service/Lists/Auftraege?fdOpen=new&fd_Title=Jahresservice%20W%C3%A4rmepumpe&fd_Prioritaet=Hoch&fd_FaelligAm=today+5
```

Öffnet die Liste und darin das Formular «Neu» im Panel: Titel «Jahresservice Wärmepumpe», Priorität «Hoch», fällig in fünf Tagen. Leerzeichen und Umlaute müssen URL-kodiert sein (`%20`, `%C3%A4`) – «Link erzeugen» im Designer erledigt das automatisch.

**2. Personen, Nachschlagen und Metadaten**

```
…/Lists/Auftraege?fdOpen=new&fd_Kunde=12&fd_Verantwortlich=anna.muster@contoso.ch;peter.beispiel@contoso.ch&fd_Standort=Altdorf
```

`fd_Kunde=12` wählt den Kunden mit der ID 12 (alternativ den Titel, z. B. `fd_Kunde=Muster%20AG`). Mehrere Personen werden mit `;` getrennt, der Standort per Begriff oder Begriffs-ID.

**3. Schaltfläche in einer anderen Liste (Spaltenformatierung)**

In der Kundenliste eine Spalte mit dieser Formatierung anlegen. Jede Zeile erhält einen Link «Auftrag erfassen», der das Formular mit dem jeweiligen Kunden vorbelegt öffnet:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
  "elmType": "a",
  "txtContent": "Auftrag erfassen",
  "attributes": {
    "href": "='/sites/service/Lists/Auftraege?fdOpen=new&fd_Kunde=' + [$ID]",
    "target": "_self",
    "class": "ms-fontColor-themePrimary ms-fontWeight-semibold"
  }
}
```

**4. Link in einer Power-Automate-Mail**

Neues Formular mit dem Betreff einer eingegangenen Mail als Titel:

```
concat('https://contoso.sharepoint.com/sites/service/Lists/Auftraege?fdOpen=new&quelle=mail&fd_Title=', encodeUriComponent(triggerOutputs()?['body/subject']))
```

Bestehendes Element zum Bearbeiten öffnen:

```
https://contoso.sharepoint.com/sites/service/Lists/Auftraege?fdOpen=edit&fdId=@{triggerOutputs()?['body/ID']}
```

**5. URL-Parameter im Formular-Skript auswerten**

Parameter ohne `fd_` belegen keine Felder vor, stehen aber im Skript zur Verfügung. So lassen sich etwa Herkunft oder Vorgaben steuern:

```js
form.onReady(() => {
  // Aufruf aus der Mail aus Beispiel 4
  if (form.urlParams.quelle === 'mail') {
    form.field('Kanal').value = 'E-Mail';
    form.field('Kanal').disabled = true;
  }
  // Vorbelegten Kunden nicht mehr ändern lassen
  if (form.urlParams.fd_Kunde) {
    form.field('Kunde').disabled = true;
  }
});
```

`form.urlParams` enthält die Parameter mit der Schreibweise aus der URL.

**6. Direkt auf die Formularseite statt ins Panel**

```
https://contoso.sharepoint.com/sites/service/_layouts/15/SPListForm.aspx?PageType=8&List=%7B<ListId>%7D&ContentTypeId=<InhaltstypId>&RootFolder=%2Fsites%2Fservice%2FLists%2FAuftraege&fd_Title=Test
```

Sinnvoll, wenn die Panel-Erweiterung nicht auf der Liste registriert ist oder das Formular in einem eigenen Tab aufgehen soll. Listen- und Inhaltstyp-ID setzt «Link erzeugen» ein (Option «Über die Listenansicht öffnen» abwählen).

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
| `form.urlParams` | Alle Query-Parameter der Formular-URL |
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
- Die Oberfläche unterstützt Deutsch und Englisch; das Diagnoseprotokoll in der Konsole ist deutsch.
- Begriffe aus verwalteten Metadaten werden bis zu sechs Ebenen tief geladen.
- Formular-Skript und Formeln verwenden `new Function`. Sollte SharePoint Online `unsafe-eval` per Content Security Policy unterbinden, braucht es dafür eine Alternative.
