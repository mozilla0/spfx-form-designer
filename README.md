# Form Designer

**English** · [Deutsch](README.de.md)

WYSIWYG form designer for SharePoint Online lists. Forms are designed in the browser with drag and drop and deployed as an **SPFx List Form Customizer** – no Windows desktop tool required.
<img width="1094" height="989" alt="Bildschirmfoto_20261103_194146" src="https://github.com/user-attachments/assets/0bb932e2-c540-409d-a405-ed7e11ab5616" />
<img width="1102" height="1122" alt="Bildschirmfoto_20261103_194232" src="https://github.com/user-attachments/assets/a3120ce6-1c81-45da-85a0-099feb80afe5" />
<img width="1261" height="989" alt="Bildschirmfoto_20261103_194206" src="https://github.com/user-attachments/assets/ce98024b-dba7-49da-a56b-c5bfdcee868c" />

## Components

| Component | Type | Purpose |
|---|---|---|
| `FormDesignerWebPart` | Web part | Designer: layout, properties, header, form script, preview, deployment |
| `FormDesignerRuntimeFormCustomizer` | Form customizer | Renders the New, Edit and Display forms as a full page |
| `FormDesignerPanelCommandSet` | ListView command set | Opens the forms from the list view in a panel or dialog |

All three ship in the same package (`form-designer.sppkg`) so they always share the same definition schema.

## Features

- Tabs and sections (one or two columns), fields at half or full width
- Drag and drop (dnd-kit, keyboard accessible): drag columns into the form, reorder fields, drop onto tabs, drag back to the column list to remove, move sections
- Title per form type with placeholders, e.g. `Edit {Title}`
- Header with style, subtitle, icon, logo and accent colour
- Required, read-only, help text, conditional visibility
- Validation: length, format (e-mail, Swiss phone number, Swiss postcode, custom regular expression), number range, date range, custom error message
- Default values in the New form: fixed value, today, today + n (working) days, next working day, end of month, current user, JavaScript formula
- Form script with a `form` API (see below)
- Preview that runs the form script, optionally with an existing item
- Prefill via URL parameters, with link generator
- Display mode per form type: panel, dialog or full page, with configurable width
- Multilingual: user interface in German and English, translatable form content
- Deployment with backup of the previous configuration and "Reset"

### Supported column types

| Column type | Editing | Value in the form script |
|---|---|---|
| Single/multiple lines of text | Text field | `'Text'` |
| Number, currency | Number field | `'42'` (input text) |
| Choice / multiple choice | Dropdown | `'Open'` / `['A', 'B']` |
| Yes/No | Checkbox | `true` / `false` |
| Date / date and time | Date picker (+ time) | `'2026-11-20'` / `'2026-11-20T14:30'` |
| Person (single/multiple) | People picker | `{ id, title, email }` / array of these |
| Lookup (single/multiple) | Dropdown | `{ id, title }` / array of these |
| Managed metadata (single/multiple) | Term picker with search | `{ termGuid, label }` / array of these |
| Hyperlink or picture | URL and description | `{ url, description }` |
| Calculated | Display only | `'Text'` (formatted by SharePoint) |

Attachments, location, image columns and other types with their own UI (e.g. "Image" as file upload) are displayed but not yet editable.

Managed metadata and hyperlinks are saved via `ValidateUpdateListItem`; depending on versioning settings, saving therefore creates an additional version.

## Multilingual support

**User interface:** the designer, forms, panel, dialog and all messages appear in German or English depending on the language of the SharePoint page (German for all `de-*` cultures, English otherwise). Dates are displayed and entered in the culture's format, e.g. `03.11.2026` (de-CH) or `11/03/2026` (en-US).

**Form content:** in the designer's "Translations" view you can translate titles, subtitle, tabs, sections, field labels, help texts and custom error messages, and set the language in which the texts were entered. If a translation is missing, the entered text is shown. The preview lets you switch languages.

Technically, the German texts double as keys (`t('Speichern')`); the English texts live in `src/common/i18n.en.ts`. A unit test checks that every text has a translation.

## Prefill via URL

Fields can be prefilled with URL parameters: `fd_<InternalName>=value`. The field name is case-insensitive.

```
…/_layouts/15/SPListForm.aspx?PageType=8&List={ListId}&ContentTypeId=…&fd_Title=Heating&fd_Priority=High&fd_DueDate=today+10
…/Lists/Orders?fdOpen=new&fd_Title=Heating        (list view, opens the form in the panel/dialog)
…/Lists/Orders?fdOpen=edit&fdId=8                 (edit item 8)
```

| Column type | Format |
|---|---|
| Text, number, choice | Value, e.g. `Open` |
| Multiple choice | `A;B` |
| Yes/No | `yes`/`no`, `ja`/`nein`, `1`/`0`, `true`/`false` |
| Date | `YYYY-MM-DD`, `today`, `today+7`, `today-1` |
| Date and time | `YYYY-MM-DDTHH:mm` or as date |
| Person (single/multiple) | E-mail address or login name; multiple separated by `;` |
| Lookup (single/multiple) | ID or title; multiple separated by `;` |
| Managed metadata | Term or term ID; multiple separated by `;` |
| Hyperlink or picture | `URL` or `URL|description` |

- Setting under "Title and tabs" › "Prefill via URL": off, New form only (default), or New and Edit.
- Order: default values → URL parameters → form script. Invalid values are shown as a warning in the form.
- All URL parameters are available in the script as `form.urlParams`, e.g. `form.urlParams.source`.
- "Create link" in the designer builds a link with prefilled values (form page or list view with panel).
- `fdOpen` is removed from the address bar after opening so the form does not reopen on reload.
- The designer preview has a field for URL parameters.

### Examples

The examples use a list "Orders" at `https://contoso.sharepoint.com/sites/service` with the columns `Title`, `Customer` (lookup), `Priority` (choice), `DueDate` (date), `Owners` (person, multiple), `Location` (managed metadata) and `Channel` (choice).

**1. Simple prefill link**

```
https://contoso.sharepoint.com/sites/service/Lists/Orders?fdOpen=new&fd_Title=Annual%20heat%20pump%20service&fd_Priority=High&fd_DueDate=today+5
```

Opens the list and the New form in the panel: title "Annual heat pump service", priority "High", due in five days. Spaces and special characters must be URL-encoded (`%20`, …) – "Create link" in the designer does this automatically.

**2. People, lookups and managed metadata**

```
…/Lists/Orders?fdOpen=new&fd_Customer=12&fd_Owners=anna.example@contoso.com;peter.sample@contoso.com&fd_Location=Zurich
```

`fd_Customer=12` selects the customer with ID 12 (or use the title, e.g. `fd_Customer=Contoso%20Ltd`). Separate multiple people with `;`; the location takes a term or term ID.

**3. Button in another list (column formatting)**

Add a column with this formatting to the customer list. Each row gets an "Create order" link that opens the form prefilled with that customer:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
  "elmType": "a",
  "txtContent": "Create order",
  "attributes": {
    "href": "='/sites/service/Lists/Orders?fdOpen=new&fd_Customer=' + [$ID]",
    "target": "_self",
    "class": "ms-fontColor-themePrimary ms-fontWeight-semibold"
  }
}
```

**4. Link in a Power Automate e-mail**

New form with the subject of an incoming e-mail as title:

```
concat('https://contoso.sharepoint.com/sites/service/Lists/Orders?fdOpen=new&source=mail&fd_Title=', encodeUriComponent(triggerOutputs()?['body/subject']))
```

Open an existing item for editing:

```
https://contoso.sharepoint.com/sites/service/Lists/Orders?fdOpen=edit&fdId=@{triggerOutputs()?['body/ID']}
```

**5. Using URL parameters in the form script**

Parameters without `fd_` do not prefill fields but are available in the script, e.g. to track the origin or lock values:

```js
form.onReady(() => {
  // Opened from the e-mail in example 4
  if (form.urlParams.source === 'mail') {
    form.field('Channel').value = 'E-mail';
    form.field('Channel').disabled = true;
  }
  // Do not allow changing a prefilled customer
  if (form.urlParams.fd_Customer) {
    form.field('Customer').disabled = true;
  }
});
```

`form.urlParams` keeps the spelling used in the URL.

**6. Form page instead of the panel**

```
https://contoso.sharepoint.com/sites/service/_layouts/15/SPListForm.aspx?PageType=8&List=%7B<ListId>%7D&ContentTypeId=<ContentTypeId>&RootFolder=%2Fsites%2Fservice%2FLists%2FOrders&fd_Title=Test
```

Useful if the panel extension is not registered on the list or the form should open in its own tab. "Create link" fills in the list and content type IDs (clear the option "Open via the list view").

## Header

Under "Title and tabs" › "Header": style (plain, accent line, coloured band), subtitle with placeholders, Fluent UI icon, logo or image (URL or upload to "Site Assets/FormDesigner/Bilder" (folder name is fixed)), image position and accent colour. On the coloured band the text colour is chosen automatically for contrast.

Microsoft's JSON list form configuration (header/body/footer) only applies to SharePoint's standard form, not to forms rendered by Form Designer.

## Panel, dialog or full page

For each form type (New, Edit, Display) you choose under "Title and tabs" how the form opens from the list view: **panel**, **dialog** (movable and resizable) or **full page**, plus the width of panel and dialog.

How it works:

- On deployment the designer registers `FormDesignerPanelCommandSet` as a custom action on the list. If all form types are set to full page, it is removed again; the same happens on "Reset".
- The extension loads with the list's command bar and intercepts form openings: navigations (Navigation API), clicks on form links and `window.open`. It recognises `SPListForm.aspx`, `listform.aspx` and the list's `NewForm/EditForm/DispForm.aspx`. Duplicate navigation events are handled once, and at most one form is open at a time.
- Instead of navigating to the form page, it opens the form in its own Fluent UI panel or dialog. After saving, the list view reloads.
- If interception does not apply (direct link, list web part without command bar, browser without Navigation API, load error), the form opens as a full page through the form customizer.

The interception is not documented by Microsoft and may need adjustments after SharePoint updates.

In addition, the form customizer sets `isPanelExperienceEnabled: true`. Once Microsoft enables the native list panel for form customizers (currently a bug, sp-dev-docs #10996), SharePoint opens the form in its own panel. For that case the runtime experimentally applies the configured width to the panel container.

### Diagnostics

Append a parameter to the list view URL (applies to the browser session). Output appears in the browser console with the prefix `[Form Designer]`. Tip: enable "Preserve log" in the console settings and filter for `Form Designer`.

| Parameter | Effect |
|---|---|
| `?fdDebug=log` | Log only: every click with automation IDs and every recognised form, without redirecting |
| `?fdDebug=on` | Log and redirect |
| `?fdDebug=off` | Back to normal operation |

In diagnostic mode the command bar shows "Form Designer: Diagnostics". Clicking it prints the loaded configuration.

## Where definitions are stored

In the site's "Site Assets" library, folder `FormDesigner`:

| File | Content |
|---|---|
| `{ListId}_{ContentTypeId}.draft.json` | Draft from the designer |
| `{ListId}_{ContentTypeId}.json` | Deployed definition, read by the runtime and the panel extension |
| `{ListId}_{ContentTypeId}.backup.json` | Content type form configuration before the first deployment |
| `Bilder/…` | Uploaded logos and images for the header |

The content type gets `New/Edit/DisplayFormClientSideComponentId` (runtime ID) and the path to the definition in `…ClientSideComponentProperties` (`{"definitionUrl": "…"}`).

Site members need read access to "Site Assets" (default).

## Form script

```js
form.onReady(() => {
  if (form.type === 'new') form.field('DueDate').value = form.date.addWorkdays(form.date.today(), 5);
});

form.field('Status').onChange(value => {
  form.field('InternalNotes').visible = value === 'Completed';
  form.field('Technician').required = value === 'In progress';
});

form.field('DueDate').validate(value => {
  if (form.field('Priority').value === 'High' && !value) return 'A due date is required for high priority.';
});

form.onBeforeSave(async () => {
  // Returning a string cancels saving.
});
```

| API | Description |
|---|---|
| `form.field(name).value` | Read/write a value (dates: `'YYYY-MM-DD'` or a Date object) |
| `.visible` · `.required` · `.disabled` | Show/hide, required, disable |
| `.onChange(fn)` · `.validate(fn)` | Change handlers and custom validation |
| `form.onReady(fn)` · `form.onBeforeSave(fn)` | Load and save |
| `form.showMessage(text, kind)` | `'info'`, `'warning'`, `'error'`, `'success'` |
| `form.type` · `form.title` | Form type and title |
| `form.urlParams` | All query parameters of the form URL |
| `form.date` | `today`, `addDays`, `addWorkdays`, `endOfMonth`, `format`, `iso` |
| `form.currentUser` · `form.sp.get(path)` / `.post(path, body)` | Current user and REST access to the site |

The script runs with the permissions of the signed-in user. Place the designer only on a page restricted to site owners.

## Development

Requirements: Node.js 22 (≥ 22.14), SPFx 1.23.2 (Heft).

```bash
npm install
npm run start   # local development
npm run build   # tests, production build and sharepoint/solution/form-designer.sppkg
```

For `npm run start`, enter your tenant and list in `config/serve.json`.

## Deployment

1. Upload `form-designer.sppkg` to the app catalog and deploy it.
2. Add the "Form Designer" app to the site.
3. Add the "Form Designer" web part to a page for site owners (full width recommended).
4. Select list and content type, design the form, "Save draft", then "Deploy".
5. If a Power Apps form is active on the list: List settings › Form settings › "Use the default SharePoint form".

## Known limitations

- `addWorkdays` does not take public holidays into account.
- Dates are displayed in the browser's time zone.
- Multi-value person and lookup fields, URL, managed metadata and attachment fields are not yet editable.
- The user interface supports German and English; the diagnostic console log is in German.
- Managed metadata terms are loaded up to six levels deep.
- Form script and formulas use `new Function`. If SharePoint Online blocks `unsafe-eval` through its Content Security Policy, an alternative will be needed.
