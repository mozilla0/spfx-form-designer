# Form Designer

**English** · [Deutsch](README.de.md)

WYSIWYG form designer for SharePoint Online lists. Forms are designed in the browser with drag and drop and deployed as an **SPFx List Form Customizer** – no Windows desktop tool required.

Author: Matthias Hummler

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
- Display mode per form type: panel, dialog or full page, with configurable width
- Deployment with backup of the previous configuration and "Reset"

### Supported column types

Single and multiple lines of text, number and currency, choice, multi-choice, yes/no, date, date and time, person (single), lookup (single). Other types are displayed but not yet editable.

## Header

Under "Titel und Tabs" › "Kopfbereich": style (plain, accent line, coloured band), subtitle with placeholders, Fluent UI icon, logo or image (URL or upload to "Site Assets/FormDesigner/Bilder"), image position and accent colour. On the coloured band the text colour is chosen automatically for contrast.

Microsoft's JSON list form configuration (header/body/footer) only applies to SharePoint's standard form, not to forms rendered by Form Designer.

## Panel, dialog or full page

For each form type (New, Edit, Display) you choose under "Titel und Tabs" how the form opens from the list view: **panel**, **dialog** (movable and resizable) or **full page**, plus the width of panel and dialog.

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

In diagnostic mode the command bar shows "Form Designer: Diagnose". Clicking it prints the loaded configuration.

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
4. Select list and content type, design the form, "Entwurf speichern" (save draft), then "Bereitstellen" (deploy).
5. If a Power Apps form is active on the list: List settings › Form settings › "Use the default SharePoint form".

## Known limitations

- `addWorkdays` does not take public holidays into account.
- Dates are displayed in the browser's time zone.
- Multi-value person and lookup fields, URL, managed metadata and attachment fields are not yet editable.
- The user interface is currently German only.
- Form script and formulas use `new Function`. If SharePoint Online blocks `unsafe-eval` through its Content Security Policy, an alternative will be needed.
