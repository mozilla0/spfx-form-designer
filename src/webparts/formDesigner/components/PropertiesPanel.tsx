import * as React from 'react';
import { Checkbox, ChoiceGroup, DefaultButton, Dropdown, Icon, IDropdownOption, IconButton, Stack, Text, TextField } from '@fluentui/react';
import { DateRule, FormMode, IFieldSettings, IFormDefinition, PatternKind, VisibilityOperator } from '../../../common/model';
import { FIELD_KIND_LABELS, ICurrentUser, ISpField, valueToText } from '../../../common/spTypes';
import { computeDefault, defaultOptionsFor } from '../../../common/defaults';
import { PATTERNS } from '../../../common/validation';
import { formatDateWithWeekday } from '../../../common/dateUtils';
import { PANEL_SIZES } from '../../../common/panelSize';
import { DisplayMode, displayModeOf, HeaderStyle, IFormHeader, PanelSize } from '../../../common/model';
import { DEFAULT_ACCENT, isHexColor } from '../../../common/components/FormHeader';
import {
  deleteSection,
  deleteTab,
  moveFieldBy,
  removeField,
  renameTab,
  Selection,
  SelectionChange,
  updateFieldSettings,
  updateSection
} from './designerState';
import styles from './FormDesigner.module.scss';

export interface IPropertiesPanelProps {
  definition: IFormDefinition;
  fieldsByName: Record<string, ISpField>;
  selection: Selection;
  formMode: FormMode;
  currentUser: ICurrentUser;
  onChange: (def: IFormDefinition, selection?: SelectionChange) => void;
  onUploadImage: (file: File) => Promise<string>;
}

const MODE_LABELS: Record<FormMode, string> = { new: 'Neu', edit: 'Bearbeiten', display: 'Anzeigen' };

const VISIBILITY_OPTIONS: IDropdownOption[] = [
  { key: 'always', text: 'Immer' },
  { key: 'eq', text: 'Wenn Feld gleich Wert' },
  { key: 'ne', text: 'Wenn Feld ungleich Wert' },
  { key: 'notEmpty', text: 'Wenn Feld ausgefüllt' },
  { key: 'empty', text: 'Wenn Feld leer' },
  { key: 'modeIs', text: 'Nur in einem Formulartyp' }
];

const PATTERN_OPTIONS: IDropdownOption[] = [
  { key: 'none', text: 'Keines' },
  { key: 'email', text: PATTERNS.email.label },
  { key: 'phoneCH', text: PATTERNS.phoneCH.label },
  { key: 'plzCH', text: PATTERNS.plzCH.label },
  { key: 'custom', text: 'Eigener regulärer Ausdruck' }
];

const DATE_RULE_OPTIONS: IDropdownOption[] = [
  { key: 'none', text: 'Keine Einschränkung' },
  { key: 'notPast', text: 'Nicht in der Vergangenheit' },
  { key: 'future', text: 'Nur in der Zukunft' },
  { key: 'notFuture', text: 'Nicht in der Zukunft' }
];

const FORMULA_HELP: string =
  "Verfügbar: today(), addDays(d, n), addWorkdays(d, n), endOfMonth(d), field('Name'), user.displayName. Wird beim Öffnen des Formulars «Neu» einmal berechnet.";

function numOrUndefined(v: string | undefined): number | undefined {
  if (v === undefined || v.trim() === '') {
    return undefined;
  }
  const n = Number(v);
  return isNaN(n) ? undefined : n;
}

function numText(v: number | undefined): string {
  return v === undefined || v === null ? '' : String(v);
}

const Group: React.FC<{ legend: string }> = ({ legend, children }) => (
  <fieldset className={styles.group}>
    <legend className={styles.groupLegend}>{legend}</legend>
    {children}
  </fieldset>
);

// ---------- Kopfbereich ----------

const ICON_SUGGESTIONS: string[] = ['Repair', 'ShoppingCart', 'Contact', 'People', 'Calendar', 'TaskManager', 'Money', 'Ticket', 'Mail', 'Settings', 'Info', 'Warning'];

const HeaderSettings: React.FC<IPropertiesPanelProps> = (p) => {
  const def = p.definition;
  const h: IFormHeader = def.header || {};
  const [uploading, setUploading] = React.useState(false);
  const [uploadError, setUploadError] = React.useState('');
  const fileRef = React.useRef<HTMLInputElement>(null);
  const set = (patch: Partial<IFormHeader>): void => p.onChange({ ...def, header: { ...h, ...patch } });

  return (
    <Group legend="Kopfbereich">
      <Dropdown
        label="Stil"
        selectedKey={h.style || 'plain'}
        options={[
          { key: 'plain', text: 'Schlicht' },
          { key: 'line', text: 'Mit Akzentlinie' },
          { key: 'band', text: 'Farbiges Band' }
        ]}
        onChange={(_, o) => o && set({ style: o.key as HeaderStyle })}
      />
      <TextField label="Untertitel" value={h.subtitle || ''} placeholder="Optional, z. B. Auftrag {Title}" onChange={(_, v) => set({ subtitle: v || '' })} />

      <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
        <TextField
          label="Icon (Fluent-UI-Name)"
          value={h.icon || ''}
          placeholder="z. B. Repair"
          styles={{ root: { flexGrow: 1 } }}
          onChange={(_, v) => set({ icon: (v || '').trim() })}
        />
        <div style={{ width: 32, height: 32, display: 'flex', alignItems: 'center', justifyContent: 'center', border: '1px solid #e1dfdd', borderRadius: 4 }} aria-hidden="true">
          {h.icon && <Icon iconName={h.icon} style={{ fontSize: 18, color: isHexColor(h.accentColor) ? h.accentColor : DEFAULT_ACCENT }} />}
        </div>
      </Stack>
      <div className={styles.chips}>
        {ICON_SUGGESTIONS.map((name) => (
          <IconButton key={name} iconProps={{ iconName: name }} title={name} ariaLabel={`Icon ${name} verwenden`} checked={h.icon === name} onClick={() => set({ icon: h.icon === name ? '' : name })} />
        ))}
      </div>
      <span className={styles.hint}>Alle Namen findet man im Fluent UI Icons Tool (aka.ms/fluentui-icons, Version 8).</span>

      <TextField label="Logo oder Bild (URL)" value={h.imageUrl || ''} placeholder="/sites/…/SiteAssets/logo.png" onChange={(_, v) => set({ imageUrl: (v || '').trim() })} />
      <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
        <input
          ref={fileRef}
          type="file"
          accept="image/png,image/jpeg,image/svg+xml,image/gif,image/webp"
          style={{ display: 'none' }}
          onChange={(e) => {
            const file = e.target.files && e.target.files[0];
            e.target.value = '';
            if (!file) {
              return;
            }
            setUploading(true);
            setUploadError('');
            p.onUploadImage(file)
              .then((url) => set({ imageUrl: url }))
              .catch((err: Error) => setUploadError(err.message))
              .then(() => setUploading(false))
              .catch(() => undefined);
          }}
        />
        <DefaultButton text={uploading ? 'Wird hochgeladen …' : 'Bild hochladen'} iconProps={{ iconName: 'Upload' }} disabled={uploading} onClick={() => fileRef.current && fileRef.current.click()} />
        {h.imageUrl && <DefaultButton text="Entfernen" onClick={() => set({ imageUrl: '' })} />}
      </Stack>
      {uploadError && <Text styles={{ root: { color: '#a4262c' } }}>{uploadError}</Text>}
      {h.imageUrl && (
        <ChoiceGroup
          label="Position des Bilds"
          selectedKey={h.imagePosition || 'left'}
          options={[
            { key: 'left', text: 'Links' },
            { key: 'right', text: 'Rechts' }
          ]}
          onChange={(_, o) => o && set({ imagePosition: o.key === 'right' ? 'right' : 'left' })}
        />
      )}

      <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
        <TextField
          label="Akzentfarbe"
          value={h.accentColor || ''}
          placeholder={DEFAULT_ACCENT}
          styles={{ root: { flexGrow: 1 } }}
          errorMessage={h.accentColor && !isHexColor(h.accentColor) ? 'Format #RRGGBB' : undefined}
          onChange={(_, v) => set({ accentColor: (v || '').trim() })}
        />
        <input
          type="color"
          aria-label="Akzentfarbe wählen"
          value={isHexColor(h.accentColor) ? (h.accentColor as string) : DEFAULT_ACCENT}
          onChange={(e) => set({ accentColor: e.target.value })}
          style={{ width: 40, height: 32, padding: 0, border: '1px solid #e1dfdd', borderRadius: 4, background: '#fff' }}
        />
      </Stack>
      <span className={styles.hint}>Beim farbigen Band wird die Schrift automatisch schwarz oder weiss gesetzt, je nachdem, was besser lesbar ist.</span>
    </Group>
  );
};

// ---------- Formular: Titel und Tabs ----------

const FormProperties: React.FC<IPropertiesPanelProps> = (p) => {
  const def = p.definition;
  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <TextField
        label={`Titel – Formular «${MODE_LABELS[p.formMode]}»`}
        value={def.titles[p.formMode]}
        onChange={(_, v) => p.onChange({ ...def, titles: { ...def.titles, [p.formMode]: v || '' } })}
        description="Platzhalter mit internem Feldnamen in geschweiften Klammern, z. B. {Title}. Oben den Formulartyp wechseln, um die anderen Titel zu bearbeiten. Per Skript mit form.title überschreibbar."
      />
      <HeaderSettings {...p} />
      <Group legend="Darstellung aus der Listenansicht">
        {(['new', 'edit', 'display'] as FormMode[]).map((m) => (
          <Dropdown
            key={m}
            label={`Formular «${MODE_LABELS[m]}»`}
            selectedKey={displayModeOf(def.layout, m)}
            options={[
              { key: 'panel', text: 'Panel' },
              { key: 'dialog', text: 'Dialog' },
              { key: 'page', text: 'Ganze Seite' }
            ]}
            onChange={(_, o) =>
              o &&
              p.onChange({
                ...def,
                layout: {
                  ...(def.layout || { panelSize: 'default' }),
                  displayModes: { ...((def.layout && def.layout.displayModes) || {}), [m]: o.key as DisplayMode }
                }
              })
            }
          />
        ))}
        <span className={styles.hint}>
          Panel und Dialog greifen beim Öffnen aus der Listenansicht (Befehlsleiste muss sichtbar sein). Direkte Links öffnen das Formular als ganze Seite.
        </span>
      </Group>
      <Group legend="Breite von Panel und Dialog">
        <Dropdown
          ariaLabel="Breite von Panel und Dialog"
          selectedKey={(def.layout && def.layout.panelSize) || 'default'}
          options={PANEL_SIZES.map((s) => ({ key: s.key, text: s.text }))}
          onChange={(_, o) =>
            o &&
            p.onChange({
              ...def,
              layout: { ...(def.layout || { panelSize: 'default' }), panelSize: o.key as PanelSize, panelWidth: (def.layout && def.layout.panelWidth) || 720 }
            })
          }
        />
        {def.layout && def.layout.panelSize === 'custom' && (
          <TextField
            label="Breite in Pixeln"
            type="number"
            min={320}
            value={numText(def.layout.panelWidth)}
            onChange={(_, v) => p.onChange({ ...def, layout: { ...(def.layout || { panelSize: 'custom' }), panelWidth: numOrUndefined(v) } })}
            errorMessage={def.layout.panelWidth !== undefined && def.layout.panelWidth < 320 ? 'Mindestens 320 Pixel.' : undefined}
          />
        )}
        <span className={styles.hint}>
          Gilt für das Panel und den Dialog des Form Designers. «Standard» entspricht 640 px im Panel und 720 px im Dialog. Auf kleinen Bildschirmen bleiben beide höchstens bildschirmbreit.
        </span>
      </Group>
      <Group legend="Tabs">
        {def.tabs.map((t) => {
          const hasSections = def.sections.some((s) => s.tabId === t.id);
          return (
            <Stack key={t.id} horizontal verticalAlign="end" tokens={{ childrenGap: 4 }}>
              <TextField ariaLabel="Tab-Bezeichnung" value={t.label} styles={{ root: { flexGrow: 1 } }} onChange={(_, v) => p.onChange(renameTab(def, t.id, v || ''))} />
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title={hasSections || def.tabs.length === 1 ? 'Nur leere Tabs können gelöscht werden' : 'Tab löschen'}
                ariaLabel="Tab löschen"
                disabled={hasSections || def.tabs.length === 1}
                onClick={() => p.onChange(deleteTab(def, t.id), { kind: 'form' })}
              />
            </Stack>
          );
        })}
      </Group>
    </Stack>
  );
};

// ---------- Abschnitt ----------

const SectionProperties: React.FC<IPropertiesPanelProps & { sectionId: string }> = (p) => {
  const def = p.definition;
  const section = def.sections.filter((s) => s.id === p.sectionId)[0];
  if (!section) {
    return null;
  }
  const sectionsInTab = def.sections.filter((s) => s.tabId === section.tabId).length;
  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <TextField label="Abschnittstitel" value={section.title} onChange={(_, v) => p.onChange(updateSection(def, section.id, { title: v || '' }))} />
      <ChoiceGroup
        label="Spalten"
        selectedKey={String(section.columns)}
        options={[
          { key: '1', text: 'Einspaltig' },
          { key: '2', text: 'Zweispaltig' }
        ]}
        onChange={(_, o) => o && p.onChange(updateSection(def, section.id, { columns: o.key === '1' ? 1 : 2 }))}
      />
      <Dropdown
        label="Tab"
        selectedKey={section.tabId}
        options={def.tabs.map((t) => ({ key: t.id, text: t.label }))}
        onChange={(_, o) => o && p.onChange(updateSection(def, section.id, { tabId: String(o.key) }))}
      />
      <DefaultButton
        text="Abschnitt löschen"
        iconProps={{ iconName: 'Delete' }}
        disabled={section.fields.length > 0 || sectionsInTab < 2}
        title={section.fields.length > 0 ? 'Zuerst alle Felder aus dem Abschnitt entfernen' : undefined}
        onClick={() => p.onChange(deleteSection(def, section.id), 'clear')}
      />
    </Stack>
  );
};

// ---------- Feld ----------

const FieldProperties: React.FC<IPropertiesPanelProps & { name: string }> = (p) => {
  const def = p.definition;
  const field = p.fieldsByName[p.name];
  const s = def.fields[p.name];
  if (!s) {
    return null;
  }
  const set = (patch: Partial<IFieldSettings>): void => p.onChange(updateFieldSettings(def, p.name, patch));
  const kind = field ? field.kind : 'unsupported';
  const isText = kind === 'text' || kind === 'note';
  const isNumber = kind === 'number';
  const isDate = kind === 'date' || kind === 'datetime';
  const otherFields: IDropdownOption[] = Object.keys(p.fieldsByName)
    .filter((n) => n !== p.name)
    .map((n) => ({ key: n, text: `${p.fieldsByName[n].title} (${n})` }));
  const defaultOptions = defaultOptionsFor(kind);
  const preview = field ? computeDefault(field, s, {}, p.currentUser) : { has: false };

  let previewText = '';
  if (preview.has) {
    if (preview.error) {
      previewText = preview.error;
    } else if (isDate && typeof preview.value === 'string' && preview.value) {
      previewText = `Ergebnis heute: ${formatDateWithWeekday(preview.value)}${kind === 'datetime' ? ` ${preview.value.slice(11, 16)}` : ''}`;
    } else {
      previewText = `Ergebnis heute: ${valueToText(preview.value) || '(leer)'}`;
    }
  }

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <TextField label="Bezeichnung" placeholder={field ? field.title : ''} value={s.label || ''} onChange={(_, v) => set({ label: v || '' })} description="Leer lassen, um den Spaltennamen zu verwenden." />
      <Stack horizontal tokens={{ childrenGap: 16 }}>
        <div>
          <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block' } }}>Interner Name</Text>
          <Text styles={{ root: { fontFamily: 'Consolas, monospace' } }}>{p.name}</Text>
        </div>
        <div>
          <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block' } }}>Spaltentyp</Text>
          <Text>{field ? FIELD_KIND_LABELS[kind] : 'Spalte fehlt'}</Text>
        </div>
      </Stack>
      <ChoiceGroup
        label="Breite"
        selectedKey={s.width}
        options={[
          { key: 'half', text: 'Halbe Breite' },
          { key: 'full', text: 'Volle Breite' }
        ]}
        onChange={(_, o) => o && set({ width: o.key === 'full' ? 'full' : 'half' })}
      />
      <Checkbox
        label={field && field.required ? 'Pflichtfeld (in SharePoint als erforderlich definiert)' : 'Pflichtfeld'}
        checked={s.required || (!!field && field.required)}
        disabled={!!field && field.required}
        onChange={(_, c) => set({ required: !!c })}
      />
      <Checkbox label="Schreibgeschützt" checked={s.readOnly} onChange={(_, c) => set({ readOnly: !!c })} />
      <TextField label="Hilfetext" value={s.help || ''} placeholder="Optional" onChange={(_, v) => set({ help: v || '' })} />

      <Group legend="Sichtbarkeit">
        <Dropdown
          ariaLabel="Sichtbar wenn"
          selectedKey={s.visibility.operator}
          options={VISIBILITY_OPTIONS}
          onChange={(_, o) => o && set({ visibility: { ...s.visibility, operator: o.key as VisibilityOperator } })}
        />
        {['eq', 'ne', 'empty', 'notEmpty'].indexOf(s.visibility.operator) >= 0 && (
          <Dropdown label="Feld" selectedKey={s.visibility.field} options={otherFields} onChange={(_, o) => o && set({ visibility: { ...s.visibility, field: String(o.key) } })} />
        )}
        {(s.visibility.operator === 'eq' || s.visibility.operator === 'ne') && (
          <TextField label="Wert" value={s.visibility.value || ''} onChange={(_, v) => set({ visibility: { ...s.visibility, value: v || '' } })} />
        )}
        {s.visibility.operator === 'modeIs' && (
          <Dropdown
            label="Formulartyp"
            selectedKey={s.visibility.mode || 'edit'}
            options={[
              { key: 'new', text: 'Neu' },
              { key: 'edit', text: 'Bearbeiten' },
              { key: 'display', text: 'Anzeigen' }
            ]}
            onChange={(_, o) => o && set({ visibility: { ...s.visibility, mode: o.key as FormMode } })}
          />
        )}
        <span className={styles.hint}>Für komplexere Logik das Formular-Skript verwenden.</span>
      </Group>

      <Group legend="Validierung">
        {isText && (
          <>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <TextField label="Min. Länge" type="number" value={numText(s.validation.minLength)} onChange={(_, v) => set({ validation: { ...s.validation, minLength: numOrUndefined(v) } })} />
              <TextField label="Max. Länge" type="number" value={numText(s.validation.maxLength)} onChange={(_, v) => set({ validation: { ...s.validation, maxLength: numOrUndefined(v) } })} />
            </Stack>
            <Dropdown label="Format" selectedKey={s.validation.pattern} options={PATTERN_OPTIONS} onChange={(_, o) => o && set({ validation: { ...s.validation, pattern: o.key as PatternKind } })} />
            {s.validation.pattern === 'custom' && (
              <TextField
                label="Regulärer Ausdruck"
                value={s.validation.regex || ''}
                placeholder="z. B. ^SA-[0-9]+$"
                styles={{ field: { fontFamily: 'Consolas, monospace' } }}
                onChange={(_, v) => set({ validation: { ...s.validation, regex: v || '' } })}
              />
            )}
          </>
        )}
        {isNumber && (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <TextField label="Minimum" type="number" value={numText(s.validation.min)} onChange={(_, v) => set({ validation: { ...s.validation, min: numOrUndefined(v) } })} />
            <TextField label="Maximum" type="number" value={numText(s.validation.max)} onChange={(_, v) => set({ validation: { ...s.validation, max: numOrUndefined(v) } })} />
          </Stack>
        )}
        {isDate && (
          <Dropdown label="Erlaubter Bereich" selectedKey={s.validation.dateRule} options={DATE_RULE_OPTIONS} onChange={(_, o) => o && set({ validation: { ...s.validation, dateRule: o.key as DateRule } })} />
        )}
        {(isText || isNumber || isDate) && (
          <TextField label="Eigene Fehlermeldung" placeholder="Standardmeldung verwenden" value={s.validation.message || ''} onChange={(_, v) => set({ validation: { ...s.validation, message: v || '' } })} />
        )}
        {!isText && !isNumber && !isDate && <span className={styles.hint}>Für diesen Spaltentyp gibt es Pflichtfeld und Skript-Prüfungen.</span>}
        <span className={styles.hint}>Abhängige Regeln per Skript mit .validate(fn).</span>
      </Group>

      {defaultOptions.length > 1 && (
        <Group legend="Standardwert (Formular «Neu»)">
          <Dropdown
            ariaLabel="Vorbelegung"
            selectedKey={s.defaultValue.mode}
            options={defaultOptions.map((o) => ({ key: o.key, text: o.text }))}
            onChange={(_, o) => {
              if (!o) {
                return;
              }
              const mode = o.key as IFieldSettings['defaultValue']['mode'];
              const next = { ...s.defaultValue, mode };
              if (mode === 'formula' && !next.formula) {
                next.formula = isDate ? 'addDays(today(), 7)' : "field('Title')";
              }
              if (mode === 'fixed' && kind === 'choice' && !next.value && field) {
                next.value = field.choices[0] || '';
              }
              set({ defaultValue: next });
            }}
          />
          {(s.defaultValue.mode === 'todayPlus' || s.defaultValue.mode === 'workdaysPlus') && (
            <TextField label="Anzahl Tage" type="number" value={numText(s.defaultValue.days)} onChange={(_, v) => set({ defaultValue: { ...s.defaultValue, days: numOrUndefined(v) } })} />
          )}
          {s.defaultValue.mode === 'fixed' && kind !== 'choice' && (
            <TextField label="Wert" value={s.defaultValue.value || ''} onChange={(_, v) => set({ defaultValue: { ...s.defaultValue, value: v || '' } })} />
          )}
          {s.defaultValue.mode === 'fixed' && kind === 'choice' && field && (
            <Dropdown
              label="Wert"
              selectedKey={s.defaultValue.value}
              options={field.choices.map((c) => ({ key: c, text: c }))}
              onChange={(_, o) => o && set({ defaultValue: { ...s.defaultValue, value: String(o.key) } })}
            />
          )}
          {s.defaultValue.mode === 'formula' && (
            <>
              <label htmlFor={`formula-${p.name}`} className={styles.groupLegend} style={{ padding: 0 }}>
                Formel (JavaScript-Ausdruck)
              </label>
              <textarea
                id={`formula-${p.name}`}
                className={`${styles.code} ${styles.codeSmall}`}
                spellCheck={false}
                value={s.defaultValue.formula || ''}
                onChange={(e) => set({ defaultValue: { ...s.defaultValue, formula: e.target.value } })}
              />
              <span className={styles.hint}>{FORMULA_HELP}</span>
            </>
          )}
          {previewText && <div className={preview.error ? `${styles.resultOk} ${styles.resultError}` : styles.resultOk}>{previewText}</div>}
        </Group>
      )}

      <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
        <DefaultButton text="Nach oben" iconProps={{ iconName: 'Up' }} onClick={() => p.onChange(moveFieldBy(def, p.name, -1))} />
        <DefaultButton text="Nach unten" iconProps={{ iconName: 'Down' }} onClick={() => p.onChange(moveFieldBy(def, p.name, 1))} />
        <DefaultButton text="Entfernen" iconProps={{ iconName: 'Delete' }} onClick={() => p.onChange(removeField(def, p.name), 'clear')} />
      </Stack>
    </Stack>
  );
};

/** Eigenschaften der aktuellen Auswahl. */
export const PropertiesPanel: React.FC<IPropertiesPanelProps> = (p) => {
  const sel = p.selection;
  return (
    <aside className={`${styles.panel} ${styles.properties}`} aria-label="Eigenschaften">
      <p className={styles.panelTitle}>Eigenschaften</p>
      {!sel && <span className={styles.hint}>Feld, Abschnitt oder «Titel und Tabs» anklicken, um es zu bearbeiten.</span>}
      {sel && sel.kind === 'form' && <FormProperties {...p} />}
      {sel && sel.kind === 'section' && <SectionProperties {...p} sectionId={sel.id} />}
      {sel && sel.kind === 'field' && <FieldProperties {...p} name={sel.name} />}
    </aside>
  );
};
