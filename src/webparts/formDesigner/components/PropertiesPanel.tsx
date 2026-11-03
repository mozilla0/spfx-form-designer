import * as React from 'react';
import { Checkbox, ChoiceGroup, DefaultButton, Dropdown, Icon, IDropdownOption, IconButton, Stack, Text, TextField } from '@fluentui/react';
import { DateRule, FormMode, IFieldSettings, IFormDefinition, PatternKind, VisibilityOperator } from '../../../common/model';
import { fieldKindLabel, ICurrentUser, ISpField, valueToText } from '../../../common/spTypes';
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
import { t } from '../../../common/i18n';

export interface IPropertiesPanelProps {
  definition: IFormDefinition;
  fieldsByName: Record<string, ISpField>;
  selection: Selection;
  formMode: FormMode;
  currentUser: ICurrentUser;
  onChange: (def: IFormDefinition, selection?: SelectionChange) => void;
  onUploadImage: (file: File) => Promise<string>;
}

function modeLabel(m: FormMode): string {
  return m === 'new' ? t('Neu') : m === 'edit' ? t('Bearbeiten') : t('Anzeigen');
}

const VISIBILITY_OPTIONS = (): IDropdownOption[] => [
  { key: 'always', text: t('Immer') },
  { key: 'eq', text: t('Wenn Feld gleich Wert') },
  { key: 'ne', text: t('Wenn Feld ungleich Wert') },
  { key: 'notEmpty', text: t('Wenn Feld ausgefüllt') },
  { key: 'empty', text: t('Wenn Feld leer') },
  { key: 'modeIs', text: t('Nur in einem Formulartyp') }
];

const PATTERN_OPTIONS = (): IDropdownOption[] => [
  { key: 'none', text: t('Keines') },
  { key: 'email', text: PATTERNS.email.label() },
  { key: 'phoneCH', text: PATTERNS.phoneCH.label() },
  { key: 'plzCH', text: PATTERNS.plzCH.label() },
  { key: 'custom', text: t('Eigener regulärer Ausdruck') }
];

const DATE_RULE_OPTIONS = (): IDropdownOption[] => [
  { key: 'none', text: t('Keine Einschränkung') },
  { key: 'notPast', text: t('Nicht in der Vergangenheit') },
  { key: 'future', text: t('Nur in der Zukunft') },
  { key: 'notFuture', text: t('Nicht in der Zukunft') }
];

const FORMULA_HELP = (): string =>
  t("Verfügbar: today(), addDays(d, n), addWorkdays(d, n), endOfMonth(d), field('Name'), user.displayName. Wird beim Öffnen des Formulars «Neu» einmal berechnet.");

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
    <Group legend={t('Kopfbereich')}>
      <Dropdown
        label={t('Stil')}
        selectedKey={h.style || 'plain'}
        options={[
          { key: 'plain', text: t('Schlicht') },
          { key: 'line', text: t('Mit Akzentlinie') },
          { key: 'band', text: t('Farbiges Band') }
        ]}
        onChange={(_, o) => o && set({ style: o.key as HeaderStyle })}
      />
      <TextField label={t('Untertitel')} value={h.subtitle || ''} placeholder="Optional, z. B. Auftrag {Title}" onChange={(_, v) => set({ subtitle: v || '' })} />

      <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
        <TextField
          label={t('Icon (Fluent-UI-Name)')}
          value={h.icon || ''}
          placeholder={t('z. B. Repair')}
          styles={{ root: { flexGrow: 1 } }}
          onChange={(_, v) => set({ icon: (v || '').trim() })}
        />
        <div style={{ width: 32, height: 32, display: 'flex', alignItems: 'center', justifyContent: 'center', border: '1px solid #e1dfdd', borderRadius: 4 }} aria-hidden="true">
          {h.icon && <Icon iconName={h.icon} style={{ fontSize: 18, color: isHexColor(h.accentColor) ? h.accentColor : DEFAULT_ACCENT }} />}
        </div>
      </Stack>
      <div className={styles.chips}>
        {ICON_SUGGESTIONS.map((name) => (
          <IconButton key={name} iconProps={{ iconName: name }} title={name} ariaLabel={t('Icon {0} verwenden', name)} checked={h.icon === name} onClick={() => set({ icon: h.icon === name ? '' : name })} />
        ))}
      </div>
      <span className={styles.hint}>{t('Alle Namen findet man im Fluent UI Icons Tool (aka.ms/fluentui-icons, Version 8).')}</span>

      <TextField label={t('Logo oder Bild (URL)')} value={h.imageUrl || ''} placeholder={t('/sites/…/SiteAssets/logo.png')} onChange={(_, v) => set({ imageUrl: (v || '').trim() })} />
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
        <DefaultButton text={uploading ? t('Wird hochgeladen …') : t('Bild hochladen')} iconProps={{ iconName: 'Upload' }} disabled={uploading} onClick={() => fileRef.current && fileRef.current.click()} />
        {h.imageUrl && <DefaultButton text={t('Entfernen')} onClick={() => set({ imageUrl: '' })} />}
      </Stack>
      {uploadError && <Text styles={{ root: { color: '#a4262c' } }}>{uploadError}</Text>}
      {h.imageUrl && (
        <ChoiceGroup
          label={t('Position des Bilds')}
          selectedKey={h.imagePosition || 'left'}
          options={[
            { key: 'left', text: t('Links') },
            { key: 'right', text: t('Rechts') }
          ]}
          onChange={(_, o) => o && set({ imagePosition: o.key === 'right' ? 'right' : 'left' })}
        />
      )}

      <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
        <TextField
          label={t('Akzentfarbe')}
          value={h.accentColor || ''}
          placeholder={DEFAULT_ACCENT}
          styles={{ root: { flexGrow: 1 } }}
          errorMessage={h.accentColor && !isHexColor(h.accentColor) ? t('Format #RRGGBB') : undefined}
          onChange={(_, v) => set({ accentColor: (v || '').trim() })}
        />
        <input
          type="color"
          aria-label={t('Akzentfarbe wählen')}
          value={isHexColor(h.accentColor) ? (h.accentColor as string) : DEFAULT_ACCENT}
          onChange={(e) => set({ accentColor: e.target.value })}
          style={{ width: 40, height: 32, padding: 0, border: '1px solid #e1dfdd', borderRadius: 4, background: '#fff' }}
        />
      </Stack>
      <span className={styles.hint}>{t('Beim farbigen Band wird die Schrift automatisch schwarz oder weiss gesetzt, je nachdem, was besser lesbar ist.')}</span>
    </Group>
  );
};

// ---------- Formular: Titel und Tabs ----------

const FormProperties: React.FC<IPropertiesPanelProps> = (p) => {
  const def = p.definition;
  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <TextField
        label={t('Titel – Formular «{0}»', modeLabel(p.formMode))}
        value={def.titles[p.formMode]}
        onChange={(_, v) => p.onChange({ ...def, titles: { ...def.titles, [p.formMode]: v || '' } })}
        description={t('Platzhalter mit internem Feldnamen in geschweiften Klammern, z. B. {Title}. Oben den Formulartyp wechseln, um die anderen Titel zu bearbeiten. Per Skript mit form.title überschreibbar.')}
      />
      <HeaderSettings {...p} />
      <Group legend={t('Darstellung aus der Listenansicht')}>
        {(['new', 'edit', 'display'] as FormMode[]).map((m) => (
          <Dropdown
            key={m}
            label={t('Formular «{0}»', modeLabel(m))}
            selectedKey={displayModeOf(def.layout, m)}
            options={[
              { key: 'panel', text: t('Panel') },
              { key: 'dialog', text: t('Dialog') },
              { key: 'page', text: t('Ganze Seite') }
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
          {t('Panel und Dialog greifen beim Öffnen aus der Listenansicht (Befehlsleiste muss sichtbar sein). Direkte Links öffnen das Formular als ganze Seite.')}
        </span>
      </Group>
      <Group legend={t('Breite von Panel und Dialog')}>
        <Dropdown
          ariaLabel={t('Breite von Panel und Dialog')}
          selectedKey={(def.layout && def.layout.panelSize) || 'default'}
          options={PANEL_SIZES.map((s) => ({ key: s.key, text: s.text() }))}
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
            label={t('Breite in Pixeln')}
            type="number"
            min={320}
            value={numText(def.layout.panelWidth)}
            onChange={(_, v) => p.onChange({ ...def, layout: { ...(def.layout || { panelSize: 'custom' }), panelWidth: numOrUndefined(v) } })}
            errorMessage={def.layout.panelWidth !== undefined && def.layout.panelWidth < 320 ? t('Mindestens 320 Pixel.') : undefined}
          />
        )}
        <span className={styles.hint}>
          {t('Gilt für das Panel und den Dialog des Form Designers. «Standard» entspricht 640 px im Panel und 720 px im Dialog. Auf kleinen Bildschirmen bleiben beide höchstens bildschirmbreit.')}
        </span>
      </Group>
      <Group legend={t('Tabs')}>
        {def.tabs.map((tab) => {
          const hasSections = def.sections.some((s) => s.tabId === tab.id);
          return (
            <Stack key={tab.id} horizontal verticalAlign="end" tokens={{ childrenGap: 4 }}>
              <TextField ariaLabel={t('Tab-Bezeichnung')} value={tab.label} styles={{ root: { flexGrow: 1 } }} onChange={(_, v) => p.onChange(renameTab(def, tab.id, v || ''))} />
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title={hasSections || def.tabs.length === 1 ? t('Nur leere Tabs können gelöscht werden') : t('Tab löschen')}
                ariaLabel={t('Tab löschen')}
                disabled={hasSections || def.tabs.length === 1}
                onClick={() => p.onChange(deleteTab(def, tab.id), { kind: 'form' })}
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
      <TextField label={t('Abschnittstitel')} value={section.title} onChange={(_, v) => p.onChange(updateSection(def, section.id, { title: v || '' }))} />
      <ChoiceGroup
        label={t('Spalten')}
        selectedKey={String(section.columns)}
        options={[
          { key: '1', text: t('Einspaltig') },
          { key: '2', text: t('Zweispaltig') }
        ]}
        onChange={(_, o) => o && p.onChange(updateSection(def, section.id, { columns: o.key === '1' ? 1 : 2 }))}
      />
      <Dropdown
        label={t('Tab')}
        selectedKey={section.tabId}
        options={def.tabs.map((tab) => ({ key: tab.id, text: tab.label }))}
        onChange={(_, o) => o && p.onChange(updateSection(def, section.id, { tabId: String(o.key) }))}
      />
      <DefaultButton
        text={t('Abschnitt löschen')}
        iconProps={{ iconName: 'Delete' }}
        disabled={section.fields.length > 0 || sectionsInTab < 2}
        title={section.fields.length > 0 ? t('Zuerst alle Felder aus dem Abschnitt entfernen') : undefined}
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
      previewText = t('Ergebnis heute: {0}', `${formatDateWithWeekday(preview.value)}${kind === 'datetime' ? ` ${preview.value.slice(11, 16)}` : ''}`);
    } else {
      previewText = t('Ergebnis heute: {0}', valueToText(preview.value) || t('(leer)'));
    }
  }

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <TextField label={t('Bezeichnung')} placeholder={field ? field.title : ''} value={s.label || ''} onChange={(_, v) => set({ label: v || '' })} description={t('Leer lassen, um den Spaltennamen zu verwenden.')} />
      <Stack horizontal tokens={{ childrenGap: 16 }}>
        <div>
          <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block' } }}>{t('Interner Name')}</Text>
          <Text styles={{ root: { fontFamily: 'Consolas, monospace' } }}>{p.name}</Text>
        </div>
        <div>
          <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block' } }}>{t('Spaltentyp')}</Text>
          <Text>{field ? fieldKindLabel(kind) : t('Spalte fehlt')}</Text>
        </div>
      </Stack>
      <ChoiceGroup
        label={t('Breite')}
        selectedKey={s.width}
        options={[
          { key: 'half', text: t('Halbe Breite') },
          { key: 'full', text: t('Volle Breite') }
        ]}
        onChange={(_, o) => o && set({ width: o.key === 'full' ? 'full' : 'half' })}
      />
      <Checkbox
        label={field && field.required ? t('Pflichtfeld (in SharePoint als erforderlich definiert)') : t('Pflichtfeld')}
        checked={s.required || (!!field && field.required)}
        disabled={!!field && field.required}
        onChange={(_, c) => set({ required: !!c })}
      />
      <Checkbox label={t('Schreibgeschützt')} checked={s.readOnly} onChange={(_, c) => set({ readOnly: !!c })} />
      <TextField label={t('Hilfetext')} value={s.help || ''} placeholder={t('Optional')} onChange={(_, v) => set({ help: v || '' })} />

      <Group legend={t('Sichtbarkeit')}>
        <Dropdown
          ariaLabel={t('Sichtbar wenn')}
          selectedKey={s.visibility.operator}
          options={VISIBILITY_OPTIONS()}
          onChange={(_, o) => o && set({ visibility: { ...s.visibility, operator: o.key as VisibilityOperator } })}
        />
        {['eq', 'ne', 'empty', 'notEmpty'].indexOf(s.visibility.operator) >= 0 && (
          <Dropdown label={t('Feld')} selectedKey={s.visibility.field} options={otherFields} onChange={(_, o) => o && set({ visibility: { ...s.visibility, field: String(o.key) } })} />
        )}
        {(s.visibility.operator === 'eq' || s.visibility.operator === 'ne') && (
          <TextField label={t('Wert')} value={s.visibility.value || ''} onChange={(_, v) => set({ visibility: { ...s.visibility, value: v || '' } })} />
        )}
        {s.visibility.operator === 'modeIs' && (
          <Dropdown
            label={t('Formulartyp')}
            selectedKey={s.visibility.mode || 'edit'}
            options={[
              { key: 'new', text: t('Neu') },
              { key: 'edit', text: t('Bearbeiten') },
              { key: 'display', text: t('Anzeigen') }
            ]}
            onChange={(_, o) => o && set({ visibility: { ...s.visibility, mode: o.key as FormMode } })}
          />
        )}
        <span className={styles.hint}>{t('Für komplexere Logik das Formular-Skript verwenden.')}</span>
      </Group>

      <Group legend={t('Validierung')}>
        {isText && (
          <>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <TextField label={t('Min. Länge')} type="number" value={numText(s.validation.minLength)} onChange={(_, v) => set({ validation: { ...s.validation, minLength: numOrUndefined(v) } })} />
              <TextField label={t('Max. Länge')} type="number" value={numText(s.validation.maxLength)} onChange={(_, v) => set({ validation: { ...s.validation, maxLength: numOrUndefined(v) } })} />
            </Stack>
            <Dropdown label={t('Format')} selectedKey={s.validation.pattern} options={PATTERN_OPTIONS()} onChange={(_, o) => o && set({ validation: { ...s.validation, pattern: o.key as PatternKind } })} />
            {s.validation.pattern === 'custom' && (
              <TextField
                label={t('Regulärer Ausdruck')}
                value={s.validation.regex || ''}
                placeholder={t('z. B. ^SA-[0-9]+$')}
                styles={{ field: { fontFamily: 'Consolas, monospace' } }}
                onChange={(_, v) => set({ validation: { ...s.validation, regex: v || '' } })}
              />
            )}
          </>
        )}
        {isNumber && (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <TextField label={t('Minimum')} type="number" value={numText(s.validation.min)} onChange={(_, v) => set({ validation: { ...s.validation, min: numOrUndefined(v) } })} />
            <TextField label={t('Maximum')} type="number" value={numText(s.validation.max)} onChange={(_, v) => set({ validation: { ...s.validation, max: numOrUndefined(v) } })} />
          </Stack>
        )}
        {isDate && (
          <Dropdown label={t('Erlaubter Bereich')} selectedKey={s.validation.dateRule} options={DATE_RULE_OPTIONS()} onChange={(_, o) => o && set({ validation: { ...s.validation, dateRule: o.key as DateRule } })} />
        )}
        {(isText || isNumber || isDate) && (
          <TextField label={t('Eigene Fehlermeldung')} placeholder={t('Standardmeldung verwenden')} value={s.validation.message || ''} onChange={(_, v) => set({ validation: { ...s.validation, message: v || '' } })} />
        )}
        {!isText && !isNumber && !isDate && <span className={styles.hint}>{t('Für diesen Spaltentyp gibt es Pflichtfeld und Skript-Prüfungen.')}</span>}
        <span className={styles.hint}>{t('Abhängige Regeln per Skript mit .validate(fn).')}</span>
      </Group>

      {defaultOptions.length > 1 && (
        <Group legend={t('Standardwert (Formular «Neu»)')}>
          <Dropdown
            ariaLabel={t('Vorbelegung')}
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
            <TextField label={t('Anzahl Tage')} type="number" value={numText(s.defaultValue.days)} onChange={(_, v) => set({ defaultValue: { ...s.defaultValue, days: numOrUndefined(v) } })} />
          )}
          {s.defaultValue.mode === 'fixed' && kind !== 'choice' && (
            <TextField label={t('Wert')} value={s.defaultValue.value || ''} onChange={(_, v) => set({ defaultValue: { ...s.defaultValue, value: v || '' } })} />
          )}
          {s.defaultValue.mode === 'fixed' && kind === 'choice' && field && (
            <Dropdown
              label={t('Wert')}
              selectedKey={s.defaultValue.value}
              options={field.choices.map((c) => ({ key: c, text: c }))}
              onChange={(_, o) => o && set({ defaultValue: { ...s.defaultValue, value: String(o.key) } })}
            />
          )}
          {s.defaultValue.mode === 'formula' && (
            <>
              <label htmlFor={`formula-${p.name}`} className={styles.groupLegend} style={{ padding: 0 }}>
                {t('Formel (JavaScript-Ausdruck)')}
              </label>
              <textarea
                id={`formula-${p.name}`}
                className={`${styles.code} ${styles.codeSmall}`}
                spellCheck={false}
                value={s.defaultValue.formula || ''}
                onChange={(e) => set({ defaultValue: { ...s.defaultValue, formula: e.target.value } })}
              />
              <span className={styles.hint}>{FORMULA_HELP()}</span>
            </>
          )}
          {previewText && <div className={preview.error ? `${styles.resultOk} ${styles.resultError}` : styles.resultOk}>{previewText}</div>}
        </Group>
      )}

      <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
        <DefaultButton text={t('Nach oben')} iconProps={{ iconName: 'Up' }} onClick={() => p.onChange(moveFieldBy(def, p.name, -1))} />
        <DefaultButton text={t('Nach unten')} iconProps={{ iconName: 'Down' }} onClick={() => p.onChange(moveFieldBy(def, p.name, 1))} />
        <DefaultButton text={t('Entfernen')} iconProps={{ iconName: 'Delete' }} onClick={() => p.onChange(removeField(def, p.name), 'clear')} />
      </Stack>
    </Stack>
  );
};

/** Eigenschaften der aktuellen Auswahl. */
export const PropertiesPanel: React.FC<IPropertiesPanelProps> = (p) => {
  const sel = p.selection;
  return (
    <aside className={`${styles.panel} ${styles.properties}`} aria-label={t('Eigenschaften')}>
      <p className={styles.panelTitle}>{t('Eigenschaften')}</p>
      {!sel && <span className={styles.hint}>{t('Feld, Abschnitt oder «Titel und Tabs» anklicken, um es zu bearbeiten.')}</span>}
      {sel && sel.kind === 'form' && <FormProperties {...p} />}
      {sel && sel.kind === 'section' && <SectionProperties {...p} sectionId={sel.id} />}
      {sel && sel.kind === 'field' && <FieldProperties {...p} name={sel.name} />}
    </aside>
  );
};
