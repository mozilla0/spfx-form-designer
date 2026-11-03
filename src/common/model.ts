/**
 * Datenmodell der Formulardefinition.
 * Die Definition wird als JSON-Datei in der Bibliothek «Websiteobjekte» gespeichert
 * und von der Runtime (Form Customizer) gelesen.
 */

export type FormMode = 'new' | 'edit' | 'display';
export type FieldWidth = 'half' | 'full';
export type PatternKind = 'none' | 'email' | 'phoneCH' | 'plzCH' | 'custom';
export type DateRule = 'none' | 'notPast' | 'future' | 'notFuture';
export type DefaultMode =
  | 'none'
  | 'fixed'
  | 'fixedTrue'
  | 'today'
  | 'todayPlus'
  | 'workdaysPlus'
  | 'nextWorkday'
  | 'endOfMonth'
  | 'currentUser'
  | 'formula';
export type VisibilityOperator = 'always' | 'eq' | 'ne' | 'empty' | 'notEmpty' | 'modeIs';

export interface IVisibilityRule {
  operator: VisibilityOperator;
  field?: string;
  value?: string;
  mode?: FormMode;
}

export interface IValidationRule {
  minLength?: number;
  maxLength?: number;
  pattern: PatternKind;
  regex?: string;
  min?: number;
  max?: number;
  dateRule: DateRule;
  message?: string;
}

export interface IDefaultValueRule {
  mode: DefaultMode;
  value?: string;
  days?: number;
  formula?: string;
}

export interface IFieldSettings {
  label?: string;
  required: boolean;
  readOnly: boolean;
  width: FieldWidth;
  help?: string;
  visibility: IVisibilityRule;
  validation: IValidationRule;
  defaultValue: IDefaultValueRule;
}

export interface ITab {
  id: string;
  label: string;
}

export interface ISection {
  id: string;
  tabId: string;
  title: string;
  columns: 1 | 2;
  fields: string[];
}

export type PanelSize = 'default' | 'small' | 'medium' | 'large' | 'xlarge' | 'wide' | 'custom';

/** Wo sich ein Formular öffnet, wenn es aus der Listenansicht aufgerufen wird. */
export type DisplayMode = 'panel' | 'dialog' | 'page';

export interface IFormLayout {
  /** Breite von Panel und Dialog. */
  panelSize: PanelSize;
  /** Darstellung je Formulartyp; fehlt ein Eintrag, gilt 'panel'. */
  displayModes?: Partial<Record<FormMode, DisplayMode>>;
  /** Breite in Pixeln bei panelSize 'custom'. */
  panelWidth?: number;
}

export type HeaderStyle = 'plain' | 'line' | 'band';

/** Kopfbereich des Formulars: Icon, Logo, Untertitel und Akzentfarbe. */
export interface IFormHeader {
  /** Name eines Fluent-UI-Icons, z. B. «Repair» oder «ShoppingCart». */
  icon?: string;
  /** Serverrelative oder absolute URL eines Bilds (Logo). */
  imageUrl?: string;
  imagePosition?: 'left' | 'right';
  /** Untertitel mit Platzhaltern wie im Titel, z. B. «Auftrag {Title}». */
  subtitle?: string;
  /** Akzentfarbe als Hex-Wert; leer = Designfarbe der Website. */
  accentColor?: string;
  style?: HeaderStyle;
}

/** Sprache der Formularinhalte. */
export type ContentLanguage = 'de' | 'en';

export interface IFormDefinition {
  schemaVersion: 1;
  layout?: IFormLayout;
  header?: IFormHeader;
  /** Sprache, in der Titel, Bezeichnungen usw. im Designer erfasst sind. */
  baseLanguage?: ContentLanguage;
  /** Übersetzungen der Formularinhalte je Sprache, Schlüssel siehe translate.ts. */
  translations?: Partial<Record<ContentLanguage, Record<string, string>>>;
  titles: Record<FormMode, string>;
  tabs: ITab[];
  sections: ISection[];
  fields: Record<string, IFieldSettings>;
  script: string;
}

/** Eigenschaften, die im Inhaltstyp (…FormClientSideComponentProperties) hinterlegt werden. */
export interface IRuntimeProperties {
  definitionUrl?: string;
}

/** Gesicherte Formularkonfiguration eines Inhaltstyps vor der Bereitstellung. */
export interface IFormBackup {
  savedAt: string;
  NewFormClientSideComponentId: string | null;
  NewFormClientSideComponentProperties: string | null;
  EditFormClientSideComponentId: string | null;
  EditFormClientSideComponentProperties: string | null;
  DisplayFormClientSideComponentId: string | null;
  DisplayFormClientSideComponentProperties: string | null;
}

export function displayModeOf(layout: IFormLayout | undefined, mode: FormMode): DisplayMode {
  const m = layout && layout.displayModes ? layout.displayModes[mode] : undefined;
  return m || 'panel';
}

export function newId(prefix: string): string {
  return `${prefix}${Math.random().toString(36).slice(2, 9)}`;
}

export function createFieldSettings(): IFieldSettings {
  return {
    required: false,
    readOnly: false,
    width: 'half',
    help: '',
    visibility: { operator: 'always' },
    validation: { pattern: 'none', dateRule: 'none' },
    defaultValue: { mode: 'none', days: 7 }
  };
}

export function createDefinition(fieldNames: string[]): IFormDefinition {
  const tabId = newId('tab');
  const fields: Record<string, IFieldSettings> = {};
  fieldNames.forEach((name) => {
    fields[name] = createFieldSettings();
  });
  if (fields.Title) {
    fields.Title.width = 'full';
  }
  return {
    schemaVersion: 1,
    titles: { new: 'Neues Element', edit: '{Title} bearbeiten', display: '{Title}' },
    tabs: [{ id: tabId, label: 'Allgemein' }],
    sections: [{ id: newId('sec'), tabId, title: 'Allgemein', columns: 2, fields: fieldNames.slice() }],
    fields,
    script: '',
    layout: { panelSize: 'default' }
  };
}

/** Ergänzt fehlende Einstellungen, z. B. nach dem Hinzufügen neuer Spalten. */
export function normalizeDefinition(def: IFormDefinition): IFormDefinition {
  const fields: Record<string, IFieldSettings> = {};
  Object.keys(def.fields || {}).forEach((name) => {
    const base = createFieldSettings();
    const current = def.fields[name];
    fields[name] = {
      ...base,
      ...current,
      visibility: { ...base.visibility, ...(current.visibility || {}) },
      validation: { ...base.validation, ...(current.validation || {}) },
      defaultValue: { ...base.defaultValue, ...(current.defaultValue || {}) }
    };
  });
  def.sections.forEach((s) =>
    s.fields.forEach((name) => {
      if (!fields[name]) {
        fields[name] = createFieldSettings();
      }
    })
  );
  return {
    schemaVersion: 1,
    titles: { ...{ new: '', edit: '', display: '' }, ...(def.titles || {}) },
    tabs: def.tabs && def.tabs.length ? def.tabs : [{ id: newId('tab'), label: 'Allgemein' }],
    sections: def.sections || [],
    fields,
    script: def.script || '',
    layout: { panelSize: 'default', ...(def.layout || {}) },
    header: { imagePosition: 'left', style: 'plain', ...(def.header || {}) },
    baseLanguage: def.baseLanguage || 'de',
    translations: def.translations || {}
  };
}
