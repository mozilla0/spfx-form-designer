import { ContentLanguage, FormMode, IFormDefinition } from './model';
import { ISpField } from './spTypes';
import { t } from './i18n';

/**
 * Übersetzung der Formularinhalte. Schlüssel:
 * title.new|edit|display · subtitle · tab.<id> · section.<id> · field.<name>.label|help|message
 */

export interface ITranslatableText {
  key: string;
  group: string;
  description: string;
  base: string;
}

/** Liefert eine Kopie der Definition mit den Texten in der gewünschten Sprache. */
export function localizeDefinition(def: IFormDefinition, lang: ContentLanguage): IFormDefinition {
  if (!def.translations || (def.baseLanguage || 'de') === lang) {
    return def;
  }
  const tr = def.translations[lang];
  if (!tr) {
    return def;
  }
  const pick = (key: string, base: string | undefined): string | undefined => (tr[key] && tr[key].trim() ? tr[key] : base);
  const titles = { ...def.titles };
  (['new', 'edit', 'display'] as FormMode[]).forEach((m) => {
    titles[m] = pick(`title.${m}`, titles[m]) || '';
  });
  const fields: IFormDefinition['fields'] = {};
  Object.keys(def.fields).forEach((name) => {
    const f = def.fields[name];
    fields[name] = {
      ...f,
      label: pick(`field.${name}.label`, f.label),
      help: pick(`field.${name}.help`, f.help),
      validation: { ...f.validation, message: pick(`field.${name}.message`, f.validation.message) }
    };
  });
  return {
    ...def,
    titles,
    header: def.header ? { ...def.header, subtitle: pick('subtitle', def.header.subtitle) } : def.header,
    tabs: def.tabs.map((tab) => ({ ...tab, label: pick(`tab.${tab.id}`, tab.label) || tab.label })),
    sections: def.sections.map((s) => ({ ...s, title: pick(`section.${s.id}`, s.title) || '' })),
    fields
  };
}

/** Alle übersetzbaren Texte der Definition für den Übersetzungs-Editor. */
export function translatableTexts(def: IFormDefinition, fieldsByName: Record<string, ISpField>): ITranslatableText[] {
  const out: ITranslatableText[] = [];
  const modeLabel: Record<FormMode, string> = { new: t('Neu'), edit: t('Bearbeiten'), display: t('Anzeigen') };
  (['new', 'edit', 'display'] as FormMode[]).forEach((m) => {
    out.push({ key: `title.${m}`, group: t('Titel und Kopfbereich'), description: t('Titel – Formular «{0}»', modeLabel[m]), base: def.titles[m] || '' });
  });
  out.push({ key: 'subtitle', group: t('Titel und Kopfbereich'), description: t('Untertitel'), base: (def.header && def.header.subtitle) || '' });
  def.tabs.forEach((tab) => out.push({ key: `tab.${tab.id}`, group: t('Tabs und Abschnitte'), description: t('Tab'), base: tab.label }));
  def.sections.forEach((s) => out.push({ key: `section.${s.id}`, group: t('Tabs und Abschnitte'), description: t('Abschnitt'), base: s.title }));
  def.sections.forEach((s) =>
    s.fields.forEach((name) => {
      const settings = def.fields[name];
      if (!settings) {
        return;
      }
      const column = fieldsByName[name] ? fieldsByName[name].title : name;
      out.push({ key: `field.${name}.label`, group: t('Felder'), description: t('{0} – Bezeichnung', column), base: settings.label || column });
      if (settings.help) {
        out.push({ key: `field.${name}.help`, group: t('Felder'), description: t('{0} – Hilfetext', column), base: settings.help });
      }
      if (settings.validation.message) {
        out.push({ key: `field.${name}.message`, group: t('Felder'), description: t('{0} – Fehlermeldung', column), base: settings.validation.message });
      }
    })
  );
  return out;
}
