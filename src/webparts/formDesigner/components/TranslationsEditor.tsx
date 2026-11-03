import * as React from 'react';
import { Dropdown, Stack, Text, TextField } from '@fluentui/react';
import { ContentLanguage, IFormDefinition } from '../../../common/model';
import { ISpField } from '../../../common/spTypes';
import { translatableTexts } from '../../../common/translate';
import { t } from '../../../common/i18n';
import styles from './FormDesigner.module.scss';

export interface ITranslationsEditorProps {
  definition: IFormDefinition;
  fieldsByName: Record<string, ISpField>;
  onChange: (def: IFormDefinition) => void;
}

function languageName(lang: ContentLanguage): string {
  return lang === 'de' ? t('Deutsch') : t('Englisch');
}

/** Übersetzt Titel, Tabs, Abschnitte, Feldbezeichnungen, Hilfetexte und Fehlermeldungen. */
export const TranslationsEditor: React.FC<ITranslationsEditorProps> = ({ definition, fieldsByName, onChange }) => {
  const base: ContentLanguage = definition.baseLanguage || 'de';
  const target: ContentLanguage = base === 'de' ? 'en' : 'de';
  const texts = translatableTexts(definition, fieldsByName).filter((x) => x.base.trim() !== '');
  const current = (definition.translations && definition.translations[target]) || {};
  const done = texts.filter((x) => current[x.key] && current[x.key].trim()).length;

  const setText = (key: string, value: string): void => {
    const next = { ...current, [key]: value };
    if (!value) {
      delete next[key];
    }
    onChange({ ...definition, translations: { ...(definition.translations || {}), [target]: next } });
  };

  let lastGroup = '';
  return (
    <div className={styles.panel} style={{ flex: '1 1 100%', minWidth: 0 }}>
      <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 16 }} styles={{ root: { marginBottom: 12 } }}>
        <Dropdown
          label={t('Sprache der erfassten Texte')}
          selectedKey={base}
          options={[
            { key: 'de', text: t('Deutsch') },
            { key: 'en', text: t('Englisch') }
          ]}
          styles={{ root: { width: 220 } }}
          onChange={(_, o) => o && onChange({ ...definition, baseLanguage: o.key as ContentLanguage })}
        />
        <Text styles={{ root: { paddingBottom: 6 } }}>
          {t('Übersetzung nach {0}: {1} von {2} Texten', languageName(target), done, texts.length)}
        </Text>
      </Stack>
      <p className={styles.hint} style={{ margin: '0 0 12px' }}>
        {t('Formulare erscheinen in der Sprache der SharePoint-Seite. Fehlt eine Übersetzung, wird der erfasste Text angezeigt. Spaltennamen ohne eigene Bezeichnung übersetzt SharePoint selbst, sofern die Spalte mehrsprachig gepflegt ist.')}
      </p>
      <div role="table" aria-label={t('Übersetzungen')} style={{ display: 'grid', gridTemplateColumns: 'minmax(160px, 1fr) minmax(160px, 1.2fr) minmax(200px, 1.5fr)', gap: '6px 12px', alignItems: 'center' }}>
        <Text role="columnheader" styles={{ root: { fontWeight: 600 } }}>{t('Text')}</Text>
        <Text role="columnheader" styles={{ root: { fontWeight: 600 } }}>{languageName(base)}</Text>
        <Text role="columnheader" styles={{ root: { fontWeight: 600 } }}>{languageName(target)}</Text>
        {texts.map((x) => {
          const header =
            x.group !== lastGroup ? (
              <Text key={`g-${x.key}`} role="rowheader" variant="mediumPlus" styles={{ root: { gridColumn: '1 / -1', fontWeight: 600, marginTop: 10 } }}>
                {x.group}
              </Text>
            ) : null;
          lastGroup = x.group;
          return (
            <React.Fragment key={x.key}>
              {header}
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{x.description}</Text>
              <Text styles={{ root: { overflowWrap: 'anywhere' } }}>{x.base}</Text>
              <TextField ariaLabel={`${x.description} – ${languageName(target)}`} value={current[x.key] || ''} placeholder={x.base} onChange={(_, v) => setText(x.key, v || '')} />
            </React.Fragment>
          );
        })}
      </div>
    </div>
  );
};
