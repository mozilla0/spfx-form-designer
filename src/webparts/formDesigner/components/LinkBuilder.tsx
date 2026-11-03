import * as React from 'react';
import { Checkbox, DefaultButton, Dialog, DialogFooter, DialogType, MessageBar, MessageBarType, Stack, Text, TextField } from '@fluentui/react';
import { IFormDefinition } from '../../../common/model';
import { fieldKindLabel, ISpField } from '../../../common/spTypes';
import { buildFormLink } from '../../../common/urlPrefill';
import { t } from '../../../common/i18n';

export interface ILinkBuilderProps {
  webUrl: string;
  listId: string;
  listUrl: string;
  contentTypeId: string;
  definition: IFormDefinition;
  fieldsByName: Record<string, ISpField>;
  onDismiss: () => void;
}

function formatHint(field: ISpField): string {
  switch (field.kind) {
    case 'multichoice':
      return t('Mehrere Werte mit ; trennen');
    case 'boolean':
      return t('ja/nein, 1/0, true/false');
    case 'date':
    case 'datetime':
      return t('JJJJ-MM-TT, today oder today+7');
    case 'user':
    case 'usermulti':
      return t('E-Mail-Adresse; mehrere mit ; trennen');
    case 'lookup':
    case 'lookupmulti':
      return t('ID oder Titel; mehrere mit ; trennen');
    case 'taxonomy':
    case 'taxonomymulti':
      return t('Begriff oder Begriffs-ID; mehrere mit ; trennen');
    case 'url':
      return t('URL|Beschreibung');
    default:
      return '';
  }
}

/** Erzeugt einen Link auf das Formular «Neu» mit vorbelegten Feldern. */
export const LinkBuilder: React.FC<ILinkBuilderProps> = (p) => {
  const [values, setValues] = React.useState<Record<string, string>>({});
  const [viaPanel, setViaPanel] = React.useState<boolean>(true);
  const [copied, setCopied] = React.useState<boolean>(false);
  const names: string[] = [];
  p.definition.sections.forEach((s) =>
    s.fields.forEach((n) => {
      const f = p.fieldsByName[n];
      if (f && !f.readOnly && f.kind !== 'calculated' && f.kind !== 'unsupported') {
        names.push(n);
      }
    })
  );
  const link = buildFormLink({ webUrl: p.webUrl, listId: p.listId, listUrl: p.listUrl, contentTypeId: p.contentTypeId, values, viaListPanel: viaPanel });

  return (
    <Dialog
      hidden={false}
      onDismiss={p.onDismiss}
      minWidth={560}
      maxWidth={760}
      dialogContentProps={{ type: DialogType.normal, title: t('Link mit Vorbelegung erzeugen'), subText: t('Leere Felder werden nicht in den Link übernommen.') }}
    >
      <Stack tokens={{ childrenGap: 10 }} styles={{ root: { maxHeight: '50vh', overflowY: 'auto', paddingRight: 4 } }}>
        {names.map((n) => {
          const f = p.fieldsByName[n];
          const settings = p.definition.fields[n];
          const hint = formatHint(f);
          return (
            <TextField
              key={n}
              label={`${(settings && settings.label) || f.title} (fd_${n})`}
              description={hint ? `${fieldKindLabel(f.kind)} · ${hint}` : fieldKindLabel(f.kind)}
              value={values[n] || ''}
              onChange={(_, v) => {
                setCopied(false);
                setValues({ ...values, [n]: v || '' });
              }}
            />
          );
        })}
      </Stack>
      <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 16 } }}>
        <Checkbox
          label={t('Über die Listenansicht öffnen (Panel bzw. Dialog gemäss Einstellung)')}
          checked={viaPanel}
          onChange={(_, c) => {
            setCopied(false);
            setViaPanel(!!c);
          }}
        />
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          {viaPanel
            ? t('Öffnet die Liste und darin das Formular «Neu». Voraussetzung: Die Panel-Erweiterung ist auf der Liste registriert.')
            : t('Öffnet direkt die Formularseite (ganze Seite).')}
        </Text>
        <TextField label={t('Link')} value={link} readOnly multiline autoAdjustHeight styles={{ field: { fontFamily: 'Consolas, monospace', fontSize: 12 } }} />
        {copied && <MessageBar messageBarType={MessageBarType.success}>{t('Link kopiert.')}</MessageBar>}
      </Stack>
      <DialogFooter>
        <DefaultButton
          text={t('Kopieren')}
          iconProps={{ iconName: 'Copy' }}
          onClick={() => {
            navigator.clipboard
              .writeText(link)
              .then(() => setCopied(true))
              .catch(() => setCopied(false));
          }}
        />
        <DefaultButton text={t('Testen')} iconProps={{ iconName: 'OpenInNewWindow' }} onClick={() => window.open(link, '_blank', 'noopener')} />
        <DefaultButton text={t('Schliessen')} onClick={p.onDismiss} />
      </DialogFooter>
    </Dialog>
  );
};
