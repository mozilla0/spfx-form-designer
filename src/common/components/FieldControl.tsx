import * as React from 'react';
import {
  Checkbox,
  ComboBox,
  DatePicker,
  DayOfWeek,
  defaultDatePickerStrings,
  Dropdown,
  IComboBoxOption,
  IDatePickerStrings,
  IDropdownOption,
  IPersonaProps,
  Label,
  Link,
  NormalPeoplePicker,
  Stack,
  Text,
  TextField
} from '@fluentui/react';
import { FieldValue, ILookupValue, ISpField, ITermValue, IUrlValue, IUserValue, valueToText } from '../spTypes';
import { ITermOption, SpService } from '../spService';
import { datePlaceholder, formatDate, parseDate, toIsoDate, toLocalDate } from '../dateUtils';
import { getLanguage, t } from '../i18n';

function datePickerStrings(): IDatePickerStrings {
  if (getLanguage() !== 'de') {
    return { ...defaultDatePickerStrings, invalidInputErrorMessage: t('Ungültiges Datum.'), isRequiredErrorMessage: '' };
  }
  return {
    ...defaultDatePickerStrings,
    months: ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'],
    shortMonths: ['Jan', 'Feb', 'Mär', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'],
    days: ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'],
    shortDays: ['S', 'M', 'D', 'M', 'D', 'F', 'S'],
    goToToday: 'Heute',
    prevMonthAriaLabel: 'Vorheriger Monat',
    nextMonthAriaLabel: 'Nächster Monat',
    prevYearAriaLabel: 'Vorheriges Jahr',
    nextYearAriaLabel: 'Nächstes Jahr',
    closeButtonAriaLabel: 'Schliessen',
    invalidInputErrorMessage: 'Ungültiges Datum.',
    isRequiredErrorMessage: ''
  };
}

export interface IFieldControlProps {
  field: ISpField;
  label: string;
  value: FieldValue;
  required: boolean;
  disabled: boolean;
  displayOnly: boolean;
  error?: string;
  help?: string;
  sp?: SpService;
  onChange: (value: FieldValue) => void;
  onBlur?: () => void;
}

const ERROR_STYLE = { root: { color: '#a4262c', display: 'block', marginTop: 4 } };

const ErrorText: React.FC<{ text?: string }> = ({ text }) =>
  text ? (
    <Text variant="small" styles={ERROR_STYLE} role="alert">
      {text}
    </Text>
  ) : null;

function displayText(field: ISpField, value: FieldValue): string {
  if (field.kind === 'date' && typeof value === 'string' && value) {
    return formatDate(value);
  }
  if (field.kind === 'datetime' && typeof value === 'string' && value) {
    return `${formatDate(value)} ${value.slice(11, 16)}`;
  }
  return valueToText(value);
}

// ---------- Nachschlagen (einzeln und mehrfach) ----------

const LookupField: React.FC<IFieldControlProps> = (p) => {
  const multi = p.field.kind === 'lookupmulti';
  const [options, setOptions] = React.useState<ILookupValue[] | undefined>(undefined);
  const [loadError, setLoadError] = React.useState<string>('');
  React.useEffect(() => {
    let alive = true;
    if (p.sp) {
      p.sp
        .getLookupItems(p.field)
        .then((items) => alive && setOptions(items))
        .catch((e: Error) => alive && setLoadError(e.message));
    } else {
      setOptions([]);
    }
    return () => {
      alive = false;
    };
  }, [p.field.internalName]);

  const selected: ILookupValue[] = multi
    ? Array.isArray(p.value)
      ? (p.value as ILookupValue[])
      : []
    : p.value && typeof p.value === 'object' && !Array.isArray(p.value)
      ? [p.value as ILookupValue]
      : [];
  const known = options || [];
  const list = known.concat(selected.filter((s) => !known.some((o) => o.id === s.id)));
  const placeholder = options ? t('Eintrag wählen') : t('Einträge werden geladen …');

  if (multi) {
    return (
      <Dropdown
        label={p.label}
        required={p.required}
        disabled={p.disabled}
        multiSelect
        placeholder={placeholder}
        options={list.map((o) => ({ key: o.id, text: o.title }))}
        selectedKeys={selected.map((s) => s.id)}
        errorMessage={p.error || loadError || undefined}
        onChange={(_, o) => {
          if (!o) {
            return;
          }
          const id = Number(o.key);
          p.onChange(o.selected ? selected.concat([{ id, title: o.text }]) : selected.filter((s) => s.id !== id));
        }}
        onBlur={p.onBlur}
      />
    );
  }
  const dropdownOptions: IDropdownOption[] = [{ key: 0, text: t('(Keine Auswahl)') }].concat(list.map((o) => ({ key: o.id, text: o.title })));
  return (
    <Dropdown
      label={p.label}
      required={p.required}
      disabled={p.disabled}
      placeholder={placeholder}
      options={dropdownOptions}
      selectedKey={selected[0] ? selected[0].id : 0}
      errorMessage={p.error || loadError || undefined}
      onChange={(_, o) => p.onChange(!o || o.key === 0 ? null : { id: Number(o.key), title: o.text })}
      onBlur={p.onBlur}
    />
  );
};

// ---------- Personen (einzeln und mehrfach) ----------

const PeopleField: React.FC<IFieldControlProps> = (p) => {
  const multi = p.field.kind === 'usermulti';
  const [busy, setBusy] = React.useState<boolean>(false);
  const current: IUserValue[] = multi
    ? Array.isArray(p.value)
      ? (p.value as IUserValue[])
      : []
    : p.value && typeof p.value === 'object' && !Array.isArray(p.value)
      ? [p.value as IUserValue]
      : [];
  const selected: IPersonaProps[] = current.map((u) => ({ key: `id:${u.id}`, text: u.title, secondaryText: u.email }));
  const inputId = `fd-${p.field.internalName}`;

  const resolve = async (filter: string): Promise<IPersonaProps[]> => {
    if (!p.sp || filter.length < 2) {
      return [];
    }
    const users = await p.sp.searchUsers(filter);
    return users.map((u) => ({ key: u.loginName || u.title, text: u.title, secondaryText: u.email }));
  };

  const onChange = (items: IPersonaProps[] | undefined): void => {
    const list = items || [];
    if (!p.sp) {
      return;
    }
    // Bereits bekannte Personen behalten, neue über ensureUser auflösen.
    const keep = current.filter((u) => list.some((i) => i.key === `id:${u.id}`));
    const added = list.filter((i) => String(i.key).indexOf('id:') !== 0);
    if (!added.length) {
      p.onChange(multi ? keep : keep[0] || null);
      return;
    }
    setBusy(true);
    Promise.all(added.map((i) => (p.sp as SpService).ensureUser(String(i.key))))
      .then((users) => {
        setBusy(false);
        const next = keep.concat(users);
        p.onChange(multi ? next : next[next.length - 1] || null);
      })
      .catch(() => {
        setBusy(false);
        p.onChange(multi ? keep : null);
      });
  };

  return (
    <div>
      <Label htmlFor={inputId} required={p.required} disabled={p.disabled}>
        {p.label}
      </Label>
      <NormalPeoplePicker
        inputProps={{ id: inputId, onBlur: p.onBlur, placeholder: t('Name oder E-Mail eingeben') }}
        itemLimit={multi ? undefined : 1}
        disabled={p.disabled || busy}
        selectedItems={selected}
        onResolveSuggestions={(filter) => resolve(filter)}
        pickerSuggestionsProps={{ noResultsFoundText: t('Keine Personen gefunden'), loadingText: t('Suche …') }}
        onChange={onChange}
      />
      <ErrorText text={p.error} />
    </div>
  );
};

// ---------- Verwaltete Metadaten ----------

const TaxonomyField: React.FC<IFieldControlProps> = (p) => {
  const multi = p.field.kind === 'taxonomymulti';
  const [terms, setTerms] = React.useState<ITermOption[] | undefined>(undefined);
  const [loadError, setLoadError] = React.useState<string>('');
  React.useEffect(() => {
    let alive = true;
    if (p.sp) {
      p.sp
        .getTerms(p.field)
        .then((x) => alive && setTerms(x))
        .catch((e: Error) => alive && setLoadError(t('Begriffe konnten nicht geladen werden: {0}', e.message)));
    } else {
      setTerms([]);
    }
    return () => {
      alive = false;
    };
  }, [p.field.internalName]);

  const selected: ITermValue[] = multi
    ? Array.isArray(p.value)
      ? (p.value as ITermValue[])
      : []
    : p.value && typeof p.value === 'object' && !Array.isArray(p.value) && 'termGuid' in p.value
      ? [p.value as ITermValue]
      : [];
  const known = terms || [];
  const options: IComboBoxOption[] = known
    .map((o) => ({ key: o.termGuid, text: o.path }))
    .concat(selected.filter((s) => !known.some((k) => k.termGuid === s.termGuid)).map((s) => ({ key: s.termGuid, text: s.label })));
  const labelOf = (guid: string, fallback: string): string => {
    const term = known.filter((k) => k.termGuid === guid)[0];
    return term ? term.label : fallback;
  };

  return (
    <ComboBox
      label={p.label}
      required={p.required}
      disabled={p.disabled}
      multiSelect={multi}
      allowFreeform={false}
      autoComplete="on"
      placeholder={terms ? t('Begriff wählen') : t('Begriffe werden geladen …')}
      options={options}
      selectedKey={multi ? selected.map((s) => s.termGuid) : selected[0] ? selected[0].termGuid : null}
      errorMessage={p.error || loadError || undefined}
      useComboBoxAsMenuWidth
      onChange={(_, o) => {
        if (!o) {
          return;
        }
        const guid = String(o.key);
        if (multi) {
          p.onChange(o.selected ? selected.concat([{ termGuid: guid, label: labelOf(guid, o.text) }]) : selected.filter((s) => s.termGuid !== guid));
        } else {
          p.onChange({ termGuid: guid, label: labelOf(guid, o.text) });
        }
      }}
      onBlur={p.onBlur}
    />
  );
};

// ---------- Link oder Bild ----------

const UrlField: React.FC<IFieldControlProps> = (p) => {
  const v: IUrlValue = p.value && typeof p.value === 'object' && !Array.isArray(p.value) && 'url' in p.value ? (p.value as IUrlValue) : { url: '', description: '' };
  return (
    <div>
      <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
        <TextField
          label={p.label}
          required={p.required}
          disabled={p.disabled}
          type="url"
          placeholder="https://…"
          value={v.url}
          styles={{ root: { flex: '2 1 220px' } }}
          onChange={(_, url) => p.onChange({ url: url || '', description: v.description })}
          onBlur={p.onBlur}
        />
        <TextField
          ariaLabel={t('{0} – Beschreibung', p.label)}
          placeholder={p.field.urlFormat === 'image' ? t('Alternativtext') : t('Beschreibung (optional)')}
          disabled={p.disabled}
          value={v.description}
          styles={{ root: { flex: '1 1 160px' } }}
          onChange={(_, description) => p.onChange({ url: v.url, description: description || '' })}
          onBlur={p.onBlur}
        />
      </Stack>
      <ErrorText text={p.error} />
    </div>
  );
};

function UrlDisplay({ field, value }: { field: ISpField; value: FieldValue }): JSX.Element {
  const v = value && typeof value === 'object' && !Array.isArray(value) && 'url' in value ? (value as IUrlValue) : undefined;
  if (!v || !v.url) {
    return <Text>–</Text>;
  }
  if (field.urlFormat === 'image') {
    return <img src={v.url} alt={v.description} style={{ maxWidth: '100%', maxHeight: 240 }} />;
  }
  return (
    <Link href={v.url} target="_blank" rel="noopener noreferrer">
      {v.description || v.url}
    </Link>
  );
}

/** Rendert ein einzelnes Feld abhängig vom Spaltentyp. */
export const FieldControl: React.FC<IFieldControlProps> = (p) => {
  const { field, value } = p;

  if (p.displayOnly || field.kind === 'calculated') {
    return (
      <div>
        <Label>{p.label}</Label>
        {field.kind === 'url' ? (
          <UrlDisplay field={field} value={value} />
        ) : (
          <Text styles={{ root: { whiteSpace: field.kind === 'note' ? 'pre-wrap' : 'normal', display: 'block', minHeight: 20 } }}>{displayText(field, value) || '–'}</Text>
        )}
        {p.help && (
          <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block', marginTop: 4 } }}>
            {p.help}
          </Text>
        )}
      </div>
    );
  }

  const errorMessage = p.error || undefined;
  const str = typeof value === 'string' ? value : '';

  let control: JSX.Element;
  switch (field.kind) {
    case 'text':
    case 'note':
      control = (
        <TextField
          label={p.label}
          required={p.required}
          disabled={p.disabled}
          multiline={field.kind === 'note'}
          rows={field.kind === 'note' ? 4 : undefined}
          autoAdjustHeight={field.kind === 'note'}
          value={str}
          errorMessage={errorMessage}
          onChange={(_, v) => p.onChange(v ?? '')}
          onBlur={p.onBlur}
        />
      );
      break;
    case 'number':
      control = (
        <TextField label={p.label} required={p.required} disabled={p.disabled} type="number" value={str} errorMessage={errorMessage} onChange={(_, v) => p.onChange(v ?? '')} onBlur={p.onBlur} />
      );
      break;
    case 'choice': {
      const options: IDropdownOption[] = (p.required ? [] : [{ key: '', text: t('(Keine Auswahl)') }]).concat(field.choices.map((c) => ({ key: c, text: c })));
      control = (
        <Dropdown
          label={p.label}
          required={p.required}
          disabled={p.disabled}
          options={options}
          selectedKey={str}
          placeholder={t('Auswahl treffen')}
          errorMessage={errorMessage}
          onChange={(_, o) => p.onChange(o ? String(o.key) : '')}
          onBlur={p.onBlur}
        />
      );
      break;
    }
    case 'multichoice': {
      const selected = Array.isArray(value) ? (value as string[]) : [];
      control = (
        <Dropdown
          label={p.label}
          required={p.required}
          disabled={p.disabled}
          multiSelect
          options={field.choices.map((c) => ({ key: c, text: c }))}
          selectedKeys={selected}
          placeholder={t('Auswahl treffen')}
          errorMessage={errorMessage}
          onChange={(_, o) => {
            if (!o) {
              return;
            }
            const key = String(o.key);
            p.onChange(o.selected ? selected.concat([key]) : selected.filter((k) => k !== key));
          }}
          onBlur={p.onBlur}
        />
      );
      break;
    }
    case 'boolean':
      control = (
        <Stack tokens={{ childrenGap: 6 }}>
          <Label required={p.required} disabled={p.disabled}>
            {p.label}
          </Label>
          <Checkbox label={t('Ja')} checked={value === true} disabled={p.disabled} onChange={(_, c) => p.onChange(!!c)} />
          <ErrorText text={errorMessage} />
        </Stack>
      );
      break;
    case 'date':
    case 'datetime': {
      const datePart = str ? toLocalDate(str) : undefined;
      const timePart = field.kind === 'datetime' && str.length >= 16 ? str.slice(11, 16) : '00:00';
      const picker = (
        <DatePicker
          label={p.label}
          isRequired={p.required}
          disabled={p.disabled}
          allowTextInput
          firstDayOfWeek={getLanguage() === 'de' ? DayOfWeek.Monday : DayOfWeek.Sunday}
          strings={datePickerStrings()}
          placeholder={datePlaceholder()}
          value={datePart}
          formatDate={(d) => (d ? formatDate(d) : '')}
          parseDateFromString={(s) => parseDate(s)}
          onSelectDate={(d) => {
            if (!d) {
              p.onChange('');
            } else {
              p.onChange(field.kind === 'datetime' ? `${toIsoDate(d)}T${timePart}` : toIsoDate(d));
            }
          }}
          onAfterMenuDismiss={p.onBlur}
          styles={{ root: { flexGrow: 1 } }}
        />
      );
      control =
        field.kind === 'date' ? (
          <div>
            {picker}
            <ErrorText text={errorMessage} />
          </div>
        ) : (
          <div>
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
              {picker}
              <TextField
                ariaLabel={t('{0} – Uhrzeit', p.label)}
                type="time"
                disabled={p.disabled || !datePart}
                value={timePart}
                styles={{ root: { width: 110 } }}
                onChange={(_, time) => datePart && p.onChange(`${toIsoDate(datePart)}T${time || '00:00'}`)}
                onBlur={p.onBlur}
              />
            </Stack>
            <ErrorText text={errorMessage} />
          </div>
        );
      break;
    }
    case 'user':
    case 'usermulti':
      control = <PeopleField {...p} />;
      break;
    case 'lookup':
    case 'lookupmulti':
      control = <LookupField {...p} />;
      break;
    case 'taxonomy':
    case 'taxonomymulti':
      control = <TaxonomyField {...p} />;
      break;
    case 'url':
      control = <UrlField {...p} />;
      break;
    default:
      control = (
        <div>
          <Label>{p.label}</Label>
          <Text variant="small">{t('{0} (Spaltentyp «{1}» wird im Formular noch nicht bearbeitet)', displayText(field, value) || '–', field.typeAsString)}</Text>
        </div>
      );
  }

  return (
    <div>
      {control}
      {p.help && (
        <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block', marginTop: 4 } }}>
          {p.help}
        </Text>
      )}
    </div>
  );
};
