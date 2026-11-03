import * as React from 'react';
import {
  Checkbox,
  DatePicker,
  DayOfWeek,
  defaultDatePickerStrings,
  Dropdown,
  IDatePickerStrings,
  IDropdownOption,
  IPersonaProps,
  Label,
  NormalPeoplePicker,
  Stack,
  Text,
  TextField
} from '@fluentui/react';
import { FieldValue, ILookupValue, ISpField, IUserValue, valueToText } from '../spTypes';
import { SpService } from '../spService';
import { formatDate, parseSwissDate, toIsoDate, toLocalDate } from '../dateUtils';

export const DATE_PICKER_STRINGS: IDatePickerStrings = {
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
  invalidInputErrorMessage: 'Ungültiges Datum. Format: TT.MM.JJJJ',
  isRequiredErrorMessage: ''
};

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

function displayText(field: ISpField, value: FieldValue): string {
  if (field.kind === 'date' && typeof value === 'string' && value) {
    return formatDate(value);
  }
  if (field.kind === 'datetime' && typeof value === 'string' && value) {
    return `${formatDate(value)} ${value.slice(11, 16)}`;
  }
  return valueToText(value);
}

const LookupDropdown: React.FC<IFieldControlProps> = (p) => {
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

  const current = p.value && typeof p.value === 'object' && !Array.isArray(p.value) ? (p.value as ILookupValue) : undefined;
  const list = options || (current ? [current] : []);
  const dropdownOptions: IDropdownOption[] = [{ key: 0, text: '(Keine Auswahl)' }].concat(list.map((o) => ({ key: o.id, text: o.title })));
  return (
    <Dropdown
      label={p.label}
      required={p.required}
      disabled={p.disabled}
      placeholder={options ? 'Eintrag wählen' : 'Einträge werden geladen …'}
      options={dropdownOptions}
      selectedKey={current ? current.id : 0}
      errorMessage={p.error || loadError || undefined}
      onChange={(_, o) => {
        if (!o || o.key === 0) {
          p.onChange(null);
        } else {
          p.onChange({ id: Number(o.key), title: o.text });
        }
      }}
      onBlur={p.onBlur}
    />
  );
};

const PeopleField: React.FC<IFieldControlProps> = (p) => {
  const [busy, setBusy] = React.useState<boolean>(false);
  const current = p.value && typeof p.value === 'object' && !Array.isArray(p.value) ? (p.value as IUserValue) : undefined;
  const selected: IPersonaProps[] = current ? [{ key: String(current.id), text: current.title, secondaryText: current.email }] : [];
  const inputId = `fd-${p.field.internalName}`;

  const resolve = async (filter: string): Promise<IPersonaProps[]> => {
    if (!p.sp || filter.length < 2) {
      return [];
    }
    const users = await p.sp.searchUsers(filter);
    return users.map((u) => ({ key: u.loginName || u.title, text: u.title, secondaryText: u.email }));
  };

  return (
    <div>
      <Label htmlFor={inputId} required={p.required} disabled={p.disabled}>
        {p.label}
      </Label>
      <NormalPeoplePicker
        inputProps={{ id: inputId, onBlur: p.onBlur, placeholder: 'Name oder E-Mail eingeben' }}
        itemLimit={1}
        disabled={p.disabled || busy}
        selectedItems={selected}
        onResolveSuggestions={(filter) => resolve(filter)}
        pickerSuggestionsProps={{ noResultsFoundText: 'Keine Personen gefunden', loadingText: 'Suche …' }}
        onChange={(items) => {
          const item = items && items[0];
          if (!item) {
            p.onChange(null);
            return;
          }
          if (!p.sp) {
            return;
          }
          setBusy(true);
          p.sp
            .ensureUser(String(item.key))
            .then((u) => {
              setBusy(false);
              p.onChange(u);
            })
            .catch(() => {
              setBusy(false);
              p.onChange(null);
            });
        }}
      />
      {p.error && (
        <Text variant="small" styles={{ root: { color: '#a4262c' } }} role="alert">
          {p.error}
        </Text>
      )}
    </div>
  );
};

/** Rendert ein einzelnes Feld abhängig vom Spaltentyp. */
export const FieldControl: React.FC<IFieldControlProps> = (p) => {
  const { field, value } = p;

  if (p.displayOnly) {
    return (
      <div>
        <Label>{p.label}</Label>
        <Text styles={{ root: { whiteSpace: field.kind === 'note' ? 'pre-wrap' : 'normal', display: 'block', minHeight: 20 } }}>
          {displayText(field, value) || '–'}
        </Text>
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
        <TextField
          label={p.label}
          required={p.required}
          disabled={p.disabled}
          type="number"
          value={str}
          errorMessage={errorMessage}
          onChange={(_, v) => p.onChange(v ?? '')}
          onBlur={p.onBlur}
        />
      );
      break;
    case 'choice': {
      const options: IDropdownOption[] = (p.required ? [] : [{ key: '', text: '(Keine Auswahl)' }]).concat(field.choices.map((c) => ({ key: c, text: c })));
      control = (
        <Dropdown
          label={p.label}
          required={p.required}
          disabled={p.disabled}
          options={options}
          selectedKey={str}
          placeholder="Auswahl treffen"
          errorMessage={errorMessage}
          onChange={(_, o) => p.onChange(o ? String(o.key) : '')}
          onBlur={p.onBlur}
        />
      );
      break;
    }
    case 'multichoice': {
      const selected = Array.isArray(value) ? value : [];
      control = (
        <Dropdown
          label={p.label}
          required={p.required}
          disabled={p.disabled}
          multiSelect
          options={field.choices.map((c) => ({ key: c, text: c }))}
          selectedKeys={selected}
          placeholder="Auswahl treffen"
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
          <Checkbox label="Ja" checked={value === true} disabled={p.disabled} onChange={(_, c) => p.onChange(!!c)} />
          {errorMessage && (
            <Text variant="small" styles={{ root: { color: '#a4262c' } }} role="alert">
              {errorMessage}
            </Text>
          )}
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
          firstDayOfWeek={DayOfWeek.Monday}
          strings={DATE_PICKER_STRINGS}
          placeholder="TT.MM.JJJJ"
          value={datePart}
          formatDate={(d) => (d ? formatDate(d) : '')}
          parseDateFromString={(s) => parseSwissDate(s)}
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
            {errorMessage && (
              <Text variant="small" styles={{ root: { color: '#a4262c' } }} role="alert">
                {errorMessage}
              </Text>
            )}
          </div>
        ) : (
          <div>
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
              {picker}
              <TextField
                ariaLabel={`${p.label} – Uhrzeit`}
                type="time"
                disabled={p.disabled || !datePart}
                value={timePart}
                styles={{ root: { width: 110 } }}
                onChange={(_, t) => datePart && p.onChange(`${toIsoDate(datePart)}T${t || '00:00'}`)}
                onBlur={p.onBlur}
              />
            </Stack>
            {errorMessage && (
              <Text variant="small" styles={{ root: { color: '#a4262c' } }} role="alert">
                {errorMessage}
              </Text>
            )}
          </div>
        );
      break;
    }
    case 'user':
      control = <PeopleField {...p} />;
      break;
    case 'lookup':
      control = <LookupDropdown {...p} />;
      break;
    default:
      control = (
        <div>
          <Label>{p.label}</Label>
          <Text variant="small">
            {displayText(field, value) || '–'} (Spaltentyp «{field.typeAsString}» wird im Formular noch nicht bearbeitet)
          </Text>
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
