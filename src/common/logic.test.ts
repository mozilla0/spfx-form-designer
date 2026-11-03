import { addWorkdays, endOfMonth, parseSwissDate, toIsoDate } from './dateUtils';
import { validateFieldValue } from './validation';
import { computeDefault } from './defaults';
import { createFieldSettings } from './model';
import { ISpField } from './spTypes';
import { checkScriptSyntax, FormScriptHost, IScriptHostBridge } from './scriptEngine';
import { panelCssWidth } from './panelSize';
import { parseFormUrl } from './formUrl';
import { isHexColor, textColorOn } from './components/FormHeader';

function field(kind: ISpField['kind'], name: string = 'F'): ISpField {
  return { id: name, internalName: name, title: name, typeAsString: kind, kind, required: false, readOnly: false, choices: ['Offen', 'Erledigt'], description: '' };
}

describe('dateUtils', () => {
  it('überspringt Wochenenden bei Werktagen', () => {
    // Freitag, 6.11.2026 + 1 Werktag = Montag, 9.11.2026
    expect(toIsoDate(addWorkdays(new Date(2026, 10, 6), 1))).toBe('2026-11-09');
  });
  it('liefert das Monatsende', () => {
    expect(toIsoDate(endOfMonth(new Date(2028, 1, 10)))).toBe('2028-02-29');
  });
  it('liest Schweizer Datumsformat', () => {
    expect(toIsoDate(parseSwissDate('3.11.26') as Date)).toBe('2026-11-03');
    expect(parseSwissDate('31-12-2026')).toBeNull();
  });
});

describe('validation', () => {
  it('prüft Pflichtfelder', () => {
    expect(validateFieldValue(field('text'), createFieldSettings(), '', true)).toBe('Dieses Feld ist erforderlich.');
  });
  it('prüft E-Mail-Format und eigene Meldung', () => {
    const s = createFieldSettings();
    s.validation.pattern = 'email';
    expect(validateFieldValue(field('text'), s, 'kein-mail', false)).toContain('E-Mail');
    expect(validateFieldValue(field('text'), s, 'a@b.ch', false)).toBe('');
    s.validation.message = 'Bitte Firmenadresse';
    expect(validateFieldValue(field('text'), s, 'x', false)).toBe('Bitte Firmenadresse');
  });
  it('prüft Zahlenbereiche', () => {
    const s = createFieldSettings();
    s.validation.min = 1000;
    s.validation.max = 99999;
    expect(validateFieldValue(field('number'), s, '999', false)).toContain('mindestens');
    expect(validateFieldValue(field('number'), s, '5000', false)).toBe('');
  });
  it('prüft Schweizer Telefonnummern', () => {
    const s = createFieldSettings();
    s.validation.pattern = 'phoneCH';
    expect(validateFieldValue(field('text'), s, '041 123 45 67', false)).toBe('');
    expect(validateFieldValue(field('text'), s, '+41 41 123 45 67', false)).toBe('');
    expect(validateFieldValue(field('text'), s, '123', false)).not.toBe('');
  });
});

describe('defaults', () => {
  it('berechnet Formeln mit Feldwerten', () => {
    const s = createFieldSettings();
    s.defaultValue = { mode: 'formula', formula: "field('Prio') === 'Hoch' ? 'eilig' : 'normal'" };
    expect(computeDefault(field('text'), s, { Prio: 'Hoch' }, undefined).value).toBe('eilig');
  });
  it('meldet Formelfehler', () => {
    const s = createFieldSettings();
    s.defaultValue = { mode: 'formula', formula: 'gibtsNicht()' };
    expect(computeDefault(field('text'), s, {}, undefined).error).toContain('Formelfehler');
  });
  it('wandelt Datumsergebnisse um', () => {
    const s = createFieldSettings();
    s.defaultValue = { mode: 'formula', formula: 'addDays(today(), 0)' };
    expect(computeDefault(field('date'), s, {}, undefined).value).toBe(toIsoDate(new Date()));
  });
});

describe('scriptEngine', () => {
  it('führt Handler aus und meldet Syntaxfehler', () => {
    const values: Record<string, unknown> = { Status: 'Offen' };
    const flags: Record<string, boolean> = {};
    const errors: string[] = [];
    const bridge: IScriptHostBridge = {
      mode: 'edit',
      hasField: () => true,
      fieldKind: () => 'text',
      getValue: (n) => values[n] as string,
      setValue: (n, v) => {
        values[n] = v;
      },
      getFlag: (n, f) => !!flags[`${n}.${f}`],
      setFlag: (n, f, v) => {
        flags[`${n}.${f}`] = v;
      },
      getTitle: () => '',
      setTitle: () => undefined,
      showMessage: () => undefined,
      reportError: (t) => errors.push(t)
    };
    const host = new FormScriptHost(bridge);
    host.run("form.field('Status').onChange(v => { form.field('Notiz').visible = v === 'Erledigt'; });\nform.field('Notiz').validate(v => v ? '' : 'Notiz fehlt');");
    host.fireChange('Status', 'Erledigt');
    expect(flags['Notiz.visible']).toBe(true);
    expect(host.validate('Notiz', '')).toBe('Notiz fehlt');
    expect(errors).toHaveLength(0);
    expect(checkScriptSyntax('form.onReady(() => {')).not.toBe('');
  });
});

describe('panelSize', () => {
  it('liefert die Breite je Einstellung', () => {
    expect(panelCssWidth(undefined)).toBeUndefined();
    expect(panelCssWidth({ panelSize: 'default' })).toBeUndefined();
    expect(panelCssWidth({ panelSize: 'large' })).toBe('880px');
    expect(panelCssWidth({ panelSize: 'wide' })).toBe('90vw');
    expect(panelCssWidth({ panelSize: 'custom', panelWidth: 720 })).toBe('720px');
    expect(panelCssWidth({ panelSize: 'custom', panelWidth: 200 })).toBeUndefined();
  });
});

describe('formUrl', () => {
  it('erkennt SPListForm.aspx (Bearbeiten)', () => {
    const r = parseFormUrl(
      'https://contoso.sharepoint.com/sites/test/_layouts/15/SPListForm.aspx?PageType=6&List=57dc254f%2Db1a2%2D4764%2Da0b8%2D55ced2864615&ID=8&Source=https%3A%2F%2Fcontoso%2Esharepoint%2Ecom%2Fsites%2Ftest%2FLists%2FOrders%2FAllItems%2Easpx&ContentTypeId=0x0100ABC&RootFolder=%2Fsites%2Ftest%2FLists%2FOrders&loadFormWithNucleus=1'
    );
    expect(r && r.mode).toBe('edit');
    expect(r && r.listId).toBe('57dc254f-b1a2-4764-a0b8-55ced2864615');
    expect(r && r.itemId).toBe(8);
    expect(r && r.contentTypeId).toBe('0x0100ABC');
  });
  it('erkennt Neu mit Klammern in der Listen-ID', () => {
    const r = parseFormUrl('https://contoso.sharepoint.com/_layouts/15/SPListForm.aspx?PageType=8&List=%7B57DC254F-B1A2-4764-A0B8-55CED2864615%7D');
    expect(r && r.mode).toBe('new');
    expect(r && r.listId).toBe('57dc254f-b1a2-4764-a0b8-55ced2864615');
  });
  it('erkennt klassische Formularpfade', () => {
    const r = parseFormUrl('https://contoso.sharepoint.com/sites/test/Lists/Orders/DispForm.aspx?ID=3');
    expect(r && r.mode).toBe('display');
    expect(r && r.listPath).toBe('/sites/test/lists/orders');
  });
  it('ignoriert andere Seiten', () => {
    expect(parseFormUrl('https://contoso.sharepoint.com/sites/test/Lists/Orders/AllItems.aspx')).toBeUndefined();
  });
});

describe('FormHeader', () => {
  it('wählt lesbare Schrift auf farbigem Band', () => {
    expect(textColorOn('#0f6cbd')).toBe('#ffffff');
    expect(textColorOn('#ffd335')).toBe('#1b1a19');
    expect(isHexColor('#12ab9F')).toBe(true);
    expect(isHexColor('blau')).toBe(false);
  });
});
