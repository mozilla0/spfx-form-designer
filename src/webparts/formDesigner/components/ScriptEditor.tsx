import * as React from 'react';
import { DefaultButton, MessageBar, MessageBarType, PrimaryButton, Stack } from '@fluentui/react';
import { checkScriptSyntax } from '../../../common/scriptEngine';
import { ISpField } from '../../../common/spTypes';
import styles from './FormDesigner.module.scss';

export interface IScriptEditorProps {
  script: string;
  fields: ISpField[];
  onChange: (script: string) => void;
  onTest: () => void;
}

const SNIPPETS: Array<{ label: string; code: string }> = [
  { label: 'Beim Laden ausführen', code: "form.onReady(() => {\n  \n});\n" },
  { label: 'Bei Feldänderung reagieren', code: "form.field('Title').onChange(value => {\n  \n});\n" },
  { label: 'Feld ein-/ausblenden', code: "form.field('Title').visible = false;\n" },
  { label: 'Pflichtfeld abhängig setzen', code: "form.field('Title').required = true;\n" },
  { label: 'Eigene Feldprüfung', code: "form.field('Title').validate(value => {\n  if (value && value.length < 5) return 'Der Titel ist zu kurz.';\n});\n" },
  { label: 'Prüfung vor dem Speichern', code: "form.onBeforeSave(() => {\n  if (!form.field('Title').value) return 'Der Titel fehlt.';\n});\n" },
  { label: 'Datum vorbelegen', code: "form.onReady(() => {\n  if (form.type === 'new') form.field('Datum').value = form.date.addWorkdays(form.date.today(), 5);\n});\n" },
  { label: 'Titel dynamisch setzen', code: "form.title = 'Auftrag ' + form.field('Title').value;\n" },
  { label: 'Hinweis anzeigen', code: "form.showMessage('Text', 'info');\n" }
];

/** Editor für das Formular-Skript mit Bausteinen, Feldnamen und API-Referenz. */
export const ScriptEditor: React.FC<IScriptEditorProps> = (p) => {
  const ref = React.useRef<HTMLTextAreaElement>(null);
  const [check, setCheck] = React.useState<{ ok: boolean; text: string } | undefined>(undefined);

  const insert = (text: string): void => {
    const ta = ref.current;
    const start = ta ? ta.selectionStart : p.script.length;
    const end = ta ? ta.selectionEnd : p.script.length;
    p.onChange(p.script.slice(0, start) + text + p.script.slice(end));
    setCheck(undefined);
    window.setTimeout(() => {
      if (ref.current) {
        ref.current.focus();
        ref.current.selectionStart = ref.current.selectionEnd = start + text.length;
      }
    }, 0);
  };

  return (
    <>
      <div className={styles.scriptArea}>
        <div className={styles.panel}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 10 } }}>
            <label htmlFor="fd-script" className={styles.panelTitle}>
              Formular-Skript
            </label>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton
                text="Syntax prüfen"
                onClick={() => {
                  const err = checkScriptSyntax(p.script);
                  setCheck(err ? { ok: false, text: `Syntaxfehler: ${err}` } : { ok: true, text: 'Syntax in Ordnung.' });
                }}
              />
              <PrimaryButton text="In Vorschau testen" onClick={p.onTest} />
            </Stack>
          </Stack>
          <textarea
            id="fd-script"
            ref={ref}
            className={styles.code}
            spellCheck={false}
            value={p.script}
            placeholder="// Beispiel: form.field('Status').onChange(value => { … });"
            onChange={(e) => {
              p.onChange(e.target.value);
              setCheck(undefined);
            }}
            onKeyDown={(e) => {
              if (e.key === 'Tab' && !e.shiftKey) {
                e.preventDefault();
                insert('  ');
              }
            }}
          />
          {check && (
            <MessageBar messageBarType={check.ok ? MessageBarType.success : MessageBarType.error} styles={{ root: { marginTop: 8 } }}>
              {check.text}
            </MessageBar>
          )}
        </div>
      </div>
      <aside className={styles.properties} aria-label="Skript-Hilfe">
        <div className={styles.panel}>
          <p className={styles.panelTitle}>Bausteine einfügen</p>
          <Stack tokens={{ childrenGap: 6 }} styles={{ root: { marginTop: 8 } }}>
            {SNIPPETS.map((s) => (
              <DefaultButton key={s.label} text={s.label} onClick={() => insert(s.code)} styles={{ root: { justifyContent: 'flex-start' }, flexContainer: { justifyContent: 'flex-start' } }} />
            ))}
          </Stack>
        </div>
        <div className={styles.panel}>
          <p className={styles.panelTitle}>Felder (interne Namen)</p>
          <div className={styles.chips} style={{ marginTop: 8 }}>
            {p.fields.map((f) => (
              <button key={f.internalName} type="button" className={styles.chip} title={f.title} aria-label={`form.field('${f.internalName}') einfügen`} onClick={() => insert(`form.field('${f.internalName}')`)}>
                {f.internalName}
              </button>
            ))}
          </div>
        </div>
        <div className={styles.panel}>
          <p className={styles.panelTitle}>API-Referenz</p>
          <div className={styles.apiRef} style={{ marginTop: 8 }}>
            <div><code>form.field(&apos;Name&apos;).value</code><div className={styles.hint}>Wert lesen oder setzen</div></div>
            <div><code>.visible · .required · .disabled</code><div className={styles.hint}>Feld ein-/ausblenden, Pflicht, sperren</div></div>
            <div><code>.onChange(fn)</code><div className={styles.hint}>Reagiert auf Wertänderungen</div></div>
            <div><code>.validate(fn)</code><div className={styles.hint}>Ein zurückgegebener Text erscheint als Fehler am Feld</div></div>
            <div><code>form.onReady(fn)</code><div className={styles.hint}>Nach dem Laden und nach den Standardwerten</div></div>
            <div><code>form.onBeforeSave(fn)</code><div className={styles.hint}>Ein zurückgegebener Text bricht das Speichern ab; async möglich</div></div>
            <div><code>form.showMessage(text, art)</code><div className={styles.hint}>art: &apos;info&apos;, &apos;warning&apos;, &apos;error&apos;, &apos;success&apos;</div></div>
            <div><code>form.type · form.title</code><div className={styles.hint}>Formulartyp (&apos;new&apos;, &apos;edit&apos;, &apos;display&apos;) und Titel</div></div>
            <div><code>form.date.today() · addDays · addWorkdays · endOfMonth · format</code><div className={styles.hint}>Datumshilfen; ein Date-Objekt kann direkt einem Datumsfeld zugewiesen werden</div></div>
            <div><code>form.currentUser · form.sp.get(pfad)</code><div className={styles.hint}>Angemeldeter Benutzer und REST-Zugriff auf die Website</div></div>
          </div>
        </div>
      </aside>
    </>
  );
};
