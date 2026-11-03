import * as React from 'react';
import { DefaultButton, MessageBar, MessageBarType, Pivot, PivotItem, PrimaryButton, Spinner, SpinnerSize, Stack, Text } from '@fluentui/react';
import { FormMode, IFieldSettings, IFormDefinition, ISection } from '../model';
import { FieldValue, FieldValues, ICurrentUser, ISpField, isEmptyValue, valueToText } from '../spTypes';
import { SpService } from '../spService';
import { computeDefault } from '../defaults';
import { validateFieldValue } from '../validation';
import { FormScriptHost, IScriptHostBridge, MessageKind } from '../scriptEngine';
import { FieldControl } from './FieldControl';
import { FormHeader } from './FormHeader';

export interface IFormHostProps {
  definition: IFormDefinition;
  fields: ISpField[];
  mode: FormMode;
  initialValues: FieldValues;
  currentUser?: ICurrentUser;
  sp?: SpService;
  /** Vorschau im Designer: Hinweis anzeigen, nichts speichern. */
  preview?: boolean;
  onSubmit: (values: FieldValues) => Promise<void>;
  onCancel: () => void;
  onRequestEdit?: () => void;
}

interface IFormHostState {
  activeTab: string;
  saving: boolean;
  revision: number;
}

interface IMessage {
  text: string;
  kind: MessageKind;
}

type Flag = 'visible' | 'required' | 'disabled';

const FALLBACK_TITLES: Record<FormMode, string> = {
  new: 'Neues Element',
  edit: 'Element bearbeiten',
  display: 'Element anzeigen'
};

const MESSAGE_TYPES: Record<MessageKind, MessageBarType> = {
  info: MessageBarType.info,
  warning: MessageBarType.warning,
  error: MessageBarType.error,
  success: MessageBarType.success
};

/**
 * Rendert ein Formular aus einer Formulardefinition – in der Runtime und in der Vorschau des Designers.
 */
export class FormHost extends React.Component<IFormHostProps, IFormHostState> {
  private values: FieldValues;
  private readonly overrides: Record<Flag, Record<string, boolean | undefined>> = { visible: {}, required: {}, disabled: {} };
  private errors: Record<string, string> = {};
  private messages: IMessage[] = [];
  private scriptErrors: string[] = [];
  private titleOverride: string | undefined;
  private readonly fieldsByName: Record<string, ISpField> = {};
  private readonly host: FormScriptHost;

  constructor(props: IFormHostProps) {
    super(props);
    props.fields.forEach((f) => {
      this.fieldsByName[f.internalName] = f;
    });
    this.values = { ...props.initialValues };
    if (props.mode === 'new') {
      this.applyDefaults();
    }
    this.host = new FormScriptHost(this.createBridge());
    const firstTab = props.definition.tabs[0];
    this.state = { activeTab: firstTab ? firstTab.id : '', saving: false, revision: 0 };
  }

  public componentDidMount(): void {
    this.host.run(this.props.definition.script);
    this.host.fireReady();
    this.refresh();
  }

  public render(): React.ReactElement<IFormHostProps> {
    const { definition, mode } = this.props;
    const displayOnly = mode === 'display';
    const tabs = definition.tabs.filter((t) => definition.sections.some((s) => s.tabId === t.id));
    const activeTab = tabs.some((t) => t.id === this.state.activeTab) ? this.state.activeTab : tabs[0] ? tabs[0].id : '';

    return (
      <Stack tokens={{ childrenGap: 16 }} styles={{ root: { padding: '16px 20px', maxWidth: 960, margin: '0 auto', boxSizing: 'border-box', width: '100%' } }}>
        <FormHeader title={this.title()} subtitle={this.subtitle()} header={this.props.definition.header} />

        {this.props.preview && (
          <MessageBar messageBarType={MessageBarType.info}>Vorschau mit ausgeführtem Formular-Skript – es wird nichts in SharePoint gespeichert.</MessageBar>
        )}
        {this.scriptErrors.map((e, i) => (
          <MessageBar key={`se${i}`} messageBarType={MessageBarType.severeWarning}>
            {e}
          </MessageBar>
        ))}
        {this.messages.map((m, i) => (
          <MessageBar key={`m${i}`} messageBarType={MESSAGE_TYPES[m.kind]} onDismiss={() => this.dismissMessage(i)} dismissButtonAriaLabel="Schliessen">
            {m.text}
          </MessageBar>
        ))}

        {tabs.length > 1 && (
          <Pivot selectedKey={activeTab} onLinkClick={(item) => item && this.setState({ activeTab: String(item.props.itemKey) })} headersOnly>
            {tabs.map((t) => {
              const n = this.errorCount(t.id);
              return <PivotItem key={t.id} itemKey={t.id} headerText={n ? `${t.label} (${n} Fehler)` : t.label} />;
            })}
          </Pivot>
        )}

        {definition.sections
          .filter((s) => s.tabId === activeTab)
          .map((s) => this.renderSection(s, displayOnly))}

        <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }} styles={{ root: { borderTop: '1px solid #edebe9', paddingTop: 16 } }}>
          {this.state.saving && <Spinner size={SpinnerSize.small} label="Wird gespeichert …" labelPosition="right" />}
          {displayOnly ? (
            <>
              {this.props.onRequestEdit && <PrimaryButton text="Bearbeiten" onClick={this.props.onRequestEdit} />}
              <DefaultButton text="Schliessen" onClick={this.props.onCancel} />
            </>
          ) : (
            <>
              <PrimaryButton text="Speichern" disabled={this.state.saving} onClick={() => this.submit().catch(() => undefined)} />
              <DefaultButton text="Abbrechen" disabled={this.state.saving} onClick={this.props.onCancel} />
            </>
          )}
        </Stack>
      </Stack>
    );
  }

  private renderSection(section: ISection, displayOnly: boolean): JSX.Element | null {
    const visibleFields = section.fields.filter((name) => this.fieldsByName[name] && this.isVisible(name));
    if (!visibleFields.length) {
      return null;
    }
    return (
      <section key={section.id} aria-label={section.title}>
        {section.title && (
          <Text as="h2" variant="large" styles={{ root: { margin: '0 0 8px', fontWeight: 600 } }}>
            {section.title}
          </Text>
        )}
        <div style={{ display: 'grid', gridTemplateColumns: section.columns === 2 ? 'repeat(auto-fit, minmax(260px, 1fr))' : 'minmax(0, 1fr)', gap: '12px 20px' }}>
          {visibleFields.map((name) => {
            const field = this.fieldsByName[name];
            const settings = this.settings(name);
            return (
              <div key={name} style={{ gridColumn: section.columns === 2 && settings.width === 'full' ? '1 / -1' : undefined, minWidth: 0 }}>
                <FieldControl
                  field={field}
                  label={settings.label || field.title}
                  value={this.values[name]}
                  required={this.isRequired(name)}
                  disabled={this.isDisabled(name) || this.state.saving}
                  displayOnly={displayOnly}
                  error={this.errors[name]}
                  help={settings.help}
                  sp={this.props.sp}
                  onChange={(v) => this.handleChange(name, v)}
                  onBlur={() => this.handleBlur(name)}
                />
              </div>
            );
          })}
        </div>
      </section>
    );
  }

  // ---------- Zustand ----------

  private refresh(): void {
    this.setState((s) => ({ revision: s.revision + 1 }));
  }

  private settings(name: string): IFieldSettings {
    return this.props.definition.fields[name] || { required: false, readOnly: false, width: 'half', visibility: { operator: 'always' }, validation: { pattern: 'none', dateRule: 'none' }, defaultValue: { mode: 'none' } };
  }

  private applyDefaults(): void {
    Object.keys(this.props.definition.fields).forEach((name) => {
      const field = this.fieldsByName[name];
      if (!field) {
        return;
      }
      const r = computeDefault(field, this.settings(name), this.values, this.props.currentUser);
      if (!r.has) {
        return;
      }
      if (r.error) {
        this.scriptErrors.push(`Standardwert «${this.settings(name).label || field.title}»: ${r.error}`);
      } else if (isEmptyValue(this.values[name])) {
        this.values[name] = r.value;
      }
    });
  }

  private isVisible(name: string): boolean {
    const o = this.overrides.visible[name];
    if (o !== undefined) {
      return o;
    }
    const rule = this.settings(name).visibility;
    switch (rule.operator) {
      case 'eq':
        return valueToText(this.values[rule.field || '']) === (rule.value || '');
      case 'ne':
        return valueToText(this.values[rule.field || '']) !== (rule.value || '');
      case 'empty':
        return isEmptyValue(this.values[rule.field || '']);
      case 'notEmpty':
        return !isEmptyValue(this.values[rule.field || '']);
      case 'modeIs':
        return this.props.mode === rule.mode;
      default:
        return true;
    }
  }

  private isRequired(name: string): boolean {
    const field = this.fieldsByName[name];
    const o = this.overrides.required[name];
    const base = !!field && field.required;
    return o !== undefined ? o || base : this.settings(name).required || base;
  }

  private isDisabled(name: string): boolean {
    const field = this.fieldsByName[name];
    return this.settings(name).readOnly || (!!field && field.readOnly) || this.overrides.disabled[name] === true;
  }

  private resolve(template: string): string {
    return template.replace(/\{([A-Za-z0-9_]+)\}/g, (_, n: string) => valueToText(this.values[n])).trim();
  }

  private title(): string {
    if (this.titleOverride !== undefined) {
      return this.titleOverride;
    }
    return this.resolve(this.props.definition.titles[this.props.mode] || '') || FALLBACK_TITLES[this.props.mode];
  }

  private subtitle(): string | undefined {
    const h = this.props.definition.header;
    return h && h.subtitle ? this.resolve(h.subtitle) || undefined : undefined;
  }

  private errorCount(tabId: string): number {
    let n = 0;
    this.props.definition.sections.forEach((s) => {
      if (s.tabId === tabId) {
        s.fields.forEach((f) => {
          if (this.errors[f]) {
            n++;
          }
        });
      }
    });
    return n;
  }

  private dismissMessage(i: number): void {
    this.messages = this.messages.filter((_, idx) => idx !== i);
    this.refresh();
  }

  // ---------- Validierung ----------

  private validate(name: string): string {
    const field = this.fieldsByName[name];
    if (!field || !this.isVisible(name)) {
      return '';
    }
    const value = this.values[name];
    return validateFieldValue(field, this.settings(name), value, this.isRequired(name)) || this.host.validate(name, value);
  }

  /** Prüft das geänderte Feld und alle Felder, die bereits einen Fehler anzeigen. */
  private revalidate(name: string): void {
    const names = Object.keys(this.errors);
    if (names.indexOf(name) < 0) {
      names.push(name);
    }
    names.forEach((n) => {
      const msg = this.validate(n);
      if (msg) {
        this.errors[n] = msg;
      } else {
        delete this.errors[n];
      }
    });
  }

  private handleChange(name: string, value: FieldValue): void {
    this.values[name] = value;
    this.host.fireChange(name, value);
    this.revalidate(name);
    this.refresh();
  }

  private handleBlur(name: string): void {
    this.revalidate(name);
    this.refresh();
  }

  private async submit(): Promise<void> {
    const { definition } = this.props;
    this.messages = [];
    this.errors = {};
    let firstTab: string | undefined;
    definition.sections.forEach((s) =>
      s.fields.forEach((name) => {
        const msg = this.validate(name);
        if (msg) {
          this.errors[name] = msg;
          firstTab = firstTab || s.tabId;
        }
      })
    );
    const count = Object.keys(this.errors).length;
    if (count) {
      this.messages.push({ kind: 'error', text: count === 1 ? 'Ein Feld ist ungültig. Bitte die Markierung prüfen.' : `${count} Felder sind ungültig. Bitte die Markierungen prüfen.` });
      if (firstTab && this.errorCount(this.state.activeTab) === 0) {
        this.setState({ activeTab: firstTab });
      }
      this.refresh();
      return;
    }

    const stop = await this.host.runBeforeSave();
    if (stop) {
      this.messages.push({ kind: 'error', text: stop });
      this.refresh();
      return;
    }

    this.setState({ saving: true });
    try {
      await this.props.onSubmit({ ...this.values });
      if (this.props.preview) {
        this.messages.push({ kind: 'success', text: 'Gespeichert (Vorschau – es wurde nichts geschrieben).' });
      }
    } catch (e) {
      this.messages.push({ kind: 'error', text: `Speichern fehlgeschlagen: ${(e as Error).message}` });
    } finally {
      this.setState({ saving: false });
      this.refresh();
    }
  }

  // ---------- Brücke zum Formular-Skript ----------

  private createBridge(): IScriptHostBridge {
    return {
      mode: this.props.mode,
      currentUser: this.props.currentUser,
      sp: this.props.sp,
      hasField: (name) => !!this.fieldsByName[name],
      fieldKind: (name) => (this.fieldsByName[name] ? this.fieldsByName[name].kind : undefined),
      getValue: (name) => this.values[name],
      setValue: (name, value) => {
        this.values[name] = value;
        this.refresh();
      },
      getFlag: (name, flag) => {
        if (flag === 'visible') {
          return this.isVisible(name);
        }
        return flag === 'required' ? this.isRequired(name) : this.isDisabled(name);
      },
      setFlag: (name, flag, value) => {
        this.overrides[flag][name] = value;
        if (flag === 'visible' && !value) {
          delete this.errors[name];
        }
        this.refresh();
      },
      getTitle: () => this.title(),
      setTitle: (t) => {
        this.titleOverride = t;
        this.refresh();
      },
      showMessage: (text, kind) => {
        this.messages.push({ text, kind });
        this.refresh();
      },
      reportError: (text) => {
        this.scriptErrors.push(text);
        console.error(`[Form Designer] ${text}`);
        this.refresh();
      }
    };
  }
}
