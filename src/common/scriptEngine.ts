import { FormMode } from './model';
import { FieldValue, ICurrentUser } from './spTypes';
import { scriptDateHelpers, toIsoDate } from './dateUtils';
import { SpService } from './spService';
import { t } from './i18n';

export type MessageKind = 'info' | 'warning' | 'error' | 'success';

/** Schnittstelle zwischen Skript-Engine und Formular. */
export interface IScriptHostBridge {
  mode: FormMode;
  currentUser?: ICurrentUser;
  sp?: SpService;
  hasField(name: string): boolean;
  fieldKind(name: string): string | undefined;
  getValue(name: string): FieldValue;
  setValue(name: string, value: FieldValue): void;
  getFlag(name: string, flag: 'visible' | 'required' | 'disabled'): boolean;
  setFlag(name: string, flag: 'visible' | 'required' | 'disabled', value: boolean): void;
  getTitle(): string;
  setTitle(title: string): void;
  showMessage(text: string, kind: MessageKind): void;
  reportError(text: string): void;
}

type ChangeHandler = (value: FieldValue) => void;
type Validator = (value: FieldValue) => unknown;
type SaveHandler = () => unknown;

/**
 * Führt das Formular-Skript aus und verwaltet die registrierten Ereignisse.
 * Das Skript erhält ausschliesslich das Objekt «form» und greift nie direkt aufs DOM zu.
 */
export class FormScriptHost {
  private readonly ready: Array<() => void> = [];
  private readonly change: Record<string, ChangeHandler[]> = {};
  private readonly validators: Record<string, Validator[]> = {};
  private readonly beforeSave: SaveHandler[] = [];

  constructor(private readonly bridge: IScriptHostBridge) {}

  public run(script: string): void {
    if (!script || !script.trim()) {
      return;
    }
    let runner: (form: unknown) => void;
    try {
      // eslint-disable-next-line no-new-func
      runner = new Function('form', `"use strict";\n${script}`) as (form: unknown) => void;
    } catch (e) {
      this.bridge.reportError(t('Syntaxfehler im Formular-Skript: {0}', (e as Error).message));
      return;
    }
    this.safe(() => runner(this.createApi()));
  }

  public fireReady(): void {
    this.ready.forEach((fn) => this.safe(fn));
  }

  public fireChange(name: string, value: FieldValue): void {
    (this.change[name] || []).forEach((fn) => this.safe(() => fn(value)));
  }

  public validate(name: string, value: FieldValue): string {
    for (const fn of this.validators[name] || []) {
      const r = this.safe(() => fn(value));
      if (typeof r === 'string' && r) {
        return r;
      }
    }
    return '';
  }

  /** Führt alle onBeforeSave-Handler aus. Ein zurückgegebener Text bricht das Speichern ab. */
  public async runBeforeSave(): Promise<string> {
    for (const fn of this.beforeSave) {
      try {
        const r = await Promise.resolve(fn());
        if (typeof r === 'string' && r) {
          return r;
        }
      } catch (e) {
        this.bridge.reportError(t('Skriptfehler: {0}', (e as Error).message));
        return t('Das Formular-Skript hat das Speichern wegen eines Fehlers abgebrochen.');
      }
    }
    return '';
  }

  private safe<T>(fn: () => T): T | undefined {
    try {
      return fn();
    } catch (e) {
      this.bridge.reportError(t('Skriptfehler: {0}', (e as Error).message));
      return undefined;
    }
  }

  private createApi(): unknown {
    const bridge = this.bridge;
    const host = this;
    const field = (name: string): unknown => {
      if (!bridge.hasField(name)) {
        bridge.reportError(t("Formular-Skript: Das Feld '{0}' gibt es in diesem Formular nicht.", name));
      }
      return {
        get value(): FieldValue {
          return bridge.getValue(name);
        },
        set value(v: FieldValue) {
          let next: unknown = v;
          if (next instanceof Date) {
            const kind = bridge.fieldKind(name);
            next = kind === 'datetime' ? `${toIsoDate(next)}T00:00` : toIsoDate(next);
          }
          bridge.setValue(name, next as FieldValue);
        },
        get visible(): boolean {
          return bridge.getFlag(name, 'visible');
        },
        set visible(b: boolean) {
          bridge.setFlag(name, 'visible', !!b);
        },
        get required(): boolean {
          return bridge.getFlag(name, 'required');
        },
        set required(b: boolean) {
          bridge.setFlag(name, 'required', !!b);
        },
        get disabled(): boolean {
          return bridge.getFlag(name, 'disabled');
        },
        set disabled(b: boolean) {
          bridge.setFlag(name, 'disabled', !!b);
        },
        onChange(fn: ChangeHandler): void {
          (host.change[name] = host.change[name] || []).push(fn);
        },
        validate(fn: Validator): void {
          (host.validators[name] = host.validators[name] || []).push(fn);
        }
      };
    };

    return {
      field,
      type: bridge.mode,
      get title(): string {
        return bridge.getTitle();
      },
      set title(t: string) {
        bridge.setTitle(String(t));
      },
      currentUser: bridge.currentUser
        ? { id: bridge.currentUser.id, displayName: bridge.currentUser.title, email: bridge.currentUser.email, loginName: bridge.currentUser.loginName }
        : undefined,
      date: scriptDateHelpers,
      sp: bridge.sp
        ? {
            webUrl: bridge.sp.webUrl,
            get: (path: string) => bridge.sp!.get(path),
            post: (path: string, body?: unknown) => bridge.sp!.post(path, body)
          }
        : undefined,
      onReady(fn: () => void): void {
        host.ready.push(fn);
      },
      onBeforeSave(fn: SaveHandler): void {
        host.beforeSave.push(fn);
      },
      showMessage(text: string, kind?: MessageKind): void {
        bridge.showMessage(String(text), kind && ['info', 'warning', 'error', 'success'].indexOf(kind) >= 0 ? kind : 'info');
      }
    };
  }
}

/** Prüft nur die Syntax eines Skripts, ohne es auszuführen. */
export function checkScriptSyntax(script: string): string {
  try {
    // eslint-disable-next-line no-new-func
    const compiled = new Function('form', `"use strict";\n${script}`);
    return typeof compiled === 'function' ? '' : 'Unbekannter Fehler';
  } catch (e) {
    return (e as Error).message;
  }
}
