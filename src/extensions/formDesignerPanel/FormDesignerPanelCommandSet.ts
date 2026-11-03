import { Log } from '@microsoft/sp-core-library';
import { BaseListViewCommandSet, type Command, type IListViewCommandSetExecuteEventParameters } from '@microsoft/sp-listview-extensibility';

import { RUNTIME_COMPONENT_ID } from '../../common/constants';
import { displayModeOf, IFormDefinition, IRuntimeProperties, normalizeDefinition } from '../../common/model';
import { IFormRequest } from '../../common/formUrl';
import { fdLog, getDebugMode, registerList, unregisterList } from '../../common/interceptor';
import { panelCssWidth } from '../../common/panelSize';
import { readUrlParams } from '../../common/urlPrefill';
import { FormMode } from '../../common/model';
import { IContentTypeInfo, SpService } from '../../common/spService';
import { ICurrentUser } from '../../common/spTypes';
import { IListFormConfig, isOverlayOpen, openFormOverlay } from '../../common/components/FormOverlay';
import { setCulture } from '../../common/i18n';

export interface IFormDesignerPanelCommandSetProperties {}

const LOG_SOURCE: string = 'FormDesignerPanelCommandSet';

function isOurs(id: string | null): boolean {
  return !!id && id.replace(/[{}]/g, '').toLowerCase() === RUNTIME_COMPONENT_ID;
}

function definitionUrlOf(ct: IContentTypeInfo): string | undefined {
  const pairs: Array<[string | null, string | null]> = [
    [ct.NewFormClientSideComponentId, ct.NewFormClientSideComponentProperties],
    [ct.EditFormClientSideComponentId, ct.EditFormClientSideComponentProperties],
    [ct.DisplayFormClientSideComponentId, ct.DisplayFormClientSideComponentProperties]
  ];
  for (const [id, props] of pairs) {
    if (isOurs(id) && props) {
      try {
        const parsed = JSON.parse(props) as IRuntimeProperties;
        if (parsed.definitionUrl) {
          return parsed.definitionUrl;
        }
      } catch {
        // ignorieren
      }
    }
  }
  return undefined;
}

/**
 * Öffnet Form-Designer-Formulare aus der Listenansicht im Panel oder Dialog.
 * Wird vom Designer beim Bereitstellen auf der Liste registriert.
 */
export default class FormDesignerPanelCommandSet extends BaseListViewCommandSet<IFormDesignerPanelCommandSetProperties> {
  private _sp!: SpService;
  private _configs: IListFormConfig[] = [];
  private _ready: boolean = false;
  private _listId: string = '';

  public onInit(): Promise<void> {
    const ctx = this.context.pageContext;
    setCulture(ctx.cultureInfo.currentUICultureName);
    const list = ctx.list;
    if (!list) {
      return Promise.resolve();
    }
    this._listId = list.id.toString().toLowerCase();
    this._sp = new SpService(this.context.spHttpClient, ctx.web.absoluteUrl, ctx.web.serverRelativeUrl);
    registerList(this._listId, list.serverRelativeUrl, (req, via) => this.handle(req, via));

    const diag: Command | undefined = this.tryGetCommand('FD_DIAG');
    if (diag) {
      diag.visible = getDebugMode() !== 'off';
    }

    // Konfiguration im Hintergrund laden, damit die Befehlsleiste nicht wartet.
    this.loadConfigs()
      .then(() => {
        this._ready = true;
        this.openFromPageUrl();
        fdLog(`Bereit: ${this._configs.length} Inhaltstyp(en) mit Form-Designer-Formular`, this._configs.map((c) => c.contentTypeId));
      })
      .catch((e: Error) => {
        Log.error(LOG_SOURCE, e);
        fdLog('Konfiguration konnte nicht geladen werden – Formulare öffnen als ganze Seite', e.message);
      });
    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId === 'FD_DIAG') {
      fdLog('Diagnose', {
        liste: this._listId,
        bereit: this._ready,
        modus: getDebugMode(),
        navigationApi: 'navigation' in window,
        formulare: this._configs.map((c) => ({
          inhaltstyp: c.contentTypeId,
          darstellung: {
            neu: displayModeOf(c.definition.layout, 'new'),
            bearbeiten: displayModeOf(c.definition.layout, 'edit'),
            anzeigen: displayModeOf(c.definition.layout, 'display')
          },
          breite: panelCssWidth(c.definition.layout) || 'Standard'
        }))
      });
    }
  }

  public onDispose(): void {
    unregisterList(this._listId);
    super.onDispose();
  }

  private async loadConfigs(): Promise<void> {
    const cts = await this._sp.getContentTypes(this._listId);
    const configs: IListFormConfig[] = [];
    for (const ct of cts) {
      const url = definitionUrlOf(ct);
      if (!url) {
        continue;
      }
      const def = await this._sp.readJson<IFormDefinition>(url);
      if (def) {
        configs.push({ contentTypeId: ct.id, definition: normalizeDefinition(def) });
      }
    }
    this._configs = configs;
  }

  /** Öffnet ein Formular, wenn die Listenansicht mit ?fdOpen=new|edit|display (&fdId=) aufgerufen wurde. */
  private openFromPageUrl(): void {
    const params = readUrlParams(window.location.href);
    const open = (params.fdOpen || params.fdopen || '').toLowerCase();
    const modes: Record<string, FormMode> = { new: 'new', edit: 'edit', display: 'display' };
    const mode = modes[open];
    if (!mode) {
      return;
    }
    const id = Number(params.fdId || params.fdid || 0);
    // Parameter entfernen, damit das Formular beim Neuladen nicht erneut aufgeht.
    try {
      const clean = new URL(window.location.href);
      ['fdOpen', 'fdopen', 'fdId', 'fdid'].forEach((k) => clean.searchParams.delete(k));
      Object.keys(params)
        .filter((k) => k.toLowerCase().indexOf('fd_') === 0)
        .forEach((k) => clean.searchParams.delete(k));
      window.history.replaceState(window.history.state, '', clean.href);
    } catch {
      // ignorieren
    }
    if (mode !== 'new' && !id) {
      return;
    }
    this.handle({ mode, url: window.location.href, listId: this._listId, itemId: id || undefined }, 'fdOpen', params);
  }

  private handle(req: IFormRequest, via: string, extraParams?: Record<string, string>): boolean {
    if (!this._ready || !this._configs.length) {
      return false;
    }
    if (isOverlayOpen()) {
      // Ein Formular ist bereits offen: weitere Navigationen zu Formularen nicht zulassen.
      return true;
    }
    let candidates = this._configs;
    if (req.contentTypeId) {
      const ct = req.contentTypeId.toLowerCase();
      candidates = this._configs.filter((c) => c.contentTypeId.toLowerCase().indexOf(ct) === 0 || ct.indexOf(c.contentTypeId.toLowerCase()) === 0);
      if (!candidates.length) {
        return false;
      }
    }
    const layout = candidates[0].definition.layout;
    const mode = displayModeOf(layout, req.mode);
    // Ein expliziter Aufruf per fdOpen öffnet auch dann ein Panel, wenn «Ganze Seite» eingestellt ist.
    if (mode === 'page' && via !== 'fdOpen') {
      return false;
    }
    openFormOverlay({
      sp: this._sp,
      listId: this._listId,
      request: req,
      configs: this._configs,
      currentUser: this.currentUser(),
      displayMode: mode === 'page' ? 'panel' : mode,
      cssWidth: panelCssWidth(layout),
      urlParams: { ...readUrlParams(req.url), ...(extraParams || {}) }
    });
    fdLog(`Formular «${req.mode}» als ${mode === 'dialog' ? 'Dialog' : 'Panel'} geöffnet (${via})`);
    return true;
  }

  private currentUser(): ICurrentUser {
    const ctx = this.context.pageContext;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const legacy = (ctx.legacyPageContext || {}) as any;
    return { id: Number(legacy.userId) || 0, title: ctx.user.displayName, email: ctx.user.email, loginName: ctx.user.loginName };
  }
}
