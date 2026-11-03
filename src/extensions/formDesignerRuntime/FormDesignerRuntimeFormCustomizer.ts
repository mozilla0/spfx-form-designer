import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { FormDisplayMode, Log } from '@microsoft/sp-core-library';
import { BaseFormCustomizer } from '@microsoft/sp-listview-extensibility';

import { FormMode, IFormDefinition, IRuntimeProperties, normalizeDefinition } from '../../common/model';
import { SpService } from '../../common/spService';
import { FieldValues, ICurrentUser, ISpField } from '../../common/spTypes';
import { applyPanelWidth, panelCssWidth } from '../../common/panelSize';
import FormDesignerRuntime, { IFormDesignerRuntimeProps } from './components/FormDesignerRuntime';

const LOG_SOURCE: string = 'FormDesignerRuntimeFormCustomizer';

/**
 * Runtime des Form Designers.
 * Liest die Formulardefinition (Pfad in den ClientSideComponentProperties des Inhaltstyps),
 * lädt Spalten und Element und rendert das Formular.
 */
export default class FormDesignerRuntimeFormCustomizer extends BaseFormCustomizer<IRuntimeProperties> {
  private _sp!: SpService;
  private _definition: IFormDefinition | undefined;
  private _fields: ISpField[] = [];
  private _values: FieldValues = {};
  private _loadError: string = '';
  /** Wechsel von «Anzeigen» zu «Bearbeiten» im selben Formular (Seite oder Panel). */
  private _editRequested: boolean = false;
  private _resetPanelWidth: (() => void) | undefined;

  public async onInit(): Promise<void> {
    const web = this.context.pageContext.web;
    this._sp = new SpService(this.context.spHttpClient, web.absoluteUrl, web.serverRelativeUrl);
    try {
      if (!this.properties || !this.properties.definitionUrl) {
        throw new Error('Für dieses Formular ist keine Formulardefinition hinterlegt. Bitte im Form Designer erneut bereitstellen.');
      }
      const listId = this.context.list.guid.toString();
      const [definition, fields] = await Promise.all([
        this._sp.readJson<IFormDefinition>(this.properties.definitionUrl),
        this._sp.getFields(listId, this.context.contentType.id)
      ]);
      if (!definition) {
        throw new Error(`Die Formulardefinition wurde nicht gefunden: ${this.properties.definitionUrl}`);
      }
      this._definition = normalizeDefinition(definition);
      this._fields = fields;
      if (this.context.itemId) {
        this._values = await this._sp.getItem(listId, this.context.itemId, fields);
      }
    } catch (e) {
      this._loadError = (e as Error).message;
      Log.error(LOG_SOURCE, e as Error);
    }
  }

  public render(): void {
    const element: React.ReactElement<IFormDesignerRuntimeProps> = React.createElement(FormDesignerRuntime, {
      definition: this._definition,
      fields: this._fields,
      values: this._values,
      mode: this.mode(),
      loadError: this._loadError,
      currentUser: this.currentUser(),
      sp: this._sp,
      onSubmit: (values: FieldValues) => this.save(values),
      onClose: () => this.formClosed(),
      onRequestEdit: this.context.itemId ? () => this.switchToEdit() : undefined
    });
    ReactDOM.render(element, this.domElement);

    // Experimentell: Panelbreite gemäss Formulardefinition (nur im Listen-Panel wirksam).
    if (!this._resetPanelWidth && this._definition) {
      this._resetPanelWidth = applyPanelWidth(this.domElement, panelCssWidth(this._definition.layout));
    }
  }

  public onDispose(): void {
    if (this._resetPanelWidth) {
      this._resetPanelWidth();
      this._resetPanelWidth = undefined;
    }
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  private mode(): FormMode {
    if (this._editRequested) {
      return 'edit';
    }
    switch (this.displayMode) {
      case FormDisplayMode.New:
        return 'new';
      case FormDisplayMode.Edit:
        return 'edit';
      default:
        return 'display';
    }
  }

  private currentUser(): ICurrentUser {
    const ctx = this.context.pageContext;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const legacy = (ctx.legacyPageContext || {}) as any;
    return {
      id: Number(legacy.userId) || 0,
      title: ctx.user.displayName,
      email: ctx.user.email,
      loginName: ctx.user.loginName
    };
  }

  private async save(values: FieldValues): Promise<void> {
    await this._sp.saveItem(this.context.list.guid.toString(), this.context.itemId, this._fields, values);
    this.formSaved();
  }

  /** Bleibt im selben Formular – funktioniert damit auch im Listen-Panel. */
  private switchToEdit(): void {
    this._editRequested = true;
    this.render();
  }
}
