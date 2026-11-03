import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { SpService } from '../../common/spService';
import FormDesigner, { IFormDesignerProps } from './components/FormDesigner';
import { setCulture } from '../../common/i18n';

export interface IFormDesignerWebPartProps {}

/** Webpart mit dem Form Designer. Gehört auf eine Seite, auf die nur Website-Besitzer Zugriff haben. */
export default class FormDesignerWebPart extends BaseClientSideWebPart<IFormDesignerWebPartProps> {
  protected onInit(): Promise<void> {
    setCulture(this.context.pageContext.cultureInfo.currentUICultureName);
    return super.onInit();
  }

  public render(): void {
    const ctx = this.context.pageContext;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const legacy = (ctx.legacyPageContext || {}) as any;
    const element: React.ReactElement<IFormDesignerProps> = React.createElement(FormDesigner, {
      sp: new SpService(this.context.spHttpClient, ctx.web.absoluteUrl, ctx.web.serverRelativeUrl),
      currentUser: {
        id: Number(legacy.userId) || 0,
        title: ctx.user.displayName,
        email: ctx.user.email,
        loginName: ctx.user.loginName
      }
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [] };
  }
}
