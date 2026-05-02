import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PactAppWebPartStrings';
import PactApp from './components/PactApp';
import { IPactAppProps } from './components/IPactAppProps';

export interface IPactAppWebPartProps {
  description: string;
}

export default class PactAppWebPart extends BaseClientSideWebPart<IPactAppWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _workbenchHost: HTMLDivElement | null = null;

  public render(): void {
    let isWorkbench = false;
    try {
      isWorkbench = typeof window !== 'undefined' && /workbench/i.test(window.location.href);
    } catch {
      isWorkbench = false;
    }

    let renderTarget: HTMLDivElement | HTMLElement = this.domElement;
    if (isWorkbench) {
      try {
        renderTarget = this._ensureWorkbenchHost();
      } catch (error) {
        // If the overlay cannot be created, fall back to the normal web part container.
        // The app should still render rather than fail the whole bundle.
        // eslint-disable-next-line no-console
        console.error('PACT workbench overlay failed, falling back to normal mount', error);
        renderTarget = this.domElement;
      }
    }

    if (!renderTarget) {
      return;
    }

    if (!isWorkbench && this._workbenchHost) {
      // eslint-disable-next-line @rushstack/pair-react-dom-render-unmount
      ReactDom.unmountComponentAtNode(this._workbenchHost);
      this._workbenchHost.remove();
      this._workbenchHost = null;
    }

    this.domElement.style.display = 'block';
    this.domElement.style.maxWidth = 'none';
    this.domElement.style.position = 'relative';
    this.domElement.style.width = '100%';
    this.domElement.style.minWidth = '1000px';
    this.domElement.style.height = 'auto';
    this.domElement.style.overflowX = 'auto';
    this.domElement.style.overflowY = 'hidden';

    const element: React.ReactElement<IPactAppProps> = React.createElement(
      PactApp,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );

    // eslint-disable-next-line @rushstack/pair-react-dom-render-unmount
    ReactDom.unmountComponentAtNode(renderTarget);
    // eslint-disable-next-line @rushstack/pair-react-dom-render-unmount
    ReactDom.render(element, renderTarget);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || '');
      this.domElement.style.setProperty('--link', semanticColors.link || '');
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || '');
    }

  }

  protected onDispose(): void {
    if (this._workbenchHost) {
      // eslint-disable-next-line @rushstack/pair-react-dom-render-unmount
      ReactDom.unmountComponentAtNode(this._workbenchHost);
      this._workbenchHost.remove();
      this._workbenchHost = null;
    }
    // eslint-disable-next-line @rushstack/pair-react-dom-render-unmount
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  private _ensureWorkbenchHost(): HTMLDivElement {
    if (this._workbenchHost) {
      return this._workbenchHost;
    }

    if (typeof document === 'undefined' || !document.body) {
      throw new Error('Workbench host cannot be created before document.body is available');
    }

    const host = document.createElement('div');
    host.id = 'pact-workbench-overlay';
    host.className = 'pact-workbench-overlay glass-panel';
    host.style.position = 'fixed';
    host.style.top = '40px';
    host.style.left = '0';
    host.style.right = '360px';
    host.style.bottom = '0';
    host.style.width = 'auto';
    host.style.height = 'auto';
    host.style.minWidth = '1280px';
    host.style.maxWidth = 'none';
    host.style.zIndex = '2000';
    host.style.background = 'var(--bg-base)';
    host.style.overflow = 'auto';
    host.style.boxShadow = 'none';
    host.style.border = 'none';
    host.style.borderRadius = '0';
    host.style.margin = '0';
    host.style.padding = '16px';

    document.body.appendChild(host);
    this._workbenchHost = host;
    return host;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
