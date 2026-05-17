import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
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
    // Inject CSS to force SharePoint Workbench and Canvas to be full-width
    const styleId = 'pact-full-bleed-styles';
    if (!document.getElementById(styleId)) {
      const style = document.createElement('style');
      style.id = styleId;
      style.innerHTML = `
        #workbenchPageContent,
        .SPCanvas,
        .Canvas,
        .CanvasZone,
        .CanvasSection,
        .CanvasSection-col,
        .CanvasComponent,
        .CanvasControl,
        .CanvasZoneContainer,
        .ControlZone,
        [data-automation-id='CanvasZone'],
        [data-automation-id='CanvasControl'],
        [data-automation-id='CanvasSection'],
        [data-automation-id='CanvasControlWebPart'],
        [data-sp-webpart],
        .pact-full-width-container {
          width: 100% !important;
          max-width: none !important;
          min-width: 0 !important;
          margin-left: 0 !important;
          margin-right: 0 !important;
          padding-left: 0 !important;
          padding-right: 0 !important;
        }

        .ControlZone,
        .CanvasSection-col,
        .pact-full-width-container {
          flex: 1 1 auto !important;
        }
      `;
      document.head.appendChild(style);
    }

    this.domElement.classList.add('pact-full-width-container');
    this.domElement.style.width = '100%';
    this.domElement.style.maxWidth = 'none';
    this.domElement.style.minWidth = '0';

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

    ReactDom.render(element, this.domElement);
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
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "PACT Platform Configuration"
          },
          groups: [
            {
              groupName: "Interface Settings",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "App Description"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
