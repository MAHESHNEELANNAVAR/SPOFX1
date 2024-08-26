import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SharePointOnlineFrameworkSoln1WebPart.module.scss';
import * as strings from 'SharePointOnlineFrameworkSoln1WebPartStrings';

export interface IsharePointOnlineFrameworkSoln1WebPartProps {
  description: string;
}

export default class sharePointOnlineFrameworkSoln1WebPart extends BaseClientSideWebPart<IsharePointOnlineFrameworkSoln1WebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <h1>welcome to https://www.maheshneelannavar.pro</h1>
      <section class="${styles.sharePointOnlineFrameworkSoln1} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
        <div class="${styles.welcome}">
          <p class="${styles.welcome}">Title: ${escape(this.context.pageContext.web.title)}</p>
           <p class="${styles.welcome}">Display Name: ${escape(this.context.pageContext.user.displayName)}</p>
        </div>
      </section>`;
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(() => {});
  }

  private _getEnvironmentMessage(): Promise<void> {
    if (!!this.context.sdks.microsoftTeams) { // Check if running in Teams, Office.com, or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': 
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': 
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        })
        .then(() => {});
    }

    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || '');
      this.domElement.style.setProperty('--link', semanticColors.link || '');
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || '');
    }
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
