import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpFxFetch.module.scss';
import * as strings from 'spFxFetchStrings';
import { ISpFxFetchWebPartProps } from './ISpFxFetchWebPartProps';

import { SPFxFetcher } from '../services/SPFxFetcher';
import { IODataList } from '@microsoft/sp-odata-types';

export default class SpFxFetchWebPart extends BaseClientSideWebPart<ISpFxFetchWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;

      (async () => {

        try {
          
          const _sPFXFetcher: SPFxFetcher = new SPFxFetcher(this.context.spHttpClient);
          
          const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;
          
          const _list: IODataList = await _sPFXFetcher.get(`${currentWebUrl}/_api/lists/GetByTitle('Documents1')`);
          
          console.log(_list);

        } catch (error) {
          console.log(error);
        }

      })();

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
