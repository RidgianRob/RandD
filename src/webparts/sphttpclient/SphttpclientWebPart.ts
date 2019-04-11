import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SphttpclientWebPart.module.scss';
import * as strings from 'SphttpclientWebPartStrings';

// Interfaces (models)
import { IIntroSourceListItem } from '../../models';

// Services
import { ClaasService } from '../../services';

export interface ISphttpclientWebPartProps {
  description: string;
}

export default class SphttpclientWebPart extends BaseClientSideWebPart<ISphttpclientWebPartProps> {

  private claasService: ClaasService;

  private claasIntroSourceDetailElement: HTMLElement;

  protected onInit(): Promise<any> {
    this.claasService = new ClaasService (
      this.context.pageContext.web.absoluteUrl,
      this.context.spHttpClient
    );

    return Promise.resolve();
  }

  public render(): void {
    if (!this.renderedOnce) {
      this.domElement.innerHTML = `
        <div class="${ styles.sphttpclient }">
          <div class="${ styles.container }">
            <div class="${ styles.row }">
              <div class="${ styles.column }">
                <span class="${ styles.title }">Claas Intro Source List</span>
                <p class="${ styles.subTitle }">Demonstrating SharePoint HTTP Client.</p>
                <button id="getClaasIntroSources" class="${styles.button}">Get Claas Intro Sources</button>
                <button id="getClaasIntroSource" class="${styles.button}">Get Claas Intro Source</button>
                <button id="getLastClaasIntroSource" class="${styles.button}">Get Last Claas Intro Source</button>
                <button id="createClaasIntroSource" class="${styles.button}">Create Claas Intro Source</button>
                <button id="updateClaasIntroSource" class="${styles.button}">Update Claas Intro Source</button>
                <button id="deleteClaasIntroSource" class="${styles.button}">Delete Claas Intro Source</button>
                <div id="claasIntroSources"></div>
              </div>
            </div>
          </div>
        </div>`;

        this.claasIntroSourceDetailElement = document.getElementById('claasIntroSources');

        document.getElementById('getClaasIntroSources')
          .addEventListener('click', () => {
            this._getClaasIntroSources();
          });

          document.getElementById('getClaasIntroSource')
            .addEventListener('click', () => {
              this._getClaasIntroSource();
          });

          document.getElementById('getLastClaasIntroSource')
            .addEventListener('click', () => {
              this._getLastClaasIntroSource();
          });

          document.getElementById('createClaasIntroSource')
            .addEventListener('click', () => {
              this._createClaasIntroSource();
          });          

          document.getElementById('updateClaasIntroSource')
            .addEventListener('click', () => {
              this._updateLastClaasIntroSource();
          }); 

          document.getElementById('deleteClaasIntroSource')
            .addEventListener('click', () => {
              this._deleteLastClaasIntroSource();
          }); 

    }
  }

  private _renderClaasIntroSources(element: HTMLElement, introSources: IIntroSourceListItem[]): void {
    let introSourceList: string = '';

    if (introSources && introSources.length && introSources.length > 0) {
      introSources.forEach((introSource: IIntroSourceListItem) => {
        introSourceList = introSourceList + `<tr>
          <td>${introSource.Id}</td>
          <td>${introSource.Title}</td>
          <td>${introSource.introSourceJV}</td>
          <td>${introSource.dealer}</td>
          <td>${introSource.introSource}</td>
          <td>${introSource.reviewDate}</td>
          <td>${introSource.subArea}</td>
        </tr>`;
      });
    }

    element.innerHTML = `<table border=1>
      <tr>
        <th>Id</th>
        <th>Title</th>
        <th>Intro Source JV</th>
        <th>Dealer</th>
        <th>Intro Source</th>
        <th>Review Date</th>
        <th>Sub Area</th>
        <tbody>${introSourceList}</tbody>
      </tr>
    </table>`;
  }

  private _getClaasIntroSources(): void {
    this.claasService.getClaasIntroSources()
      .then((introSources: IIntroSourceListItem[]) => {
        this._renderClaasIntroSources(this.claasIntroSourceDetailElement, introSources);
      });

  }

  private _getClaasIntroSource(): void {
    this.claasService.getClaasIntroSource(57)
      .then((introSource: IIntroSourceListItem) => {
        this._renderClaasIntroSources(this.claasIntroSourceDetailElement, [introSource]);
      });

  }

  private _getLastClaasIntroSource(): void {
    this.claasService.getLastClaasIntroSource()
      .then((introSource: IIntroSourceListItem) => {
        this._renderClaasIntroSources(this.claasIntroSourceDetailElement, [introSource]);
      });

  }

  private _createClaasIntroSource(): void {
    const newClaasIntroSource: IIntroSourceListItem = <IIntroSourceListItem> {
      Title: 'Robs Intro Source',
      introSourceJV: 'Leasing Solutions',
      dealer: 'Robs Dealer',
      introSource: 'Robs Intro Source',
      subArea: 'Robs Sub Area'
    };

    this._renderClaasIntroSources(this.claasIntroSourceDetailElement, null);

    this.claasService.createClaasIntroSource(newClaasIntroSource)
      .then(() => {
        this._getClaasIntroSources();
      });

  }

  private _updateLastClaasIntroSource(): void {
    this.claasService.getLastClaasIntroSource()
      .then((introSource: IIntroSourceListItem) => {
        introSource.reviewDate = new Date().toISOString();
          return this.claasService.updateClaasIntroSource(introSource);
      });
  }

  private _deleteLastClaasIntroSource(): void {
    this.claasService.getLastClaasIntroSource()
      .then((introSource: IIntroSourceListItem) => {
          return this.claasService.deleteClaasIntroSource(introSource);
      })
      .then(() => {
        this._getClaasIntroSources();
      });
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
