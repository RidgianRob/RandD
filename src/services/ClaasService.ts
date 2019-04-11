import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';

import { IIntroSourceListItem } from '../models';
import SphttpclientWebPart from '../../lib/webparts/sphttpclient/SphttpclientWebPart';

const LIST_API_ENDPOINT: string = `/_api/web/lists/getbytitle('Claas Intro Source')`;
const SELECT_QUERY: string = `$select=Id,Title,introSource,introSourceJV,dealer,subArea,reviewDate`;
const QUERY_ORDER_MAX: string = `&$orderby=Id desc&$top=1`;

export class ClaasService {

  private _spHttpOptions: any = {
    getNoMetadata: <ISPHttpClientOptions> {
      headers: { 'ACCEPT': 'application/json; odata.metadata=none' }
    },
    getFullMetadata: <ISPHttpClientOptions> {
      headers: { 'ACCEPT': 'application/json; odata.metadata=full' }
    },
    postNoMetadata: <ISPHttpClientOptions> {
      headers: {
        'ACCEPT': 'application/json; odata.metadata=none',
        'CONTENT-TYPE' : 'application/json' 
      }
    },
    updateNoMetadata: <ISPHttpClientOptions> {
      headers: {
        'ACCEPT': 'application/json; odata.metadata=none',
        'CONTENT-TYPE': 'application/json',
        'X-HTTP-Method': 'MERGE'
      }
    },
    deleteNoMetadata: <ISPHttpClientOptions> {
      headers: {
        'ACCEPT': 'application/json; odata.metadata=none',
        'X-HTTP-Method': 'DELETE'
      }
    }
  };

  constructor(private siteAbsUrl: string, private client: SPHttpClient) {}

  public getClaasIntroSources(): Promise<IIntroSourceListItem[]> {
    let promise: Promise<IIntroSourceListItem[]> = new Promise<IIntroSourceListItem[]>((resolve, reject) => {
      let query: string = `${this.siteAbsUrl}${LIST_API_ENDPOINT}/items?${SELECT_QUERY}`;

      this.client.get(
        query,
        SPHttpClient.configurations.v1,
        this._spHttpOptions.getNoMetadata
      )
        .then((response: SPHttpClientResponse): Promise<{value: IIntroSourceListItem[]}> => {
          return response.json();
        })
        .then((response: { value: IIntroSourceListItem[] }) => {
          resolve(response.value);
        })
        .catch((error: any) => {
          reject(error);
        });
    });

    return promise;
  }

  public getClaasIntroSource(introSourceId: number): Promise<IIntroSourceListItem> {
    let promise: Promise<IIntroSourceListItem> = new Promise<IIntroSourceListItem>((resolve, reject) => {
      let query: string = `${this.siteAbsUrl}${LIST_API_ENDPOINT}/items(${introSourceId})?${SELECT_QUERY}`;
      this.client.get(
        query,
        SPHttpClient.configurations.v1,
        this._spHttpOptions.getFullMetadata
      )
        .then((response: SPHttpClientResponse): Promise<IIntroSourceListItem> => {
          return response.json();
        })
        .then((response: IIntroSourceListItem) => {
          resolve(response);
        })
        .catch((error: any) => {
          reject(error);
        });
    });

    return promise;
  }

  public getLastClaasIntroSource(): Promise<IIntroSourceListItem> {
    let promise: Promise<IIntroSourceListItem> = new Promise<IIntroSourceListItem>((resolve, reject) => {
      let query: string = `${this.siteAbsUrl}${LIST_API_ENDPOINT}/items?${SELECT_QUERY}${QUERY_ORDER_MAX}`;
      this.client.get(
        query,
        SPHttpClient.configurations.v1,
        this._spHttpOptions.getFullMetadata
      )
        .then((response: SPHttpClientResponse): Promise<any> => {
          return response.json();
        })
        .then((response: any) => {
          resolve(response.value[0]);
        })
        .catch((error: any) => {
          reject(error);
        });
    });

    return promise;
  }

  public _getItemEntityType(): Promise<string> {
    let promise: Promise<string> = new Promise<string>((resolve, reject) => {
      this.client.get(`${this.siteAbsUrl}${LIST_API_ENDPOINT}?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
        this._spHttpOptions.getNoMetadata
      )
        .then((response: SPHttpClientResponse): Promise<{ ListItemEntityTypeFullName: string}> => {
          return response.json();
        })
        .then((response: { ListItemEntityTypeFullName: string }): void => {
          resolve(response.ListItemEntityTypeFullName);
        })
        .catch((error: any) => {
          reject(error);
        });
    });

    return promise;
  }

  public createClaasIntroSource(newClassIntroSource: IIntroSourceListItem): Promise<void> {
    let promise: Promise<void> = new Promise<void>((resolve, reject) => {
      this._getItemEntityType()
        .then((spEntityType: string) => {
          // create list item
          let newListItem: IIntroSourceListItem = newClassIntroSource;

          // add SP required metadata
          newListItem['@odata.type'] = spEntityType;

          // build request
          let requestDetails: any = this._spHttpOptions.postNoMetadata;
          requestDetails.body = JSON.stringify(newListItem);

          // create item
          return this.client.post(`${this.siteAbsUrl}${LIST_API_ENDPOINT}/items`,
            SPHttpClient.configurations.v1,
            requestDetails
          );
        })
        .then((response: SPHttpClientResponse): Promise<IIntroSourceListItem> => {
          return response.json();
        })
        .then((newSPListItem: IIntroSourceListItem): void => {
          resolve();
        })
        .catch((error: any) => {
          reject(error);
        });        
    });

    return promise;
  }

  public updateClaasIntroSource(introSourceToUpdate: IIntroSourceListItem): Promise<void> {
    let promise: Promise<void> = new Promise<void>((resolve, reject) => {
      let requestDetails: any = this._spHttpOptions.updateNoMetadata;

      requestDetails.headers['IF-MATCH'] = introSourceToUpdate['@odata.etag'];
      requestDetails.body = JSON.stringify(introSourceToUpdate);

      this.client.post(`${this.siteAbsUrl}${LIST_API_ENDPOINT}/items(${introSourceToUpdate.Id})`,
        SPHttpClient.configurations.v1,
        requestDetails
      )
      .then(() => {
        resolve();
      })
    });

    return promise;
  }

  public deleteClaasIntroSource(introSourceToDelete: IIntroSourceListItem): Promise<void> {
    let promise: Promise<void> = new Promise<void>((resolve, reject) => {
      let requestDetails: any = this._spHttpOptions.deleteNoMetadata;

      // Check to make sure we're updating the latest version
      requestDetails.headers['IF-MATCH'] = introSourceToDelete['@odata.etag'];
      requestDetails.body = JSON.stringify(introSourceToDelete);

      this.client.post(`${this.siteAbsUrl}${LIST_API_ENDPOINT}/items(${introSourceToDelete.Id})`,
        SPHttpClient.configurations.v1,
        requestDetails
      )
      .then(() => {
        resolve();
      })
    });

    return promise;
  }

}