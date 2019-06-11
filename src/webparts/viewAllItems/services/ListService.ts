import { IListService} from '../services/IListService';
import { IList} from '../services/IList';
import { IListColumn} from '../services/IListColumn';
import {
  SPHttpClient,
  ISPHttpClientOptions,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import {
  IWebPartContext
} from '@microsoft/sp-webpart-base';

export class ListService implements IListService {

    constructor(private context: IWebPartContext) {
    }

    public getLists(): Promise<IList[]> {
        var httpClientOptions : ISPHttpClientOptions = {};
    
        httpClientOptions.headers = {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
        };
    
        return new Promise<IList[]>((resolve: (results: IList[]) => void, reject: (error: any) => void): void => {
          this.context.spHttpClient.get(this.context.pageContext.web.serverRelativeUrl + `/_api/web/lists?$select=id,title&$filter=Hidden eq false`,
            SPHttpClient.configurations.v1,
            httpClientOptions
            )
            .then((response: SPHttpClientResponse): Promise<{ value: IList[] }> => {
              return response.json();
            })
            .then((lists: { value: IList[] }): void => {
              resolve(lists.value);
            }, (error: any): void => {
              reject(error);
            });
        });
    }

    public getColumns(listName: string): Promise<IListColumn[]> {
      var httpClientOptions : ISPHttpClientOptions = {};
  
      httpClientOptions.headers = {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
      };
  
      return new Promise<IListColumn[]>((resolve: (results: IListColumn[]) => void, reject: (error: any) => void): void => {
        this.context.spHttpClient.get(this.context.pageContext.web.serverRelativeUrl + `/_api/web/lists/GetByTitle('${listName}')/fields?$filter=TypeDisplayName ne 'Attachments' and Hidden eq false and ReadOnlyField eq false`,
          SPHttpClient.configurations.v1,
          httpClientOptions
          )
          .then((response: SPHttpClientResponse): Promise<{ value: IListColumn[] }> => {
            return response.json();
          })
          .then((listColumns: { value: IListColumn[] }): void => {
            resolve(listColumns.value);
          }, (error: any): void => {
            reject(error);
          });
      });
    }
      
}