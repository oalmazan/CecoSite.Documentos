import { ServiceScope } from "@microsoft/sp-core-library";  
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';  
import { PageContext } from '@microsoft/sp-page-context';  
 
export class BaseService<T> {
    protected _spHttpClient: SPHttpClient;
    protected _pageContext: PageContext;  
    protected _currentWebUrl: string;  

    constructor(serviceScope: ServiceScope) {  
        serviceScope.whenFinished(() => {  
            this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
            this._pageContext = serviceScope.consume(PageContext.serviceKey);  
            this._currentWebUrl = this._pageContext.web.absoluteUrl;  
        });  
    }

    public get(url: string): Promise<T> {
        return new Promise<T>((resolve: (itemId: T) => void, reject: (error: any) => void): void => {  
            this.read(url)  
              .then((items: T): void => {  
                resolve(this.process(items));  
              });  
          });
    }

    private read(url: string): Promise<T> {  
        return new Promise<T>((resolve: (itemId: T) => void, reject: (error: any) => void): void => {  
          this._spHttpClient.get(url,  
          SPHttpClient.configurations.v1,  
          {  
            headers: {  
              'Accept': 'application/json;odata=nometadata',  
              'odata-version': ''  
            }  
          })  
          .then((response: SPHttpClientResponse): Promise<{ value: T }> => {  
            return response.json();  
          })  
          .then((response: { value: T }): void => {  
            var output: any = JSON.stringify(response);  
            resolve(output); 
          }, (error: any): void => {  
            reject(error);  
          });  
        });      
    }  

    private process(objStr: any): any { 
        return JSON.parse(objStr);  
    }
}