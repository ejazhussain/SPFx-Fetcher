
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration, ODataVersion } from '@microsoft/sp-http';
import { ISPFxFetcher } from '../interfaces';

export class SPFxFetcher implements ISPFxFetcher {

    private _spHttpClient: SPHttpClient;

    constructor(spHttpClient: SPHttpClient) {
        this._spHttpClient = spHttpClient;
    }

    public get(endpoint: string, config: SPHttpClientConfiguration = SPHttpClient.configurations.v1): Promise<any> {
        
        if(endpoint.indexOf("/_api/search/") != -1){
            config = config.overrideWith({ defaultODataVersion: ODataVersion.v3 });
        }

        return this._spHttpClient.get(endpoint, config).then(this.processResponse);
    }

    private processResponse(response: SPHttpClientResponse): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            if (response.ok) {
                response.json().then((responseJSON) => {
                    resolve(responseJSON);
                });
            }
            else {
                response.text().then((responseText) => {
                    reject(responseText);
                });
            }
        });
    }
}