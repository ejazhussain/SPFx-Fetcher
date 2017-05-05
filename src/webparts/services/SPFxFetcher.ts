
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { ISPFxFetcher } from '../interfaces';

export class SPFxFetcher implements ISPFxFetcher {

    private _spHttpClient: SPHttpClient;

    constructor(spHttpClient: SPHttpClient) {
        this._spHttpClient = spHttpClient;
    }

    public get(endpoint: string, config: SPHttpClientConfiguration = SPHttpClient.configurations.v1): Promise<any> {
        return new Promise<any>((resolve, reject) => {

            return this._spHttpClient.get(endpoint, config)
                .then((response: SPHttpClientResponse) => {
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

        });


    }
}