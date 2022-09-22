import PowerPlatformCommand from './PowerPlatformCommand';
import request from '../../request';


export interface DynamicsApiUrl {
  instanceApiUrl: string;
}

export default abstract class DataverseCommand extends PowerPlatformCommand {

  protected async getDynamicsInstance(environment: string, asAdmin?: boolean): Promise<string> {
    let url: string = '';
    if (asAdmin) {
      url = `${this.resource}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${encodeURIComponent(environment)}`;
    }
    else {
      url = `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments/${encodeURIComponent(environment)}`;
    }

    const requestOptions: any = {
      url: `${url}?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<any>(requestOptions);

    if (response.properties.linkedEnvironmentMetadata.instanceApiUrl && response.properties.linkedEnvironmentMetadata.instanceApiUrl !== '') {
      return Promise.resolve(response.properties.linkedEnvironmentMetadata.instanceApiUrl);
    }
    else {
      return Promise.reject(`No Dynamics instance found for '${environment}'`);
    }
  }

}
