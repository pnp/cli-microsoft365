import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { SPOWebAppServicePrincipalPermissionGrant } from './SPOWebAppServicePrincipalPermissionGrant.js';

class SpoServicePrincipalGrantListCommand extends GraphCommand {
  private readonly spoServicePrincipalDisplayName = 'SharePoint Online Web Client Extensibility';

  public get name(): string {
    return commands.SERVICEPRINCIPAL_GRANT_LIST;
  }

  public get description(): string {
    return 'Lists permissions granted to the service principal';
  }

  public alias(): string[] | undefined {
    return [commands.SP_GRANT_LIST];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving permissions granted to the service principal '${this.spoServicePrincipalDisplayName}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/servicePrincipals?$filter=displayName eq '${this.spoServicePrincipalDisplayName}'&$select=id`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const response = await request.get<{ value: { id: string }[] }>(requestOptions);

      if (response.value.length === 0) {
        throw `Service principal '${this.spoServicePrincipalDisplayName}' not found`;
      }

      requestOptions.url = `${this.resource}/v1.0/servicePrincipals/${response.value[0].id}/oauth2PermissionGrants`;
      const result = await request.get<{ value: SPOWebAppServicePrincipalPermissionGrant[] }>(requestOptions);

      await logger.log(result.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoServicePrincipalGrantListCommand();