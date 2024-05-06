import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface MultitenantOrganization {
  createdDateTime?: string;
  displayName?: string;
  description?: string;
  id?: string;
  state?: string;
}

class EntraMultitenantGetCommand extends GraphCommand {
  public get name(): string {
    return commands.MULTITENANT_GET;
  }

  public get description(): string {
    return 'Gets a detail about the multitenant organization';
  }

  public async commandAction(logger: Logger): Promise<void> {

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/tenantRelationships/multiTenantOrganization`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const multitenantOrg = await request.get<MultitenantOrganization>(requestOptions);

      await logger.log(multitenantOrg);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraMultitenantGetCommand();