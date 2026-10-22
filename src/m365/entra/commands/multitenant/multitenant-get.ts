import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { MultitenantOrganization } from './MultitenantOrganization.js';

export const options = globalOptionsZod
  .extend({})
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraMultitenantGetCommand extends GraphCommand {
  public get name(): string {
    return commands.MULTITENANT_GET;
  }

  public get description(): string {
    return 'Gets properties of the multitenant organization';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, _args: CommandArgs): Promise<void> {

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/tenantRelationships/multiTenantOrganization`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      if (this.verbose) {
        await logger.logToStderr('Retrieving multitenant organization...');
      }
      const multitenantOrg = await request.get<MultitenantOrganization>(requestOptions);

      await logger.log(multitenantOrg);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraMultitenantGetCommand();