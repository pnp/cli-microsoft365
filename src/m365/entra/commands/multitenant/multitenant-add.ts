import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { MultitenantOrganization } from './MultitenantOrganization.js';

export const options = globalOptionsZod
  .extend({
    displayName: z.string().alias('n'),
    description: z.string().optional().alias('d')
  }).strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraMultitenantAddCommand extends GraphCommand {
  public get name(): string {
    return commands.MULTITENANT_ADD;
  }

  public get description(): string {
    return 'Creates a new multitenant organization';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Creating multitenant organization...');
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/tenantRelationships/multiTenantOrganization`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        description: args.options.description,
        displayName: args.options.displayName
      }
    };

    try {
      const multitenantOrg = await request.put<MultitenantOrganization>(requestOptions);

      await logger.log(multitenantOrg);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraMultitenantAddCommand();