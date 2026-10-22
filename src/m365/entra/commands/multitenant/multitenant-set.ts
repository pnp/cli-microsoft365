import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { MultitenantOrganization } from './MultitenantOrganization.js';

export const options = globalOptionsZod
  .extend({
    displayName: z.string().optional().alias('n'),
    description: z.string().optional().alias('d')
  }).strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraMultitenantSetCommand extends GraphCommand {
  public get name(): string {
    return commands.MULTITENANT_SET;
  }

  public get description(): string {
    return 'Updates the properties of a multitenant organization';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => options.displayName || options.description, {
        error: 'Specify either displayName or description or both'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Updating multitenant organization...');
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
      await request.patch<MultitenantOrganization>(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraMultitenantSetCommand();