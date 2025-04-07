import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { Extension } from '@microsoft/microsoft-graph-types';

const options = globalOptionsZod
  .extend({
    resourceId: zod.alias('i', z.string()),
    resourceType: zod.alias('t', z.enum(['user', 'group', 'device', 'organization'])),
    name: zod.alias('n', z.string())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class GraphOpenExtensionGetCommand extends GraphCommand {
  public get name(): string {
    return commands.OPENEXTENSION_GET;
  }

  public get description(): string {
    return 'Retrieves a specific open extension for a resource';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => options.resourceType !== 'group' && options.resourceType !== 'device' && options.resourceType !== 'organization' || (options.resourceId && validation.isValidGuid(options.resourceId)), options => ({
        message: `The '${options.resourceId}' must be a valid GUID`,
        path: ['resourceId']
      }))
      .refine(options => options.resourceType !== 'user' || (options.resourceId && (validation.isValidGuid(options.resourceId) || validation.isValidUserPrincipalName(options.resourceId))), options => ({
        message: `The '${options.resourceId}' must be a valid GUID or user principal name`,
        path: ['resourceId']
      }));
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving open extension for resource ${args.options.resourceId}...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/${args.options.resourceType}${args.options.resourceType === 'organization' ? '' : 's'}/${args.options.resourceId}/extensions/${args.options.name}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<Extension>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new GraphOpenExtensionGetCommand();