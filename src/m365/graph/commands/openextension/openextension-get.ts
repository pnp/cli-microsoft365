import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { Extension } from '@microsoft/microsoft-graph-types';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  resourceId: z.string().alias('i'),
  resourceType: z.enum(['user', 'group', 'device', 'organization']).alias('t'),
  name: z.string().alias('n')
});

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

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => options.resourceType !== 'group' && options.resourceType !== 'device' && options.resourceType !== 'organization' || (options.resourceId && validation.isValidGuid(options.resourceId)), {
        error: e => `The '${e.input}' must be a valid GUID`,
        path: ['resourceId']
      })
      .refine(options => options.resourceType !== 'user' || (options.resourceId && (validation.isValidGuid(options.resourceId) || validation.isValidUserPrincipalName(options.resourceId))), {
        error: e => `The '${e.input}' must be a valid GUID or user principal name`,
        path: ['resourceId']
      });
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