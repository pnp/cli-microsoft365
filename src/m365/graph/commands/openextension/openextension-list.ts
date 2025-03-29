import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { CliRequestOptions } from '../../../../request.js';
import { Extension } from '@microsoft/microsoft-graph-types';
import { odata } from '../../../../utils/odata.js';

const options = globalOptionsZod
  .extend({
    resourceId: zod.alias('i', z.string()),
    resourceType: zod.alias('t', z.enum(['user', 'group', 'device', 'organization']))
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class GraphOpenExtensionListCommand extends GraphCommand {
  public get name(): string {
    return commands.OPENEXTENSION_LIST;
  }

  public get description(): string {
    return 'Retrieves all open extensions for a resource';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'extensionName'];
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => options.resourceType !== 'group' && options.resourceType !== 'device' && options.resourceType !== 'organization' ||
        (options.resourceId && validation.isValidGuid(options.resourceId)), options => ({
        message: `The '${options.resourceId}' must be a valid GUID`,
        path: ['resourceId']
      }))
      .refine(options => options.resourceType !== 'user' ||
        (options.resourceId && (validation.isValidGuid(options.resourceId) || validation.isValidUserPrincipalName(options.resourceId))), options => ({
        message: `The '${options.resourceId}' must be a valid GUID or user principal name`,
        path: ['resourceId']
      }));
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/${args.options.resourceType}${args.options.resourceType === 'organization' ? '' : 's'}/${args.options.resourceId}/extensions`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      if (args.options.verbose) {
        await logger.logToStderr(`Retrieving open extensions for the ${args.options.resourceType} with id '${args.options.resourceId}'...`);
      }

      const res = await odata.getAllItems<Extension>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new GraphOpenExtensionListCommand();