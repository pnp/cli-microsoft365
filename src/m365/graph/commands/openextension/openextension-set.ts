import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { Extension } from '@microsoft/microsoft-graph-types';
import { optionsUtils } from '../../../../utils/optionsUtils.js';

const options = globalOptionsZod
  .extend({
    name: zod.alias('n', z.string()),
    resourceId: zod.alias('i', z.string()),
    resourceType: zod.alias('t', z.enum(['user', 'group', 'device', 'organization'])),
    removePropertyIfEmpty: zod.alias('r', z.boolean().optional())
  })
  .and(z.any());
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class GraphOpenExtensionSetCommand extends GraphCommand {
  private readonly commandOptions = ['removePropertyIfEmpty', 'resourceType', 'resourceId', 'name'];
  public get name(): string {
    return commands.OPENEXTENSION_SET;
  }

  public get description(): string {
    return 'Updates an open extension for a resource';
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
      const currentExtension: any = await this.getOpenExtension(logger, args);
      const currentExtensionNames: string[] = Object.getOwnPropertyNames(currentExtension);

      const requestBody: any = {};

      requestBody["@odata.type"] = '#microsoft.graph.openTypeExtension';

      const unknownOptions: any = optionsUtils.getUnknownOptions(args.options, this.options);
      const unknownOptionsNames: string[] = Object.getOwnPropertyNames(unknownOptions);

      unknownOptionsNames.forEach(async option => {
        if (this.commandOptions.includes(option)) {
          return;
        }

        const value = unknownOptions[option];
        if (value === "") {
          if (!args.options.removePropertyIfEmpty) {
            requestBody[option] = null;
          }
        }
        else {
          try {
            const jsonObject = JSON.parse(value);
            requestBody[option] = jsonObject;
          }
          catch {
            requestBody[option] = value;
          }
        }
      });

      currentExtensionNames.forEach(async name => {
        if (!unknownOptionsNames.includes(name)) {
          requestBody[name] = currentExtension[name];
        }
      });

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/${args.options.resourceType}${args.options.resourceType === 'organization' ? '' : 's'}/${args.options.resourceId}/extensions/${args.options.name}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        data: requestBody,
        responseType: 'json'
      };

      if (args.options.verbose) {
        await logger.logToStderr(`Updating open extension of the ${args.options.resourceType} with id '${args.options.resourceId}'...`);
      }

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getOpenExtension(logger: Logger, args: CommandArgs): Promise<Extension> {
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

    return await request.get<Extension>(requestOptions);
  }
}

export default new GraphOpenExtensionSetCommand();