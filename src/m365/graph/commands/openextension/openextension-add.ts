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

const options = z.looseObject({
  ...globalOptionsZod.shape,
  name: z.string().alias('n'),
  resourceId: z.string().alias('i'),
  resourceType: z.enum(['user', 'group', 'device', 'organization']).alias('t')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class GraphOpenExtensionAddCommand extends GraphCommand {
  public get name(): string {
    return commands.OPENEXTENSION_ADD;
  }

  public get description(): string {
    return 'Adds an open extension to a resource';
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
      const requestBody: any = {};

      requestBody["extensionName"] = args.options.name;

      const unknownOptions: any = optionsUtils.getUnknownOptions(args.options, zod.schemaToOptions(this.schema!));
      const unknownOptionsNames: string[] = Object.getOwnPropertyNames(unknownOptions);

      unknownOptionsNames.forEach(async o => {
        try {
          const jsonObject = JSON.parse(unknownOptions[o]);
          requestBody[o] = jsonObject;
        }
        catch {
          requestBody[o] = unknownOptions[o];
        }
      });

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/${args.options.resourceType}${args.options.resourceType === 'organization' ? '' : 's'}/${args.options.resourceId}/extensions`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        data: requestBody,
        responseType: 'json'
      };

      if (args.options.verbose) {
        await logger.logToStderr(`Adding open extension to the ${args.options.resourceType} with id '${args.options.resourceId}'...`);
      }

      const res = await request.post<Extension>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new GraphOpenExtensionAddCommand();