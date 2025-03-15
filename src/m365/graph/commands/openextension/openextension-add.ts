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
    name: zod.alias('n', z.string()),
    resourceId: zod.alias('i', z.string().refine(id => validation.isValidGuid(id), id => ({
      message: `'${id}' is not a valid GUID.`
    }))),
    resourceType: zod.alias('t', z.enum(['user', 'group', 'device', 'organization']))
  })
  .and(z.any());
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

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const requestBody: any = {};

      requestBody["extensionName"] = args.options.name;

      const unknownOptions: any = this.getUnknownZodOptions(args.options);
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
