import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import commands from '../../commands.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { odata } from '../../../../utils/odata.js';
import { ExtensionProperty } from '@microsoft/microsoft-graph-types';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  appId: z.uuid().optional(),
  appObjectId: z.uuid().optional(),
  appName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class GraphDirectoryExtensionListCommand extends GraphCommand {
  public get name(): string {
    return commands.DIRECTORYEXTENSION_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of directory extensions';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'appDisplayName'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options =>
        ([options.appId, options.appObjectId, options.appName].filter(x => x !== undefined).length <= 1), {
        error: 'Specify either appId, appObjectId, or appName, but not multiple.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.appId || args.options.appObjectId || args.options.appName) {
        const appObjectId = await this.getAppObjectId(args.options);

        const endpoint: string = `${this.resource}/v1.0/applications/${appObjectId}/extensionProperties/`;
        const items = await odata.getAllItems<ExtensionProperty>(endpoint);
        await logger.log(items);
      }
      else {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/directoryObjects/getAvailableExtensionProperties`,
          headers: {
            'content-type': 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        const res: any = await request.post(requestOptions);
        await logger.log(res.value);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppObjectId(options: Options): Promise<string> {
    if (options.appObjectId) {
      return options.appObjectId;
    }

    if (options.appId) {
      return (await entraApp.getAppRegistrationByAppId(options.appId, ["id"])).id!;
    }

    return (await entraApp.getAppRegistrationByAppName(options.appName!, ["id"])).id!;
  }
}

export default new GraphDirectoryExtensionListCommand();