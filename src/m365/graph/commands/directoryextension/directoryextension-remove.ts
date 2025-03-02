import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { zod } from '../../../../utils/zod.js';
import commands from '../../commands.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { validation } from '../../../../utils/validation.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { cli } from '../../../../cli/cli.js';
import { directoryExtension } from '../../../../utils/directoryExtension.js';

const options = globalOptionsZod
  .extend({
    id: zod.alias('i', z.string().refine(id => validation.isValidGuid(id), id => ({
      message: `'${id}' is not a valid GUID.`
    })).optional()),
    name: zod.alias('n', z.string().optional()),
    appId: z.string().refine(id => validation.isValidGuid(id), id => ({
      message: `'${id}' is not a valid GUID.`
    })).optional(),
    appObjectId: z.string().refine(id => validation.isValidGuid(id), id => ({
      message: `'${id}' is not a valid GUID.`
    })).optional(),
    appName: z.string().optional(),
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class GraphDirectoryExtensionRemoveCommand extends GraphCommand{
  public get name(): string {
    return commands.DIRECTORYEXTENSION_REMOVE;
  }

  public get description(): string {
    return 'Removes a directory extension';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => !options.id !== !options.name, {
        message: 'Specify either id or name, but not both'
      })
      .refine(options => options.id || options.name, {
        message: 'Specify either id or name'
      })
      .refine(options => Object.values([options.appId, options.appObjectId, options.appName]).filter(v => typeof v !== 'undefined').length === 1, {
        message: 'Specify either appId, appObjectId or appName, but not multiple'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeDirectoryExtension = async (): Promise<void> => {
      try {
        const appObjectId = await this.getAppObjectId(args.options);

        let schemeExtensionId = args.options.id;

        if (args.options.name) {
          schemeExtensionId = (await directoryExtension.getDirectoryExtensionByName(args.options.name, appObjectId, ['id'])).id;
        }

        if (args.options.verbose) {
          await logger.logToStderr(`Removing schema extension with ID ${schemeExtensionId} from application with ID ${appObjectId}...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/applications/${appObjectId}/extensionProperties/${schemeExtensionId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          }
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeDirectoryExtension();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove directory extension '${args.options.id || args.options.name}' from application '${args.options.appObjectId || args.options.appId || args.options.appName}'?` });

      if (result) {
        await removeDirectoryExtension();
      }
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

export default new GraphDirectoryExtensionRemoveCommand();
