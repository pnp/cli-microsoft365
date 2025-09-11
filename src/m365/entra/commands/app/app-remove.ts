import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const options = globalOptionsZod
  .extend({
    appId: z.string().uuid().optional(),
    objectId: z.string().uuid().optional(),
    name: z.string().optional(),
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAppRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_REMOVE;
  }

  public get description(): string {
    return 'Removes an Entra app registration';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.appId, options.objectId, options.name].filter(Boolean).length === 1, {
        message: 'Specify either appId, objectId, or name'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const deleteApp = async (): Promise<void> => {
      try {
        const objectId = await this.getObjectId(args, logger);

        if (this.verbose) {
          await logger.logToStderr(`Deleting Microsoft Entra app ${objectId}...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await deleteApp();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the app?` });

      if (result) {
        await deleteApp();
      }
    }
  }

  private async getObjectId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.objectId) {
      return args.options.objectId;
    }

    const { appId, name } = args.options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appId ? appId : name}...`);
    }

    if (appId) {
      const app = await entraApp.getAppRegistrationByAppId(appId, ['id']);
      return app.id!;
    }
    else {
      const app = await entraApp.getAppRegistrationByAppName(name as string, ['id']);
      return app.id!;
    }
  }
}

export default new EntraAppRemoveCommand();