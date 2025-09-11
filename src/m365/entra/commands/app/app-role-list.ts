import { AppRole } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const options = globalOptionsZod
  .extend({
    appId: z.string().uuid().optional(),
    appObjectId: z.string().uuid().optional(),
    appName: z.string().optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAppRoleListCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_ROLE_LIST;
  }

  public get description(): string {
    return 'Gets Entra app registration roles';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.appId, options.appObjectId, options.appName].filter(Boolean).length === 1, {
        message: 'Specify either appId, appObjectId, or appName but not multiple'
      });
  }

  public defaultProperties(): string[] | undefined {
    return ['displayName', 'description', 'id'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const objectId = await this.getAppObjectId(args, logger);
      const appRoles = await odata.getAllItems<AppRole>(`${this.resource}/v1.0/myorganization/applications/${objectId}/appRoles`);
      await logger.log(appRoles);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppObjectId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.appObjectId) {
      return args.options.appObjectId;
    }

    const { appId, appName } = args.options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appId ? appId : appName}...`);
    }

    if (appId) {
      const app = await entraApp.getAppRegistrationByAppId(appId, ["id"]);
      return app.id!;
    }
    else {
      const app = await entraApp.getAppRegistrationByAppName(appName!, ["id"]);
      return app.id!;
    }
  }
}

export default new EntraAppRoleListCommand();