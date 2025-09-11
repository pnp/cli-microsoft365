import { Application, AppRole } from "@microsoft/microsoft-graph-types";
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraApp } from "../../../../utils/entraApp.js";
import { formatting } from "../../../../utils/formatting.js";
import { zod } from "../../../../utils/zod.js";
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const options = globalOptionsZod
  .extend({
    appId: z.string().uuid().optional(),
    appObjectId: z.string().uuid().optional(),
    appName: z.string().optional(),
    claim: zod.alias('c', z.string().optional()),
    name: zod.alias('n', z.string().optional()),
    id: zod.alias('i', z.string().optional()),
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAppRoleRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_ROLE_REMOVE;
  }

  public get description(): string {
    return 'Removes role from the specified Entra app registration';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.appId, options.appObjectId, options.appName].filter(Boolean).length === 1, {
        message: 'Specify either appId, appObjectId, or appName'
      })
      .refine(options => [options.name, options.claim, options.id].filter(Boolean).length === 1, {
        message: 'Specify either name, claim, or id'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const deleteAppRole = async (): Promise<void> => {
      try {
        await this.processAppRoleDelete(logger, args);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await deleteAppRole();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the app role?` });

      if (result) {
        await deleteAppRole();
      }
    }
  }

  private async processAppRoleDelete(logger: Logger, args: CommandArgs): Promise<void> {
    const app = await this.getEntraApp(args, logger);

    const appRoleDeleteIdentifierNameValue = args.options.name ? `name '${args.options.name}'` : (args.options.claim ? `claim '${args.options.claim}'` : `id '${args.options.id}'`);
    if (this.verbose) {
      await logger.logToStderr(`Deleting role with ${appRoleDeleteIdentifierNameValue} from Microsoft Entra app ${app.id}...`);
    }

    // Find the role search criteria provided by the user.
    const appRoleDeleteIdentifierProperty = args.options.name ? `displayName` : (args.options.claim ? `value` : `id`);
    const appRoleDeleteIdentifierValue = args.options.name ? args.options.name : (args.options.claim ? args.options.claim : args.options.id);

    const appRoleToDelete: AppRole[] = app.appRoles!.filter((role: AppRole) => role[appRoleDeleteIdentifierProperty] === appRoleDeleteIdentifierValue);

    if (args.options.name &&
      appRoleToDelete !== undefined &&
      appRoleToDelete.length > 1) {

      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', appRoleToDelete);
      appRoleToDelete[0] = await cli.handleMultipleResultsFound<AppRole>(`Multiple roles with name '${args.options.name}' found.`, resultAsKeyValuePair);
    }
    if (appRoleToDelete.length === 0) {
      throw `No app role with ${appRoleDeleteIdentifierNameValue} found.`;
    }

    const roleToDelete: AppRole = appRoleToDelete[0];

    if (roleToDelete.isEnabled) {
      await this.disableAppRole(logger, app, roleToDelete.id!);
      await this.deleteAppRole(logger, app, roleToDelete.id!);
    }
    else {
      await this.deleteAppRole(logger, app, roleToDelete.id!);
    }
  }


  private async disableAppRole(logger: Logger, app: Application, roleId: string): Promise<void> {
    const roleIndex = app.appRoles!.findIndex((role: AppRole) => role.id === roleId);

    if (this.verbose) {
      await logger.logToStderr(`Disabling the app role`);
    }

    app.appRoles![roleIndex].isEnabled = false;

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications/${app.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        appRoles: app.appRoles
      }
    };

    return request.patch(requestOptions);
  }

  private async deleteAppRole(logger: Logger, app: Application, roleId: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Deleting the app role.`);
    }

    const updatedAppRoles = app.appRoles!.filter((role: AppRole) => role.id !== roleId);
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications/${app.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        appRoles: updatedAppRoles
      }
    };

    return request.patch(requestOptions);
  }

  private async getEntraApp(args: CommandArgs, logger: Logger): Promise<Application> {
    const { appObjectId, appId, appName } = args.options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appObjectId ? appObjectId : (appId ? appId : appName)}...`);
    }

    if (appObjectId) {
      return await entraApp.getAppRegistrationByObjectId(appObjectId, ['id', 'appRoles']);
    }
    else if (appId) {
      return await entraApp.getAppRegistrationByAppId(appId, ['id', 'appRoles']);
    }
    else {
      return await entraApp.getAppRegistrationByAppName(appName!, ['id', 'appRoles']);
    }
  }
}

export default new EntraAppRoleRemoveCommand();
