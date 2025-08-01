import { Application, AppRole } from "@microsoft/microsoft-graph-types";
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from "../../../../utils/formatting.js";
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraApp } from "../../../../utils/entraApp.js";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  appObjectId?: string;
  appName?: string;
  claim?: string;
  name?: string;
  id?: string;
}

class EntraAppRoleRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_ROLE_REMOVE;
  }

  public get description(): string {
    return 'Removes role from the specified Entra app registration';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        appId: typeof args.options.appId !== 'undefined',
        appObjectId: typeof args.options.appObjectId !== 'undefined',
        appName: typeof args.options.appName !== 'undefined',
        claim: typeof args.options.claim !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        id: typeof args.options.id !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--appId [appId]' },
      { option: '--appObjectId [appObjectId]' },
      { option: '--appName [appName]' },
      { option: '-n, --name [name]' },
      { option: '-i, --id [id]' },
      { option: '-c, --claim [claim]' },
      { option: '-f, --force' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id) {
          if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['appId', 'appObjectId', 'appName'] },
      { options: ['name', 'claim', 'id'] }
    );
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
