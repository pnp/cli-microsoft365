import { Application, AppRole } from "@microsoft/microsoft-graph-types";
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from "../../../../utils/formatting";
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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

class AadAppRoleRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_ROLE_REMOVE;
  }

  public get description(): string {
    return 'Removes role from the specified Azure AD app registration';
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
      { option: '--confirm' }
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

    if (args.options.confirm) {
      await deleteAppRole();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the app role ?`
      });

      if (result.continue) {
        await deleteAppRole();
      }
    }
  }

  private async processAppRoleDelete(logger: Logger, args: CommandArgs): Promise<void> {
    const appObjectId = await this.getAppObjectId(args, logger);
    const aadApp = await this.getAadApp(appObjectId, logger);

    const appRoleDeleteIdentifierNameValue = args.options.name ? `name '${args.options.name}'` : (args.options.claim ? `claim '${args.options.claim}'` : `id '${args.options.id}'`);
    if (this.verbose) {
      logger.logToStderr(`Deleting role with ${appRoleDeleteIdentifierNameValue} from Azure AD app ${aadApp.id}...`);
    }

    // Find the role search criteria provided by the user.
    const appRoleDeleteIdentifierProperty = args.options.name ? `displayName` : (args.options.claim ? `value` : `id`);
    const appRoleDeleteIdentifierValue = args.options.name ? args.options.name : (args.options.claim ? args.options.claim : args.options.id);

    const appRoleToDelete: AppRole[] = aadApp.appRoles!.filter((role: AppRole) => role[appRoleDeleteIdentifierProperty] === appRoleDeleteIdentifierValue);

    if (args.options.name &&
      appRoleToDelete !== undefined &&
      appRoleToDelete.length > 1) {
      throw `Multiple roles with the provided 'name' were found. Please disambiguate using the claims : ${appRoleToDelete.map(role => `${role.value}`).join(', ')}`;
    }
    if (appRoleToDelete.length === 0) {
      throw `No app role with ${appRoleDeleteIdentifierNameValue} found.`;
    }

    const roleToDelete: AppRole = appRoleToDelete[0];

    if (roleToDelete.isEnabled) {
      await this.disableAppRole(logger, aadApp, roleToDelete.id!);
      await this.deleteAppRole(logger, aadApp, roleToDelete.id!);
    }
    else {
      await this.deleteAppRole(logger, aadApp, roleToDelete.id!);
    }
  }


  private async disableAppRole(logger: Logger, aadApp: Application, roleId: string): Promise<void> {
    const roleIndex = aadApp.appRoles!.findIndex((role: AppRole) => role.id === roleId);

    if (this.verbose) {
      logger.logToStderr(`Disabling the app role`);
    }

    aadApp.appRoles![roleIndex].isEnabled = false;

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications/${aadApp.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        appRoles: aadApp.appRoles
      }
    };

    return request.patch(requestOptions);
  }

  private async deleteAppRole(logger: Logger, aadApp: Application, roleId: string): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Deleting the app role.`);
    }

    const updatedAppRoles = aadApp.appRoles!.filter((role: AppRole) => role.id !== roleId);
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications/${aadApp.id}`,
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

  private async getAadApp(appId: string, logger: Logger): Promise<Application> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving app roles information for the Azure AD app ${appId}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications/${appId}?$select=id,appRoles`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    return request.get(requestOptions);
  }

  private async getAppObjectId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.appObjectId) {
      return args.options.appObjectId;
    }

    const { appId, appName } = args.options;

    if (this.verbose) {
      logger.logToStderr(`Retrieving information about Azure AD app ${appId ? appId : appName}...`);
    }

    const filter: string = appId ?
      `appId eq '${formatting.encodeQueryParameter(appId)}'` :
      `displayName eq '${formatting.encodeQueryParameter(appName as string)}'`;

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications?$filter=${filter}&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ value: { id: string }[] }>(requestOptions);

    if (res.value.length === 1) {
      return res.value[0].id;
    }

    if (res.value.length === 0) {
      const applicationIdentifier = appId ? `ID ${appId}` : `name ${appName}`;
      throw `No Azure AD application registration with ${applicationIdentifier} found`;
    }

    throw `Multiple Azure AD application registration with name ${appName} found. Please disambiguate using app object IDs: ${res.value.map(a => a.id).join(', ')}`;
  }
}

module.exports = new AadAppRoleRemoveCommand();
