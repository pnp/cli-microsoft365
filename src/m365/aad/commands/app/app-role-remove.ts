import { Application, AppRole } from "@microsoft/microsoft-graph-types";
import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
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

  public alias(): string[] | undefined {
    return [commands.APP_ROLE_DELETE];
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appId = typeof args.options.appId !== 'undefined';
    telemetryProps.appObjectId = typeof args.options.appObjectId !== 'undefined';
    telemetryProps.appName = typeof args.options.appName !== 'undefined';
    telemetryProps.claim = typeof args.options.claim !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.id = typeof args.options.id !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this.showDeprecationWarning(logger, commands.APP_ROLE_DELETE, commands.APP_ROLE_REMOVE);

    const deleteAppRole: () => void = (): void => {
      this
        .processAppRoleDelete(logger, args)
        .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
    };

    if (args.options.confirm) {
      deleteAppRole();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the app role ?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          deleteAppRole();
        }
      });
    }
  }

  private processAppRoleDelete(logger: Logger, args: CommandArgs): Promise<void> {
    return this
      .getAppObjectId(args, logger)
      .then((appObjectId: string) => this.getAadApp(appObjectId, logger))
      .then((aadApp: Application): Promise<void> => {
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
          return Promise.reject(`Multiple roles with the provided 'name' were found. Please disambiguate using the claims : ${appRoleToDelete.map(role => `${role.value}`).join(', ')}`);
        }
        if (appRoleToDelete.length === 0) {
          return Promise.reject(`No app role with ${appRoleDeleteIdentifierNameValue} found.`);
        }

        const roleToDelete: AppRole = appRoleToDelete[0];

        if (roleToDelete.isEnabled) {
          return this
            .disableAppRole(logger, aadApp, roleToDelete.id!)
            .then(_ => this.deleteAppRole(logger, aadApp, roleToDelete.id!));
        }
        else {
          return this.deleteAppRole(logger, aadApp, roleToDelete.id!);
        }
      });
  }


  private disableAppRole(logger: Logger, aadApp: Application, roleId: string): Promise<void> {
    const roleIndex = aadApp.appRoles!.findIndex((role: AppRole) => role.id === roleId);

    if (this.verbose) {
      logger.logToStderr(`Disabling the app role`);
    }

    aadApp.appRoles![roleIndex].isEnabled = false;

    const requestOptions: any = {
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

  private deleteAppRole(logger: Logger, aadApp: Application, roleId: string): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Deleting the app role.`);
    }

    const updatedAppRoles = aadApp.appRoles!.filter((role: AppRole) => role.id !== roleId);
    const requestOptions: any = {
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

  private getAadApp(appId: string, logger: Logger): Promise<Application> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving app roles information for the Azure AD app ${appId}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/myorganization/applications/${appId}?$select=id,appRoles`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    return request.get(requestOptions);
  }

  private getAppObjectId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.appObjectId) {
      return Promise.resolve(args.options.appObjectId);
    }

    const { appId, appName } = args.options;

    if (this.verbose) {
      logger.logToStderr(`Retrieving information about Azure AD app ${appId ? appId : appName}...`);
    }

    const filter: string = appId ?
      `appId eq '${encodeURIComponent(appId)}'` :
      `displayName eq '${encodeURIComponent(appName as string)}'`;

    const requestOptions: any = {
      url: `${this.resource}/v1.0/myorganization/applications?$filter=${filter}&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: { id: string }[] }>(requestOptions)
      .then((res: { value: { id: string }[] }): Promise<string> => {
        if (res.value.length === 1) {
          return Promise.resolve(res.value[0].id);
        }

        if (res.value.length === 0) {
          const applicationIdentifier = appId ? `ID ${appId}` : `name ${appName}`;
          return Promise.reject(`No Azure AD application registration with ${applicationIdentifier} found`);
        }

        return Promise.reject(`Multiple Azure AD application registration with name ${appName} found. Please disambiguate using app object IDs: ${res.value.map(a => a.id).join(', ')}`);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '--appId [appId]' },
      { option: '--appObjectId [appObjectId]' },
      { option: '--appName [appName]' },
      { option: '-n, --name [name]' },
      { option: '-i, --id [id]' },
      { option: '-c, --claim [claim]' },
      { option: '--confirm' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const { appId, appObjectId, appName, name, id, claim } = args.options;

    if ((appId && appObjectId) ||
      (appId && appName) ||
      (appObjectId && appName)) {
      return `Specify either appId, appObjectId or appName but not multiple`;
    }

    if ((name && claim) ||
      (name && id) ||
      (claim && id)) {
      return `Specify either name, claim or id of the role but not multiple`;
    }

    if (!appId &&
      !appObjectId &&
      !appName) {
      return `Specify either appId, appObjectId or appName`;
    }

    if (!name &&
      !claim &&
      !id) {
      return `Specify either name, claim or id of the role`;
    }

    if (args.options.id) {
      if (!validation.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }
    }

    return true;
  }
}

module.exports = new AadAppRoleRemoveCommand();
