import { v4 } from 'uuid';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface AppInfo {
  appRoles: {
    allowedMemberTypes: ('User' | 'Application')[];
    description: string;
    displayName: string;
    id: string;
    value: string;
  }[];
  id: string;
}

interface Options extends GlobalOptions {
  allowedMembers: string;
  appId?: string;
  appObjectId?: string;
  appName?: string;
  claim: string;
  name: string;
  description: string;
}

class AadAppRoleAddCommand extends GraphCommand {
  private static readonly allowedMembers: string[] = ['usersGroups', 'applications', 'both'];

  public get name(): string {
    return commands.APP_ROLE_ADD;
  }

  public get description(): string {
    return 'Adds role to the specified Azure AD app registration';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appId = typeof args.options.appId !== 'undefined';
    telemetryProps.appObjectId = typeof args.options.appObjectId !== 'undefined';
    telemetryProps.appName = typeof args.options.appName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getAppObjectId(args, logger)
      .then((appId: string) => this.getAppInfo(appId, logger))
      .then((appInfo: AppInfo): Promise<void> => {
        if (this.verbose) {
          logger.logToStderr(`Adding role ${args.options.name} to Azure AD app ${appInfo.id}...`);
        }

        const requestOptions: any = {
          url: `${this.resource}/v1.0/myorganization/applications/${appInfo.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: {
            appRoles: appInfo.appRoles.concat({
              displayName: args.options.name,
              description: args.options.description,
              id: v4(),
              value: args.options.claim,
              allowedMemberTypes: this.getAllowedMemberTypes(args)
            })
          }
        };
        return request.patch(requestOptions);
      })
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  private getAppInfo(appId: string, logger: Logger): Promise<AppInfo> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information about roles for Azure AD app ${appId}...`);
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

  private getAllowedMemberTypes(args: CommandArgs): ('User' | 'Application')[] {
    switch (args.options.allowedMembers) {
      case 'usersGroups':
        return ['User'];
      case 'applications':
        return ['Application'];
      case 'both':
        return ['User', 'Application'];
      default:
        return [];
    }
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

        return Promise.reject(`Multiple Azure AD application registration with name ${appName} found. Please disambiguate (app object IDs): ${res.value.map(a => a.id).join(', ')}`);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '--appId [appId]' },
      { option: '--appObjectId [appObjectId]' },
      { option: '--appName [appName]' },
      { option: '-n, --name <name>' },
      { option: '-d, --description <description>' },
      { option: '-m, --allowedMembers <allowedMembers>', autocomplete: AadAppRoleAddCommand.allowedMembers },
      { option: '-c, --claim <claim>' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const { appId, appObjectId, appName, allowedMembers, claim } = args.options;

    if ((appId && appObjectId) ||
      (appId && appName) ||
      (appObjectId && appName)) {
      return `Specify either appId, appObjectId or appName but not multiple`;
    }

    if (!appId && !appObjectId && !appName) {
      return `Specify either appId, appObjectId or appName`;
    }

    if (AadAppRoleAddCommand.allowedMembers.indexOf(allowedMembers) < 0) {
      return `${allowedMembers} is not a valid value for allowedMembers. Valid values are ${AadAppRoleAddCommand.allowedMembers.join(', ')}`;
    }

    if (claim.length > 120) {
      return `Claim must not be longer than 120 characters`;
    }

    if (claim.startsWith('.')) {
      return 'Claim must not begin with .';
    }

    if (!/^[\w:!#$%&'()*+,-.\/:;<=>?@\[\]^+_`{|}~]+$/.test(claim)) {
      return `Claim can contain only the following characters a-z, A-Z, 0-9, :!#$%&'()*+,-./:;<=>?@[]^+_\`{|}~]+`;
    }

    return true;
  }
}

module.exports = new AadAppRoleAddCommand();