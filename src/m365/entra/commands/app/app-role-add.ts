import { Application } from "@microsoft/microsoft-graph-types";
import { v4 } from 'uuid';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraApp } from '../../../../utils/entraApp.js';

interface CommandArgs {
  options: Options;
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

class EntraAppRoleAddCommand extends GraphCommand {
  private static readonly allowedMembers: string[] = ['usersGroups', 'applications', 'both'];

  public get name(): string {
    return commands.APP_ROLE_ADD;
  }

  public get description(): string {
    return 'Adds role to the specified Entra app registration';
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
        appName: typeof args.options.appName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--appId [appId]' },
      { option: '--appObjectId [appObjectId]' },
      { option: '--appName [appName]' },
      { option: '-n, --name <name>' },
      { option: '-d, --description <description>' },
      {
        option: '-m, --allowedMembers <allowedMembers>', autocomplete: EntraAppRoleAddCommand.allowedMembers
      },
      { option: '-c, --claim <claim>' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const { allowedMembers, claim } = args.options;

        if (EntraAppRoleAddCommand.allowedMembers.indexOf(allowedMembers) < 0) {
          return `${allowedMembers} is not a valid value for allowedMembers. Valid values are ${EntraAppRoleAddCommand.allowedMembers.join(', ')}`;
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
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['appId', 'appObjectId', 'appName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appInfo = await this.getAppInfo(args, logger);

      if (this.verbose) {
        await logger.logToStderr(`Adding role ${args.options.name} to Microsoft Entra app ${appInfo.id}...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/myorganization/applications/${appInfo.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          appRoles: appInfo.appRoles!.concat({
            displayName: args.options.name,
            description: args.options.description,
            id: v4(),
            value: args.options.claim,
            allowedMemberTypes: this.getAllowedMemberTypes(args)
          })
        }
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
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

  private async getAppInfo(args: CommandArgs, logger: Logger): Promise<Application> {
    const { appObjectId, appId, appName } = args.options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appObjectId ? appObjectId : (appId ? appId : appName) }...`);
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

export default new EntraAppRoleAddCommand();