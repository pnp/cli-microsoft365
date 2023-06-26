import { Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { aadGroup } from '../../../../utils/aadGroup';
import { validation } from '../../../../utils/validation';
import PowerAppsCommand from '../../../base/PowerAppsCommand';
import commands from '../../commands';
import { aadUser } from '../../../../utils/aadUser';
import { accessToken } from '../../../../utils/accessToken';
import Auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';


interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appName: string;
  userId?: string;
  userName?: string;
  groupId?: string;
  groupName?: string;
  tenant?: boolean;
  asAdmin?: boolean;
  environmentName?: string;
  confirm?: boolean;
}

class PaAppPermissionRemoveCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.APP_PERMISSION_REMOVE;
  }

  public get description(): string {
    return 'Removes permissions to a Power Apps app';
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
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        groupId: typeof args.options.userId !== 'undefined',
        groupName: typeof args.options.userName !== 'undefined',
        tenant: !!args.options.tenant,
        asAdmin: !!args.options.asAdmin,
        environmentName: typeof args.options.environmentName !== 'undefined',
        confirm: !!args.options.confirm
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--appName <appName>'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '--groupId [groupId]'
      },
      {
        option: '--groupName [groupName]'
      },
      {
        option: '--tenant'
      },
      {
        option: '--asAdmin'
      },
      {
        option: '-e, --environmentName [environmentName]'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.appName)) {
          return `${args.options.appName} is not a valid GUID`;
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
          return `${args.options.groupId} is not a valid GUID`;
        }

        if (args.options.environmentName && !args.options.asAdmin) {
          return 'please use asAdmin when using environmentName';
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['userId', 'userName', 'groupId', 'groupName', 'tenant'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Removing permissions to a Power Apps app ${args.options.appName}...`);
    }

    try {
      if (args.options.confirm) {
        await this.removeAppPermission(args.options);
      }
      else {
        const result = await Cli.prompt<{ continue: boolean }>({
          type: 'confirm',
          name: 'continue',
          default: false,
          message: `Are you sure you want to remove the permissions for the Power App '${args.options.appName}'?`
        });

        if (result.continue) {
          await this.removeAppPermission(args.options);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async removeAppPermission(options: Options): Promise<void> {
    const principalId = await this.getPrincipalId(options);
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/providers/Microsoft.PowerApps/${options.asAdmin ? `scopes/admin/environments/${options.environmentName}/` : ''}apps/${options.appName}/modifyPermissions?api-version=2022-11-01`,
      headers: {
        accept: 'application/json'
      },
      data: {
        delete: [
          {
            id: principalId
          }
        ]
      },
      responseType: 'json'
    };

    await request.post<any>(requestOptions);
  }

  private async getPrincipalId(options: Options): Promise<string | undefined> {
    if (options.groupId) {
      return options.groupId;
    }
    else if (options.userId) {
      return options.userId;
    }
    else if (options.groupName) {
      const group: Group = await aadGroup.getGroupByDisplayName(options.groupName!);
      return group.id!;
    }
    else if (options.userName) {
      const userId: string = await aadUser.getUserIdByUpn(options.userName);
      return userId;
    }

    return `tenant-${accessToken.getTenantIdFromAccessToken(Auth.service.accessTokens[Auth.defaultResource].accessToken)}`;
  }
}

module.exports = new PaAppPermissionRemoveCommand();