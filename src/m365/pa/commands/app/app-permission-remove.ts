import Auth from '../../../../Auth.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';

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
  force?: boolean;
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
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        tenant: !!args.options.tenant,
        asAdmin: !!args.options.asAdmin,
        environmentName: typeof args.options.environmentName !== 'undefined',
        force: !!args.options.force
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
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.appName)) {
          return `${args.options.appName} is not a valid GUID for appName.`;
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID for userId.`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid user principal name (UPN) for userName.`;
        }

        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
          return `${args.options.groupId} is not a valid GUID for groupId.`;
        }

        if (args.options.environmentName && !args.options.asAdmin) {
          return 'Specifying environmentName is only allowed when using asAdmin';
        }

        if (args.options.asAdmin && !args.options.environmentName) {
          return 'Specifying asAdmin is only allowed when using environmentName';
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['userId', 'userName', 'groupId', 'groupName', 'tenant'] });
  }

  #initTypes(): void {
    this.types.string.push('appName', 'userId', 'userName', 'groupId', 'groupName', 'environmentName');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.force) {
        await this.removeAppPermission(logger, args.options);
      }
      else {
        const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the permissions of '${args.options.userId || args.options.userName || args.options.groupId || args.options.groupName || (args.options.tenant && 'everyone')}' from the Power App '${args.options.appName}'?` });

        if (result) {
          await this.removeAppPermission(logger, args.options);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async removeAppPermission(logger: Logger, options: Options): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing permissions for '${options.userId || options.userName || options.groupId || options.groupName || (options.tenant && 'everyone')}' for the Power Apps app ${options.appName}...`);
    }

    const principalId: string = await this.getPrincipalId(options);
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

  private async getPrincipalId(options: Options): Promise<string> {
    if (options.groupId) {
      return options.groupId;
    }
    if (options.userId) {
      return options.userId;
    }
    if (options.groupName) {
      const group = await entraGroup.getGroupByDisplayName(options.groupName);
      return group.id!;
    }
    if (options.userName) {
      const userId = await entraUser.getUserIdByUpn(options.userName);
      return userId;
    }

    return `tenant-${accessToken.getTenantIdFromAccessToken(Auth.service.accessTokens[Auth.defaultResource].accessToken)}`;
  }
}

export default new PaAppPermissionRemoveCommand();