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


interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appName: string;
  roleName: string;
  userId?: string;
  userName?: string;
  groupId?: string;
  groupName?: string;
  tenant?: boolean;
  asAdmin?: boolean;
  environmentName?: string;
  sendInvitationMail?: boolean;
}

class PaAppPermissionEnsureCommand extends PowerAppsCommand {
  private static readonly roleNames = ['CanEdit', 'CanView'];
  public get name(): string {
    return commands.APP_PERMISSION_ENSURE;
  }

  public get description(): string {
    return 'Assigns/updates permissions to a Power Apps app';
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
        sendInvitationMail: !!args.options.sendInvitationMail
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--appName <appName>'
      },
      {
        option: '--roleName <roleName>',
        autocomplete: PaAppPermissionEnsureCommand.roleNames
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
        option: '--sendInvitationMail'
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

        if (args.options.roleName && PaAppPermissionEnsureCommand.roleNames.indexOf(args.options.roleName) < 0) {
          return `${args.options.roleName} is not a valid roleName. Allowed values are ${PaAppPermissionEnsureCommand.roleNames.join(', ')}`;
        }

        if (args.options.environmentName && !args.options.asAdmin) {
          return 'please use asAdmin when using environmentName';
        }

        if (args.options.tenant && args.options.roleName !== 'CanView') {
          return 'You can only use the tenant option when using CanView as roleName';
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
      logger.logToStderr(`Assigning/updating permissions to the Power Apps app ${args.options.appName}...`);
    }

    try {
      const principalId = await this.getPrincipalId(args.options);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/providers/Microsoft.PowerApps/${args.options.asAdmin ? `scopes/admin/environments/${args.options.environmentName}/` : ''}apps/${args.options.appName}/modifyPermissions?api-version=2022-11-01`,
        headers: {
          accept: 'application/json'
        },
        data: {
          put: [
            {
              properties: {
                principal: {
                  id: principalId,
                  type: this.getPrincipalType(args.options)
                },
                NotifyShareTargetOption: args.options.sendInvitationMail ? 'Notify' : 'DoNotNotify',
                roleName: args.options.roleName
              }
            }
          ]
        },
        responseType: 'json'
      };

      await request.post<any>(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getPrincipalType(options: Options): string {
    if (options.userId || options.userName) {
      return 'User';
    }
    else if (options.groupId || options.groupName) {
      return 'Group';
    }

    return 'Tenant';
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

    return accessToken.getTenantIdFromAccessToken(Auth.service.accessTokens[Auth.defaultResource].accessToken);
  }
}

module.exports = new PaAppPermissionEnsureCommand();