import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { aadUser } from '../../../../utils/aadUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  flowName: string;
  environmentName: string;
  roleName: string;
  userId?: string;
  userName?: string;
  groupId?: string;
  groupName?: string;
  asAdmin?: boolean;
}

class FlowOwnerEnsureCommand extends PowerAutomateCommand {
  private static readonly allowedRoleNames: string[] = ['CanView', 'CanEdit'];

  public get name(): string {
    return commands.OWNER_ENSURE;
  }

  public get description(): string {
    return 'Assigns/updates permissions to a Power Automate flow';
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
        asAdmin: !!args.options.asAdmin,
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--flowName <flowName>'
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
        option: '--roleName <roleName>',
        autocomplete: FlowOwnerEnsureCommand.allowedRoleNames
      },
      {
        option: '--asAdmin'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['userId', 'userName', 'groupId', 'groupName'] });
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.flowName)) {
          return `${args.options.flowName} is not a valid GUID.`;
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID.`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid userName.`;
        }

        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
          return `${args.options.groupId} is not a valid GUID.`;
        }

        if (FlowOwnerEnsureCommand.allowedRoleNames.indexOf(args.options.roleName) === -1) {
          return `${args.options.roleName} is not a valid roleName. Valid values are: ${FlowOwnerEnsureCommand.allowedRoleNames.join(', ')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Assigning permissions for ${args.options.userId || args.options.userName || args.options.groupId || args.options.groupName} with permissions ${args.options.roleName} to Power Automate flow ${args.options.flowName}`);
      }

      let id = '';
      if (args.options.userId) {
        id = args.options.userId;
      }
      else if (args.options.userName) {
        id = await aadUser.getUserIdByUpn(args.options.userName);
      }
      else if (args.options.groupId) {
        id = args.options.groupId;
      }
      else {
        id = await aadGroup.getGroupIdByDisplayName(args.options.groupName!);
      }

      let type: string;
      if (args.options.userId || args.options.userName) {
        type = 'User';
      }
      else {
        type = 'Group';
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/modifyPermissions?api-version=2016-11-01`,
        headers: {
          accept: 'application/json'
        },
        data: {
          put: [
            {
              properties: {
                principal: {
                  id: id,
                  type: type
                },
                roleName: args.options.roleName
              }
            }
          ]
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new FlowOwnerEnsureCommand(); 