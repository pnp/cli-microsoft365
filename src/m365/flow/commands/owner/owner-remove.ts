import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
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
  userId?: string;
  userName?: string;
  groupId?: string;
  groupName?: string;
  asAdmin?: boolean;
  force?: boolean;
}

class FlowOwnerRemoveCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.OWNER_REMOVE;
  }

  public get description(): string {
    return 'Removes owner permissions to a Power Automate flow';
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
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        asAdmin: !!args.options.asAdmin,
        force: !!args.options.force
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
        option: '--asAdmin'
      },
      {
        option: '-f, --force'
      }
    );
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

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['userId', 'userName', 'groupId', 'groupName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Removing owner ${args.options.userId || args.options.userName || args.options.groupId || args.options.groupName} from flow ${args.options.flowName} in environment ${args.options.environmentName}`);
      }

      const removeFlowOwner = async (): Promise<void> => {
        let idToRemove = '';
        if (args.options.userId) {
          idToRemove = args.options.userId;
        }
        else if (args.options.userName) {
          idToRemove = await entraUser.getUserIdByUpn(args.options.userName);
        }
        else if (args.options.groupId) {
          idToRemove = args.options.groupId;
        }
        else {
          idToRemove = await entraGroup.getGroupIdByDisplayName(args.options.groupName!);
        }

        const requestOptions: CliRequestOptions = {
          url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/modifyPermissions?api-version=2016-11-01`,
          headers: {
            accept: 'application/json'
          },
          data: {
            delete: [
              {
                id: idToRemove
              }
            ]
          },
          responseType: 'json'
        };
        await request.post(requestOptions);
      };

      if (args.options.force) {
        await removeFlowOwner();
      }
      else {
        const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove owner '${args.options.groupId || args.options.groupName || args.options.userId || args.options.userName}' from the specified flow?` });

        if (result) {
          await removeFlowOwner();
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new FlowOwnerRemoveCommand();