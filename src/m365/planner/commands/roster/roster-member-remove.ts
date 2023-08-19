import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { aadUser } from '../../../../utils/aadUser.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  rosterId: string;
  userId?: string;
  userName?: string;
  force?: boolean;
}

class PlannerRosterMemberRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_MEMBER_REMOVE;
  }

  public get description(): string {
    return 'Removes a member from a Microsoft Planner Roster';
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
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--rosterId <rosterId>'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['userId', 'userName'] }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid userName`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing member ${args.options.userName || args.options.userId} from the Microsoft Planner Roster`);
    }

    if (args.options.force) {
      await this.removeRosterMember(args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove member '${args.options.userId || args.options.userName}'?`
      });

      if (result.continue) {
        await this.removeRosterMember(args);
      }
    }
  }

  private async getUserId(args: CommandArgs): Promise<string> {
    if (args.options.userId) {
      return args.options.userId;
    }

    return aadUser.getUserIdByUpn(args.options.userName!);
  }

  private async removeRosterMember(args: CommandArgs): Promise<void> {
    try {
      const rosterMembersContinue = await this.removeLastMemberConfirmation(args);
      if (rosterMembersContinue) {
        const userId = await this.getUserId(args);

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/beta/planner/rosters/${args.options.rosterId}/members/${userId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async removeLastMemberConfirmation(args: CommandArgs): Promise<boolean> {
    if (!args.options.force) {
      const rosterMembers = await odata.getAllItems(`${this.resource}/beta/planner/rosters/${args.options.rosterId}/members?$select=Id`);
      if (rosterMembers.length === 1) {
        const result = await Cli.prompt<{ continue: boolean }>({
          type: 'confirm',
          name: 'continue',
          default: false,
          message: `You are about to remove the last member of this Roster. When this happens, the Roster and all its contents will be deleted within 30 days. Are you sure you want to proceed?`
        });

        return result.continue;
      }
    }

    return true;
  }
}

export default new PlannerRosterMemberRemoveCommand();
