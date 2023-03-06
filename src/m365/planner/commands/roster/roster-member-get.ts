import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { aadUser } from '../../../../utils/aadUser';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  rosterId: string;
  userId?: string;
  userName?: string;
}

class PlannerRosterMemberGetCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_MEMBER_GET;
  }

  public get description(): string {
    return 'Gets a member of the specified Microsoft Planner Roster';
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
        userName: typeof args.options.userName !== 'undefined'
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
      logger.logToStderr(`Retrieving member ${args.options.userName || args.options.userId} from the Microsoft Planner Roster`);
    }
    try {
      const userId = await this.getUserId(args);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/planner/rosters/${args.options.rosterId}/members/${userId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const response = await request.get(requestOptions);
      logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUserId(args: CommandArgs): Promise<string> {
    if (args.options.userId) {
      return args.options.userId;
    }

    return aadUser.getUserIdByUpn(args.options.userName!);
  }
}

module.exports = new PlannerRosterMemberGetCommand();
