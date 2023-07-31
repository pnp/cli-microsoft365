
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

class PlannerRosterMemberAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_MEMBER_ADD;
  }

  public get description(): string {
    return 'Adds a user to a Microsoft Planner Roster';
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
          return `${args.options.userName} is not a valid user principal name (UPN)`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr('Adding a user to a Microsoft Planner Roster');
    }

    try {
      const userId = await this.getUserId(logger, args);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/planner/rosters/${args.options.rosterId}/members`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: {
          userId: userId
        },
        responseType: 'json'
      };

      const response = await request.post(requestOptions);
      logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUserId(logger: Logger, args: CommandArgs): Promise<string> {
    if (this.verbose) {
      logger.logToStderr("Getting the user ID");
    }

    if (args.options.userId) {
      return args.options.userId;
    }

    const userId = await aadUser.getUserIdByUpn(args.options.userName!);

    return userId;
  }
}

module.exports = new PlannerRosterMemberAddCommand();