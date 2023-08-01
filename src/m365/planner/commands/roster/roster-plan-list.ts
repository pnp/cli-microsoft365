import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
}

class PlannerRosterPlanListCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_PLAN_LIST;
  }

  public get description(): string {
    return 'Lists all Microsoft Planner Roster plans for a specific user';
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
        userName: typeof args.options.userName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      }
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

  #initTypes(): void {
    this.types.string.push('userId', 'userName');
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'createdDateTime', 'owner'];
  }

  #initOptionSets(): void {
    this.optionSets.push({
      options: ['userId', 'userName'],
      runsWhen: (args) => args.options.userId || args.options.userName
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken);
    if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
      this.handleError(`Specify at least 'userId' or 'userName' when using application permissions.`);
    }
    else if (!isAppOnlyAccessToken && (args.options.userId || args.options.userName)) {
      this.handleError(`The options 'userId' or 'userName' cannot be used when obtaining Microsoft Planner Roster plans using delegated permissions.`);
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving Microsoft Planner Roster plans for user: ${args.options.userId || args.options.userName || 'current user'}.`);
    }

    let requestUrl: string = `${this.resource}/beta/`;
    if (args.options.userId || args.options.userName) {
      requestUrl += `users/${args.options.userId || args.options.userName}`;
    }
    else {
      requestUrl += 'me';
    }
    requestUrl += `/planner/rosterPlans`;

    try {
      const items = await odata.getAllItems(requestUrl);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PlannerRosterPlanListCommand();