import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { accessToken } from '../../../../utils/accessToken';
import { validation } from '../../../../utils/validation';
import { aadGroup } from '../../../../utils/aadGroup';
import { planner } from '../../../../utils/planner';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  ownerGroupId?: string;
  ownerGroupName?: string;
}

class PlannerPlanListCommand extends GraphCommand {
  public get name(): string {
    return commands.PLAN_LIST;
  }

  public get description(): string {
    return 'Returns a list of plans associated with a specified group';
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
        ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
        ownerGroupName: typeof args.options.ownerGroupName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: "--ownerGroupId [ownerGroupId]"
      },
      {
        option: "--ownerGroupName [ownerGroupName]"
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId as string)) {
          return `${args.options.ownerGroupId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['ownerGroupId', 'ownerGroupName'] });
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'createdDateTime', 'owner'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.');
      return;
    }

    try {
      const groupId = await this.getGroupId(args);
      const result = await planner.getPlansByGroupId(groupId);
      if (result && result.length > 0) {
        logger.log(result);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.ownerGroupId) {
      return Promise.resolve(args.options.ownerGroupId);
    }

    return aadGroup
      .getGroupByDisplayName(args.options.ownerGroupName!)
      .then(group => group.id!);
  }
}

module.exports = new PlannerPlanListCommand();