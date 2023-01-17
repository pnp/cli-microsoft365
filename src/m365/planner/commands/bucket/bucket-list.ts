import { PlannerBucket } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import { aadGroup } from '../../../../utils/aadGroup';
import { planner } from '../../../../utils/planner';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  planId?: string;
  planTitle?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
}

class PlannerBucketListCommand extends GraphCommand {
  public get name(): string {
    return commands.BUCKET_LIST;
  }

  public get description(): string {
    return 'Lists the Microsoft Planner buckets in a plan';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'planId', 'orderHint'];
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
        planId: typeof args.options.planId !== 'undefined',
        planTitle: typeof args.options.planTitle !== 'undefined',
        ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
        ownerGroupName: typeof args.options.ownerGroupName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--planId [planId]'
      },
      {
        option: "--planTitle [planTitle]"
      },
      {
        option: '--ownerGroupId [ownerGroupId]'
      },
      {
        option: '--ownerGroupName [ownerGroupName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.planTitle && !args.options.ownerGroupId && !args.options.ownerGroupName) {
          return 'Specify either ownerGroupId or ownerGroupName when using planTitle';
        }

        if (args.options.planTitle && args.options.ownerGroupId && args.options.ownerGroupName) {
          return 'Specify either ownerGroupId or ownerGroupName when using planTitle but not both';
        }

        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId as string)) {
          return `${args.options.ownerGroupId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['planId', 'planTitle'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const planId = await this.getPlanId(args);
      const buckets = await odata.getAllItems<PlannerBucket>(`${this.resource}/v1.0/planner/plans/${planId}/buckets`);
      logger.log(buckets);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getPlanId(args: CommandArgs): Promise<string> {
    if (args.options.planId) {
      return Promise.resolve(args.options.planId);
    }

    return this
      .getGroupId(args)
      .then(groupId => planner.getPlanByTitle(args.options.planTitle!, groupId))
      .then(plan => plan.id!);
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

module.exports = new PlannerBucketListCommand();