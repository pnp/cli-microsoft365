import { PlannerBucket } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import { aadGroup } from '../../../../utils/aadGroup';
import { planner } from '../../../../utils/planner';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  planId?: string;
  planTitle?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
}

class PlannerBucketGetCommand extends GraphCommand {
  public get name(): string {
    return commands.BUCKET_GET;
  }

  public get description(): string {
    return 'Gets the Microsoft Planner bucket in a plan';
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
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
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
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
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
        if (args.options.id) {
          if (args.options.planId || args.options.planTitle || args.options.ownerGroupId || args.options.ownerGroupName) {
            return 'Don\'t specify planId, planTitle, ownerGroupId or ownerGroupName when using id';
          }
        }

        if (args.options.name) {
          if (!args.options.planId && !args.options.planTitle) {
            return 'Specify either planId or planTitle when using name';
          }

          if (args.options.planId && args.options.planTitle) {
            return 'Specify either planId or planTitle when using name but not both';
          }

          if (args.options.planTitle) {
            if (!args.options.ownerGroupId && !args.options.ownerGroupName) {
              return 'Specify either ownerGroupId or ownerGroupName when using planTitle';
            }

            if (args.options.ownerGroupId && args.options.ownerGroupName) {
              return 'Specify either ownerGroupId or ownerGroupName when using planTitle but not both';
            }

            if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
              return `${args.options.ownerGroupId} is not a valid GUID`;
            }
          }

          if (args.options.planId) {
            if (args.options.ownerGroupId || args.options.ownerGroupName) {
              return 'Don\'t specify ownerGroupId or ownerGroupName when using planId';
            }
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'name'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const bucketId = await this.getBucketId(args);
      const bucket = await this.getBucketById(bucketId);
      logger.log(bucket);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getBucketId(args: CommandArgs): Promise<string> {
    const { id, name } = args.options;
    if (id) {
      return id;
    }

    const planId = await this.getPlanId(args);
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/plans/${planId}/buckets`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const buckets = await request.get<{ value: PlannerBucket[] }>(requestOptions);

    const filteredBuckets = buckets.value.filter(b => name!.toLowerCase() === b.name!.toLowerCase());

    if (!filteredBuckets.length) {
      throw `The specified bucket ${name} does not exist`;
    }

    if (filteredBuckets.length > 1) {
      throw `Multiple buckets with name ${name} found: ${filteredBuckets.map(x => x.id)}`;
    }

    return filteredBuckets[0].id!.toString();
  }

  private async getPlanId(args: CommandArgs): Promise<string> {
    const { planId, planTitle } = args.options;

    if (planId) {
      return planId;
    }

    const groupId = await this.getGroupId(args);
    const plan = await planner.getPlanByTitle(planTitle!, groupId);
    return plan.id!;
  }

  private async getBucketById(id: string): Promise<PlannerBucket> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/buckets/${id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<PlannerBucket>(requestOptions);
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    const { ownerGroupId, ownerGroupName } = args.options;

    if (ownerGroupId) {
      return ownerGroupId;
    }

    const group = await aadGroup.getGroupByDisplayName(ownerGroupName!);
    return group.id!;
  }
}

module.exports = new PlannerBucketGetCommand();