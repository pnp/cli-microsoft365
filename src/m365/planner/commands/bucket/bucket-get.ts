import { PlannerBucket } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  planId?: string;
  planTitle?: string;
  rosterId?: string;
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
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        planId: typeof args.options.planId !== 'undefined',
        planTitle: typeof args.options.planTitle !== 'undefined',
        rosterId: typeof args.options.rosterId !== 'undefined',
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
        option: '--rosterId [rosterId]'
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
          if (args.options.planId || args.options.planTitle || args.options.rosterId || args.options.ownerGroupId || args.options.ownerGroupName) {
            return 'Don\'t specify planId, planTitle, rosterId, ownerGroupId or ownerGroupName when using id';
          }
        }

        if (args.options.name) {
          if (args.options.planTitle) {
            if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
              return `${args.options.ownerGroupId} is not a valid GUID`;
            }
          }

          if (args.options.planId) {
            if (args.options.ownerGroupId || args.options.ownerGroupName) {
              return 'Don\'t specify ownerGroupId or ownerGroupName when using planId';
            }
          }

          if (args.options.rosterId) {
            if (args.options.ownerGroupId || args.options.ownerGroupName) {
              return 'Don\'t specify ownerGroupId or ownerGroupName when using rosterId';
            }
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'name'] },
      {
        options: ['planId', 'planTitle', 'rosterId'],
        runsWhen: (args) => args.options.name !== undefined
      },
      {
        options: ['ownerGroupId', 'ownerGroupName'],
        runsWhen: (args) => args.options.planTitle !== undefined
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('id', 'name', 'planId', 'planTitle', 'ownerGroupId', 'ownerGroupName', 'rosterId ');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const bucket = await this.getBucket(args);
      await logger.log(bucket);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getBucket(args: CommandArgs): Promise<PlannerBucket> {
    const { id, name } = args.options;
    if (id) {
      return this.getBucketById(id);
    }

    const planId = await this.getPlanId(args);
    return planner.getBucketByTitle(name!, planId);
  }

  private async getPlanId(args: CommandArgs): Promise<string> {
    const { planId, planTitle, rosterId } = args.options;

    if (planId) {
      return planId;
    }

    if (planTitle) {
      const groupId: string = await this.getGroupId(args);
      return planner.getPlanIdByTitle(planTitle, groupId);
    }

    return planner.getPlanIdByRosterId(rosterId!);
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

    return entraGroup.getGroupIdByDisplayName(ownerGroupName!);
  }
}

export default new PlannerBucketGetCommand();