import { PlannerBucket } from '@microsoft/microsoft-graph-types';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { formatting } from '../../../../utils/formatting.js';

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
  force?: boolean;
}

class PlannerBucketRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.BUCKET_REMOVE;
  }

  public get description(): string {
    return 'Removes the Microsoft Planner bucket from a plan';
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
        ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
        force: args.options.force || false
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
      },
      {
        option: '-f, --force'
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
        else {
          if (args.options.planTitle) {
            if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
              return `${args.options.ownerGroupId} is not a valid GUID`;
            }
          }
          else if (args.options.planId) {
            if (args.options.ownerGroupId || args.options.ownerGroupName) {
              return 'Don\'t specify ownerGroupId or ownerGroupName when using planId';
            }
          }
          else {
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
    const removeBucket = async (): Promise<void> => {
      try {
        const bucket = await this.getBucket(args);

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/planner/buckets/${bucket.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'if-match': (bucket as any)['@odata.etag']
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeBucket();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the bucket ${args.options.id || args.options.name}?` });

      if (result) {
        await removeBucket();
      }
    }
  }

  private async getBucket(args: CommandArgs): Promise<PlannerBucket> {
    if (args.options.id) {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/planner/buckets/${args.options.id}`,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      return await request.get<PlannerBucket>(requestOptions);
    }

    const planId = await this.getPlanId(args);
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/plans/${planId}/buckets`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    const buckets = await request.get<{ value: PlannerBucket[] }>(requestOptions);
    const filteredBuckets = buckets.value.filter(b => args.options.name!.toLowerCase() === b.name!.toLowerCase());

    if (!filteredBuckets.length) {
      throw `The specified bucket ${args.options.name} does not exist`;
    }

    if (filteredBuckets.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', filteredBuckets);
      return await cli.handleMultipleResultsFound<PlannerBucket>(`Multiple buckets with name '${args.options.name}' found.`, resultAsKeyValuePair);
    }

    return filteredBuckets[0];
  }

  private async getPlanId(args: CommandArgs): Promise<string> {
    const { planId, planTitle, rosterId } = args.options;

    if (planId) {
      return planId;
    }

    if (planTitle) {
      const groupId: string = await this.getGroupId(args);
      const plan = await planner.getPlanByTitle(planTitle, groupId);
      return plan.id!;
    }

    const plan = await planner.getPlanByRosterId(rosterId!);
    return plan.id!;
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

export default new PlannerBucketRemoveCommand();