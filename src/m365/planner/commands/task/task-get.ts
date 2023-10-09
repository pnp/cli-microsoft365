import { PlannerBucket, PlannerPlan, PlannerTask, PlannerTaskDetails } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { Cli } from '../../../../cli/Cli.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  bucketId?: string;
  bucketName?: string;
  planId?: string;
  planTitle?: string;
  rosterId?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
}

class PlannerTaskGetCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_GET;
  }

  public get description(): string {
    return 'Retrieve the specified planner task';
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
        title: typeof args.options.title !== 'undefined',
        bucketId: typeof args.options.bucketId !== 'undefined',
        bucketName: typeof args.options.bucketName !== 'undefined',
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
      { option: '-i, --id [id]' },
      { option: '-t, --title [title]' },
      { option: '--bucketId [bucketId]' },
      { option: '--bucketName [bucketName]' },
      { option: '--planId [planId]' },
      { option: '--planTitle [planTitle]' },
      { option: '--rosterId [rosterId]' },
      { option: '--ownerGroupId [ownerGroupId]' },
      { option: '--ownerGroupName [ownerGroupName]' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id) {
          if (args.options.bucketId || args.options.bucketName ||
            args.options.planId || args.options.planTitle || args.options.rosterId ||
            args.options.ownerGroupId || args.options.ownerGroupName) {
            return 'Don\'t specify bucketId, bucketName, planId, planTitle, rosterId, ownerGroupId or ownerGroupName when using id';
          }
        }

        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId as string)) {
          return `${args.options.ownerGroupId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'title'] },
      {
        options: ['planId', 'planTitle', 'rosterId'],
        runsWhen: (args) => {
          return args.options.id === undefined;
        }
      },
      {
        options: ['bucketId', 'bucketName'],
        runsWhen: (args) => {
          return args.options.title !== undefined;
        }
      },
      {
        options: ['planId', 'planTitle'],
        runsWhen: (args) => {
          return args.options.bucketName !== undefined && args.options.rosterId === undefined;
        }
      },
      {
        options: ['ownerGroupId', 'ownerGroupName'],
        runsWhen: (args) => {
          return args.options.planTitle !== undefined;
        }
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('id', 'title', 'planId', 'planTitle', 'ownerGroupId', 'ownerGroupName', 'bucketId', 'bucketName', 'rosterId');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const taskId = await this.getTaskId(args.options);
      const task = await this.getTask(taskId);
      const res = await this.getTaskDetails(task);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTask(taskId: string): Promise<PlannerTask> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(taskId)}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return await request.get<PlannerTask>(requestOptions);
  }

  private async getTaskDetails(task: PlannerTask): Promise<PlannerTask & PlannerTaskDetails> {
    const requestOptionsTaskDetails: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/tasks/${task.id}/details`,
      headers: {
        'accept': 'application/json;odata.metadata=none',
        'Prefer': 'return=representation'
      },
      responseType: 'json'
    };

    const taskDetails = await request.get(requestOptionsTaskDetails);
    return { ...task, ...taskDetails as PlannerTaskDetails };
  }

  private async getTaskId(options: Options): Promise<string> {
    if (options.id) {
      return options.id;
    }

    const bucketId = await this.getBucketId(options);
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/buckets/${bucketId}/tasks?$select=id,title`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: PlannerTask[] }>(requestOptions);

    const title = options.title as string;
    const tasks: PlannerTask[] | undefined = response.value.filter(val => val.title?.toLocaleLowerCase() === title.toLocaleLowerCase());

    if (!tasks.length) {
      throw `The specified task ${options.title} does not exist`;
    }

    if (tasks.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', tasks);
      const result = (await Cli.handleMultipleResultsFound<PlannerTask>(`Multiple tasks with title '${options.title}' found.`, resultAsKeyValuePair));
      return result.id!;
    }

    return tasks[0].id as string;
  }

  private async getBucketId(options: Options): Promise<string> {
    if (options.bucketId) {
      return options.bucketId;
    }

    const planId = await this.getPlanId(options);
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/plans/${planId}/buckets?$select=id,name`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: PlannerBucket[] }>(requestOptions);
    const bucketName = options.bucketName as string;
    const buckets: PlannerBucket[] | undefined = response.value.filter(val => val.name?.toLocaleLowerCase() === bucketName.toLocaleLowerCase());

    if (!buckets.length) {
      throw `The specified bucket ${options.bucketName} does not exist`;
    }

    if (buckets.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', buckets);
      const result = await Cli.handleMultipleResultsFound<PlannerBucket>(`Multiple buckets with name '${options.bucketName}' found.`, resultAsKeyValuePair);
      return result.id!;
    }

    return buckets[0].id as string;
  }

  private async getPlanId(options: Options): Promise<string> {
    if (options.planId) {
      return options.planId;
    }

    if (options.rosterId) {
      const plan: PlannerPlan = await planner.getPlanByRosterId(options.rosterId);
      return plan.id!;
    }
    else {
      const groupId: string = await this.getGroupId(options);
      const plan: PlannerPlan = await planner.getPlanByTitle(options.planTitle!, groupId);
      return plan.id!;
    }
  }

  private async getGroupId(options: Options): Promise<string> {
    if (options.ownerGroupId) {
      return options.ownerGroupId;
    }

    const group = await aadGroup.getGroupByDisplayName(options.ownerGroupName!);
    return group.id!;
  }
}

export default new PlannerTaskGetCommand();