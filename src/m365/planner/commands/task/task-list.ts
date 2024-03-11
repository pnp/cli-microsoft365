import { PlannerTask } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  bucketId?: string;
  bucketName?: string;
  planId?: string;
  planTitle?: string;
  rosterId?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
}

class PlannerTaskListCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_LIST;
  }

  public get description(): string {
    return 'Lists planner tasks in a bucket, plan, or tasks for the currently logged in user';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'startDateTime', 'dueDateTime', 'completedDateTime'];
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
      {
        option: '--bucketId [bucketId]'
      },
      {
        option: '--bucketName [bucketName]'
      },
      {
        option: '--planId [planId]'
      },
      {
        option: '--planTitle [planTitle]'
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
        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId as string)) {
          return `${args.options.ownerGroupId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['bucketId', 'bucketName'],
        runsWhen: (args) => {
          return args.options.bucketId !== undefined || args.options.bucketName !== undefined;
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
    this.types.string.push('planId', 'planTitle', 'ownerGroupId', 'ownerGroupName', 'bucketId', 'bucketName', 'rosterId');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const bucketName: string | undefined = args.options.bucketName;
    let bucketId: string | undefined = args.options.bucketId;
    const planTitle: string | undefined = args.options.planTitle;
    let planId: string | undefined = args.options.planId;
    let taskItems: PlannerTask[] = [];

    try {
      if (bucketId || bucketName) {
        bucketId = await this.getBucketId(args);
        taskItems = await odata.getAllItems<PlannerTask>(`${this.resource}/v1.0/planner/buckets/${bucketId}/tasks`);
        const betaTasks = await odata.getAllItems<PlannerTask>(`${this.resource}/beta/planner/buckets/${bucketId}/tasks`);

        await logger.log(this.mergeTaskPriority(taskItems, betaTasks));
      }
      else if (planId || planTitle) {
        planId = await this.getPlanId(args);
        taskItems = await odata.getAllItems<PlannerTask>(`${this.resource}/v1.0/planner/plans/${planId}/tasks`);
        const betaTasks = await odata.getAllItems<PlannerTask>(`${this.resource}/beta/planner/plans/${planId}/tasks`);

        await logger.log(this.mergeTaskPriority(taskItems, betaTasks));
      }
      else {
        taskItems = await odata.getAllItems<PlannerTask>(`${this.resource}/v1.0/me/planner/tasks`);
        const betaTasks = await odata.getAllItems<PlannerTask>(`${this.resource}/beta/me/planner/tasks`);

        await logger.log(this.mergeTaskPriority(taskItems, betaTasks));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getBucketId(args: CommandArgs): Promise<string> {
    if (args.options.bucketId) {
      return formatting.encodeQueryParameter(args.options.bucketId);
    }

    const planId = await this.getPlanId(args);
    return planner.getBucketIdByTitle(args.options.bucketName!, planId);
  }

  private async getPlanId(args: CommandArgs): Promise<string> {
    if (args.options.planId) {
      return formatting.encodeQueryParameter(args.options.planId);
    }

    if (args.options.rosterId) {
      return planner.getPlanIdByRosterId(args.options.rosterId);
    }
    else {
      const groupId: string = await this.getGroupId(args);
      return planner.getPlanIdByTitle(args.options.planTitle!, groupId);
    }
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.ownerGroupId) {
      return formatting.encodeQueryParameter(args.options.ownerGroupId);
    }

    return entraGroup.getGroupIdByDisplayName(args.options.ownerGroupName!);
  }

  private mergeTaskPriority(taskItems: PlannerTask[], betaTaskItems: PlannerTask[]): PlannerTask[] {
    const findBetaTask = (id: string): PlannerTask | undefined => betaTaskItems.find(task => task.id === id);

    taskItems.forEach(task => {
      const betaTaskItem = findBetaTask(task.id as string);
      if (betaTaskItem) {
        const { priority } = betaTaskItem;
        Object.assign(task, { priority });
      }
    });

    return taskItems;
  }
}

export default new PlannerTaskListCommand();