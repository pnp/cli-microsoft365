import { PlannerTask } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  bucketId: z.string().optional(),
  bucketName: z.string().optional(),
  planId: z.string().optional(),
  planTitle: z.string().optional(),
  rosterId: z.string().optional(),
  ownerGroupId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .optional(),
  ownerGroupName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerTaskListCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_LIST;
  }

  public get description(): string {
    return 'Lists planner tasks in a bucket, plan, or tasks for the currently logged in user';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'startDateTime', 'dueDateTime', 'completedDateTime'];
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => !(opts.bucketId && opts.bucketName), {
        message: `Specify either 'bucketId' or 'bucketName', but not both.`,
        params: { customCode: 'optionSet', options: ['bucketId', 'bucketName'] }
      })
      .refine(opts => !(opts.bucketName && !opts.rosterId) || [opts.planId, opts.planTitle].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'planId' or 'planTitle'.`,
        params: { customCode: 'optionSet', options: ['planId', 'planTitle'] }
      })
      .refine(opts => !opts.planTitle || [opts.ownerGroupId, opts.ownerGroupName].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'ownerGroupId' or 'ownerGroupName'.`,
        params: { customCode: 'optionSet', options: ['ownerGroupId', 'ownerGroupName'] }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const bucketName: string | undefined = args.options.bucketName;
    let bucketId: string | undefined = args.options.bucketId;
    const planTitle: string | undefined = args.options.planTitle;
    let planId: string | undefined = args.options.planId;
    let taskItems: PlannerTask[];

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
