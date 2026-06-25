import { PlannerTask, PlannerTaskDetails } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().optional().alias('i'),
  title: z.string().optional().alias('t'),
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

class PlannerTaskGetCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_GET;
  }

  public get description(): string {
    return 'Retrieve the specified planner task';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => [opts.id, opts.title].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'id' or 'title'.`,
        params: { customCode: 'optionSet', options: ['id', 'title'] }
      })
      .refine(opts => opts.id !== undefined || opts.bucketId !== undefined || [opts.planId, opts.planTitle, opts.rosterId].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'planId', 'planTitle' or 'rosterId'.`,
        params: { customCode: 'optionSet', options: ['planId', 'planTitle', 'rosterId'] }
      })
      .refine(opts => opts.id !== undefined || [opts.bucketId, opts.bucketName].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'bucketId' or 'bucketName'.`,
        params: { customCode: 'optionSet', options: ['bucketId', 'bucketName'] }
      })
      .refine(opts => !(opts.bucketName && !opts.rosterId) || [opts.planId, opts.planTitle].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'planId' or 'planTitle'.`,
        params: { customCode: 'optionSet', options: ['planId', 'planTitle'] }
      })
      .refine(opts => !opts.planTitle || [opts.ownerGroupId, opts.ownerGroupName].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'ownerGroupId' or 'ownerGroupName'.`,
        params: { customCode: 'optionSet', options: ['ownerGroupId', 'ownerGroupName'] }
      })
      .refine(opts => !opts.id || !(opts.bucketId || opts.bucketName || opts.planId || opts.planTitle || opts.rosterId || opts.ownerGroupId || opts.ownerGroupName), {
        message: `Don't specify bucketId, bucketName, planId, planTitle, rosterId, ownerGroupId or ownerGroupName when using id.`
      });
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
      const result = await cli.handleMultipleResultsFound<PlannerTask>(`Multiple tasks with title '${options.title}' found.`, resultAsKeyValuePair);
      return result.id!;
    }

    return tasks[0].id as string;
  }

  private async getBucketId(options: Options): Promise<string> {
    if (options.bucketId) {
      return options.bucketId;
    }

    const planId = await this.getPlanId(options);
    return planner.getBucketIdByTitle(options.bucketName!, planId);
  }

  private async getPlanId(options: Options): Promise<string> {
    if (options.planId) {
      return options.planId;
    }

    if (options.rosterId) {
      return planner.getPlanIdByRosterId(options.rosterId);
    }
    else {
      const groupId: string = await this.getGroupId(options);
      return planner.getPlanIdByTitle(options.planTitle!, groupId);
    }
  }

  private async getGroupId(options: Options): Promise<string> {
    if (options.ownerGroupId) {
      return options.ownerGroupId;
    }

    return entraGroup.getGroupIdByDisplayName(options.ownerGroupName!);
  }
}

export default new PlannerTaskGetCommand();
