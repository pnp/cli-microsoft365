import { PlannerTask } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { odata } from '../../../../utils/odata.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { formatting } from '../../../../utils/formatting.js';

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
  ownerGroupName: z.string().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerTaskRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_REMOVE;
  }

  public get description(): string {
    return 'Removes the Microsoft Planner task from a plan';
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
      .refine(opts => opts.id !== undefined || [opts.planId, opts.planTitle, opts.rosterId].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'planId', 'planTitle' or 'rosterId'.`,
        params: { customCode: 'optionSet', options: ['planId', 'planTitle', 'rosterId'] }
      })
      .refine(opts => opts.title === undefined || [opts.bucketId, opts.bucketName].filter(x => x !== undefined).length === 1, {
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
    const removeTask = async (): Promise<void> => {
      try {
        const task = await this.getTask(args.options);

        if (this.verbose) {
          await logger.logToStderr(`Removing task '${task.title}' ...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/planner/tasks/${task.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'if-match': (task as any)['@odata.etag']
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
      await removeTask();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the task ${args.options.id || args.options.title}?` });

      if (result) {
        await removeTask();
      }
    }
  }

  private async getTask(options: Options): Promise<PlannerTask> {
    const { id, title } = options;

    if (id) {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/planner/tasks/${id}`,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      return await request.get<PlannerTask>(requestOptions);
    }

    const bucketId = await this.getBucketId(options);

    const tasks = await odata.getAllItems<PlannerTask>(`${this.resource}/v1.0/planner/buckets/${bucketId}/tasks?$select=title,id`, 'minimal');
    const filteredTasks = tasks.filter(b => title!.toLocaleLowerCase() === b.title!.toLocaleLowerCase());

    if (filteredTasks.length === 0) {
      throw `The specified task ${title} does not exist`;
    }

    if (filteredTasks.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', filteredTasks);
      return await cli.handleMultipleResultsFound<PlannerTask>(`Multiple tasks with title '${title}' found.`, resultAsKeyValuePair);
    }

    return filteredTasks[0];
  }

  private async getBucketId(options: Options): Promise<string> {
    const { bucketId, bucketName } = options;

    if (bucketId) {
      return bucketId;
    }

    const planId = await this.getPlanId(options);
    return planner.getBucketIdByTitle(bucketName!, planId);
  }

  private async getPlanId(options: Options): Promise<string> {
    const { planId, planTitle, rosterId } = options;

    if (planId) {
      return planId;
    }

    if (options.rosterId) {
      return planner.getPlanIdByRosterId(rosterId as string);
    }
    else {
      const groupId = await this.getGroupId(options);
      return planner.getPlanIdByTitle(planTitle!, groupId);
    }
  }

  private async getGroupId(options: Options): Promise<string> {
    const { ownerGroupId, ownerGroupName } = options;

    if (ownerGroupId) {
      return ownerGroupId;
    }

    return entraGroup.getGroupIdByDisplayName(ownerGroupName!);
  }
}

export default new PlannerTaskRemoveCommand();
