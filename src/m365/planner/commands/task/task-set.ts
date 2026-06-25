import { PlannerTask, PlannerTaskDetails, User } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { AppliedCategories } from '../../AppliedCategories.js';
import commands from '../../commands.js';
import { taskPriority } from '../../taskPriority.js';

const allowedAppliedCategories = ['category1', 'category2', 'category3', 'category4', 'category5', 'category6'];

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i'),
  title: z.string().optional().alias('t'),
  planId: z.string().optional(),
  planTitle: z.string().optional(),
  rosterId: z.string().optional(),
  ownerGroupId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .optional(),
  ownerGroupName: z.string().optional(),
  bucketId: z.string().optional(),
  bucketName: z.string().optional(),
  startDateTime: z.string()
    .refine(val => validation.isValidISODateTime(val), {
      message: 'The startDateTime is not a valid ISO date string.'
    })
    .optional(),
  dueDateTime: z.string()
    .refine(val => validation.isValidISODateTime(val), {
      message: 'The dueDateTime is not a valid ISO date string.'
    })
    .optional(),
  percentComplete: z.string()
    .refine(val => !isNaN(Number(val)) && Number(val) >= 0 && Number(val) <= 100, {
      message: 'percentComplete should be between 0 and 100.'
    })
    .optional(),
  assignedToUserIds: z.string()
    .superRefine((val, ctx) => {
      const result = validation.isValidGuidArray(val);
      if (result !== true) {
        ctx.addIssue({ code: z.ZodIssueCode.custom, message: `The following GUIDs are invalid for the option 'assignedToUserIds': ${result}.` });
      }
    })
    .optional(),
  assignedToUserNames: z.string()
    .superRefine((val, ctx) => {
      const result = validation.isValidUserPrincipalNameArray(val);
      if (result !== true) {
        ctx.addIssue({ code: z.ZodIssueCode.custom, message: `The following user principal names are invalid for the option 'assignedToUserNames': ${result}.` });
      }
    })
    .optional(),
  assigneePriority: z.string().optional(),
  description: z.string().optional(),
  appliedCategories: z.string()
    .refine(val => val.split(',').every(cat => allowedAppliedCategories.includes(cat.toLocaleLowerCase().trim())), {
      message: `The appliedCategories contains invalid value. Specify either ${allowedAppliedCategories.join(', ')} as properties.`
    })
    .optional(),
  orderHint: z.string().optional(),
  priority: z.string()
    .refine(val => {
      const num = Number(val);
      if (!isNaN(num)) {
        return num >= 0 && num <= 10 && Number.isInteger(num);
      }
      return taskPriority.priorityValues.map(l => l.toLowerCase()).includes(val.toLowerCase());
    }, {
      message: `The value is not a valid priority. Allowed values are 0-10 or ${taskPriority.priorityValues.join('|')}.`
    })
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerTaskSetCommand extends GraphCommand {
  private assignments: { [userId: string]: { [property: string]: string; }; } | undefined;
  private bucketId: string | undefined;

  public get name(): string {
    return commands.TASK_SET;
  }

  public get description(): string {
    return 'Updates a Microsoft Planner Task';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => !(opts.bucketId && opts.bucketName), {
        message: `Specify either 'bucketId' or 'bucketName', but not both.`,
        params: { customCode: 'optionSet', options: ['bucketId', 'bucketName'] }
      })
      .refine(opts => !opts.bucketName || [opts.planId, opts.planTitle, opts.rosterId].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'planId', 'planTitle' or 'rosterId'.`,
        params: { customCode: 'optionSet', options: ['planId', 'planTitle', 'rosterId'] }
      })
      .refine(opts => !opts.planTitle || [opts.ownerGroupId, opts.ownerGroupName].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'ownerGroupId' or 'ownerGroupName'.`,
        params: { customCode: 'optionSet', options: ['ownerGroupId', 'ownerGroupName'] }
      })
      .refine(opts => !(opts.assignedToUserIds && opts.assignedToUserNames), {
        message: 'Specify either assignedToUserIds or assignedToUserNames but not both.',
        params: { customCode: 'optionSet', options: ['assignedToUserIds', 'assignedToUserNames'] }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      this.bucketId = await this.getBucketId(args.options);
      this.assignments = await this.generateUserAssignments(args.options);
      const etag = await this.getTaskEtag(args.options.id);
      const appliedCategories = this.generateAppliedCategories(args.options);
      const data = this.mapRequestBody(args.options, appliedCategories);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/planner/tasks/${args.options.id}`,
        headers: {
          'accept': 'application/json;odata.metadata=none',
          'If-Match': etag,
          'Prefer': 'return=representation'
        },
        responseType: 'json',
        data: data
      };

      const newTask = await request.patch<PlannerTask>(requestOptions);
      const result = await this.updateTaskDetails(args.options, newTask);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async updateTaskDetails(options: Options, newTask: PlannerTask): Promise<PlannerTask & PlannerTaskDetails> {
    if (!options.description) {
      return newTask;
    }

    const taskId = newTask.id as string;

    const etag = await this.getTaskDetailsEtag(taskId);
    const requestOptionsTaskDetails: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/tasks/${taskId}/details`,
      headers: {
        'accept': 'application/json;odata.metadata=none',
        'If-Match': etag,
        'Prefer': 'return=representation'
      },
      responseType: 'json',
      data: {
        description: options.description
      }
    };

    const taskDetails = await request.patch(requestOptionsTaskDetails);
    return { ...newTask, ...taskDetails as PlannerTaskDetails };
  }

  private async getTaskDetailsEtag(taskId: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(taskId)}/details`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    const response = await request.get<any>(requestOptions);
    return response['@odata.etag'];
  }

  private async getTaskEtag(taskId: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(taskId)}`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    const response = await request.get<any>(requestOptions);
    return response['@odata.etag'];
  }

  private generateAppliedCategories(options: Options): AppliedCategories {
    if (!options.appliedCategories) {
      return {};
    }

    const categories: AppliedCategories = {};
    options.appliedCategories.toLocaleLowerCase().split(',').forEach(x => categories[x.trim()] = true);
    return categories;
  }

  private async generateUserAssignments(options: Options): Promise<{ [userId: string]: { [property: string]: string }; }> {
    const assignments: { [userId: string]: { [property: string]: string } } = {};

    if (!options.assignedToUserIds && !options.assignedToUserNames) {
      return assignments;
    }

    const userIds = await this.getUserIds(options);
    userIds.forEach(x => assignments[x] = {
      '@odata.type': '#microsoft.graph.plannerAssignment',
      orderHint: ' !'
    });

    return assignments;
  }

  private async getUserIds(options: Options): Promise<string[]> {
    if (options.assignedToUserIds) {
      return options.assignedToUserIds.split(',').map(o => o.trim());
    }

    const userNames = options.assignedToUserNames as string;
    const userArr: string[] = userNames.split(',').map(o => o.trim());

    const promises: Promise<{ value: User[] }>[] = userArr.map(user => {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user)}'&$select=id,userPrincipalName`,
        headers: {
          'accept ': 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      return request.get(requestOptions);
    });

    const usersRes = await Promise.all(promises);

    let userUpns: string[] = [];

    userUpns = usersRes.map(res => res.value[0]?.userPrincipalName as string);
    const userIds = usersRes.map(res => res.value[0]?.id as string);

    const invalidUsers = userArr.filter(user => !userUpns.some((upn) => upn?.toLowerCase() === user.toLowerCase()));

    if (invalidUsers && invalidUsers.length > 0) {
      throw `Cannot proceed with planner task update. The following users provided are invalid : ${invalidUsers.join(',')}`;
    }

    return userIds;
  }

  private async getBucketId(options: Options): Promise<string | undefined> {
    if (options.bucketId) {
      return options.bucketId;
    }

    if (!options.bucketName) {
      return undefined;
    }

    const planId = await this.getPlanId(options);
    return planner.getBucketIdByTitle(options.bucketName, planId);
  }

  private async getPlanId(options: Options): Promise<string> {
    if (options.planId) {
      return options.planId;
    }

    if (options.rosterId) {
      return planner.getPlanIdByRosterId(options.rosterId);
    }
    else {
      const groupId = await this.getGroupId(options);
      return planner.getPlanIdByTitle(options.planTitle!, groupId);
    }
  }

  private async getGroupId(options: Options): Promise<string> {
    if (options.ownerGroupId) {
      return options.ownerGroupId;
    }

    return entraGroup.getGroupIdByDisplayName(options.ownerGroupName!);
  }

  private mapRequestBody(options: Options, appliedCategories: AppliedCategories): any {
    const requestBody: any = {};

    if (options.title) {
      requestBody.title = options.title;
    }

    if (this.bucketId) {
      requestBody.bucketId = this.bucketId;
    }

    if (options.startDateTime) {
      requestBody.startDateTime = options.startDateTime;
    }

    if (options.dueDateTime) {
      requestBody.dueDateTime = options.dueDateTime;
    }

    if (options.percentComplete) {
      requestBody.percentComplete = Number(options.percentComplete);
    }

    if (this.assignments && Object.keys(this.assignments).length > 0) {
      requestBody.assignments = this.assignments;
    }

    if (options.assigneePriority) {
      requestBody.assigneePriority = options.assigneePriority;
    }

    if (appliedCategories && Object.keys(appliedCategories).length > 0) {
      requestBody.appliedCategories = appliedCategories;
    }

    if (options.orderHint) {
      requestBody.orderHint = options.orderHint;
    }

    if (options.priority !== undefined) {
      requestBody.priority = taskPriority.getPriorityValue(options.priority);
    }

    return requestBody;
  }
}

export default new PlannerTaskSetCommand();
