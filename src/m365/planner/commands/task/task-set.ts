import { PlannerBucket, PlannerTask, PlannerTaskDetails, User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import { aadGroup } from '../../../../utils/aadGroup';
import { planner } from '../../../../utils/planner';
import GraphCommand from '../../../base/GraphCommand';
import { AppliedCategories } from '../../AppliedCategories';
import commands from '../../commands';
import { taskPriority } from '../../taskPriority';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  title?: string;
  planId?: string;
  planTitle?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
  bucketId?: string;
  bucketName?: string;
  startDateTime?: string;
  dueDateTime?: string;
  percentComplete?: number;
  assignedToUserIds?: string;
  assignedToUserNames?: string;
  assigneePriority?: string;
  assignments?: string;
  description?: string;
  appliedCategories?: string;
  orderHint?: string;
  priority?: number | string;
}

class PlannerTaskSetCommand extends GraphCommand {
  private assignments: { [userId: string]: { [property: string]: string; }; } | undefined;
  private bucketId: string | undefined;
  private allowedAppliedCategories: string[] = ['category1', 'category2', 'category3', 'category4', 'category5', 'category6'];

  public get name(): string {
    return commands.TASK_SET;
  }

  public get description(): string {
    return 'Updates a Microsoft Planner Task';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        title: typeof args.options.title !== 'undefined',
        planId: typeof args.options.planId !== 'undefined',
        planTitle: typeof args.options.planTitle !== 'undefined',
        ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
        ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
        bucketId: typeof args.options.bucketId !== 'undefined',
        bucketName: typeof args.options.bucketName !== 'undefined',
        startDateTime: typeof args.options.startDateTime !== 'undefined',
        dueDateTime: typeof args.options.dueDateTime !== 'undefined',
        percentComplete: typeof args.options.percentComplete !== 'undefined',
        assignedToUserIds: typeof args.options.assignedToUserIds !== 'undefined',
        assignedToUserNames: typeof args.options.assignedToUserNames !== 'undefined',
        assigneePriority: typeof args.options.assigneePriority !== 'undefined',
        description: typeof args.options.description !== 'undefined',
        appliedCategories: typeof args.options.appliedCategories !== 'undefined',
        orderHint: typeof args.options.orderHint !== 'undefined',
        priority: typeof args.options.priority !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-i, --id <id>' },
      { option: '-t, --title [title]' },
      { option: '--planId [planId]' },
      { option: '--planTitle [planTitle]' },
      { option: '--ownerGroupId [ownerGroupId]' },
      { option: '--ownerGroupName [ownerGroupName]' },
      { option: '--bucketId [bucketId]' },
      { option: '--bucketName [bucketName]' },
      { option: '--startDateTime [startDateTime]' },
      { option: '--dueDateTime [dueDateTime]' },
      { option: '--percentComplete [percentComplete]' },
      { option: '--assignedToUserIds [assignedToUserIds]' },
      { option: '--assignedToUserNames [assignedToUserNames]' },
      { option: '--assigneePriority [assigneePriority]' },
      { option: '--description [description]' },
      { option: '--appliedCategories [appliedCategories]' },
      { option: '--orderHint [orderHint]' },
      { option: '--priority [priority]', autocomplete: taskPriority.priorityValues }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.bucketId && args.options.bucketName) {
          return 'Specify either bucketId or bucketName but not both';
        }

        if (args.options.bucketName && !args.options.planId && !args.options.planTitle) {
          return 'Specify either planId or planTitle when using bucketName';
        }

        if (args.options.bucketName && args.options.planId && args.options.planTitle) {
          return 'Specify either planId or planTitle when using bucketName but not both';
        }

        if (args.options.planTitle && !args.options.ownerGroupId && !args.options.ownerGroupName) {
          return 'Specify either ownerGroupId or ownerGroupName when using planTitle';
        }

        if (args.options.planTitle && args.options.ownerGroupId && args.options.ownerGroupName) {
          return 'Specify either ownerGroupId or ownerGroupName when using planTitle but not both';
        }

        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId as string)) {
          return `${args.options.ownerGroupId} is not a valid GUID`;
        }

        if (args.options.startDateTime && !validation.isValidISODateTime(args.options.startDateTime)) {
          return 'The startDateTime is not a valid ISO date string';
        }

        if (args.options.dueDateTime && !validation.isValidISODateTime(args.options.dueDateTime)) {
          return 'The dueDateTime is not a valid ISO date string';
        }

        if (args.options.percentComplete && isNaN(args.options.percentComplete)) {
          return `percentComplete is not a number`;
        }

        if (args.options.percentComplete && (args.options.percentComplete < 0 || args.options.percentComplete > 100)) {
          return `percentComplete should be between 0 and 100`;
        }

        if (args.options.assignedToUserIds && !validation.isValidGuidArray(args.options.assignedToUserIds.split(','))) {
          return 'assignedToUserIds contains invalid GUID';
        }

        if (args.options.assignedToUserIds && args.options.assignedToUserNames) {
          return 'Specify either assignedToUserIds or assignedToUserNames but not both';
        }

        if (args.options.appliedCategories && args.options.appliedCategories.split(',').filter(category => this.allowedAppliedCategories.indexOf(category.toLocaleLowerCase()) < 0).length !== 0) {
          return 'The appliedCategories contains invalid value. Specify either category1, category2, category3, category4, category5 and/or category6 as properties';
        }

        if (args.options.priority !== undefined) {
          if (typeof args.options.priority === "number") {
            if (isNaN(args.options.priority) || args.options.priority < 0 || args.options.priority > 10 || !Number.isInteger(args.options.priority)) {
              return 'priority should be an integer between 0 and 10.';
            }
          }
          else if (taskPriority.priorityValues.map(l => l.toLowerCase()).indexOf(args.options.priority.toString().toLowerCase()) === -1) {
            return `${args.options.priority} is not a valid priority value. Allowed values are ${taskPriority.priorityValues.join('|')}.`;
          }
        }

        return true;
      }
    );
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
      logger.log(result);
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
    options.appliedCategories.toLocaleLowerCase().split(',').forEach(x => categories[x] = true);
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

    // Hitting this section means assignedToUserNames won't be undefined
    const userNames = options.assignedToUserNames as string;
    const userArr: string[] = userNames.split(',').map(o => o.trim());
    let userIds: string[] = [];

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
    userIds = usersRes.map(res => res.value[0]?.id as string);

    // Find the members where no graph response was found
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
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/plans/${planId}/buckets?$select=id,name`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: PlannerBucket[] }>(requestOptions);

    const bucket: PlannerBucket | undefined = response.value.find(val => val.name === options.bucketName);

    if (!bucket) {
      throw 'The specified bucket does not exist';
    }

    return bucket.id as string;

  }

  private async getPlanId(options: Options): Promise<string> {
    if (options.planId) {
      return options.planId;
    }

    const groupId = await this.getGroupId(options);
    const plan = await planner.getPlanByTitle(options.planTitle!, groupId);
    return plan.id!;
  }

  private async getGroupId(options: Options): Promise<string> {
    if (options.ownerGroupId) {
      return options.ownerGroupId;
    }

    const group = await aadGroup.getGroupByDisplayName(options.ownerGroupName!);
    return group.id!;
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
      requestBody.percentComplete = options.percentComplete;
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

module.exports = new PlannerTaskSetCommand();