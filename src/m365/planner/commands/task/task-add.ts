import { PlannerBucket, PlannerTask, PlannerTaskDetails, User } from '@microsoft/microsoft-graph-types';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { accessToken, formatting, validation } from '../../../../utils';
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
  planId?: string;
  planName?: string;
  planTitle?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
  bucketId?: string;
  bucketName?: string;
  title: string;
  startDateTime?: string;
  dueDateTime?: string;
  percentComplete?: number;
  assignedToUserIds?: string;
  assignedToUserNames?: string;
  assigneePriority?: string;
  description?: string;
  appliedCategories?: string;
  previewType?: string;
  orderHint?: string;
  priority?: number | string;
}

class PlannerTaskAddCommand extends GraphCommand {
  private planId: string | undefined;
  private bucketId: string | undefined;
  private allowedAppliedCategories: string[] = ['category1', 'category2', 'category3', 'category4', 'category5', 'category6'];
  private allowedPreviewTypes: string[] = ['automatic', 'nopreview', 'checklist', 'description', 'reference'];

  public get name(): string {
    return commands.TASK_ADD;
  }

  public get description(): string {
    return 'Adds a new Microsoft Planner Task';
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
        planId: typeof args.options.planId !== 'undefined',
        planName: typeof args.options.planName !== 'undefined',
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
        previewType: typeof args.options.previewType !== 'undefined',
        orderHint: typeof args.options.orderHint !== 'undefined',
        priority: typeof args.options.priority !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-t, --title <title>' },
      { option: '--planId [planId]' },
      { option: '--planName [planName]' },
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
      {
        option: '--appliedCategories [appliedCategories]',
        autocomplete: this.allowedAppliedCategories
      },
      {
        option: '--previewType [previewType]',
        autocomplete: this.allowedPreviewTypes
      },
      { option: '--orderHint [orderHint]' },
      { option: '--priority [priority]', autocomplete: taskPriority.priorityValues }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if ((args.options.planName || args.options.planTitle) && !args.options.ownerGroupId && !args.options.ownerGroupName) {
          return 'Specify either ownerGroupId or ownerGroupName when using planTitle';
        }

        if ((args.options.planName || args.options.planTitle) && args.options.ownerGroupId && args.options.ownerGroupName) {
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
          return `The appliedCategories contains invalid value. Specify either ${this.allowedAppliedCategories.join(', ')} as properties`;
        }

        if (args.options.previewType && this.allowedPreviewTypes.indexOf(args.options.previewType.toLocaleLowerCase()) === -1) {
          return `${args.options.previewType} is not a valid preview type value. Allowed values are ${this.allowedPreviewTypes.join(', ')}`;
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

  #initOptionSets(): void {
    this.optionSets.push(
      ['planId', 'planTitle'],
      ['bucketId', 'bucketName']
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (args.options.planName) {
      args.options.planTitle = args.options.planName;

      this.warn(logger, `Option 'planName' is deprecated. Please use 'planTitle' instead`);
    }

    if (accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.', logger, cb);
      return;
    }

    this
      .getPlanId(args)
      .then(planId => {
        this.planId = planId;
        return this.getBucketId(args, planId);
      })
      .then(bucketId => {
        this.bucketId = bucketId;
        return this.generateUserAssignments(args);
      })
      .then(assignments => {
        const appliedCategories = this.generateAppliedCategories(args.options);

        const requestOptions: any = {
          url: `${this.resource}/v1.0/planner/tasks`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: {
            planId: this.planId,
            bucketId: this.bucketId,
            title: args.options.title,
            startDateTime: args.options.startDateTime,
            dueDateTime: args.options.dueDateTime,
            percentComplete: args.options.percentComplete,
            assignments: assignments,
            orderHint: args.options.orderHint,
            assigneePriority: args.options.assigneePriority,
            appliedCategories: appliedCategories,
            priority: taskPriority.getPriorityValue(args.options.priority)
          }
        };

        return request.post<PlannerTask>(requestOptions);
      })
      .then(newTask => this.updateTaskDetails(args.options, newTask))
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getTaskDetailsEtag(taskId: string): Promise<string> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/tasks/${encodeURIComponent(taskId)}/details`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    return request
      .get(requestOptions)
      .then((response: any) => response['@odata.etag']);
  }

  private generateAppliedCategories(options: Options): AppliedCategories {
    if (!options.appliedCategories) {
      return {};
    }

    const categories: AppliedCategories = {};
    options.appliedCategories.toLocaleLowerCase().split(',').forEach(x => categories[x] = true);
    return categories;
  }

  private updateTaskDetails(options: Options, newTask: PlannerTask): Promise<PlannerTask & PlannerTaskDetails> {
    const taskId = newTask.id as string;

    if (!options.description && !options.previewType) {
      return Promise.resolve(newTask);
    }

    return this
      .getTaskDetailsEtag(taskId)
      .then(etag => {
        const requestOptionsTaskDetails: any = {
          url: `${this.resource}/v1.0/planner/tasks/${taskId}/details`,
          headers: {
            'accept': 'application/json;odata.metadata=none',
            'If-Match': etag,
            'Prefer': 'return=representation'
          },
          responseType: 'json',
          data: {
            description: options.description,
            previewType: options.previewType
          }
        };

        return request.patch(requestOptionsTaskDetails);
      })
      .then(taskDetails => {
        return { ...newTask, ...taskDetails as PlannerTaskDetails };
      });
  }

  private generateUserAssignments(args: CommandArgs): Promise<{ [userId: string]: { [property: string]: string }; }> {
    const assignments: { [userId: string]: { [property: string]: string } } = {};

    if (!args.options.assignedToUserIds && !args.options.assignedToUserNames) {
      return Promise.resolve(assignments);
    }

    return this
      .getUserIds(args.options)
      .then((userIds) => {
        userIds.map(x => assignments[x] = {
          '@odata.type': '#microsoft.graph.plannerAssignment',
          orderHint: ' !'
        });

        return Promise.resolve(assignments);
      });
  }

  private getBucketId(args: CommandArgs, planId: string): Promise<string> {
    if (args.options.bucketId) {
      return Promise.resolve(args.options.bucketId);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/plans/${planId}/buckets`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: PlannerBucket[] }>(requestOptions)
      .then((response) => {
        const bucket: PlannerBucket | undefined = response.value.find(val => val.name === args.options.bucketName);

        if (!bucket) {
          return Promise.reject(`The specified bucket does not exist`);
        }

        return Promise.resolve(bucket.id as string);
      });
  }

  private getPlanId(args: CommandArgs): Promise<string> {
    if (args.options.planId) {
      return Promise.resolve(args.options.planId);
    }

    return this
      .getGroupId(args)
      .then((groupId: string) => planner.getPlanByTitle(args.options.planTitle!, groupId))
      .then(plan => plan.id!);
  }

  private getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.ownerGroupId) {
      return Promise.resolve(args.options.ownerGroupId);
    }

    return aadGroup
      .getGroupByDisplayName(args.options.ownerGroupName!)
      .then(group => group.id!);
  }

  private getUserIds(options: Options): Promise<string[]> {
    if (options.assignedToUserIds) {
      return Promise.resolve(options.assignedToUserIds.split(','));
    }

    // Hitting this section means assignedToUserNames won't be undefined
    const userNames = options.assignedToUserNames as string;
    const userArr: string[] = userNames.split(',').map(o => o.trim());
    let userIds: string[] = [];

    const promises: Promise<{ value: User[] }>[] = userArr.map(user => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user)}'&$select=id,userPrincipalName`,
        headers: {
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      return request.get(requestOptions);
    });

    return Promise
      .all(promises)
      .then((usersRes: { value: User[] }[]): Promise<string[]> => {
        let userUpns: string[] = [];

        userUpns = usersRes.map(res => res.value[0]?.userPrincipalName as string);
        userIds = usersRes.map(res => res.value[0]?.id as string);

        // Find the members where no graph response was found
        const invalidUsers = userArr.filter(user => !userUpns.some((upn) => upn?.toLowerCase() === user.toLowerCase()));

        if (invalidUsers && invalidUsers.length > 0) {
          return Promise.reject(`Cannot proceed with planner task creation. The following users provided are invalid : ${invalidUsers.join(',')}`);
        }

        return Promise.resolve(userIds);
      });
  }
}

module.exports = new PlannerTaskAddCommand();