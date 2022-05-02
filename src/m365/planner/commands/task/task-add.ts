import { Group, PlannerBucket, PlannerPlan, PlannerTask, PlannerTaskDetails, User } from '@microsoft/microsoft-graph-types';
import Auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { accessToken, formatting, validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  planId?: string;
  planName?: string;
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
  description?: string;
  orderHint?: string;
}

class PlannerTaskAddCommand extends GraphCommand {
  private planId: string | undefined;
  private bucketId: string | undefined;

  public get name(): string {
    return commands.TASK_ADD;
  }

  public get description(): string {
    return 'Adds a new Microsoft Planner Task';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.planId = typeof args.options.planId !== 'undefined';
    telemetryProps.planName = typeof args.options.planName !== 'undefined';
    telemetryProps.ownerGroupId = typeof args.options.ownerGroupId !== 'undefined';
    telemetryProps.ownerGroupName = typeof args.options.ownerGroupName !== 'undefined';
    telemetryProps.bucketId = typeof args.options.bucketId !== 'undefined';
    telemetryProps.bucketName = typeof args.options.bucketName !== 'undefined';
    telemetryProps.startDateTime = typeof args.options.startDateTime !== 'undefined';
    telemetryProps.dueDateTime = typeof args.options.dueDateTime !== 'undefined';
    telemetryProps.percentComplete = typeof args.options.percentComplete !== 'undefined';
    telemetryProps.assignedToUserIds = typeof args.options.assignedToUserIds !== 'undefined';
    telemetryProps.assignedToUserNames = typeof args.options.assignedToUserNames !== 'undefined';
    telemetryProps.description = typeof args.options.description !== 'undefined';
    telemetryProps.orderHint = typeof args.options.orderHint !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (accessToken.isAppOnlyAccessToken(Auth.service.accessTokens[this.resource].accessToken)) {
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
        const requestOptions: any = {
          url: `${this.resource}/v1.0/planner/tasks`,
          headers: {
            'accept': 'application/json;odata.metadata=none'
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
            orderHint: args.options.orderHint
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
      .then((response: any) => {
        const etag: string | undefined = response ? response['@odata.etag'] : undefined;

        if (!etag) {
          return Promise.reject(`Error fetching task details`);
        }

        return Promise.resolve(etag);
      });
  }

  private updateTaskDetails(options: Options, newTask: PlannerTask): Promise<PlannerTask & PlannerTaskDetails> {
    const taskId = newTask.id as string;

    if (!options.description) {
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
            description: options.description
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
      .then((groupId: string) => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/planner/plans?$filter=(owner eq '${groupId}')`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        return request.get<{ value: PlannerPlan[] }>(requestOptions);
      }).then((response) => {
        const plan: PlannerPlan | undefined = response.value.find(val => val.title === args.options.planName);

        if (!plan) {
          return Promise.reject(`The specified plan does not exist`);
        }

        return Promise.resolve(plan.id as string);
      });
  }

  private getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.ownerGroupId) {
      return Promise.resolve(args.options.ownerGroupId);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups?$filter=displayName eq '${encodeURIComponent(args.options.ownerGroupName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: Group[] }>(requestOptions)
      .then(response => {
        const group: Group | undefined = response.value[0];

        if (!group) {
          return Promise.reject(`The specified owner group does not exist`);
        }

        return Promise.resolve(group.id as string);
      });
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '-t, --title <title>' },
      { option: '--planId [planId]' },
      { option: '--planName [planName]' },
      { option: '--ownerGroupId [ownerGroupId]' },
      { option: '--ownerGroupName [ownerGroupName]' },
      { option: '--bucketId [bucketId]' },
      { option: '--bucketName [bucketName]' },
      { option: '--startDateTime [startDateTime]' },
      { option: '--dueDateTime [dueDateTime]' },
      { option: '--percentComplete [percentComplete]' },
      { option: '--assignedToUserIds [assignedToUserIds]' },
      { option: '--assignedToUserNames [assignedToUserNames]' },
      { option: '--description [description]' },
      { option: '--orderHint [orderHint]' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.planId && !args.options.planName) {
      return 'Specify either planId or planName';
    }

    if (args.options.planId && args.options.planName) {
      return 'Specify either planId or planName but not both';
    }

    if (args.options.planName && !args.options.ownerGroupId && !args.options.ownerGroupName) {
      return 'Specify either ownerGroupId or ownerGroupName when using planName';
    }

    if (args.options.planName && args.options.ownerGroupId && args.options.ownerGroupName) {
      return 'Specify either ownerGroupId or ownerGroupName when using planName but not both';
    }

    if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId as string)) {
      return `${args.options.ownerGroupId} is not a valid GUID`;
    }

    if (!args.options.bucketId && !args.options.bucketName) {
      return 'Specify either bucketId or bucketName';
    }

    if (args.options.bucketId && args.options.bucketName) {
      return 'Specify either bucketId or bucketName but not both';
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
      return `percentComplete should be between 0 and 100 `;
    }

    if (args.options.assignedToUserIds && !validation.isValidGuidArray(args.options.assignedToUserIds.split(','))) {
      return 'assignedToUserIds contains invalid GUID';
    }

    if (args.options.assignedToUserIds && args.options.assignedToUserNames) {
      return 'Specify either assignedToUserIds or assignedToUserNames but not both';
    }

    return true;
  }
}

module.exports = new PlannerTaskAddCommand();