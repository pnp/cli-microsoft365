import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { PlannerTask, User, PlannerAssignment, PlannerPlan, Group } from '@microsoft/microsoft-graph-types';

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
    this.getPlanId(args)
      .then(planId => {
        args.options.planId = planId;
        return this.getBucketId(args, planId);
      })
      .then(bucketId => {
        args.options.bucketId = bucketId;
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
            planId: args.options.planId,
            bucketId: args.options.bucketId,
            title: args.options.title,
            startDateTime: args.options.startDateTime,
            dueDateTime: args.options.dueDateTime,
            percentComplete: args.options.percentComplete,
            assignments: assignments,
            orderHint: args.options.orderHint
          }
        };

        return request.post(requestOptions) as PlannerTask;
      })
      .then(newTask => { 
        const taskId = newTask.id as string;
  
        if (args.options.description) {
          return this.getTaskDetails(taskId, true)
            .then(taskDetails => {
              const requestOptionsTaskDetails: any = {
                url: `${this.resource}/v1.0/planner/tasks/${taskId}/details`,
                headers: {
                  'accept': 'application/json;odata.metadata=none',
                  'If-Match': taskDetails['@odata.etag']
                },
                responseType: 'json',
                data: {
                  description: args.options.description
                }
              };
      
              return request.patch(requestOptionsTaskDetails);              
            })
            .then(() => {
              return this.getTaskDetails(taskId);
            })
            .then(taskDetails => {
              return { ...newTask, ...taskDetails };
            });
        } 
        else {
          return newTask;
        }  
      })
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getTaskDetails(taskId: string, withOdata: boolean = false): Promise<any> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/tasks/${encodeURIComponent(taskId)}/details`,
      headers: {
        accept: `application/json${withOdata ? '' : ';odata.metadata=none'}`
      },
      responseType: 'json'
    };

    return request
      .get(requestOptions)
      .then((response: any) => {
        const taskDetails: any | undefined = response;

        if (!taskDetails) {
          return Promise.reject(`Error fetching task details`);
        }

        return Promise.resolve(taskDetails);
      });
  }

  private generateUserAssignments(args: CommandArgs): Promise<{ [userId: string]: PlannerAssignment; }> {
    const assignments: { [userId: string]: PlannerAssignment; } = {};

    if (args.options.assignedToUserNames) {
      return this.getUserIds(args.options.assignedToUserNames)
        .then((userIds) => {
          userIds.map(x => assignments[x] = {
            orderHint: " !"
          });

          return Promise.resolve(assignments);
        });
    }
    else if (args.options.assignedToUserIds) {
      args.options.assignedToUserIds.split(',').map(x => assignments[x] = {
        orderHint: " !"
      });

      return Promise.resolve(assignments);
    } 
    else {
      return Promise.resolve(assignments);
    }
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
      .get<{ value: { id: string; name: string; }[] }>(requestOptions)
      .then((response) => {
        const bucket: { id: string; name: string; } | undefined = response.value.find(val => val.name === args.options.bucketName);
    
        if (!bucket) {
          return Promise.reject(`The specified bucket does not exist`);
        }
    
        return Promise.resolve(bucket.id);
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

  private getUserIds(users: string): Promise<string[]> {
    const userArr: string[] = users.split(',').map(o => o.trim());
    let userIds: string[] = [];

    const promises: Promise<{ value: User[] }>[] = userArr.map(user => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/users?$filter=userPrincipalName eq '${Utils.encodeQueryParameter(user)}'&$select=id,userPrincipalName`,
        headers: {
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      return request.get(requestOptions);
    });

    return Promise.all(promises).then((usersRes: { value: User[] }[]): Promise<string[]> => {
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
      {
        option: "-t, --title <title>"
      },
      {
        option: "--planId [planId]"
      },
      {
        option: "--planName [planName]"
      },
      {
        option: "--ownerGroupId [ownerGroupId]"
      },
      {
        option: "--ownerGroupName [ownerGroupName]"
      },
      {
        option: "--bucketId [bucketId]"
      },
      {
        option: "--bucketName [bucketName]"
      },
      {
        option: "--startDateTime [startDateTime]"
      },
      {
        option: "--dueDateTime [dueDateTime]"
      },
      {
        option: "--percentComplete [percentComplete]"
      },
      {
        option: "--assignedToUserIds [assignedToUserIds]"
      },
      {
        option: "--assignedToUserNames [assignedToUserNames]"
      },
      {
        option: "--description [description]"
      },
      {
        option: "--orderHint [orderHint]"
      }
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

    if (args.options.ownerGroupId && !Utils.isValidGuid(args.options.ownerGroupId as string)) {
      return `${args.options.ownerGroupId} is not a valid GUID`;
    }
    
    if (!args.options.bucketId && !args.options.bucketName) {
      return 'Specify either bucketId or bucketName';
    }

    if (args.options.bucketId && args.options.bucketName) {
      return 'Specify either bucketId or bucketName but not both';
    }

    if (args.options.startDateTime && !Utils.isValidISODateTime(args.options.startDateTime)) {
      return 'The startDateTime is not a valid ISO date string';
    }

    if (args.options.dueDateTime && !Utils.isValidISODateTime(args.options.dueDateTime)) {
      return 'The dueDateTime is not a valid ISO date string';
    }

    if (args.options.percentComplete && isNaN(args.options.percentComplete)) {
      return `percentComplete is not a number`;
    }

    if (args.options.percentComplete && (args.options.percentComplete < 0 || args.options.percentComplete > 100)) {
      return `percentComplete should be between 0 and 100 `;
    }

    if (args.options.assignedToUserIds && !Utils.isValidGuidArray(args.options.assignedToUserIds.split(','))) {
      return 'assignedToUserIds contains invalid GUID';
    }

    if (args.options.assignedToUserIds && args.options.assignedToUserNames) {
      return 'Specify either assignedToUserIds or assignedToUserNames but not both';
    }

    return true;
  }
}

module.exports = new PlannerTaskAddCommand();