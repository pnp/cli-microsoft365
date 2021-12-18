import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { PlannerAssignment, PlannerTask, User } from '@microsoft/microsoft-graph-types';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  title?: string;
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
  conversationThreadId?: string;
  appliedCategories?: string;
  orderHint?: string;
}

class PlannerTaskSetCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_SET;
  }

  public get description(): string {
    return 'Updates a Microsoft Planner Task';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.title = typeof args.options.title !== 'undefined';
    telemetryProps.bucketId = typeof args.options.bucketId !== 'undefined';
    telemetryProps.bucketName = typeof args.options.bucketName !== 'undefined';
    telemetryProps.startDateTime = typeof args.options.startDateTime !== 'undefined';
    telemetryProps.dueDateTime = typeof args.options.dueDateTime !== 'undefined';
    telemetryProps.percentComplete = typeof args.options.percentComplete !== 'undefined';
    telemetryProps.assignedToUserIds = typeof args.options.assignedToUserIds !== 'undefined';
    telemetryProps.assignedToUserNames = typeof args.options.assignedToUserNames !== 'undefined';
    telemetryProps.assigneePriority = typeof args.options.assigneePriority !== 'undefined';
    telemetryProps.description = typeof args.options.description !== 'undefined';
    telemetryProps.conversationThreadId = typeof args.options.conversationThreadId !== 'undefined';
    telemetryProps.appliedCategories = typeof args.options.appliedCategories !== 'undefined';
    telemetryProps.orderHint = typeof args.options.orderHint !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let assignments: { [userId: string]: PlannerAssignment; } = {};

    this.generateUserAssignments(args)
      .then(genAssignments => {
        assignments = genAssignments;

        return this.getTaskEtag(args.options.id);
      })
      .then(etag => {
        const genAppliedcategories = this.generateAppliedCategories(args);
        const data = this.mapRequestBody(args.options, assignments, genAppliedcategories);

        const requestOptions: any = {
          url: `${this.resource}/v1.0/planner/tasks/${args.options.id}`,
          headers: {
            'accept': 'application/json;odata.metadata=none',
            'If-Match': etag,
            'Prefer': 'return=representation'
          },
          responseType: 'json',
          data: data
        };

        logger.log(requestOptions);

        return request.patch(requestOptions) as PlannerTask;
      })
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getTaskEtag(taskId: string): Promise<string> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/tasks/${encodeURIComponent(taskId)}`,
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
          return Promise.reject(`Error fetching task`);
        }

        return Promise.resolve(etag);
      });
  }

  private generateAppliedCategories(args: CommandArgs): { [category: string]: boolean } {
    const categories: { [category: string]: boolean } = {};

    if (args.options.assignedToUserIds) {
      args.options.assignedToUserIds.split(',').map(x => categories[x] = true);

      return categories;
    } 
    else {
      return categories;
    }
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
        return Promise.reject(`Cannot proceed with planner task update. The following users provided are invalid : ${invalidUsers.join(',')}`);
      }
      return Promise.resolve(userIds);
    });
  }

  private mapRequestBody(options: Options, assignments: { [userId: string]: PlannerAssignment }, appliedcategories: { [category: string]: boolean }): any {
    const requestBody: any = {};

    if (options.title) 
      requestBody.title = options.title;

    if (options.bucketId)
      requestBody.bucketId = options.bucketId;

    if (options.startDateTime)
      requestBody.startDateTime = options.startDateTime;

    if (options.dueDateTime)
      requestBody.dueDateTime = options.dueDateTime;

    if (options.percentComplete)
      requestBody.percentComplete = options.percentComplete;

    if (assignments.length != {})
      requestBody.assignments = assignments;

    if (options.assigneePriority)
      requestBody.assigneePriority = options.assigneePriority;

    if (options.conversationThreadId)
      requestBody.conversationThreadId = options.conversationThreadId;

    if (appliedcategories != {})
      requestBody.appliedCategories = options.assigneePriority;

    if (options.orderHint)
      requestBody.orderHint = options.orderHint;

    return requestBody;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: "-i, --id <id>"
      },
      {
        option: "-t, --title [title]"
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
        option: "--assigneePriority [assigneePriority]"
      },
      {
        option: "--description [description]"
      },
      {
        option: "--conversationThreadId [conversationThreadId]"
      },
      {
        option: "--appliedCategories [appliedCategories]"
      },
      {
        option: "--orderHint [orderHint]"
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {    
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
      return `percentComplete should be between 0 and 100`;
    }

    if (args.options.assignedToUserIds && !Utils.isValidGuidArray(args.options.assignedToUserIds.split(','))) {
      return 'assignedToUserIds contains invalid GUID';
    }

    if (args.options.assignedToUserIds && args.options.assignedToUserNames) {
      return 'Specify either assignedToUserIds or assignedToUserNames but not both';
    }
    if (args.options.appliedCategories && args.options.appliedCategories.split(',').filter(category => 
        category.toLocaleLowerCase() != "category1" &&
        category.toLocaleLowerCase() != "category2" &&
        category.toLocaleLowerCase() != "category3" &&
        category.toLocaleLowerCase() != "category4" &&
        category.toLocaleLowerCase() != "category5" &&
        category.toLocaleLowerCase() != "category6"
      ).length != 0) {
      return 'The appliedCategories contains invalid value. Specify either category1, category2, category3, category4, category5 and/or category6 as properties';
    }

    return true;
  }
}

module.exports = new PlannerTaskSetCommand();