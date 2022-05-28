import { PlannerTask } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import { accessToken, odata, validation } from '../../../../utils';
import Auth from '../../../../Auth';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { planner } from '../../../../utils/planner';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  bucketId?: string;
  bucketName?: string;
  planId?: string;
  planName?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
}

interface BetaPlannerTask extends PlannerTask {
  priority?: number;
}

class PlannerTaskListCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_LIST;
  }

  public get description(): string {
    return 'Lists planner tasks in a bucket, plan, or tasks for the currently logged in user';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.bucketId = typeof args.options.bucketId !== 'undefined';
    telemetryProps.bucketName = typeof args.options.bucketName !== 'undefined';
    telemetryProps.planId = typeof args.options.planId !== 'undefined';
    telemetryProps.planName = typeof args.options.planName !== 'undefined';
    telemetryProps.ownerGroupId = typeof args.options.ownerGroupId !== 'undefined';
    telemetryProps.ownerGroupName = typeof args.options.ownerGroupName !== 'undefined';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'startDateTime', 'dueDateTime', 'completedDateTime'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (accessToken.isAppOnlyAccessToken(Auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.', logger, cb);
      return;
    }
    
    const bucketName: string | undefined = args.options.bucketName;
    let bucketId: string | undefined = args.options.bucketId;
    const planName: string | undefined = args.options.planName;
    let planId: string | undefined = args.options.planId;
    let taskItems: PlannerTask[] = [];

    if (bucketId || bucketName) {
      this
        .getBucketId(args)
        .then((retrievedBucketId: string): Promise<PlannerTask[]> => {
          bucketId = retrievedBucketId;
          return odata.getAllItems<PlannerTask>(`${this.resource}/v1.0/planner/buckets/${bucketId}/tasks`);
        })
        .then((tasks): Promise<BetaPlannerTask[]> => {
          taskItems = tasks;
          return odata.getAllItems<BetaPlannerTask>(`${this.resource}/beta/planner/buckets/${bucketId}/tasks`);
        })
        .then((betaTasks): void => {
          logger.log(this.mergeTaskPriority(taskItems, betaTasks));
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }
    else if (planId || planName) {
      this
        .getPlanId(args)
        .then((retrievedPlanId: string): Promise<PlannerTask[]> => {
          planId = retrievedPlanId;
          return odata.getAllItems<PlannerTask>(`${this.resource}/v1.0/planner/plans/${planId}/tasks`);
        })
        .then((tasks): Promise<BetaPlannerTask[]> => {
          taskItems = tasks;
          return odata.getAllItems<BetaPlannerTask>(`${this.resource}/beta/planner/plans/${planId}/tasks`);
        })
        .then((betaTasks): void => {
          logger.log(this.mergeTaskPriority(taskItems, betaTasks));
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }
    else {
      odata
        .getAllItems<PlannerTask>(`${this.resource}/v1.0/me/planner/tasks`)
        .then((tasks): Promise<BetaPlannerTask[]> => {
          taskItems = tasks;
          return odata.getAllItems<BetaPlannerTask>(`${this.resource}/beta/me/planner/tasks`);
        })
        .then((betaTasks): void => {
          logger.log(this.mergeTaskPriority(taskItems, betaTasks));
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }
  }

  private getBucketId(args: CommandArgs): Promise<string> {
    if (args.options.bucketId) {
      return Promise.resolve(encodeURIComponent(args.options.bucketId));
    }

    return this
      .getPlanId(args)
      .then((planId: string) => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/planner/plans/${planId}/buckets`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        return request.get<{ value: { id: string; name: string; }[] }>(requestOptions);
      })
      .then(response => {
        const bucket: { id: string; name: string; } | undefined = response.value.find(val => val.name === args.options.bucketName);

        if (!bucket) {
          return Promise.reject(`The specified bucket does not exist`);
        }

        return Promise.resolve(bucket.id);
      });
  }

  private getPlanId(args: CommandArgs): Promise<string> {
    if (args.options.planId) {
      return Promise.resolve(encodeURIComponent(args.options.planId));
    }

    return this
      .getGroupId(args)
      .then((groupId: string) => planner.getPlanByName(args.options.planName!, groupId))
      .then(plan => plan.id!);
  }

  private getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.ownerGroupId) {
      return Promise.resolve(encodeURIComponent(args.options.ownerGroupId));
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups?$filter=displayName eq '${encodeURIComponent(args.options.ownerGroupName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: { id: string; }[] }>(requestOptions)
      .then(response => {
        const group: { id: string; } | undefined = response.value[0];
        if (!group) {
          return Promise.reject(`The specified owner group does not exist`);
        }

        return Promise.resolve(group.id);
      });
  }

  private mergeTaskPriority(taskItems: PlannerTask[], betaTaskItems: BetaPlannerTask[]): BetaPlannerTask[] {
    const findBetaTask = (id: string) => betaTaskItems.find(task => task.id === id);

    taskItems.forEach(task => {
      const betaTaskItem = findBetaTask(task.id as string);
      if (betaTaskItem) {
        const { priority } = betaTaskItem;
        Object.assign(task, { priority });
      }
    });

    return taskItems;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
        option: '--planName [planName]'
      },
      {
        option: '--ownerGroupId [ownerGroupId]'
      },
      {
        option: '--ownerGroupName [ownerGroupName]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.bucketId && args.options.bucketName) {
      return 'To retrieve tasks from a bucket, specify bucketId or bucketName, but not both';
    }

    if (args.options.bucketName && !args.options.planId && !args.options.planName) {
      return 'Specify either planId or planName when using bucketName';
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

    return true;
  }
}

module.exports = new PlannerTaskListCommand();