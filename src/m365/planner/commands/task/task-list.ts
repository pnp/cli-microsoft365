import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import { Task } from '../../Task';

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

class PlannerTaskListCommand extends GraphItemsListCommand<Task> {
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
    const bucketId: string | undefined = args.options.bucketId;
    const bucketName: string | undefined = args.options.bucketName;
    const planId: string | undefined = args.options.planId;
    const planName: string | undefined = args.options.planName;

    if (bucketId || bucketName) {
      this
        .getBucketId(args)
        .then((bucketId: string): Promise<void> => this.getAllItems(`${this.resource}/v1.0/planner/buckets/${bucketId}/tasks`, logger, true))
        .then((): void => {
          logger.log(this.items);
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }
    else if (planId || planName) {
      this
        .getPlanId(args)
        .then((planId: string): Promise<void> => this.getAllItems(`${this.resource}/v1.0/planner/plans/${planId}/tasks`, logger, true))
        .then((): void => {
          logger.log(this.items);
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }
    else {
      this
        .getAllItems(`${this.resource}/v1.0/me/planner/tasks`, logger, true)
        .then((): void => {
          logger.log(this.items);
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
      .then((groupId: string) => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/planner/plans?$filter=(owner eq '${groupId}')`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        return request.get<{ value: { id: string; title: string; }[] }>(requestOptions);
      })
      .then(response => {
        const plan: { id: string; title: string; } | undefined = response.value.find(val => val.title === args.options.planName);

        if (!plan) {
          return Promise.reject(`The specified plan does not exist`);
        }

        return Promise.resolve(plan.id);
      });
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

    if (args.options.ownerGroupId && !Utils.isValidGuid(args.options.ownerGroupId as string)) {
      return `${args.options.ownerGroupId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new PlannerTaskListCommand();