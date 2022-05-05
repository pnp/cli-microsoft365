import { Group, PlannerBucket, PlannerPlan, PlannerTask } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  bucketId?: string;
  bucketName?: string;
  planId?: string;
  planName?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
}

class PlannerTaskGetCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_GET;
  }

  public get description(): string {
    return 'Retrieve the the specified planner task';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getTaskId(args.options)
      .then(taskId => {
        const requestOptions: any = {
          url: `${this.resource}/beta/planner/tasks/${encodeURIComponent(taskId)}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        }; 

        return request.get<{ value: PlannerTask }>(requestOptions);
      })
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getTaskId(options: Options): Promise<string> {
    if (options.id) {
      return Promise.resolve(options.id);
    }

    return this
      .getBucketId(options)
      .then(bucketId => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/planner/buckets/${bucketId}/tasks?$select=id,title`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        return request.get<{ value: PlannerTask[] }>(requestOptions);
      })
      .then((response) => {
        const title = options.title as string;
        const tasks: PlannerTask[] | undefined = response.value.filter(val => val.title?.toLocaleLowerCase() === title.toLocaleLowerCase());
        
        if (!tasks.length) {
          return Promise.reject(`The specified task ${options.title} does not exist`);
        }

        if (tasks.length > 1) {
          return Promise.reject(`Multiple tasks with title ${options.title} found: ${tasks.map(x => x.id)}`);
        }

        return Promise.resolve(tasks[0].id as string);
      });
  }

  private getBucketId(options: Options): Promise<string> {
    if (options.bucketId) {
      return Promise.resolve(options.bucketId);
    }

    return this
      .getPlanId(options)
      .then(planId => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/planner/plans/${planId}/buckets?$select=id,name`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        return request.get<{ value: PlannerBucket[] }>(requestOptions);
      })
      .then((response) => {
        const bucketName = options.bucketName as string;
        const buckets: PlannerBucket[] | undefined = response.value.filter(val => val.name?.toLocaleLowerCase() === bucketName.toLocaleLowerCase());
        
        if (!buckets.length) {
          return Promise.reject(`The specified bucket ${options.bucketName} does not exist`);
        }

        if (buckets.length > 1) {
          return Promise.reject(`Multiple buckets with name ${options.bucketName} found: ${buckets.map(x => x.id)}`);
        }

        return Promise.resolve(buckets[0].id as string);
      });
  }

  private getPlanId(options: Options): Promise<string> {
    if (options.planId) {
      return Promise.resolve(options.planId);
    }

    return this
      .getGroupId(options)
      .then((groupId: string) => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/planner/plans?$filter=owner eq '${groupId}'&$select=id,title`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        return request.get<{ value: PlannerPlan[] }>(requestOptions);
      })
      .then((response) => {
        const planName = options.planName as string;
        const plans: PlannerPlan[] | undefined = response.value.filter(val => val.title?.toLocaleLowerCase() === planName.toLocaleLowerCase());
        
        if (!plans.length) {
          return Promise.reject(`The specified plan ${options.planName} does not exist`);
        }

        if (plans.length > 1) {
          return Promise.reject(`Multiple plans with name ${options.planName} found: ${plans.map(x => x.id)}`);
        }

        return Promise.resolve(plans[0].id as string);
      });
  }

  private getGroupId(options: Options): Promise<string> {
    if (options.ownerGroupId) {
      return Promise.resolve(options.ownerGroupId);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups?$filter=displayName eq '${encodeURIComponent(options.ownerGroupName as string)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: Group[] }>(requestOptions)
      .then(response => {
        const groups: Group[] | undefined = response.value;
        
        if (!groups.length) {
          return Promise.reject(`The specified ownerGroup ${options.ownerGroupName} does not exist`);
        }

        if (groups.length > 1) {
          return Promise.reject(`Multiple ownerGroups with name ${options.ownerGroupName} found: ${groups.map(x => x.id)}`);
        }

        return Promise.resolve(groups[0].id as string);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '-i, --id [id]' },
      { option: '-t, --title [title]' },
      { option: '--bucketId [bucketId]' },
      { option: '--bucketName [bucketName]' },
      { option: '--planId [planId]' },
      { option: '--planName [planName]' },
      { option: '--ownerGroupId [ownerGroupId]' },
      { option: '--ownerGroupName [ownerGroupName]' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
  
  public validate(args: CommandArgs): boolean | string {
    if (args.options.title && !args.options.bucketId && !args.options.bucketName) {
      return 'Specify either bucketId or bucketName when using title';
    }

    if (args.options.title && args.options.bucketId && args.options.bucketName) {
      return 'Specify either bucketId or bucketName when using title but not both';
    }

    if (args.options.bucketName && !args.options.planId && !args.options.planName) {
      return 'Specify either planId or planName when using bucketName';
    }

    if (args.options.bucketName && args.options.planId && args.options.planName) {
      return 'Specify either planId or planName when using bucketName but not both';
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

module.exports = new PlannerTaskGetCommand();
