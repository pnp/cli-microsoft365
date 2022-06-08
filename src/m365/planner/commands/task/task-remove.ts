import { Group, PlannerBucket, PlannerTask } from '@microsoft/microsoft-graph-types';
import { AxiosRequestConfig } from 'axios';
import * as os from 'os';
import { Cli, Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { odata, validation } from '../../../../utils';
import { planner } from '../../../../utils/planner';
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
  planTitle?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
  confirm?: boolean;
}

class PlannerTaskRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_REMOVE;
  }

  public get description(): string {
    return 'Removes the Microsoft Planner task from a plan';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeTask: () => void = (): void => {
      this
        .getTask(args.options)
        .then(task => {
          const requestOptions: AxiosRequestConfig = {
            url: `${this.resource}/v1.0/planner/tasks/${task.id}`,
            headers: {
              accept: 'application/json;odata.metadata=none',
              'if-match': (task as any)['@odata.etag']
            },
            responseType: 'json'
          };

          return request.delete(requestOptions);
        })
        .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeTask();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the task ${args.options.id || args.options.title}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeTask();
        }
      });
    }
  }

  private getTask(options: Options): Promise<PlannerTask> {
    const { id, title } = options;

    if (id) {
      const requestOptions: AxiosRequestConfig = {
        url: `${this.resource}/v1.0/planner/tasks/${id}`,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      return request.get<PlannerTask>(requestOptions);
    }

    return this
      .getBucketId(options)
      .then(bucketId => {        
        // $filter is not working on the buckets/{bucketId}/tasks endpoint, hence it is not being used.
        return odata.getAllItems<PlannerTask>(`${this.resource}/v1.0/planner/buckets/${bucketId}/tasks?$select=title,id`, 'minimal');
      })
      .then(tasks => {
        const filteredtasks = tasks.filter(b => title!.toLocaleLowerCase() === b.title!.toLocaleLowerCase());

        if (filteredtasks.length === 0) {
          return Promise.reject(`The specified task ${title} does not exist`);
        }

        if (filteredtasks.length > 1) {
          return Promise.reject(`Multiple tasks with title ${title} found: Please disambiguate: ${os.EOL}${filteredtasks.map(f => `- ${f.id}`).join(os.EOL)}`);
        }

        return Promise.resolve(filteredtasks[0]);
      });
  }

  private getBucketId(options: Options): Promise<string> {
    const { bucketId, bucketName } = options;

    if (bucketId) {
      return Promise.resolve(bucketId);
    }

    return this
      .getPlanId(options)
      .then(planId => {
        const requestOptions: AxiosRequestConfig = {
          url: `${this.resource}/v1.0/planner/plans/${planId}/buckets?$select=id,name`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        return request.get<{ value: PlannerBucket[] }>(requestOptions);
      })
      .then((buckets) => {
        const filteredBuckets = buckets.value.filter(b => bucketName!.toLocaleLowerCase() === b.name!.toLocaleLowerCase());

        if (filteredBuckets.length === 0) {
          return Promise.reject(`The specified bucket ${bucketName} does not exist`);
        }

        if (filteredBuckets.length > 1) {
          return Promise.reject(`Multiple buckets with name ${bucketName} found: Please disambiguate:${os.EOL}${filteredBuckets.map(f => `- ${f.id}`).join(os.EOL)}`);
        }

        return Promise.resolve(filteredBuckets[0].id!);
      });
  }

  private getPlanId(options: Options): Promise<string> {
    const { planId, planTitle } = options;

    if (planId) {
      return Promise.resolve(planId);
    }

    return this
      .getGroupId(options)
      .then(groupId => planner.getPlanByName(planTitle!, groupId))
      .then(plan => plan.id!);
  }

  private getGroupId(options: Options): Promise<string> {
    const { ownerGroupId, ownerGroupName } = options;

    if (ownerGroupId) {
      return Promise.resolve(ownerGroupId);
    }

    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/groups?$filter=displayName eq '${encodeURIComponent(ownerGroupName!)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: Group[] }>(requestOptions)
      .then(response => {
        if (response.value.length === 0) {
          return Promise.reject(`The specified owner group ${ownerGroupName} does not exist`);
        }

        if (response.value.length > 1) {
          return Promise.reject(`Multiple owner groups with name ${ownerGroupName} found: ${response.value.map(x => x.id)}`);
        }

        return Promise.resolve(response.value[0].id!);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '-i, --id [id]' },
      { option: '-t, --title [title]' },
      { option: '--bucketId [bucketId]' },
      { option: '--bucketName [bucketName]' },
      { option: '--planId [planId]' },
      { option: '--planTitle [planTitle]' },
      { option: '--ownerGroupId [ownerGroupId]' },
      { option: '--ownerGroupName [ownerGroupName]' },
      { option: '--confirm' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public optionSets(): string[][] | undefined {
    return [
      ['id', 'title']
    ];
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.id) {
      if (args.options.bucketId || args.options.bucketName || args.options.planId || args.options.planTitle || args.options.ownerGroupId || args.options.ownerGroupName) {
        return 'Don\'t specify bucketId,bucketName, planId, planTitle, ownerGroupId or ownerGroupName when using id';
      }
    }

    if (args.options.title) {
      if (!args.options.bucketId && !args.options.bucketName) {
        return 'Specify either bucketId or bucketName when using title';
      }

      if (args.options.bucketId && args.options.bucketName) {
        return 'Specify either bucketId or bucketName when using title but not both';
      }

      if (args.options.bucketName) {
        if (!args.options.planId && !args.options.planTitle) {
          return 'Specify either planId or planTitle when using bucketName';
        }

        if (args.options.planId && args.options.planTitle) {
          return 'Specify either planId or planTitle when using bucketName but not both';
        }
      }

      if (args.options.planTitle) {
        if (!args.options.ownerGroupId && !args.options.ownerGroupName) {
          return 'Specify either ownerGroupId or ownerGroupName when using planTitle';
        }
        if (args.options.ownerGroupId && args.options.ownerGroupName) {
          return 'Specify either ownerGroupId or ownerGroupName when using planTitle but not both';
        }
      }
      
      if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId as string)) {
        return `${args.options.ownerGroupId} is not a valid GUID`;
      }
    }
    return true;
  }
}

module.exports = new PlannerTaskRemoveCommand();