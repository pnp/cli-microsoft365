import { PlannerBucket, PlannerTask } from '@microsoft/microsoft-graph-types';
import * as os from 'os';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { aadGroup } from '../../../../utils/aadGroup';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
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
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        bucketId: typeof args.options.bucketId !== 'undefined',
        bucketName: typeof args.options.bucketName !== 'undefined',
        planId: typeof args.options.planId !== 'undefined',
        planTitle: typeof args.options.planTitle !== 'undefined',
        ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
        ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
        confirm: !!args.options.confirm
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-i, --id [id]' },
      { option: '-t, --title [title]' },
      { option: '--bucketId [bucketId]' },
      { option: '--bucketName [bucketName]' },
      { option: '--planId [planId]' },
      { option: '--planTitle [planTitle]' },
      { option: '--ownerGroupId [ownerGroupId]' },
      { option: '--ownerGroupName [ownerGroupName]' },
      { option: '--confirm' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
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
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'title'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeTask = async (): Promise<void> => {
      try {
        const task = await this.getTask(args.options);

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/planner/tasks/${task.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'if-match': (task as any)['@odata.etag']
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await removeTask();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the task ${args.options.id || args.options.title}?`
      });

      if (result.continue) {
        await removeTask();
      }
    }
  }

  private async getTask(options: Options): Promise<PlannerTask> {
    const { id, title } = options;

    if (id) {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/planner/tasks/${id}`,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      return await request.get<PlannerTask>(requestOptions);
    }

    const bucketId = await this.getBucketId(options);

    // $filter is not working on the buckets/{bucketId}/tasks endpoint, hence it is not being used.
    const tasks = await odata.getAllItems<PlannerTask>(`${this.resource}/v1.0/planner/buckets/${bucketId}/tasks?$select=title,id`, 'minimal');
    const filteredtasks = tasks.filter(b => title!.toLocaleLowerCase() === b.title!.toLocaleLowerCase());

    if (filteredtasks.length === 0) {
      throw `The specified task ${title} does not exist`;
    }

    if (filteredtasks.length > 1) {
      throw `Multiple tasks with title ${title} found: Please disambiguate: ${os.EOL}${filteredtasks.map(f => `- ${f.id}`).join(os.EOL)}`;
    }

    return filteredtasks[0];
  }

  private async getBucketId(options: Options): Promise<string> {
    const { bucketId, bucketName } = options;

    if (bucketId) {
      return bucketId;
    }

    const planId = await this.getPlanId(options);
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/plans/${planId}/buckets?$select=id,name`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const buckets = await request.get<{ value: PlannerBucket[] }>(requestOptions);
    const filteredBuckets = buckets.value.filter(b => bucketName!.toLocaleLowerCase() === b.name!.toLocaleLowerCase());

    if (filteredBuckets.length === 0) {
      throw `The specified bucket ${bucketName} does not exist`;
    }

    if (filteredBuckets.length > 1) {
      throw `Multiple buckets with name ${bucketName} found: Please disambiguate:${os.EOL}${filteredBuckets.map(f => `- ${f.id}`).join(os.EOL)}`;
    }

    return filteredBuckets[0].id!;
  }

  private async getPlanId(options: Options): Promise<string> {
    const { planId, planTitle } = options;

    if (planId) {
      return planId;
    }

    const groupId = await this.getGroupId(options);
    const plan = await planner.getPlanByTitle(planTitle!, groupId);
    return plan.id!;
  }

  private async getGroupId(options: Options): Promise<string> {
    const { ownerGroupId, ownerGroupName } = options;

    if (ownerGroupId) {
      return ownerGroupId;
    }

    const group = await aadGroup.getGroupByDisplayName(ownerGroupName!);
    return group.id!;
  }
}

module.exports = new PlannerTaskRemoveCommand();