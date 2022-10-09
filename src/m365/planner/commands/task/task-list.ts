import { PlannerTask } from '@microsoft/microsoft-graph-types';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { accessToken } from '../../../../utils/accessToken';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import { aadGroup } from '../../../../utils/aadGroup';
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
  planTitle?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
}

class PlannerTaskListCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_LIST;
  }

  public get description(): string {
    return 'Lists planner tasks in a bucket, plan, or tasks for the currently logged in user';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'startDateTime', 'dueDateTime', 'completedDateTime'];
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
        bucketId: typeof args.options.bucketId !== 'undefined',
        bucketName: typeof args.options.bucketName !== 'undefined',
        planId: typeof args.options.planId !== 'undefined',
        planName: typeof args.options.planName !== 'undefined',
        planTitle: typeof args.options.planTitle !== 'undefined',
        ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
        ownerGroupName: typeof args.options.ownerGroupName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
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
        option: '--planTitle [planTitle]'
      },
      {
        option: '--ownerGroupId [ownerGroupId]'
      },
      {
        option: '--ownerGroupName [ownerGroupName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.bucketId && args.options.bucketName) {
	      return 'To retrieve tasks from a bucket, specify bucketId or bucketName, but not both';
	    }

	    if (args.options.bucketName && !args.options.planId && !args.options.planName && !args.options.planTitle) {
	      return 'Specify either planId or planTitle when using bucketName';
	    }

	    if (args.options.planId && (args.options.planName || args.options.planTitle)) {
	      return 'Specify either planId or planTitle but not both';
	    }

	    if ((args.options.planName || args.options.planTitle) && !args.options.ownerGroupId && !args.options.ownerGroupName) {
	      return 'Specify either ownerGroupId or ownerGroupName when using planTitle';
	    }

	    if ((args.options.planName || args.options.planTitle) && args.options.ownerGroupId && args.options.ownerGroupName) {
	      return 'Specify either ownerGroupId or ownerGroupName when using planTitle but not both';
	    }

	    if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId as string)) {
	      return `${args.options.ownerGroupId} is not a valid GUID`;
	    }

	    return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.planName) {
      args.options.planTitle = args.options.planName;

      this.warn(logger, `Option 'planName' is deprecated. Please use 'planTitle' instead`);
    }

    if (accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.');
      return;
    }

    const bucketName: string | undefined = args.options.bucketName;
    let bucketId: string | undefined = args.options.bucketId;
    const planTitle: string | undefined = args.options.planTitle;
    let planId: string | undefined = args.options.planId;
    let taskItems: PlannerTask[] = [];

    if (bucketId || bucketName) {
      try {
        bucketId = await this.getBucketId(args);
        taskItems = await odata.getAllItems<PlannerTask>(`${this.resource}/v1.0/planner/buckets/${bucketId}/tasks`);
        const betaTasks = await odata.getAllItems<PlannerTask>(`${this.resource}/beta/planner/buckets/${bucketId}/tasks`);

        logger.log(this.mergeTaskPriority(taskItems, betaTasks));
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    }
    else if (planId || planTitle) {
      try {
        planId = await this.getPlanId(args);
        taskItems = await odata.getAllItems<PlannerTask>(`${this.resource}/v1.0/planner/plans/${planId}/tasks`);
        const betaTasks = await odata.getAllItems<PlannerTask>(`${this.resource}/beta/planner/plans/${planId}/tasks`);

        logger.log(this.mergeTaskPriority(taskItems, betaTasks));
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    }
    else {
      try {
        taskItems = await odata.getAllItems<PlannerTask>(`${this.resource}/v1.0/me/planner/tasks`);
        const betaTasks = await odata.getAllItems<PlannerTask>(`${this.resource}/beta/me/planner/tasks`);
        logger.log(this.mergeTaskPriority(taskItems, betaTasks));
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
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
      .then((groupId: string) => planner.getPlanByTitle(args.options.planTitle!, groupId))
      .then(plan => plan.id!);
  }

  private getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.ownerGroupId) {
      return Promise.resolve(encodeURIComponent(args.options.ownerGroupId));
    }

    return aadGroup
      .getGroupByDisplayName(args.options.ownerGroupName!)
      .then(group => group.id!);
  }

  private mergeTaskPriority(taskItems: PlannerTask[], betaTaskItems: PlannerTask[]): PlannerTask[] {
    const findBetaTask = (id: string): PlannerTask | undefined => betaTaskItems.find(task => task.id === id);

    taskItems.forEach(task => {
      const betaTaskItem = findBetaTask(task.id as string);
      if (betaTaskItem) {
        const { priority } = betaTaskItem;
        Object.assign(task, { priority });
      }
    });

    return taskItems;
  }
}

module.exports = new PlannerTaskListCommand();