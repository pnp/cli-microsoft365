import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { accessToken, validation } from '../../../../utils';
import { aadGroup } from '../../../../utils/aadGroup';
import { planner } from '../../../../utils/planner';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  planId?: string;
  planName?: string;
  planTitle?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
  orderHint?: string;
}

class PlannerBucketAddCommand extends GraphCommand {
  public get name(): string {
    return commands.BUCKET_ADD;
  }

  public get description(): string {
    return 'Adds a new Microsoft Planner bucket';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'planId', 'orderHint'];
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
        planId: typeof args.options.planId !== 'undefined',
        planName: typeof args.options.planName !== 'undefined',
        planTitle: typeof args.options.planTitle !== 'undefined',
        ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
        ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
        orderHint: typeof args.options.orderHint !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: "--planId [planId]"
      },
      {
        option: "--planName [planName]"
      },
      {
        option: "--planTitle [planTitle]"
      },
      {
        option: "--ownerGroupId [ownerGroupId]"
      },
      {
        option: "--ownerGroupName [ownerGroupName]"
      },
      {
        option: "--orderHint [orderHint]"
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!args.options.planId && !args.options.planName && !args.options.planTitle) {
	      return 'Specify either planId or planTitle';
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
      .then((planId: string): Promise<any> => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/planner/buckets`,
          headers: {
            'accept': 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: {
            name: args.options.name,
            planId: planId,
            orderHint: args.options.orderHint
          }
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getPlanId(args: CommandArgs): Promise<string> {
    if (args.options.planId) {
      return Promise.resolve(args.options.planId);
    }

    return this
      .getGroupId(args)
      .then(groupId => planner.getPlanByTitle(args.options.planTitle!, groupId))
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
}

module.exports = new PlannerBucketAddCommand();