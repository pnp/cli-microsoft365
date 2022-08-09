import { PlannerBucket } from '@microsoft/microsoft-graph-types';
import { AxiosRequestConfig } from 'axios';
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
  id?: string;
  name?: string;
  planId?: string;
  planName?: string;
  planTitle?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
  newName?: string;
  orderHint?: string;
}

class PlannerBucketSetCommand extends GraphCommand {
  public get name(): string {
    return commands.BUCKET_SET;
  }

  public get description(): string {
    return 'Updates a Microsoft Planner bucket';
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
        name: typeof args.options.name !== 'undefined',
        planId: typeof args.options.planId !== 'undefined',
        planName: typeof args.options.planName !== 'undefined',
        planTitle: typeof args.options.planTitle !== 'undefined',
        ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
        ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
        newName: typeof args.options.newName !== 'undefined',
        orderHint: typeof args.options.orderHint !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '--planId [planId]'
      },
      {
        option: '--planName [planName]'
      },
      {
        option: "--planTitle [planTitle]"
      },
      {
        option: '--ownerGroupId [ownerGroupId]'
      },
      {
        option: '--ownerGroupName [ownerGroupName]'
      },
      {
        option: '--newName [newName]'
      },
      {
        option: '--orderHint [orderHint]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id) {
	      if (args.options.planId || args.options.planName || args.options.planTitle  || args.options.ownerGroupId || args.options.ownerGroupName) {
	        return 'Don\'t specify planId, planTitle, ownerGroupId or ownerGroupName when using id';
	      }
	    }

	    if (args.options.name) {
	      if (!args.options.planId && !args.options.planName && !args.options.planTitle) {
	        return 'Specify either planId or planTitle when using name';
	      }

	      if (args.options.planId && (args.options.planName || args.options.planTitle)) {
	        return 'Specify either planId or planTitle when using name but not both';
	      }

	      if (args.options.planName || args.options.planTitle) {
	        if (!args.options.ownerGroupId && !args.options.ownerGroupName) {
	          return 'Specify either ownerGroupId or ownerGroupName when using planTitle';
	        }

	        if (args.options.ownerGroupId && args.options.ownerGroupName) {
	          return 'Specify either ownerGroupId or ownerGroupName when using planTitle but not both';
	        }

	        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
	          return `${args.options.ownerGroupId} is not a valid GUID`;
	        }
	      }
	      
	      if (args.options.planId) {
	        if (args.options.ownerGroupId || args.options.ownerGroupName) {
	          return 'Don\'t specify ownerGroupId or ownerGroupName when using planId';
	        }
	      }
	    }

	    if (!args.options.newName && !args.options.orderHint) {
	      return 'Specify either newName or orderHint';
	    }

	    return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      ['id', 'name']
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
      .getBucket(args)
      .then(bucket => {
        const requestOptions: AxiosRequestConfig = {
          url: `${this.resource}/v1.0/planner/buckets/${bucket.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'if-match': (bucket as any)['@odata.etag']
          },
          responseType: 'json',
          data: {}
        };

        const { newName, orderHint } = args.options;
        if (newName) {
          requestOptions.data.name = newName;
        }
        if (orderHint) {
          requestOptions.data.orderHint = orderHint;
        }

        return request.patch(requestOptions);
      })
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  private getBucket(args: CommandArgs): Promise<PlannerBucket> {
    if (args.options.id) {
      const requestOptions: AxiosRequestConfig = {
        url: `${this.resource}/v1.0/planner/buckets/${args.options.id}`,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      return request.get<PlannerBucket>(requestOptions);
    }

    return this
      .getPlanId(args)
      .then(planId => {
        const requestOptions: AxiosRequestConfig = {
          url: `${this.resource}/v1.0/planner/plans/${planId}/buckets`,
          headers: {
            accept: 'application/json'
          },
          responseType: 'json'
        };

        return request.get<{ value: PlannerBucket[] }>(requestOptions);
      })
      .then(buckets => {
        const filteredBuckets = buckets.value.filter(b => args.options.name!.toLowerCase() === b.name!.toLowerCase());

        if (!filteredBuckets.length) {
          return Promise.reject(`The specified bucket ${args.options.name} does not exist`);
        }

        if (filteredBuckets.length > 1) {
          return Promise.reject(`Multiple buckets with name ${args.options.name} found: ${filteredBuckets.map(x => x.id)}`);
        }

        return Promise.resolve(filteredBuckets[0]);
      });
  }

  private getPlanId(args: CommandArgs): Promise<string> {
    const { planId, planTitle} = args.options;

    if (planId) {
      return Promise.resolve(planId);
    }

    return this
      .getGroupId(args)
      .then(groupId => planner.getPlanByTitle(planTitle!, groupId))
      .then(plan => plan.id!);
  }

  private getGroupId(args: CommandArgs): Promise<string> {
    const { ownerGroupId, ownerGroupName } = args.options;

    if (ownerGroupId) {
      return Promise.resolve(ownerGroupId);
    }

    return aadGroup
      .getGroupByDisplayName(ownerGroupName!)
      .then(group => group.id!);
  }
}

module.exports = new PlannerBucketSetCommand();