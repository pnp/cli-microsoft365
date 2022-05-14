import { PlannerBucket, PlannerPlan, Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import { AxiosRequestConfig } from 'axios';
import {
  CommandOption
} from '../../../../Command';
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
  name?: string;
  planId?: string;
  planName?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
}

class PlannerBucketGetCommand extends GraphCommand {
  public get name(): string {
    return commands.BUCKET_GET;
  }

  public get description(): string {
    return 'Gets the Microsoft Planner bucket in a plan';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.planId = typeof args.options.planId !== 'undefined';
    telemetryProps.planName = typeof args.options.planName !== 'undefined';
    telemetryProps.ownerGroupId = typeof args.options.ownerGroupId !== 'undefined';
    telemetryProps.ownerGroupName = typeof args.options.ownerGroupName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getBucketId(args)
      .then((bucketId: string) => this.getBucketById(bucketId))
      .then((bucket: PlannerBucket) => {
        logger.log(bucket);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getBucketId (args: CommandArgs): Promise<string> {
    const { id, name } = args.options;
    if (id) {
      return Promise.resolve(id);
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

        return request.get<{ value:PlannerBucket[] }>(requestOptions);
      })
      .then(buckets => {
        const filteredBuckets = buckets.value.filter(b => name!.toLowerCase() === b.name!.toLowerCase());

        if (!filteredBuckets.length) {
          return Promise.reject(`The specified bucket ${name} does not exist`);
        }
        
        if (filteredBuckets.length > 1) {
          return Promise.reject(`Multiple buckets with name ${name} found: ${filteredBuckets.map(x => x.id)}`);
        }

        return Promise.resolve(filteredBuckets[0].id!.toString());
      });     
  }

  private getPlanId(args: CommandArgs): Promise<string> {
    const { planId, planName } = args.options;

    if (planId) {
      return Promise.resolve(planId);
    }

    return this
      .getGroupId(args)
      .then(groupId => {
        const requestOptions: AxiosRequestConfig = {
          url: `${this.resource}/v1.0/planner/plans?$filter=owner eq '${groupId}'`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        return request.get<{ value: PlannerPlan[] }>(requestOptions);
      })
      .then(plans => {
        const filteredPlans = plans.value.filter(p => p.title!.toLowerCase() === planName!.toLowerCase());

        if (filteredPlans.length === 0) {
          return Promise.reject(`The specified plan ${planName} does not exist`);
        }

        if (filteredPlans.length > 1) {
          return Promise.reject(`Multiple plans with name ${planName} found: ${filteredPlans.map(x => x.id)}`);
        }

        return Promise.resolve(filteredPlans[0].id!);
      });
  }

  private async getBucketById(id: string): Promise<PlannerBucket> {
    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/planner/buckets/${id}`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    return request.get<PlannerBucket>(requestOptions);    
  }

  private getGroupId(args: CommandArgs): Promise<string> {
    const { ownerGroupId, ownerGroupName } = args.options;

    if (ownerGroupId) {
      return Promise.resolve(ownerGroupId);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups?$filter=displayName eq '${encodeURIComponent(ownerGroupName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: Group[] }>(requestOptions)
      .then(response => {
        if (!response.value.length) {
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
      {
        option: '-i, --id [id]'
      },
      {
        option: '--name [name]'
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
    if (args.options.id) {
      if (args.options.planId || args.options.planName || args.options.ownerGroupId || args.options.ownerGroupName) {
        return 'Don\'t specify planId, planName, ownerGroupId or ownerGroupName when using id';
      }
      if (args.options.name) {
        return 'Specify either id or name';
      }
    }

    if (args.options.name) {
      if (!args.options.planId && !args.options.planName) {
        return 'Specify either planId or planName when using name';
      }

      if (args.options.planId && args.options.planName) {
        return 'Specify either planId or planName when using name but not both';
      }

      if (args.options.planName) {
        if (!args.options.ownerGroupId && !args.options.ownerGroupName) {
          return 'Specify either ownerGroupId or ownerGroupName when using planName';
        }

        if (args.options.ownerGroupId && args.options.ownerGroupName) {
          return 'Specify either ownerGroupId or ownerGroupName when using planName but not both';
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

    if (!args.options.id && !args.options.name) {
      return 'Please specify id or name';
    }

    return true;
  }
}

module.exports = new PlannerBucketGetCommand();