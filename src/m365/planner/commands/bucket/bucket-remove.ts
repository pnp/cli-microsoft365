import { Group, PlannerBucket, PlannerPlan } from '@microsoft/microsoft-graph-types';
import { AxiosRequestConfig } from 'axios';
import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import { validation } from '../../../../utils';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
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
  confirm?: boolean;
}

class PlannerBucketRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.BUCKET_REMOVE;
  }

  public get description(): string {
    return 'Removes the Microsoft Planner bucket from a plan';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.planId = typeof args.options.planId !== 'undefined';
    telemetryProps.planName = typeof args.options.planName !== 'undefined';
    telemetryProps.ownerGroupId = typeof args.options.ownerGroupId !== 'undefined';
    telemetryProps.ownerGroupName = typeof args.options.ownerGroupName !== 'undefined';
    telemetryProps.confirm = args.options.confirm || false;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeBucket: () => void = (): void => {
      this
        .getBucket(args)
        .then(bucket => {
          const requestOptions: AxiosRequestConfig = {
            url: `${this.resource}/v1.0/planner/buckets/${bucket.id}`,
            headers: {
              accept: 'application/json;odata.metadata=none',
              'if-match': (bucket as any)['@odata.etag']
            },
            responseType: 'json'
          };
    
          return request.delete(requestOptions);
        })
        .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeBucket();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the bucket ${args.options.id || args.options.name}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeBucket();
        }
      });
    }
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

        return request.get<{ value:PlannerBucket[] }>(requestOptions);
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

        if (!filteredPlans.length) {
          return Promise.reject(`The specified plan ${planName} does not exist`);
        }

        if (filteredPlans.length > 1) {
          return Promise.reject(`Multiple plans with name ${planName} found: ${filteredPlans.map(x => x.id)}`);
        }

        return Promise.resolve(filteredPlans[0].id!);
      });
  }

  private getGroupId(args: CommandArgs): Promise<string> {
    const { ownerGroupId, ownerGroupName } = args.options;

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
        option: '--id [id]'
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
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public optionSets(): string[][] | undefined {
    return [
      ['id', 'name']
    ];
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.id) {
      if (args.options.planId || args.options.planName || args.options.ownerGroupId || args.options.ownerGroupName) {
        return 'Don\'t specify planId, planName, ownerGroupId or ownerGroupName when using id';
      }
    }
    else {
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
      else {
        if (args.options.ownerGroupId || args.options.ownerGroupName) {
          return 'Don\'t specify ownerGroupId or ownerGroupName when using planId';
        }
      }
    }

    return true;
  }
}

module.exports = new PlannerBucketRemoveCommand();