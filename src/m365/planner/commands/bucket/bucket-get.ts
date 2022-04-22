import { PlannerBucket } from '@microsoft/microsoft-graph-types';
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
  title?: string;
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
    telemetryProps.title = typeof args.options.title !== 'undefined';
    telemetryProps.planId = typeof args.options.planId !== 'undefined';
    telemetryProps.planName = typeof args.options.planName !== 'undefined';
    telemetryProps.ownerGroupId = typeof args.options.ownerGroupId !== 'undefined';
    telemetryProps.ownerGroupName = typeof args.options.ownerGroupName !== 'undefined';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'planId', 'orderHint'];
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
    if (args.options.id) {
      return Promise.resolve(args.options.id);
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
        const bucket: { id: string; name: string; } | undefined = response.value.find(val => val.name === args.options.title);

        if (!bucket) {
          return Promise.reject(`The specified bucket does not exist`);
        }

        return Promise.resolve(bucket.id);
      });     
  }

  private getPlanId(args: CommandArgs): Promise<string> {
    if (args.options.planId) {
      return Promise.resolve(args.options.planId);
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

  private async getBucketById(id: string): Promise<PlannerBucket> {
    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/planner/buckets/${encodeURIComponent(id)}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'      
    };

    return request.get<PlannerBucket>(requestOptions);    
  }

  private getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.ownerGroupId) {
      return Promise.resolve(args.options.ownerGroupId);
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
        option: '-i, --id [id]'
      },
      {
        option: '-t, --title [title]'
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
    if (!args.options.id && !args.options.title) {
      return 'Specify either id or title';
    }

    if (args.options.id && args.options.title) {
      return 'Specify either id or title';
    }

    if (args.options.title && !args.options.planId && !args.options.planName) {
      return 'Specify either planId or planName';
    }

    if (args.options.title && args.options.planId && args.options.planName) {
      return 'Specify either planId or planName but not both';
    }

    if (args.options.title && args.options.planName && !args.options.ownerGroupId && !args.options.ownerGroupName) {
      return 'Specify either ownerGroupId or ownerGroupName when using planName';
    }

    if (args.options.title && args.options.planName && args.options.ownerGroupId && args.options.ownerGroupName) {
      return 'Specify either ownerGroupId or ownerGroupName when using planName but not both';
    }

    if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId as string)) {
      return `${args.options.ownerGroupId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new PlannerBucketGetCommand();