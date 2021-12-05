import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import { Bucket } from '../../Bucket';
import commands from '../../commands';
interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  bucketId: string;
  bucketName: string;
  planId?: string;
  planName?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
  confirm?: boolean;
}

class PlannerBucketRemoveCommand extends GraphItemsListCommand<any> {
  public get name(): string {
    return commands.BUCKET_REMOVE;
  }

  public get description(): string {
    return 'Removes the Microsoft Planner bucket from a plan';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.bucketId = typeof args.options.bucketId !== 'undefined';
    telemetryProps.bucketName = typeof args.options.bucketName !== 'undefined';
    telemetryProps.planId = typeof args.options.planId !== 'undefined';
    telemetryProps.planName = typeof args.options.planName !== 'undefined';
    telemetryProps.ownerGroupId = typeof args.options.ownerGroupId !== 'undefined';
    telemetryProps.ownerGroupName = typeof args.options.ownerGroupName !== 'undefined';
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['bucketId', 'bucketName', 'planId', 'confirm'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const removeBucket: () => void = (): void => {
      if (args.options.bucketName) {

        this.getPlanId(args)
          .then((planId: string): Promise<any> => {
            const requestOptions: any = {
              url: `${this.resource}/v1.0/planner/plans/${planId}/buckets`,
              headers: {
                'accept': 'application/json',
                'prefer': 'odata.eTag'
              },
              responseType: 'json'
            };
            return request.get<{ value: { id: string; title: string; }[] }>(requestOptions);

          }).then((response) => {
            const bucket: Bucket | undefined = response.value.find((bucket: Bucket) => bucket.name === args.options.bucketName);
            if (!bucket) {
              return Promise.reject(`The specified bucket does not exist in the Microsoft Planner`);
            }
            const requestOptions: any = {
              url: `${this.resource}/v1.0/planner/buckets/${bucket.id}`,
              headers: {
                'X-HTTP-Method': 'DELETE',
                'If-Match': bucket['@odata.etag'],
                'accept': 'application/json;odata=nometadata'
              },
              responseType: 'json'
            };
            return request.delete(requestOptions);
          }).then(_ => cb(), (err: any) => {
            this.handleRejectedODataJsonPromise(err, logger, cb);
          });
      }
      if (args.options.bucketId) {
        this.getBucket(args)
          .then((res: Bucket): Promise<void> => {
            const bucketid: string = res.id;
            const requestOptions: any = {
              url: `${this.resource}/v1.0/planner/buckets/${bucketid}`,
              headers: {
                'X-HTTP-Method': 'DELETE',
                'If-Match': res['@odata.etag'],
                'accept': 'application/json;odata=nometadata'
              },
              responseType: 'json'
            };
            return request.delete(requestOptions);
          })
          .then(_ => cb(), (err: any) => {
            this.handleRejectedODataJsonPromise(err, logger, cb);
          });
      }
    };
    if (args.options.confirm) {
      removeBucket();
    }
    else {
      const bucketName = args.options.bucketName ? args.options.bucketName : args.options.bucketId;
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the bucket ${bucketName} from planner?`
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

  private getBucket(args: CommandArgs): Promise<any> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/buckets/${encodeURIComponent(args.options.bucketId)}`,
      headers: {
        accept: 'application/json',
        'prefer': 'odata.eTag'
      },
      responseType: 'json'
    };
    return request
      .get<any>(requestOptions)
      .then(response => {
        const bucketItem: any | undefined = response;
        return Promise.resolve(bucketItem);
      })
      .catch(() => {
        return Promise.reject(`The specified bucket does not exist in the Microsoft Planner`);
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
      }).then((response) => {
        const plan: { id: string; title: string; } | undefined = response.value.find(val => val.title === args.options.planName);

        if (!plan) {
          return Promise.reject(`The specified plan does not exist`);
        }

        return Promise.resolve(plan.id);
      });
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
        option: '-i, --bucketId [bucketId]'
      },
      {
        option: '-n, --bucketName [bucketName]'
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

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.bucketId && !args.options.bucketName) {
      return 'Specify either bucketId or bucketName';
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
module.exports = new PlannerBucketRemoveCommand();