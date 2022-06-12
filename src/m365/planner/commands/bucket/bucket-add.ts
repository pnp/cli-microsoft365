import Auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.planId = typeof args.options.planId !== 'undefined';
    telemetryProps.planName = typeof args.options.planName !== 'undefined';
    telemetryProps.ownerGroupId = typeof args.options.ownerGroupId !== 'undefined';
    telemetryProps.ownerGroupName = typeof args.options.ownerGroupName !== 'undefined';
    telemetryProps.orderHint = typeof args.options.orderHint !== 'undefined';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'planId', 'orderHint'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (accessToken.isAppOnlyAccessToken(Auth.service.accessTokens[this.resource].accessToken)) {
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
      .then(groupId => planner.getPlanByName(args.options.planName!, groupId))
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
        option: "--ownerGroupId [ownerGroupId]"
      },
      {
        option: "--ownerGroupName [ownerGroupName]"
      },
      {
        option: "--orderHint [orderHint]"
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.planId && !args.options.planName) {
      return 'Specify either planId or planName';
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

    if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId as string)) {
      return `${args.options.ownerGroupId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new PlannerBucketAddCommand();