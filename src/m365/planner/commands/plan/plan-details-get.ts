import { PlannerPlanDetails } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import { accessToken, validation } from '../../../../utils';
import Auth from '../../../../Auth';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { planner } from '../../../../utils/planner';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { aadGroup } from '../../../../utils/aadGroup';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  planId?: string;
  planTitle?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
}

class PlannerPlanDetailsGetCommand extends GraphCommand {
  private groupId: string = '';

  public get name(): string {
    return commands.PLAN_DETAILS_GET;
  }

  public get description(): string {
    return 'Get details of a Microsoft Planner plan';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.planId = typeof args.options.planId !== 'undefined';
    telemetryProps.planTitle = typeof args.options.planTitle !== 'undefined';
    telemetryProps.ownerGroupId = typeof args.options.ownerGroupId !== 'undefined';
    telemetryProps.ownerGroupName = typeof args.options.ownerGroupName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (accessToken.isAppOnlyAccessToken(Auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.', logger, cb);
      return;
    }
    
    this
      .getGroupId(args)
      .then((groupId: string): Promise<string> => {
        this.groupId = groupId;
        return this.getPlanId(args);
      })
      .then((planId: string): Promise<PlannerPlanDetails> => {
        args.options.planId = planId;
        return this.getPlanDetails(args);
      })
      .then((res: PlannerPlanDetails): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.planId) {
      return Promise.resolve('');
    }

    if (args.options.ownerGroupId) {
      return Promise.resolve(args.options.ownerGroupId);
    }
    
    return aadGroup
      .getGroupByDisplayName(args.options.ownerGroupName!)
      .then(group => group.id!);
  }

  private getPlanId(args: CommandArgs): Promise<string> {
    if (args.options.planId) {
      return Promise.resolve(args.options.planId);
    }

    return planner
      .getPlanByName(args.options.planTitle!, this.groupId)
      .then(plan => plan.id!);
  }

  private getPlanDetails(args: CommandArgs): Promise<PlannerPlanDetails> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/plans/${args.options.planId}/details`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<PlannerPlanDetails>(requestOptions);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --planId [planId]'
      },
      {
        option: '-t, --planTitle [planTitle]'
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
    if (!args.options.planId && !args.options.planTitle) {
      return 'Specify either planId or planTitle';
    }

    if (args.options.planId && args.options.planTitle) {
      return 'Specify either planId or planTitle';
    }

    if (args.options.planTitle && !args.options.ownerGroupId && !args.options.ownerGroupName) {
      return 'Specify either ownerGroupId or ownerGroupName';
    }

    if (args.options.planTitle && args.options.ownerGroupId && args.options.ownerGroupName) {
      return 'Specify either ownerGroupId or ownerGroupName but not both';
    }

    if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId as string)) {
      return `${args.options.ownerGroupId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new PlannerPlanDetailsGetCommand();