import { Group, PlannerPlan, PlannerPlanDetails } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import { accessToken, odata, validation } from '../../../../utils';
import Auth from '../../../../Auth';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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
        return this.getPlanId(args, logger);
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
    
    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups?$filter=displayName eq '${encodeURIComponent(args.options.ownerGroupName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: Group[] }>(requestOptions)
      .then(response => {
        const groupItem: Group | undefined = response.value[0];

        if (!groupItem) {
          return Promise.reject(`The specified ownerGroup does not exist`);
        }

        if (response.value.length > 1) {
          return Promise.reject(`Multiple ownerGroups with name ${args.options.ownerGroupName} found: Please choose between the following IDs ${response.value.map(x => x.id)}`);
        }

        return Promise.resolve(groupItem.id as string);
      });
  }

  private getPlanId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.planId) {
      return Promise.resolve(args.options.planId);
    }

    return odata
      .getAllItems<PlannerPlan>(`${this.resource}/v1.0/groups/${this.groupId}/planner/plans`, logger)
      .then((plans: PlannerPlan[]): Promise<string> => {
        const plansMatchingName = plans.filter((plan: PlannerPlan) => plan.title === args.options.planTitle);

        if (plansMatchingName && plansMatchingName.length > 0) {
          if (plansMatchingName.length > 1) {
            return Promise.reject(`Multiple plans with name ${args.options.planTitle} found: ${plansMatchingName.map(x => x.id)}`);
          }

          return Promise.resolve(plansMatchingName[0].id as string);
        }

        return Promise.reject(`The specified plan title does not exist`);
      });
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