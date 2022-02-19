import { Logger } from '../../../../cli';
import { CommandOption} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { PlannerPlan, PlannerPlanDetails  } from '@microsoft/microsoft-graph-types';

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
  private groupId: string = "";

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
      return Promise.resolve("");
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
      .get<{ value: [{ id: string, resourceProvisioningOptions: string[] }] }>(requestOptions)
      .then(response => {
        const groupItem: { id: string, resourceProvisioningOptions: string[] } | undefined = response.value[0];

        if (!groupItem) {
          return Promise.reject(`The specified ownerGroup does not exist`);
        }

        if (response.value.length > 1) {
          return Promise.reject(`Multiple ownerGroups with name ${args.options.ownerGroupName} found: Please choose between the following IDs ${response.value.map(x => x.id)}`);
        }

        return Promise.resolve(groupItem.id);
      });
  }

  private getPlanId(args: CommandArgs): Promise<string> {
    if (args.options.planId) {
      return Promise.resolve(args.options.planId);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups/${this.groupId}/planner/plans`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    
    return request
      .get<{ value: PlannerPlan [] }>(requestOptions)
      .then((response: { value: PlannerPlan[] }): Promise<string> => {
        const filteredPlan = response.value.filter((plan: PlannerPlan) => plan.title === args.options.planTitle);
        if (filteredPlan && filteredPlan.length > 0) {
          if (filteredPlan.length > 1) {
            return Promise.reject(`Multiple plans with name ${args.options.planTitle} found: ${filteredPlan.map(x => x.id)}`);
          }
          if(filteredPlan[0].id) {
            return Promise.resolve(filteredPlan[0].id);
          }
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

    return request
      .get<{ value: PlannerPlanDetails }>(requestOptions)
      .then(response => {
        const planDetailsItem: any | undefined = response;
        return Promise.resolve(planDetailsItem);
      });
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

    if (args.options.ownerGroupId && !Utils.isValidGuid(args.options.ownerGroupId as string)) {
      return `${args.options.ownerGroupId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new PlannerPlanDetailsGetCommand();