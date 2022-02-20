import * as os from 'os';
import { PlannerPlan } from "@microsoft/microsoft-graph-types";
import { Logger } from '../../../../cli';
import {
  CommandOption,
  CommandError
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { odata, validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
}

class PlannerPlanGetCommand extends GraphCommand {
  public get name(): string {
    return commands.PLAN_GET;
  }

  public get description(): string {
    return 'Get a Microsoft Planner plan';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.title = typeof args.options.title !== 'undefined';
    telemetryProps.ownerGroupId = typeof args.options.ownerGroupId !== 'undefined';
    telemetryProps.ownerGroupName = typeof args.options.ownerGroupName !== 'undefined';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'createdDateTime', 'owner', '@odata.etag'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (error?: any) => void): void {
    if (args.options.id) {
      this
        .getPlan(args)
        .then((res: PlannerPlan): void => {
          logger.log(res);
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }
    else {
      this
        .getGroupId(args)
        .then((groupId: string): Promise<PlannerPlan[]> => odata.getAllItems(`${this.resource}/v1.0/groups/${groupId}/planner/plans`, logger, 'minimal'))
        .then((plans): void => {
          const filteredPlans = plans.filter((plan: PlannerPlan) => plan.title === args.options.title);

          if (!filteredPlans.length) {
            cb(new CommandError(`No plan with the name ${args.options.title} found`));
            return;
          }

          if (filteredPlans && filteredPlans.length > 1) {
            let sameNamePlans: string = `Multiple plans with the name ${args.options.title} found. Please choose between the following IDs:`;
            filteredPlans.map((plan: PlannerPlan) => sameNamePlans += `${os.EOL}${plan.id}`);
            cb(new CommandError(`${sameNamePlans}`));
            return;
          }

          logger.log(filteredPlans[0]);

          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }
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

  private getPlan(args: CommandArgs): Promise<PlannerPlan> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/plans/${args.options.id}`,
      headers: {
        'accept': 'application/json'
      },
      responseType: 'json'
    };

    return request.get<PlannerPlan>(requestOptions);
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

    if (args.options.title && !args.options.ownerGroupId && !args.options.ownerGroupName) {
      return 'Specify either ownerGroupId or ownerGroupName';
    }

    if (args.options.title && args.options.ownerGroupId && args.options.ownerGroupName) {
      return 'Specify either ownerGroupId or ownerGroupName but not both';
    }

    if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId as string)) {
      return `${args.options.ownerGroupId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new PlannerPlanGetCommand();