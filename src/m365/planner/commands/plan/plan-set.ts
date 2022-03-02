import { PlannerPlan } from "@microsoft/microsoft-graph-types";
import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command, {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import request from '../../../../request';
import { Options as PlanGetCommandOptions } from '../plan/plan-get';
import * as planGetCommand from '../plan/plan-get';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  newTitle: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
}

interface ExtendedPlannerPlan extends PlannerPlan {
  '@odata.etag': string
}

class PlannerPlanSetCommand extends GraphCommand {
  public get name(): string {
    return commands.PLAN_SET;
  }

  public get description(): string {
    return 'Update title of a specified plan';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.title = typeof args.options.title !== 'undefined';
    telemetryProps.ownerGroupId = typeof args.options.ownerGroupId !== 'undefined';
    telemetryProps.ownerGroupName = typeof args.options.ownerGroupName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (error?: any) => void): void {
    this.getPlan(logger, args)
      .then((output: CommandOutput): void => {
        const plan: ExtendedPlannerPlan = JSON.parse(output.stdout);
        this.updatePlan(logger, args, plan)
          .then((): void => {
            cb();
          })
          .catch((err) => {
            this.handleRejectedODataJsonPromise(err, logger, cb);
          });
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getPlan(logger: Logger, args: CommandArgs): Promise<CommandOutput> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving plan ${args.options.id || args.options.title} ...`);
    }

    const options: PlanGetCommandOptions = {
      ...args.options,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    return Cli.executeCommandWithOutput(planGetCommand as Command, { options: { ...options, _: [] } });
  }

  private updatePlan(logger: Logger, args: CommandArgs, plan: ExtendedPlannerPlan): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Updating plan with id ${plan.id} ...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/plans/${plan.id}`,
      headers: {
        'accept': 'application/json;odata.metadata=none',
        'If-Match': `${plan["@odata.etag"]}`
      },
      responseType: 'json',
      data: {
        title: args.options.newTitle
      }
    };

    return request.patch(requestOptions);
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
        option: '--newTitle <newTitle>'
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
      return 'Specify either id or title but not both';
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

module.exports = new PlannerPlanSetCommand();