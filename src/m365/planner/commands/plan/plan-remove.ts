import { PlannerPlan } from '@microsoft/microsoft-graph-types';
import { AxiosRequestConfig } from 'axios';
import Auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
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
  id?: string;
  title?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
  confirm?: boolean;
}

class PlannerPlanRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.PLAN_REMOVE;
  }

  public get description(): string {
    return 'Removes the Microsoft Planner plan';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.title = typeof args.options.title !== 'undefined';
    telemetryProps.ownerGroupId = typeof args.options.ownerGroupId !== 'undefined';
    telemetryProps.ownerGroupName = typeof args.options.ownerGroupName !== 'undefined';
    telemetryProps.confirm = !!args.options.confirm;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (accessToken.isAppOnlyAccessToken(Auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.', logger, cb);
      return;
    }

    const removePlan: () => Promise<void> = async (): Promise<void> => {
      try {
        const plan = await this.getPlan(args);

        const requestOptions: AxiosRequestConfig = {
          url: `${this.resource}/v1.0/planner/plans/${plan.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'if-match': (plan as any)['@odata.etag']
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
        cb();
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err, logger, cb);
      }
    };

    if (args.options.confirm) {
      removePlan();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the plan ${args.options.id || args.options.title}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removePlan();
        }
      });
    }
  }

  private async getPlan(args: CommandArgs): Promise<PlannerPlan> {
    const { id, title } = args.options;

    if (id) {
      return planner.getPlanById(id, 'minimal');
    }

    const groupId = await this.getGroupId(args);
    return await planner.getPlanByTitle(title!, groupId);
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    const { ownerGroupId, ownerGroupName } = args.options;

    if (ownerGroupId) {
      return ownerGroupId;
    }

    const group = await aadGroup.getGroupByDisplayName(ownerGroupName!);
    return group.id!;
  }

  public optionSets(): string[][] | undefined {
    return [['id', 'title']];
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
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.title) {
      if (!args.options.ownerGroupId && !args.options.ownerGroupName) {
        return 'Specify either ownerGroupId or ownerGroupName';
      }

      if (args.options.ownerGroupId && args.options.ownerGroupName) {
        return 'Specify either ownerGroupId or ownerGroupName but not both';
      }

      if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
        return `${args.options.ownerGroupId} is not a valid GUID`;
      }
    }
    else if (args.options.ownerGroupId || args.options.ownerGroupName) {
      return 'Don\'t specify ownerGroupId or ownerGroupName when using id';
    }

    return true;
  }
}

module.exports = new PlannerPlanRemoveCommand();