import { PlannerPlan, PlannerPlanDetails } from '@microsoft/microsoft-graph-types';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { accessToken } from '../../../../utils/accessToken';
import { validation } from '../../../../utils/validation';
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
}

class PlannerPlanGetCommand extends GraphCommand {
  public get name(): string {
    return commands.PLAN_GET;
  }

  public get description(): string {
    return 'Get a Microsoft Planner plan';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'createdDateTime', 'owner', '@odata.etag'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
        ownerGroupName: typeof args.options.ownerGroupName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
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
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
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
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'title'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.');
      return;
    }

    try {
      if (args.options.id) {
        const plan = await planner.getPlanById(args.options.id);
        const result = await this.getPlanDetails(plan);
        logger.log(result);
      }
      else {
        const groupId = await this.getGroupId(args);
        const plan = await planner.getPlanByTitle(args.options.title!, groupId);
        const result = await this.getPlanDetails(plan);

        if (result) {
          logger.log(result);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getPlanDetails(plan: PlannerPlan): Promise<PlannerPlan & PlannerPlanDetails> {
    const requestOptionsTaskDetails: any = {
      url: `${this.resource}/v1.0/planner/plans/${plan.id}/details`,
      headers: {
        'accept': 'application/json;odata.metadata=none',
        'Prefer': 'return=representation'
      },
      responseType: 'json'
    };

    return request
      .get(requestOptionsTaskDetails)
      .then(planDetails => {
        return { ...plan, ...planDetails as PlannerPlanDetails };
      });
  }

  private getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.ownerGroupId) {
      return Promise.resolve(args.options.ownerGroupId);
    }

    return aadGroup
      .getGroupByDisplayName(args.options.ownerGroupName!)
      .then(group => group.id!);
  }
}

module.exports = new PlannerPlanGetCommand();