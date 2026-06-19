import { PlannerPlan, PlannerPlanDetails } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().optional().alias('i'),
  title: z.string().optional().alias('t'),
  rosterId: z.string().optional(),
  ownerGroupId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .optional(),
  ownerGroupName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerPlanGetCommand extends GraphCommand {
  public get name(): string {
    return commands.PLAN_GET;
  }

  public get description(): string {
    return 'Get a Microsoft Planner plan';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => [opts.id, opts.title, opts.rosterId].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'id', 'title' or 'rosterId'.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'title', 'rosterId']
        }
      })
      .refine(opts => !opts.title || [opts.ownerGroupId, opts.ownerGroupName].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'ownerGroupId' or 'ownerGroupName'.`,
        params: {
          customCode: 'optionSet',
          options: ['ownerGroupId', 'ownerGroupName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let plan: PlannerPlan;
      if (args.options.id) {
        plan = await planner.getPlanById(args.options.id);
      }
      else if (args.options.rosterId) {
        plan = await planner.getPlanByRosterId(args.options.rosterId);
      }
      else {
        const groupId = await this.getGroupId(args);
        plan = await planner.getPlanByTitle(args.options.title!, groupId);
      }

      const result = await this.getPlanDetails(plan);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getPlanDetails(plan: PlannerPlan): Promise<PlannerPlan & PlannerPlanDetails> {
    const requestOptionsTaskDetails: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/plans/${plan.id}/details`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        Prefer: 'return=representation'
      },
      responseType: 'json'
    };

    const planDetails = await request.get(requestOptionsTaskDetails);
    return { ...plan, ...planDetails as PlannerPlanDetails };
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.ownerGroupId) {
      return args.options.ownerGroupId;
    }

    return entraGroup.getGroupIdByDisplayName(args.options.ownerGroupName!);
  }
}

export default new PlannerPlanGetCommand();
