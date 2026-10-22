import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  ownerGroupId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .optional(),
  ownerGroupName: z.string().optional(),
  rosterId: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerPlanListCommand extends GraphCommand {
  public get name(): string {
    return commands.PLAN_LIST;
  }

  public get description(): string {
    return 'Returns a list of plans associated with a specified group or roster';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'createdDateTime', 'owner'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => [opts.ownerGroupId, opts.ownerGroupName, opts.rosterId].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'ownerGroupId', 'ownerGroupName' or 'rosterId'.`,
        params: {
          customCode: 'optionSet',
          options: ['ownerGroupId', 'ownerGroupName', 'rosterId']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let plannerPlans = [];
      if (args.options.ownerGroupId || args.options.ownerGroupName) {
        const groupId = await this.getGroupId(args);
        plannerPlans = await planner.getPlansByGroupId(groupId);
      }
      else {
        const plan = await planner.getPlanByRosterId(args.options.rosterId!);
        plannerPlans.push(plan);
      }

      await logger.log(plannerPlans);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.ownerGroupId) {
      return args.options.ownerGroupId;
    }

    return entraGroup.getGroupIdByDisplayName(args.options.ownerGroupName!);
  }
}

export default new PlannerPlanListCommand();
