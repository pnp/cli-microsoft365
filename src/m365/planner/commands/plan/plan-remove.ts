import { PlannerPlan } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
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
  ownerGroupId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .optional(),
  ownerGroupName: z.string().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerPlanRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.PLAN_REMOVE;
  }

  public get description(): string {
    return 'Removes the Microsoft Planner plan';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => [opts.id, opts.title].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'id' or 'title'.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'title']
        }
      })
      .refine(opts => !opts.title || [opts.ownerGroupId, opts.ownerGroupName].filter(x => x !== undefined).length === 1, {
        message: `Specify either ownerGroupId or ownerGroupName when using title.`,
        params: {
          customCode: 'optionSet',
          options: ['ownerGroupId', 'ownerGroupName']
        }
      })
      .refine(opts => !opts.id || (!opts.ownerGroupId && !opts.ownerGroupName), {
        message: `Don't specify ownerGroupId or ownerGroupName when using id`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removePlan = async (): Promise<void> => {
      try {
        const plan = await this.getPlan(args);

        if (this.verbose) {
          await logger.logToStderr(`Removing plan '${plan.title}' ...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/planner/plans/${plan.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'if-match': (plan as any)['@odata.etag']
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removePlan();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the plan ${args.options.id || args.options.title}?` });

      if (result) {
        await removePlan();
      }
    }
  }

  private async getPlan(args: CommandArgs): Promise<PlannerPlan> {
    const { id, title } = args.options;

    if (id) {
      return planner.getPlanById(id, 'minimal');
    }

    const groupId = await this.getGroupId(args);
    return planner.getPlanByTitle(title!, groupId, 'minimal');
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    const { ownerGroupId, ownerGroupName } = args.options;

    if (ownerGroupId) {
      return ownerGroupId;
    }

    return entraGroup.getGroupIdByDisplayName(ownerGroupName!);
  }
}

export default new PlannerPlanRemoveCommand();
