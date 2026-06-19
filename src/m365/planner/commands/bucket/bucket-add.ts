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
  name: z.string().alias('n'),
  planId: z.string().optional(),
  planTitle: z.string().optional(),
  rosterId: z.string().optional(),
  ownerGroupId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .optional(),
  ownerGroupName: z.string().optional(),
  orderHint: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerBucketAddCommand extends GraphCommand {
  public get name(): string {
    return commands.BUCKET_ADD;
  }

  public get description(): string {
    return 'Adds a new Microsoft Planner bucket';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => [opts.planId, opts.planTitle, opts.rosterId].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'planId', 'planTitle' or 'rosterId'.`,
        params: {
          customCode: 'optionSet',
          options: ['planId', 'planTitle', 'rosterId']
        }
      })
      .refine(opts => !opts.planTitle || [opts.ownerGroupId, opts.ownerGroupName].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'ownerGroupId' or 'ownerGroupName'.`,
        params: {
          customCode: 'optionSet',
          options: ['ownerGroupId', 'ownerGroupName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const planId = await this.getPlanId(args);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/planner/buckets`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          name: args.options.name,
          planId: planId,
          orderHint: args.options.orderHint
        }
      };

      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getPlanId(args: CommandArgs): Promise<string> {
    if (args.options.planId) {
      return args.options.planId;
    }

    if (args.options.planTitle) {
      const groupId: string = await this.getGroupId(args);
      return planner.getPlanIdByTitle(args.options.planTitle, groupId);
    }

    return planner.getPlanIdByRosterId(args.options.rosterId!);
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.ownerGroupId) {
      return args.options.ownerGroupId;
    }

    return entraGroup.getGroupIdByDisplayName(args.options.ownerGroupName!);
  }
}

export default new PlannerBucketAddCommand();
