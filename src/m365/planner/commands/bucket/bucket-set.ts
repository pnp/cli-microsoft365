import { PlannerBucket } from '@microsoft/microsoft-graph-types';
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
  name: z.string().optional().alias('n'),
  planId: z.string().optional(),
  planTitle: z.string().optional(),
  rosterId: z.string().optional(),
  ownerGroupId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .optional(),
  ownerGroupName: z.string().optional(),
  newName: z.string().optional(),
  orderHint: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerBucketSetCommand extends GraphCommand {
  public get name(): string {
    return commands.BUCKET_SET;
  }

  public get description(): string {
    return 'Updates a Microsoft Planner bucket';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => [opts.id, opts.name].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'id' or 'name'.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      })
      .refine(opts => !opts.name || [opts.planId, opts.planTitle, opts.rosterId].filter(x => x !== undefined).length === 1, {
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
      })
      .refine(opts => !opts.id || (!opts.planId && !opts.planTitle && !opts.ownerGroupId && !opts.ownerGroupName && !opts.rosterId), {
        message: `Don't specify planId, planTitle, ownerGroupId, ownerGroupName or rosterId when using id`
      })
      .refine(opts => !opts.name || !opts.planId || (!opts.ownerGroupId && !opts.ownerGroupName), {
        message: `Don't specify ownerGroupId or ownerGroupName when using planId`
      })
      .refine(opts => !opts.name || !opts.rosterId || (!opts.ownerGroupId && !opts.ownerGroupName), {
        message: `Don't specify ownerGroupId or ownerGroupName when using rosterId`
      })
      .refine(opts => opts.newName !== undefined || opts.orderHint !== undefined, {
        message: 'Specify either newName or orderHint'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Updating bucket...`);
    }

    try {
      const bucket = await this.getBucket(args);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/planner/buckets/${bucket.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'if-match': (bucket as any)['@odata.etag']
        },
        responseType: 'json',
        data: {
          name: args.options.newName,
          orderHint: args.options.orderHint
        }
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getBucket(args: CommandArgs): Promise<PlannerBucket> {
    if (args.options.id) {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/planner/buckets/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=minimal'
        },
        responseType: 'json'
      };

      return await request.get<PlannerBucket>(requestOptions);
    }

    const planId = await this.getPlanId(args);
    return planner.getBucketByTitle(args.options.name!, planId, 'minimal');
  }

  private async getPlanId(args: CommandArgs): Promise<string> {
    const { planId, planTitle, rosterId } = args.options;

    if (planId) {
      return planId;
    }

    if (planTitle) {
      const groupId = await this.getGroupId(args);
      return planner.getPlanIdByTitle(planTitle, groupId);
    }

    return planner.getPlanIdByRosterId(rosterId!);
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    const { ownerGroupId, ownerGroupName } = args.options;

    if (ownerGroupId) {
      return ownerGroupId;
    }

    return entraGroup.getGroupIdByDisplayName(ownerGroupName!);
  }
}

export default new PlannerBucketSetCommand();
