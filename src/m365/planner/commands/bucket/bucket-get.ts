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
  ownerGroupName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerBucketGetCommand extends GraphCommand {
  public get name(): string {
    return commands.BUCKET_GET;
  }

  public get description(): string {
    return 'Gets the Microsoft Planner bucket in a plan';
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
      .refine(opts => !opts.id || (!opts.planId && !opts.planTitle && !opts.rosterId && !opts.ownerGroupId && !opts.ownerGroupName), {
        message: `Don't specify planId, planTitle, rosterId, ownerGroupId or ownerGroupName when using id`
      })
      .refine(opts => !opts.name || !opts.planId || (!opts.ownerGroupId && !opts.ownerGroupName), {
        message: `Don't specify ownerGroupId or ownerGroupName when using planId`
      })
      .refine(opts => !opts.name || !opts.rosterId || (!opts.ownerGroupId && !opts.ownerGroupName), {
        message: `Don't specify ownerGroupId or ownerGroupName when using rosterId`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const bucket = await this.getBucket(args);
      await logger.log(bucket);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getBucket(args: CommandArgs): Promise<PlannerBucket> {
    const { id, name } = args.options;
    if (id) {
      return this.getBucketById(id);
    }

    const planId = await this.getPlanId(args);
    return planner.getBucketByTitle(name!, planId);
  }

  private async getPlanId(args: CommandArgs): Promise<string> {
    const { planId, planTitle, rosterId } = args.options;

    if (planId) {
      return planId;
    }

    if (planTitle) {
      const groupId: string = await this.getGroupId(args);
      return planner.getPlanIdByTitle(planTitle, groupId);
    }

    return planner.getPlanIdByRosterId(rosterId!);
  }

  private async getBucketById(id: string): Promise<PlannerBucket> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/buckets/${id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<PlannerBucket>(requestOptions);
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    const { ownerGroupId, ownerGroupName } = args.options;

    if (ownerGroupId) {
      return ownerGroupId;
    }

    return entraGroup.getGroupIdByDisplayName(ownerGroupName!);
  }
}

export default new PlannerBucketGetCommand();
