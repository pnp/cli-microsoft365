import { PlannerPlan, PlannerPlanDetails, User } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.object({
  ...globalOptionsZod.shape,
  id: z.string().optional().alias('i'),
  title: z.string().optional().alias('t'),
  ownerGroupId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .optional(),
  ownerGroupName: z.string().optional(),
  rosterId: z.string().optional(),
  newTitle: z.string().optional(),
  shareWithUserIds: z.string()
    .refine(val => validation.isValidGuidArray(val) === true, {
      message: 'The value contains invalid GUIDs.'
    })
    .optional(),
  shareWithUserNames: z.string()
    .refine(val => validation.isValidUserPrincipalNameArray(val) === true, {
      message: 'The value contains invalid user principal names.'
    })
    .optional()
}).catchall(z.unknown());

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerPlanSetCommand extends GraphCommand {
  public get name(): string {
    return commands.PLAN_SET;
  }

  public get description(): string {
    return 'Updates a Microsoft Planner plan';
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
      })
      .refine(opts => !opts.shareWithUserIds || !opts.shareWithUserNames, {
        message: 'Specify either shareWithUserIds or shareWithUserNames but not both'
      })
      .refine(opts => {
        const allowedCategories: string[] = Array.from({ length: 25 }, (_, i) => `category${i + 1}`);
        const keys = Object.keys(opts);
        const hasCategoryKey = keys.some(key => key.indexOf('category') !== -1);
        if (!hasCategoryKey) {
          return true;
        }
        return keys.filter(key => key.indexOf('category') !== -1).every(key => allowedCategories.indexOf(key) !== -1);
      }, {
        message: 'Specify category values between category1 to category25'
      });
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    const { ownerGroupId, ownerGroupName } = args.options;

    if (ownerGroupId) {
      return ownerGroupId;
    }

    const id = await entraGroup.getGroupIdByDisplayName(ownerGroupName!);
    return id;
  }

  private async getPlanId(args: CommandArgs): Promise<string> {
    const { id, title } = args.options;

    if (id) {
      return id;
    }

    if (args.options.rosterId) {
      const id = await planner.getPlanIdByRosterId(args.options.rosterId);
      return id;
    }
    else {
      const groupId = await this.getGroupId(args);
      const id = await planner.getPlanIdByTitle(title!, groupId);
      return id;
    }
  }

  private async getUserIds(options: Options): Promise<string[]> {
    if (options.shareWithUserIds) {
      return options.shareWithUserIds.split(',');
    }

    const userNames = options.shareWithUserNames as string;
    const userArr: string[] = userNames.split(',').map(o => o.trim());

    const promises: Promise<{ value: User[] }>[] = userArr.map(user => {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user)}'&$select=id,userPrincipalName`,
        headers: {
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      return request.get(requestOptions);
    });

    const usersRes = await Promise.all(promises);

    const userUpns = usersRes.map(res => res.value[0]?.userPrincipalName as string);
    const userIds = usersRes.map(res => res.value[0]?.id as string);

    // Find the members where no graph response was found
    const invalidUsers = userArr.filter(user => !userUpns.some((upn) => upn?.toLowerCase() === user.toLowerCase()));

    if (invalidUsers && invalidUsers.length > 0) {
      throw `Cannot proceed with planner plan creation. The following users provided are invalid: ${invalidUsers.join(',')}`;
    }

    return userIds;
  }

  private async generateSharedWith(options: Options): Promise<{ [userId: string]: boolean }> {
    const sharedWith: { [userId: string]: boolean } = {};

    const userIds = await this.getUserIds(options);
    userIds.map(x => sharedWith[x] = true);
    return sharedWith;
  }

  private async getPlanEtag(planId: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/plans/${planId}`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    const response: any = await request.get(requestOptions);
    return response['@odata.etag'];
  }

  private async getPlanDetailsEtag(planId: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/plans/${planId}/details`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    const response: any = await request.get(requestOptions);
    return response['@odata.etag'];
  }

  private async getPlanDetails(plan: PlannerPlan): Promise<PlannerPlan & PlannerPlanDetails> {
    const requestOptionsTaskDetails: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/plans/${plan.id}/details`,
      headers: {
        'accept': 'application/json;odata.metadata=none',
        'Prefer': 'return=representation'
      },
      responseType: 'json'
    };

    const planDetails: any = await request.get(requestOptionsTaskDetails);
    return { ...plan, ...planDetails as PlannerPlanDetails };
  }

  private async updatePlanDetails(options: Options, planId: string): Promise<PlannerPlan & PlannerPlanDetails> {
    const plan = await planner.getPlanById(planId);

    const categories: any = {};
    let categoriesCount: number = 0;

    Object.keys(options).forEach(key => {
      if (key.indexOf('category') !== -1) {
        categories[key] = (options as any)[key];
        categoriesCount++;
      }
    });

    if (!options.shareWithUserIds && !options.shareWithUserNames && categoriesCount === 0) {
      return this.getPlanDetails(plan);
    }

    const requestBody: any = {};

    if (options.shareWithUserIds || options.shareWithUserNames) {
      const sharedWith = await this.generateSharedWith(options);
      requestBody['sharedWith'] = sharedWith;
    }

    if (categoriesCount > 0) {
      requestBody['categoryDescriptions'] = categories;
    }

    const etag = await this.getPlanDetailsEtag(planId);

    const requestOptionsPlanDetails: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/plans/${planId}/details`,
      headers: {
        'accept': 'application/json;odata.metadata=none',
        'If-Match': etag,
        'Prefer': 'return=representation'
      },
      responseType: 'json',
      data: requestBody
    };

    const planDetails = await request.patch(requestOptionsPlanDetails);
    return { ...plan, ...planDetails as PlannerPlanDetails };
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const planId: string = await this.getPlanId(args);

      if (args.options.newTitle) {
        const etag = await this.getPlanEtag(planId);

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/planner/plans/${planId}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'If-Match': etag,
            'Prefer': 'return=representation'
          },
          responseType: 'json',
          data: {
            "title": args.options.newTitle
          }
        };

        await request.patch(requestOptions);
      }

      const result = await this.updatePlanDetails(args.options, planId);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PlannerPlanSetCommand();
