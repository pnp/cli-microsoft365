import { PlannerPlan, PlannerPlanDetails, User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import { CommandError } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { aadGroup } from '../../../../utils/aadGroup';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
  rosterId?: string;
  shareWithUserIds?: string;
  shareWithUserNames?: string;
}

class PlannerPlanAddCommand extends GraphCommand {
  public get name(): string {
    return commands.PLAN_ADD;
  }

  public get description(): string {
    return 'Adds a new Microsoft Planner plan';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'createdDateTime', 'owner'];
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
        ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
        ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
        shareWithUserIds: typeof args.options.shareWithUserIds !== 'undefined',
        shareWithUserNames: typeof args.options.shareWithUserNames !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --title <title>'
      },
      {
        option: "--ownerGroupId [ownerGroupId]"
      },
      {
        option: "--ownerGroupName [ownerGroupName]"
      },
      {
        option: "--rosterId [rosterId]"
      },
      {
        option: '--shareWithUserIds [shareWithUserIds]'
      },
      {
        option: '--shareWithUserNames [shareWithUserNames]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId as string)) {
          return `${args.options.ownerGroupId} is not a valid GUID`;
        }

        if (args.options.shareWithUserIds && args.options.shareWithUserNames) {
          return 'Specify either shareWithUserIds or shareWithUserNames but not both';
        }

        if (args.options.shareWithUserIds && !validation.isValidGuidArray(args.options.shareWithUserIds.split(','))) {
          return 'shareWithUserIds contains invalid GUID';
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['ownerGroupId', 'ownerGroupName', 'rosterId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const data: any = {
        title: args.options.title
      };
      if (args.options.rosterId) {
        data.container = {
          "url": `https://graph.microsoft.com/v1.0/planner/rosters/${args.options.rosterId}`
        };
      }
      else {
        const groupId = await this.getGroupId(args);
        data.owner = groupId;
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/planner/plans`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: data
      };

      const newPlan = await request.post<any>(requestOptions);
      const result = await this.updatePlanDetails(args.options, newPlan);
      logger.log(result);
    }
    catch (err: any) {
      if (err.error && err.error.error.message === "You do not have the required permissions to access this item, or the item may not exist.") {
        throw new CommandError("You can only add 1 plan to a Roster");
      }
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async updatePlanDetails(options: Options, newPlan: PlannerPlan): Promise<PlannerPlan & PlannerPlanDetails> {
    const planId = newPlan.id as string;

    if (!options.shareWithUserIds && !options.shareWithUserNames) {
      return newPlan;
    }

    const resArray = await Promise.all([this.generateSharedWith(options), this.getPlanDetailsEtag(planId)]);
    const sharedWith = resArray[0];
    const etag = resArray[1];

    const requestOptionsPlanDetails: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/plans/${planId}/details`,
      headers: {
        'accept': 'application/json;odata.metadata=none',
        'If-Match': etag,
        'Prefer': 'return=representation'
      },
      responseType: 'json',
      data: {
        sharedWith: sharedWith
      }
    };

    const planDetails = await request.patch(requestOptionsPlanDetails);
    return { ...newPlan, ...planDetails as PlannerPlanDetails };
  }

  private async getPlanDetailsEtag(planId: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/plans/${planId}/details`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    const response = await request.get<any>(requestOptions);
    return response['@odata.etag'];
  }

  private async generateSharedWith(options: Options): Promise<{ [userId: string]: boolean }> {
    const sharedWith: { [userId: string]: boolean } = {};

    const userIds = await this.getUserIds(options);

    userIds.map(x => sharedWith[x] = true);

    return sharedWith;
  }

  private async getUserIds(options: Options): Promise<string[]> {
    if (options.shareWithUserIds) {
      return options.shareWithUserIds.split(',');
    }

    // Hitting this section means assignedToUserNames won't be undefined
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
      throw `Cannot proceed with planner plan creation. The following users provided are invalid : ${invalidUsers.join(',')}`;
    }

    return userIds;
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.ownerGroupId) {
      return args.options.ownerGroupId;
    }

    const group = await aadGroup.getGroupByDisplayName(args.options.ownerGroupName!);
    return group.id!;
  }
}

module.exports = new PlannerPlanAddCommand();