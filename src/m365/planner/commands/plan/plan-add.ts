import { PlannerPlan, PlannerPlanDetails, User } from '@microsoft/microsoft-graph-types';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { aadGroup } from '../../../../utils/aadGroup';
import { accessToken } from '../../../../utils/accessToken';
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
    this.optionSets.push({ options: ['ownerGroupId', 'ownerGroupName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.');
      return;
    }

    try {
      const groupId = await this.getGroupId(args);
      const requestOptions: any = {
        url: `${this.resource}/v1.0/planner/plans`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          owner: groupId,
          title: args.options.title
        }
      };

      const newPlan = await request.post<any>(requestOptions);
      const result = await this.updatePlanDetails(args.options, newPlan);
      logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private updatePlanDetails(options: Options, newPlan: PlannerPlan): Promise<PlannerPlan & PlannerPlanDetails> {
    const planId = newPlan.id as string;

    if (!options.shareWithUserIds && !options.shareWithUserNames) {
      return Promise.resolve(newPlan);
    }

    return Promise
      .all([this.generateSharedWith(options), this.getPlanDetailsEtag(planId)])
      .then(resArray => {
        const sharedWith = resArray[0];
        const etag = resArray[1];
        const requestOptionsPlanDetails: any = {
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

        return request.patch(requestOptionsPlanDetails);
      })
      .then(planDetails => {
        return { ...newPlan, ...planDetails as PlannerPlanDetails };
      });
  }

  private getPlanDetailsEtag(planId: string): Promise<string> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/planner/plans/${planId}/details`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    return request
      .get(requestOptions)
      .then((response: any) => response['@odata.etag']);
  }

  private generateSharedWith(options: Options): Promise<{ [userId: string]: boolean }> {
    const sharedWith: { [userId: string]: boolean } = {};

    return this
      .getUserIds(options)
      .then((userIds) => {
        userIds.map(x => sharedWith[x] = true);

        return Promise.resolve(sharedWith);
      });
  }

  private getUserIds(options: Options): Promise<string[]> {
    if (options.shareWithUserIds) {
      return Promise.resolve(options.shareWithUserIds.split(','));
    }

    // Hitting this section means assignedToUserNames won't be undefined
    const userNames = options.shareWithUserNames as string;
    const userArr: string[] = userNames.split(',').map(o => o.trim());

    const promises: Promise<{ value: User[] }>[] = userArr.map(user => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user)}'&$select=id,userPrincipalName`,
        headers: {
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      return request.get(requestOptions);
    });

    return Promise
      .all(promises)
      .then((usersRes: { value: User[] }[]): Promise<string[]> => {
        const userUpns = usersRes.map(res => res.value[0]?.userPrincipalName as string);
        const userIds = usersRes.map(res => res.value[0]?.id as string);

        // Find the members where no graph response was found
        const invalidUsers = userArr.filter(user => !userUpns.some((upn) => upn?.toLowerCase() === user.toLowerCase()));

        if (invalidUsers && invalidUsers.length > 0) {
          return Promise.reject(`Cannot proceed with planner plan creation. The following users provided are invalid : ${invalidUsers.join(',')}`);
        }

        return Promise.resolve(userIds);
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

module.exports = new PlannerPlanAddCommand();