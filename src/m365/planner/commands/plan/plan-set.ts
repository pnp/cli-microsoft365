import { PlannerPlan, PlannerPlanDetails, User } from '@microsoft/microsoft-graph-types';
import { AxiosRequestConfig } from 'axios';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { accessToken } from '../../../../utils/accessToken';
import { formatting } from '../../../../utils/formatting';
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
  newTitle?: string;
  shareWithUserIds?: string;
  shareWithUserNames?: string;
}

class PlannerPlanSetCommand extends GraphCommand {
  public get name(): string {
    return commands.PLAN_SET;
  }

  public get description(): string {
    return 'Updates a Microsoft Planner plan';
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
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
        ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
        newTitle: typeof args.options.newTitle !== 'undefined',
        shareWithUserIds: typeof args.options.shareWithUserIds !== 'undefined',
        shareWithUserNames: typeof args.options.shareWithUserNames !== 'undefined'
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
      },
      {
        option: '--newTitle [newTitle]'
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
        if (args.options.title) {
          if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId as string)) {
            return `${args.options.ownerGroupId} is not a valid GUID`;
          }

          if (!args.options.ownerGroupId && !args.options.ownerGroupName) {
            return 'Specify either ownerGroupId or ownerGroupName when using title';
          }

          if (args.options.ownerGroupId && args.options.ownerGroupName) {
            return 'Specify either ownerGroupId or ownerGroupName when using title but not both';
          }
        }

        if (args.options.shareWithUserIds && args.options.shareWithUserNames) {
          return 'Specify either shareWithUserIds or shareWithUserNames but not both';
        }

        if (args.options.shareWithUserIds && !validation.isValidGuidArray(args.options.shareWithUserIds.split(','))) {
          return 'shareWithUserIds contains invalid GUID';
        }

        const allowedCategories: string[] = [
          'category1',
          'category2',
          'category3',
          'category4',
          'category5',
          'category6',
          'category7',
          'category8',
          'category9',
          'category10',
          'category11',
          'category12',
          'category13',
          'category14',
          'category15',
          'category16',
          'category17',
          'category18',
          'category19',
          'category20',
          'category21',
          'category22',
          'category23',
          'category24',
          'category25'
        ];

        let invalidCategoryOptions: boolean = false;
        Object.keys(args.options).forEach(key => {
          if (key.indexOf('category') !== -1 && allowedCategories.indexOf(key) === -1) {
            invalidCategoryOptions = true;
          }
        });

        if (invalidCategoryOptions) {
          return 'Specify category values between category1 to category25';
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'title']);
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    const { ownerGroupId, ownerGroupName } = args.options;

    if (ownerGroupId) {
      return ownerGroupId;
    }

    const group = await aadGroup.getGroupByDisplayName(ownerGroupName!);
    return group.id!;
  }

  private async getPlanId(args: CommandArgs): Promise<string> {
    const { id, title } = args.options;

    if (id) {
      return id;
    }

    const groupId: string = await this.getGroupId(args);
    const plan: PlannerPlan = await planner.getPlanByTitle(title!, groupId);
    return plan.id!;
  }

  private getUserIds(options: Options): Promise<string[]> {
    if (options.shareWithUserIds) {
      return Promise.resolve(options.shareWithUserIds.split(','));
    }

    const userNames = options.shareWithUserNames as string;
    const userArr: string[] = userNames.split(',').map(o => o.trim());

    const promises: Promise<{ value: User[] }>[] = userArr.map(user => {
      const requestOptions: AxiosRequestConfig = {
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
          return Promise.reject(`Cannot proceed with planner plan creation. The following users provided are invalid: ${invalidUsers.join(',')}`);
        }

        return Promise.resolve(userIds);
      });
  }

  private async generateSharedWith(options: Options): Promise<{ [userId: string]: boolean }> {
    const sharedWith: { [userId: string]: boolean } = {};

    const userIds = await this.getUserIds(options);
    userIds.map(x => sharedWith[x] = true);
    return sharedWith;
  }

  private async getPlanEtag(planId: string): Promise<string> {
    const requestOptions: AxiosRequestConfig = {
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
    const requestOptions: AxiosRequestConfig = {
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
    const requestOptionsTaskDetails: AxiosRequestConfig = {
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
        categories[key] = options[key];
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

    const requestOptionsPlanDetails: AxiosRequestConfig = {
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
    if (accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.');
      return;
    }

    try {
      const planId: string = await this.getPlanId(args);

      if (args.options.newTitle) {
        const etag = await this.getPlanEtag(planId);

        const requestOptions: AxiosRequestConfig = {
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
      logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PlannerPlanSetCommand();