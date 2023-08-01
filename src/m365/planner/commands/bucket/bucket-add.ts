import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  planId?: string;
  planTitle?: string;
  rosterId?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
  orderHint?: string;
}

class PlannerBucketAddCommand extends GraphCommand {
  public get name(): string {
    return commands.BUCKET_ADD;
  }

  public get description(): string {
    return 'Adds a new Microsoft Planner bucket';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'planId', 'orderHint'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        planId: typeof args.options.planId !== 'undefined',
        planTitle: typeof args.options.planTitle !== 'undefined',
        rosterId: typeof args.options.rosterId !== 'undefined',
        ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
        ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
        orderHint: typeof args.options.orderHint !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: "--planId [planId]"
      },
      {
        option: "--planTitle [planTitle]"
      },
      {
        option: '--rosterId [rosterId]'
      },
      {
        option: "--ownerGroupId [ownerGroupId]"
      },
      {
        option: "--ownerGroupName [ownerGroupName]"
      },
      {
        option: "--orderHint [orderHint]"
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId as string)) {
          return `${args.options.ownerGroupId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['planId', 'planTitle', 'rosterId'] },
      {
        options: ['ownerGroupId', 'ownerGroupName'],
        runsWhen: (args) => args.options.planTitle !== undefined
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('name', 'planId', 'planTitle', 'ownerGroupName', 'orderHint', 'rosterId ');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const planId = await this.getPlanId(args);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/planner/buckets`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
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
      const plan = await planner.getPlanByTitle(args.options.planTitle, groupId);
      return plan.id!;
    }

    const plan = await planner.getPlanByRosterId(args.options.rosterId!);
    return plan.id!;
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.ownerGroupId) {
      return args.options.ownerGroupId;
    }

    const group = await aadGroup.getGroupByDisplayName(args.options.ownerGroupName!);
    return group.id!;
  }
}

export default new PlannerBucketAddCommand();