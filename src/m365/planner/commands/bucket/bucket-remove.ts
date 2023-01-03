import { PlannerBucket } from '@microsoft/microsoft-graph-types';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { accessToken } from '../../../../utils/accessToken';
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
  name?: string;
  planId?: string;
  planTitle?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
  confirm?: boolean;
}

class PlannerBucketRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.BUCKET_REMOVE;
  }

  public get description(): string {
    return 'Removes the Microsoft Planner bucket from a plan';
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
        name: typeof args.options.name !== 'undefined',
        planId: typeof args.options.planId !== 'undefined',
        planTitle: typeof args.options.planTitle !== 'undefined',
        ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
        ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
        confirm: args.options.confirm || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '--planId [planId]'
      },
      {
        option: "--planTitle [planTitle]"
      },
      {
        option: '--ownerGroupId [ownerGroupId]'
      },
      {
        option: '--ownerGroupName [ownerGroupName]'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id) {
          if (args.options.planId || args.options.planTitle || args.options.ownerGroupId || args.options.ownerGroupName) {
            return 'Don\'t specify planId, planTitle, ownerGroupId or ownerGroupName when using id';
          }
        }
        else {
          if (!args.options.planId && !args.options.planTitle) {
            return 'Specify either planId or planTitle when using name';
          }

          if (args.options.planId && args.options.planTitle) {
            return 'Specify either planId or planTitle when using name but not both';
          }

          if (args.options.planTitle) {
            if (!args.options.ownerGroupId && !args.options.ownerGroupName) {
              return 'Specify either ownerGroupId or ownerGroupName when using planTitle';
            }

            if (args.options.ownerGroupId && args.options.ownerGroupName) {
              return 'Specify either ownerGroupId or ownerGroupName when using planTitle but not both';
            }

            if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
              return `${args.options.ownerGroupId} is not a valid GUID`;
            }
          }
          else {
            if (args.options.ownerGroupId || args.options.ownerGroupName) {
              return 'Don\'t specify ownerGroupId or ownerGroupName when using planId';
            }
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'name'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.');
      return;
    }

    const removeBucket: () => Promise<void> = async (): Promise<void> => {
      try {
        const bucket = await this.getBucket(args);

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/planner/buckets/${bucket.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'if-match': (bucket as any)['@odata.etag']
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await removeBucket();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the bucket ${args.options.id || args.options.name}?`
      });

      if (result.continue) {
        await removeBucket();
      }
    }
  }

  private getBucket(args: CommandArgs): Promise<PlannerBucket> {
    if (args.options.id) {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/planner/buckets/${args.options.id}`,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      return request.get<PlannerBucket>(requestOptions);
    }

    return this
      .getPlanId(args)
      .then(planId => {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/planner/plans/${planId}/buckets`,
          headers: {
            accept: 'application/json'
          },
          responseType: 'json'
        };

        return request.get<{ value: PlannerBucket[] }>(requestOptions);
      })
      .then(buckets => {
        const filteredBuckets = buckets.value.filter(b => args.options.name!.toLowerCase() === b.name!.toLowerCase());

        if (!filteredBuckets.length) {
          return Promise.reject(`The specified bucket ${args.options.name} does not exist`);
        }

        if (filteredBuckets.length > 1) {
          return Promise.reject(`Multiple buckets with name ${args.options.name} found: ${filteredBuckets.map(x => x.id)}`);
        }

        return Promise.resolve(filteredBuckets[0]);
      });
  }

  private getPlanId(args: CommandArgs): Promise<string> {
    const { planId, planTitle } = args.options;

    if (planId) {
      return Promise.resolve(planId);
    }

    return this
      .getGroupId(args)
      .then(groupId => planner.getPlanByTitle(planTitle!, groupId))
      .then(plan => plan.id!);
  }

  private getGroupId(args: CommandArgs): Promise<string> {
    const { ownerGroupId, ownerGroupName } = args.options;

    if (ownerGroupId) {
      return Promise.resolve(ownerGroupId);
    }

    return aadGroup
      .getGroupByDisplayName(ownerGroupName!)
      .then(group => group.id!);
  }
}

module.exports = new PlannerBucketRemoveCommand();