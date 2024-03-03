import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import aadCommands from '../../aadCommands.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type?: string;
}

const policyEndPoints: any = {
  activitybasedtimeout: "activityBasedTimeoutPolicies",
  authorization: "authorizationPolicy",
  claimsmapping: "claimsMappingPolicies",
  homerealmdiscovery: "homeRealmDiscoveryPolicies",
  identitysecuritydefaultsenforcement: "identitySecurityDefaultsEnforcementPolicy",
  tokenissuance: "tokenIssuancePolicies",
  tokenlifetime: "tokenLifetimePolicies"
};

class EntraPolicyListCommand extends GraphCommand {
  private static readonly supportedPolicyTypes: string[] = ['activityBasedTimeout', 'authorization', 'claimsMapping', 'homeRealmDiscovery', 'identitySecurityDefaultsEnforcement', 'tokenIssuance', 'tokenLifetime'];

  public get name(): string {
    return commands.POLICY_LIST;
  }

  public get description(): string {
    return 'Returns policies from Entra ID';
  }

  public alias(): string[] | undefined {
    return [aadCommands.POLICY_LIST];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        policyType: args.options.type || 'all'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --type [type]',
        autocomplete: EntraPolicyListCommand.supportedPolicyTypes
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.type) {
          const policyType: string = args.options.type.toLowerCase();
          if (!EntraPolicyListCommand.supportedPolicyTypes.find(p => p.toLowerCase() === policyType)) {
            return `${args.options.type} is not a valid type. Allowed values are ${EntraPolicyListCommand.supportedPolicyTypes.join(', ')}`;
          }
        }

        return true;
      }
    );
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'isOrganizationDefault'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const policyType: string = args.options.type ? args.options.type.toLowerCase() : 'all';

    try {
      if (policyType && policyType !== "all") {
        const policies = await this.getPolicies(policyType);
        await logger.log(policies);
      }
      else {
        const policyTypes: string[] = Object.keys(policyEndPoints);
        const results = await Promise.all(policyTypes.map(policyType => this.getPolicies(policyType)));
        let allPolicies: any = [];
        results.forEach((policies: any) => {
          allPolicies = allPolicies.concat(policies);
        });

        await logger.log(allPolicies);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getPolicies(policyType: string): Promise<any> {
    const endpoint = policyEndPoints[policyType];
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/policies/${endpoint}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<any>(requestOptions);

    if (endpoint === policyEndPoints.authorization ||
      endpoint === policyEndPoints.identitysecuritydefaultsenforcement) {
      return response;
    }
    else {
      return response.value;
    }
  }
}

export default new EntraPolicyListCommand();