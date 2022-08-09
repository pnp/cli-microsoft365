import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  policyType?: string;
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

class AadPolicyListCommand extends GraphCommand {
  private static readonly supportedPolicyTypes: string[] = ['activityBasedTimeout', 'authorization', 'claimsMapping', 'homeRealmDiscovery', 'identitySecurityDefaultsEnforcement', 'tokenIssuance', 'tokenLifetime'];

  public get name(): string {
    return commands.POLICY_LIST;
  }

  public get description(): string {
    return 'Returns policies from Azure AD';
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
        policyType: args.options.policyType || 'all'
      });
    });
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-p, --policyType [policyType]',
        autocomplete: AadPolicyListCommand.supportedPolicyTypes
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.policyType) {
          const policyType: string = args.options.policyType.toLowerCase();
          if (!AadPolicyListCommand.supportedPolicyTypes.find(p => p.toLowerCase() === policyType)) {
            return `${args.options.policyType} is not a valid policyType. Allowed values are ${AadPolicyListCommand.supportedPolicyTypes.join(', ')}`;
          }
        }
    
        return true;
      }
    );
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'isOrganizationDefault'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const policyType: string = args.options.policyType ? args.options.policyType.toLowerCase() : 'all';

    if (policyType && policyType !== "all") {
      this
        .getPolicies(policyType)
        .then((policies?: any): void => {
          logger.log(policies);
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }
    else {
      const policyTypes: string[] = Object.keys(policyEndPoints);
      Promise
        .all(policyTypes.map(policyType => this.getPolicies(policyType)))
        .then((results: any[]) => {
          let allPolicies: any = [];
          results.forEach((policies: any) => {
            allPolicies = allPolicies.concat(policies);
          });

          logger.log(allPolicies);
          cb();
        }, err => this.handleRejectedODataJsonPromise(err, logger, cb));
    }
  }

  private getPolicies(policyType: string): Promise<any> {
    const endpoint = policyEndPoints[policyType];
    const requestOptions: any = {
      url: `${this.resource}/v1.0/policies/${endpoint}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get(requestOptions)
      .then((response: any) => {
        if (endpoint === policyEndPoints.authorization ||
          endpoint === policyEndPoints.identitysecuritydefaultsenforcement) {
          return Promise.resolve(response);
        }
        else {
          return Promise.resolve(response.value);
        }
      });
  }
}

module.exports = new AadPolicyListCommand();
