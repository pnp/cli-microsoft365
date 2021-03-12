import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import request from '../../../../request';

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
  public get name(): string {
    return commands.POLICY_LIST;
  }

  public get description(): string {
    return 'Returns policies from Azure AD';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.policyType = args.options.policyType || 'all';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'isOrganizationDefault'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {

    const policyType: string = args.options.policyType ? args.options.policyType.toLowerCase() : 'all';
    if (policyType && policyType !== "all") {
      const endpoint = policyEndPoints[policyType];
      const url: string = `${this.resource}/v1.0/policies/${endpoint}`;
      if (endpoint === policyEndPoints.authorization || endpoint === policyEndPoints.identitysecuritydefaultsenforcement) {
        this.getPolicy(url)
          .then((policy: any): void => {
            if (policy) {
              logger.log(policy);
            }
            cb();
          }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
      }
      else {
        this.getPolicies(url)
          .then((policies: any[]): void => {
            if (policies && policies.length > 0) {
              logger.log(policies);
            }
            cb();
          }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
      }
    }
    else {
      let promiseCalls: any[] = [];
      const policyTypes: string[] = Object.keys(policyEndPoints);
      policyTypes.forEach((policyType) => {
        let endpoint = policyEndPoints[policyType];
        const url: string = `${this.resource}/v1.0/policies/${endpoint}`;
        if (endpoint === policyEndPoints.authorization || endpoint === policyEndPoints.identitysecuritydefaultsenforcement) {
          promiseCalls.push(this.getPolicy(url));
        }
        else {
          promiseCalls.push(this.getPolicies(url));
        }
      }
      );
      Promise.all(promiseCalls)
        .then(((results) => {
          let allPolicies: any = [];
          results.forEach((policies: any) => {
            allPolicies = allPolicies.concat(policies);
          });
          if (allPolicies && allPolicies.length > 0) {
            logger.log(allPolicies);
          }
          cb();
        }))
        .catch(err => {
          this.handleRejectedODataJsonPromise(err, logger, cb)
        });
    }
  }

  private async getPolicies(url: string): Promise<any[]> {
    const requestOptions: any = {
      url: url,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const response = await request
        .get<{ value: any[]; }>(requestOptions);
      return await Promise.resolve(response.value);
    } catch (error) {
      return await Promise.reject(error);
    }
  }

  private async getPolicy(url: string): Promise<any> {
    const requestOptions: any = {
      url: url,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const response = await request
        .get<{ value: any; }>(requestOptions);
      return await Promise.resolve(response);
    } catch (error) {
      return await Promise.reject(error);
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-p, --policyType [policyType]',
        autocomplete: ['activityBasedTimeout', 'authorization', 'claimsMapping', 'homeRealmDiscovery', 'identitySecurityDefaultsEnforcement', 'tokenIssuance', 'tokenLifetime']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.policyType) {
      const policyType: string = args.options.policyType.toLowerCase()
      if (policyType !== 'activitybasedtimeout' &&
        policyType !== 'authorization' &&
        policyType !== 'claimsmapping' &&
        policyType !== 'homerealmdiscovery' &&
        policyType !== 'identitysecuritydefaultsenforcement' &&
        policyType !== 'tokenissuance' &&
        policyType !== 'tokenlifetime') {
        return `${policyType} is not a valid policyType. Allowed values are activityBasedTimeout|authorization|claimsMapping|homeRealmDiscovery|identitySecurityDefaultsEnforcement|tokenIssuance|tokenLifetime`;
      }
    }

    return true;
  }
}

module.exports = new AadPolicyListCommand();
