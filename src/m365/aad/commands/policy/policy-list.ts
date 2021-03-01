import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import request from '../../../../request';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  policyType?: string;
}

class AadPolicyListCommand extends GraphItemsListCommand<any> {
  public get name(): string {
    return commands.POLICY_LIST;
  }

  public get description(): string {
    return 'Returns policies from Azure AD';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.policyType = args.options.policyType || 'All';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const policyType: string = args.options.policyType ? args.options.policyType : 'all';
    if (policyType && policyType !== "all") {
      let endpoint = this.getPolicyEndPoint[policyType];
      const url: string = `${this.resource}/v1.0/policies/${endpoint}`;
      ((): Promise<any[]> => {
        return this.getPolicies(url);
      })()
        .then((policies: any[]): void => {
          if (policies && policies.length > 0) {
            logger.log(policies);
          }
          cb();
        }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
    }
    else {
      const policyTypes: string[] = ['activityBasedTimeout', 'claimsMapping', 'homeRealmDiscovery', 'tokenLifetime', 'tokenIssuance'];
      let promiseCalls: any[] = [];
      policyTypes.forEach((policyType) => {
        let endpoint = this.getPolicyEndPoint[policyType];
        const url: string = `${this.resource}/v1.0/policies/${endpoint}`;
        promiseCalls.push(
          ((): Promise<any[]> => {
            return this.getPolicies(url);
          })()
        );
      });
      Promise.all(promiseCalls)
        .then(((results) => {
          let allPolicies: any[] = [];
          results.forEach((policies: any[]) => {
            allPolicies = allPolicies.concat(policies);
          });
          if (allPolicies && allPolicies.length > 0) {
            logger.log(allPolicies);
          }
          cb();
        }))
        .catch(err => {
          this.handleRejectedPromise(err, logger, cb)
        });
    }
  }

  private getPolicyEndPoint: { [key: string]: string } = {
    activityBasedTimeout: "activityBasedTimeoutPolicies",
    claimsMapping: "claimsMappingPolicies",
    homeRealmDiscovery: "homeRealmDiscoveryPolicies",
    tokenLifetime: "tokenLifetimePolicies",
    tokenIssuance: "tokenIssuancePolicies"
  }

  private getPolicies(url: string): Promise<any[]> {
    const requestOptions: any = {
      url: url,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return new Promise<any[]>((resolve: (list: any[]) => void, reject: (error: any) => void): void => {
      request
        .get<{ value: any[]; }>(requestOptions)
        .then((response: { value: any[] }) => {
          resolve(response.value);
        })
        .catch((error: any) => {
          reject(error);
        });
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-p, --policyType [policyType]',
        autocomplete: ['activityBasedTimeout', 'claimsMapping', 'homeRealmDiscovery', 'tokenLifetime', 'tokenIssuance']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.policyType) {
      const policyType: string = args.options.policyType.toLowerCase()
      if (policyType !== 'activitybasedtimeout' &&
        policyType !== 'claimsmapping' &&
        policyType !== 'homerealmdiscovery' &&
        policyType !== 'tokenlifetime' &&
        policyType !== 'tokenissuance') {
        return `${policyType} is not a valid policyType. Allowed values are activityBasedTimeout|claimsMapping|homeRealmDiscovery|tokenLifetime|tokenIssuance`;
      }
    }

    return true;
  }
}

module.exports = new AadPolicyListCommand();
