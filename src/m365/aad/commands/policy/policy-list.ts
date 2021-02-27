import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';

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
      const url: string = `${this.resource}/v1.0/policies/${policyType}`;

      this
        .getAllItems(url, logger, true)
        .then((): void => {
          logger.log(this.items);
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }
    else {
      const policyTypes: string[] = ['activityBasedTimeout', 'claimsMapping', 'homeRealmDiscovery', 'tokenLifetime', 'tokenIssuance', 'identitySecurityDefaultsEnforcement'];
      let policies: any = [];

      policyTypes.forEach((policyType) => {
        const url: string = `${this.resource}/v1.0/policies/${policyType}`;
        this
          .getAllItems(url, logger, true)
          .then((): void => {
            policies.push(this.items);
          }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));

      });

      logger.log(policies);
      cb();
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-p, --policyType [policyType]',
        autocomplete: ['activityBasedTimeout', 'claimsMapping', 'homeRealmDiscovery', 'tokenLifetime', 'tokenIssuance', 'identitySecurityDefaultsEnforcement']
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
        policyType !== 'tokenissuance' &&
        policyType !== 'identitysecuritydefaultsenforcement') {
        return `${policyType} is not a valid policyType. Allowed values are activityBasedTimeout|claimsMapping|homeRealmDiscovery|tokenLifetime|tokenIssuance|identitySecurityDefaultsEnforcement`;
      }
    }

    return true;
  }
}

module.exports = new AadPolicyListCommand();
