import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const policyEndPoints: any = {
  activitybasedtimeout: "activityBasedTimeoutPolicies",
  adminconsentrequest: "adminConsentRequestPolicy",
  appManagement: "appManagementPolicies",
  authenticationflows: "authenticationFlowsPolicy",
  authenticationmethods: "authenticationMethodsPolicy",
  authenticationstrength: "authenticationStrengthPolicies",
  authorization: "authorizationPolicy",
  claimsmapping: "claimsMappingPolicies",
  conditionalaccess: "conditionalAccessPolicies",
  crosstenantaccess: "crossTenantAccessPolicy",
  defaultappmanagement: "defaultAppManagementPolicy",
  deviceregistration: "deviceRegistrationPolicy",
  featurerolloutpolicy: "featureRolloutPolicies",
  homerealmdiscovery: "homeRealmDiscoveryPolicies",
  identitysecuritydefaultsenforcement: "identitySecurityDefaultsEnforcementPolicy",
  permissiongrant: "permissionGrantPolicies",
  rolemanagement: "roleManagementPolicies",
  tokenissuance: "tokenIssuancePolicies",
  tokenlifetime: "tokenLifetimePolicies"
};

const supportedPolicyTypes = ['activityBasedTimeout', 'adminConsentRequest', 'appManagement', 'authenticationFlows', 'authenticationMethods', 'authenticationStrength', 'authorization', 'claimsMapping', 'conditionalAccess', 'crossTenantAccess', 'defaultAppManagement', 'deviceRegistration', 'featureRolloutPolicy', 'homeRealmDiscovery', 'identitySecurityDefaultsEnforcement', 'permissionGrant', 'roleManagement', 'tokenIssuance', 'tokenLifetime'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  type: z.enum(supportedPolicyTypes).optional().alias('t')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraPolicyListCommand extends GraphCommand {
  public get name(): string {
    return commands.POLICY_LIST;
  }

  public get description(): string {
    return 'Returns policies from Entra ID';
  }

  public get schema(): z.ZodType | undefined {
    return options;
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

    let requestUrl = `${this.resource}/v1.0/policies/${endpoint}`;

    if (endpoint === policyEndPoints.rolemanagement) {
      // roleManagementPolicies endpoint requires $filter query parameter
      requestUrl += `?$filter=scopeId eq '/' and scopeType eq 'DirectoryRole'`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<any>(requestOptions);

    if (endpoint === policyEndPoints.adminconsentrequest ||
      endpoint === policyEndPoints.authenticationflows ||
      endpoint === policyEndPoints.authenticationmethods ||
      endpoint === policyEndPoints.authorization ||
      endpoint === policyEndPoints.crosstenantaccess ||
      endpoint === policyEndPoints.defaultappmanagement ||
      endpoint === policyEndPoints.deviceregistration ||
      endpoint === policyEndPoints.identitysecuritydefaultsenforcement) {
      return response;
    }
    else {
      return response.value;
    }
  }
}

export default new EntraPolicyListCommand();