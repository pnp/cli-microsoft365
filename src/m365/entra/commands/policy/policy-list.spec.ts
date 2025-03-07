import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './policy-list.js';

describe(commands.POLICY_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.POLICY_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'isOrganizationDefault']);
  });

  it('retrieves the specified policy', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/policies/authorizationPolicy`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authorizationPolicy/$entity",
          "@odata.id": "https://graph.microsoft.com/v2/b30f2eac-f6b4-4f87-9dcb-cdf7ae1f8923/authorizationPolicy/authorizationPolicy",
          "id": "authorizationPolicy",
          "allowInvitesFrom": "everyone",
          "allowedToSignUpEmailBasedSubscriptions": true,
          "allowedToUseSSPR": true,
          "allowEmailVerifiedUsersToJoinOrganization": true,
          "blockMsolPowerShell": null,
          "displayName": "Authorization Policy",
          "description": "Used to manage authorization related settings across the company.",
          "defaultUserRolePermissions": {
            "allowedToCreateApps": true,
            "allowedToCreateSecurityGroups": true,
            "allowedToReadOtherUsers": true,
            "permissionGrantPoliciesAssigned": [
              "ManagePermissionGrantsForSelf.microsoft-user-default-legacy"
            ]
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        type: "authorization"
      }
    });
    assert(loggerLogSpy.calledWith({
      "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authorizationPolicy/$entity",
      "@odata.id": "https://graph.microsoft.com/v2/b30f2eac-f6b4-4f87-9dcb-cdf7ae1f8923/authorizationPolicy/authorizationPolicy",
      "id": "authorizationPolicy",
      "allowInvitesFrom": "everyone",
      "allowedToSignUpEmailBasedSubscriptions": true,
      "allowedToUseSSPR": true,
      "allowEmailVerifiedUsersToJoinOrganization": true,
      "blockMsolPowerShell": null,
      "displayName": "Authorization Policy",
      "description": "Used to manage authorization related settings across the company.",
      "defaultUserRolePermissions": {
        "allowedToCreateApps": true,
        "allowedToCreateSecurityGroups": true,
        "allowedToReadOtherUsers": true,
        "permissionGrantPoliciesAssigned": [
          "ManagePermissionGrantsForSelf.microsoft-user-default-legacy"
        ]
      }
    }));
  });

  it('retrieves the specified policies', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/policies/tokenLifetimePolicies`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/tokenLifetimePolicies",
          "value": [
            {
              id: 'a457c42c-0f2e-4a25-be2a-545e840add1f',
              deletedDateTime: null,
              definition: [
                '{"TokenLifetimePolicy":{"Version":1,"AccessTokenLifetime":"8:00:00"}}'
              ],
              displayName: 'TokenLifetimePolicy1',
              isOrganizationDefault: true
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        type: "tokenLifetime"
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        id: 'a457c42c-0f2e-4a25-be2a-545e840add1f',
        deletedDateTime: null,
        definition: [
          '{"TokenLifetimePolicy":{"Version":1,"AccessTokenLifetime":"8:00:00"}}'
        ],
        displayName: 'TokenLifetimePolicy1',
        isOrganizationDefault: true
      }
    ]));
  });

  it('retrieves all policies', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/policies/activityBasedTimeoutPolicies`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/activityBasedTimeoutPolicies",
          "value": []
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/adminConsentRequestPolicy`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/adminConsentRequestPolicy/$entity",
          "isEnabled": false,
          "notifyReviewers": false,
          "remindersEnabled": false,
          "requestDurationInDays": 0,
          "version": 0,
          "reviewers": []
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/appManagementPolicies`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/appManagementPolicies",
          "value": [
            {
              "id": "db9d4b58-3488-4da4-9994-49773c454e33",
              "displayName": "Custom app management policy",
              "description": "Custom policy that enforces app management restrictions on specific applications and service principals.",
              "isEnabled": false,
              "restrictions": {
                "passwordCredentials": [
                  {
                    "restrictionType": "passwordAddition",
                    "maxLifetime": null,
                    "restrictForAppsCreatedAfterDateTime": "2019-10-19T10:37:00Z"
                  },
                  {
                    "restrictionType": "passwordLifetime",
                    "maxLifetime": "P4DT12H30M5S",
                    "restrictForAppsCreatedAfterDateTime": "2017-10-19T10:37:00Z"
                  },
                  {
                    "restrictionType": "symmetricKeyAddition",
                    "maxLifetime": null,
                    "restrictForAppsCreatedAfterDateTime": "2021-10-19T10:37:00Z"
                  },
                  {
                    "restrictionType": "symmetricKeyLifetime",
                    "maxLifetime": "P4D",
                    "restrictForAppsCreatedAfterDateTime": "2014-10-19T10:37:00Z"
                  }
                ],
                "keyCredentials": [
                  {
                    "restrictionType": "asymmetricKeyLifetime",
                    "maxLifetime": "P90D",
                    "restrictForAppsCreatedAfterDateTime": "2014-10-19T10:37:00Z"
                  }
                ]
              }
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/authenticationFlowsPolicy`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationFlowsPolicy/$entity",
          "id": "authenticationFlowsPolicy",
          "displayName": "Authentication flows policy",
          "description": "Authentication flows policy allows modification of settings related to authentication flows in AAD tenant, such as self-service sign up configuration.",
          "selfServiceSignUp": {
            "isEnabled": false
          }
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#authenticationMethodsPolicy",
          "id": "authenticationMethodsPolicy",
          "displayName": "Authentication Methods Policy",
          "description": "The tenant-wide policy that controls which authentication methods are allowed in the tenant, authentication method registration requirements, and self-service password reset settings",
          "lastModifiedDateTime": "2025-03-03T09:38:22.5409946Z",
          "policyVersion": "1.5",
          "policyMigrationState": "migrationInProgress",
          "registrationEnforcement": {
            "authenticationMethodsRegistrationCampaign": {
              "snoozeDurationInDays": 1,
              "state": "default",
              "excludeTargets": [],
              "includeTargets": [
                {
                  "id": "all_users",
                  "targetType": "group",
                  "targetedAuthenticationMethod": "microsoftAuthenticator"
                }
              ]
            }
          },
          "authenticationMethodConfigurations@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations",
          "authenticationMethodConfigurations": [
            {
              "@odata.type": "#microsoft.graph.fido2AuthenticationMethodConfiguration",
              "id": "Fido2",
              "state": "disabled",
              "isSelfServiceRegistrationAllowed": true,
              "isAttestationEnforced": false,
              "excludeTargets": [],
              "keyRestrictions": {
                "isEnforced": false,
                "enforcementType": "block",
                "aaGuids": []
              },
              "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('Fido2')/microsoft.graph.fido2AuthenticationMethodConfiguration/includeTargets",
              "includeTargets": [
                {
                  "targetType": "group",
                  "id": "all_users",
                  "isRegistrationRequired": false
                }
              ]
            },
            {
              "@odata.type": "#microsoft.graph.microsoftAuthenticatorAuthenticationMethodConfiguration",
              "id": "MicrosoftAuthenticator",
              "state": "enabled",
              "isSoftwareOathEnabled": false,
              "excludeTargets": [],
              "featureSettings": {
                "displayAppInformationRequiredState": {
                  "state": "default",
                  "includeTarget": {
                    "targetType": "group",
                    "id": "all_users"
                  },
                  "excludeTarget": {
                    "targetType": "group",
                    "id": "00000000-0000-0000-0000-000000000000"
                  }
                },
                "displayLocationInformationRequiredState": {
                  "state": "default",
                  "includeTarget": {
                    "targetType": "group",
                    "id": "all_users"
                  },
                  "excludeTarget": {
                    "targetType": "group",
                    "id": "00000000-0000-0000-0000-000000000000"
                  }
                }
              },
              "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('MicrosoftAuthenticator')/microsoft.graph.microsoftAuthenticatorAuthenticationMethodConfiguration/includeTargets",
              "includeTargets": [
                {
                  "targetType": "group",
                  "id": "all_users",
                  "isRegistrationRequired": false,
                  "authenticationMode": "any"
                }
              ]
            },
            {
              "@odata.type": "#microsoft.graph.smsAuthenticationMethodConfiguration",
              "id": "Sms",
              "state": "disabled",
              "excludeTargets": [],
              "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('Sms')/microsoft.graph.smsAuthenticationMethodConfiguration/includeTargets",
              "includeTargets": [
                {
                  "targetType": "group",
                  "id": "all_users",
                  "isRegistrationRequired": false,
                  "isUsableForSignIn": true
                }
              ]
            },
            {
              "@odata.type": "#microsoft.graph.temporaryAccessPassAuthenticationMethodConfiguration",
              "id": "TemporaryAccessPass",
              "state": "disabled",
              "defaultLifetimeInMinutes": 60,
              "defaultLength": 8,
              "minimumLifetimeInMinutes": 60,
              "maximumLifetimeInMinutes": 480,
              "isUsableOnce": false,
              "excludeTargets": [],
              "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('TemporaryAccessPass')/microsoft.graph.temporaryAccessPassAuthenticationMethodConfiguration/includeTargets",
              "includeTargets": [
                {
                  "targetType": "group",
                  "id": "all_users",
                  "isRegistrationRequired": false
                }
              ]
            },
            {
              "@odata.type": "#microsoft.graph.softwareOathAuthenticationMethodConfiguration",
              "id": "SoftwareOath",
              "state": "disabled",
              "excludeTargets": [],
              "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('SoftwareOath')/microsoft.graph.softwareOathAuthenticationMethodConfiguration/includeTargets",
              "includeTargets": [
                {
                  "targetType": "group",
                  "id": "all_users",
                  "isRegistrationRequired": false
                }
              ]
            },
            {
              "@odata.type": "#microsoft.graph.voiceAuthenticationMethodConfiguration",
              "id": "Voice",
              "state": "disabled",
              "isOfficePhoneAllowed": false,
              "excludeTargets": [],
              "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('Voice')/microsoft.graph.voiceAuthenticationMethodConfiguration/includeTargets",
              "includeTargets": [
                {
                  "targetType": "group",
                  "id": "all_users",
                  "isRegistrationRequired": false
                }
              ]
            },
            {
              "@odata.type": "#microsoft.graph.emailAuthenticationMethodConfiguration",
              "id": "Email",
              "state": "enabled",
              "allowExternalIdToUseEmailOtp": "default",
              "excludeTargets": [],
              "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('Email')/microsoft.graph.emailAuthenticationMethodConfiguration/includeTargets",
              "includeTargets": []
            },
            {
              "@odata.type": "#microsoft.graph.x509CertificateAuthenticationMethodConfiguration",
              "id": "X509Certificate",
              "state": "disabled",
              "excludeTargets": [],
              "certificateUserBindings": [
                {
                  "x509CertificateField": "PrincipalName",
                  "userProperty": "userPrincipalName",
                  "priority": 1,
                  "trustAffinityLevel": "low"
                },
                {
                  "x509CertificateField": "RFC822Name",
                  "userProperty": "userPrincipalName",
                  "priority": 2,
                  "trustAffinityLevel": "low"
                },
                {
                  "x509CertificateField": "SubjectKeyIdentifier",
                  "userProperty": "certificateUserIds",
                  "priority": 3,
                  "trustAffinityLevel": "high"
                }
              ],
              "authenticationModeConfiguration": {
                "x509CertificateAuthenticationDefaultMode": "x509CertificateSingleFactor",
                "x509CertificateDefaultRequiredAffinityLevel": "low",
                "rules": []
              },
              "crlValidationConfiguration": {
                "state": "disabled",
                "exemptedCertificateAuthoritiesSubjectKeyIdentifiers": []
              },
              "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('X509Certificate')/microsoft.graph.x509CertificateAuthenticationMethodConfiguration/includeTargets",
              "includeTargets": [
                {
                  "targetType": "group",
                  "id": "all_users",
                  "isRegistrationRequired": false
                }
              ]
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/authenticationStrengthPolicies`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationStrengthPolicies",
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000002",
              "createdDateTime": "2021-12-01T00:00:00Z",
              "modifiedDateTime": "2021-12-01T00:00:00Z",
              "displayName": "Multifactor authentication",
              "description": "Combinations of methods that satisfy strong authentication, such as a password + SMS",
              "policyType": "builtIn",
              "requirementsSatisfied": "mfa",
              "allowedCombinations": [
                "windowsHelloForBusiness",
                "fido2",
                "x509CertificateMultiFactor",
                "deviceBasedPush",
                "temporaryAccessPassOneTime",
                "temporaryAccessPassMultiUse",
                "password,microsoftAuthenticatorPush",
                "password,softwareOath",
                "password,hardwareOath",
                "password,sms",
                "password,voice",
                "federatedMultiFactor",
                "microsoftAuthenticatorPush,federatedSingleFactor",
                "softwareOath,federatedSingleFactor",
                "hardwareOath,federatedSingleFactor",
                "sms,federatedSingleFactor",
                "voice,federatedSingleFactor"
              ],
              "combinationConfigurations@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationStrengthPolicies('00000000-0000-0000-0000-000000000002')/combinationConfigurations",
              "combinationConfigurations": []
            },
            {
              "id": "00000000-0000-0000-0000-000000000003",
              "createdDateTime": "2021-12-01T00:00:00Z",
              "modifiedDateTime": "2021-12-01T00:00:00Z",
              "displayName": "Passwordless MFA",
              "description": "Passwordless methods that satisfy strong authentication, such as Passwordless sign-in with the Microsoft Authenticator",
              "policyType": "builtIn",
              "requirementsSatisfied": "mfa",
              "allowedCombinations": [
                "windowsHelloForBusiness",
                "fido2",
                "x509CertificateMultiFactor",
                "deviceBasedPush"
              ],
              "combinationConfigurations@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationStrengthPolicies('00000000-0000-0000-0000-000000000003')/combinationConfigurations",
              "combinationConfigurations": []
            },
            {
              "id": "00000000-0000-0000-0000-000000000004",
              "createdDateTime": "2021-12-01T00:00:00Z",
              "modifiedDateTime": "2021-12-01T00:00:00Z",
              "displayName": "Phishing-resistant MFA",
              "description": "Phishing-resistant, Passwordless methods for the strongest authentication, such as a FIDO2 security key",
              "policyType": "builtIn",
              "requirementsSatisfied": "mfa",
              "allowedCombinations": [
                "windowsHelloForBusiness",
                "fido2",
                "x509CertificateMultiFactor"
              ],
              "combinationConfigurations@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationStrengthPolicies('00000000-0000-0000-0000-000000000004')/combinationConfigurations",
              "combinationConfigurations": []
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/authorizationPolicy`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authorizationPolicy/$entity",
          "@odata.id": "https://graph.microsoft.com/v2/b30f2eac-f6b4-4f87-9dcb-cdf7ae1f8923/authorizationPolicy/authorizationPolicy",
          "id": "authorizationPolicy",
          "allowInvitesFrom": "everyone",
          "allowedToSignUpEmailBasedSubscriptions": true,
          "allowedToUseSSPR": true,
          "allowEmailVerifiedUsersToJoinOrganization": true,
          "blockMsolPowerShell": null,
          "displayName": "Authorization Policy",
          "description": "Used to manage authorization related settings across the company.",
          "defaultUserRolePermissions": {
            "allowedToCreateApps": true,
            "allowedToCreateSecurityGroups": true,
            "allowedToReadOtherUsers": true,
            "permissionGrantPoliciesAssigned": [
              "ManagePermissionGrantsForSelf.microsoft-user-default-legacy"
            ]
          }
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/claimsMappingPolicies`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/claimsMappingPolicies",
          "value": []
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/conditionalAccessPolicies`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/conditionalAccessPolicies",
          "value": [
            {
              "id": "e8c26533-d895-49b4-8eed-00e9f29418ef",
              "templateId": null,
              "displayName": "Default policy",
              "createdDateTime": "2024-09-27T13:14:06.7074253Z",
              "modifiedDateTime": null,
              "state": "enabledForReportingButNotEnforced",
              "sessionControls": null,
              "conditions": {
                "userRiskLevels": [],
                "signInRiskLevels": [],
                "clientAppTypes": [
                  "all"
                ],
                "servicePrincipalRiskLevels": [],
                "insiderRiskLevels": null,
                "platforms": null,
                "locations": null,
                "devices": null,
                "clientApplications": null,
                "authenticationFlows": null,
                "applications": {
                  "includeApplications": [
                    "None"
                  ],
                  "excludeApplications": [],
                  "includeUserActions": [],
                  "includeAuthenticationContextClassReferences": [],
                  "applicationFilter": null
                },
                "users": {
                  "includeUsers": [
                    "52f26d18-d151-434f-ae14-a4a83122b2b2"
                  ],
                  "excludeUsers": [],
                  "includeGroups": [],
                  "excludeGroups": [],
                  "includeRoles": [],
                  "excludeRoles": [],
                  "includeGuestsOrExternalUsers": null,
                  "excludeGuestsOrExternalUsers": null
                }
              },
              "grantControls": {
                "operator": "OR",
                "builtInControls": [
                  "compliantDevice"
                ],
                "customAuthenticationFactors": [],
                "termsOfUse": [],
                "authenticationStrength@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/conditionalAccessPolicies('e8c26533-d895-49b4-8eed-00e9f29418ef')/grantControls/authenticationStrength/$entity",
                "authenticationStrength": null
              }
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/crossTenantAccessPolicy/$entity",
          "id": "d4699183-964f-4e62-879a-463c03360364",
          "displayName": "CrossTenantAccessPolicy for 00000000-0000-0000-0000-000000000005",
          "allowedCloudEndpoints": []
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/defaultAppManagementPolicy`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/defaultAppManagementPolicy/$entity",
          "id": "00000000-0000-0000-0000-000000000000",
          "displayName": "Default app management tenant policy",
          "description": "Default tenant policy that enforces app management restrictions on applications and service principals. To apply policy to targeted resources, create a new policy under appManagementPolicies collection.",
          "isEnabled": false,
          "applicationRestrictions": {
            "passwordCredentials": [],
            "keyCredentials": []
          },
          "servicePrincipalRestrictions": {
            "passwordCredentials": [],
            "keyCredentials": []
          }
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/deviceRegistrationPolicy`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/deviceRegistrationPolicy/$entity",
          "multiFactorAuthConfiguration": "notRequired",
          "id": "deviceRegistrationPolicy",
          "displayName": "Device Registration Policy",
          "description": "Tenant-wide policy that manages initial provisioning controls using quota restrictions, additional authentication and authorization checks",
          "userDeviceQuota": 50,
          "azureADRegistration": {
            "isAdminConfigurable": false,
            "allowedToRegister": {
              "@odata.type": "#microsoft.graph.allDeviceRegistrationMembership"
            }
          },
          "azureADJoin": {
            "isAdminConfigurable": true,
            "allowedToJoin": {
              "@odata.type": "#microsoft.graph.allDeviceRegistrationMembership"
            }
          },
          "localAdminPassword": {
            "isEnabled": false
          }
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/featureRolloutPolicies`) {
        return {
          "value": [
            {
              "id": "e3c2f23a-edd2-43a8-849f-154e70794ac5",
              "displayName": "PassthroughAuthentication rollout policy",
              "description": "PassthroughAuthentication rollout policy",
              "feature": "passthroughAuthentication",
              "isEnabled": true,
              "isAppliedToOrganization": false
            },
            {
              "id": "df85e4d9-e8c4-4033-a41c-73419a95c29c",
              "displayName": "SeamlessSso rollout policy",
              "description": "SeamlessSso rollout policy",
              "feature": "seamlessSso",
              "isEnabled": true,
              "isAppliedToOrganization": false
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/homeRealmDiscoveryPolicies`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/homeRealmDiscoveryPolicies",
          "value": []
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/identitySecurityDefaultsEnforcementPolicy/$entity",
          "id": "00000000-0000-0000-0000-000000000005",
          "displayName": "Security Defaults",
          "description": "Security defaults is a set of basic identity security mechanisms recommended by Microsoft. When enabled, these recommendations will be automatically enforced in your organization. Administrators and users will be better protected from common identity related attacks.",
          "isEnabled": false
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/permissionGrantPolicies`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/permissionGrantPolicies",
          "@microsoft.graph.tips": "Use $select to choose only the properties your app needs, as this can lead to performance improvements. For example: GET policies/permissionGrantPolicies?$select=description,displayName",
          "value": [
            {
              "id": "microsoft-all-application-permissions",
              "displayName": "All application permissions, for any client app",
              "description": "Includes all application permissions (app roles), for all APIs, for any client application.",
              "includes@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/permissionGrantPolicies('microsoft-all-application-permissions')/includes",
              "includes": [
                {
                  "id": "bddda1ec-0174-44d5-84e2-47fb0ac01595",
                  "permissionClassification": "all",
                  "permissionType": "application",
                  "resourceApplication": "any",
                  "permissions": [
                    "all"
                  ],
                  "clientApplicationIds": [
                    "all"
                  ],
                  "clientApplicationTenantIds": [
                    "all"
                  ],
                  "clientApplicationPublisherIds": [
                    "all"
                  ],
                  "clientApplicationsFromVerifiedPublisherOnly": false
                }
              ],
              "excludes@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/permissionGrantPolicies('microsoft-all-application-permissions')/excludes",
              "excludes": []
            },
            {
              "id": "microsoft-user-default-recommended",
              "displayName": "Microsoft User Default Recommended Policy",
              "description": "Permissions consentable based on Microsoft's current recommendations.",
              "includes@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/permissionGrantPolicies('microsoft-user-default-recommended')/includes",
              "includes": [
                {
                  "id": "939e5649-d754-4aa4-90df-8bfb027d11cc",
                  "permissionClassification": "all",
                  "permissionType": "delegated",
                  "resourceApplication": "00000003-0000-0000-c000-000000000000",
                  "permissions": [
                    "7427e0e9-2fba-42fe-b0c0-848c9e6a8182",
                    "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
                    "37f7f235-527c-4136-accd-4a02d197296e",
                    "64a6cdd6-aab1-4aaf-94b8-3cc8405e90d0",
                    "14dad69e-099b-42c9-810b-d002981feec1"
                  ],
                  "clientApplicationIds": [
                    "all"
                  ],
                  "clientApplicationTenantIds": [
                    "all"
                  ],
                  "clientApplicationPublisherIds": [
                    "all"
                  ],
                  "clientApplicationsFromVerifiedPublisherOnly": true
                }
              ],
              "excludes@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/permissionGrantPolicies('microsoft-user-default-recommended')/excludes",
              "excludes": []
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/roleManagementPolicies?$filter=scopeId eq '/' and scopeType eq 'DirectoryRole'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/roleManagementPolicies",
          "value": [
            {
              "id": "DirectoryRole_a457c42c-0f2e-4a25-be2a-545e840add1f_7ace6474-d11c-4a14-bc8f-3c9fdfc34930",
              "displayName": "DirectoryRole",
              "description": "DirectoryRole",
              "isOrganizationDefault": false,
              "scopeId": "/",
              "scopeType": "DirectoryRole",
              "lastModifiedDateTime": null,
              "lastModifiedBy": {
                "displayName": null,
                "id": null
              }
            },
            {
              "id": "DirectoryRole_a457c42c-0f2e-4a25-be2a-545e840add1f_c1001179-7988-4481-98b8-f641310eb7de",
              "displayName": "DirectoryRole",
              "description": "DirectoryRole",
              "isOrganizationDefault": false,
              "scopeId": "/",
              "scopeType": "DirectoryRole",
              "lastModifiedDateTime": null,
              "lastModifiedBy": {
                "displayName": null,
                "id": null
              }
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/tokenLifetimePolicies`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/tokenLifetimePolicies",
          "value": [
            {
              id: 'a457c42c-0f2e-4a25-be2a-545e840add1f',
              deletedDateTime: null,
              definition: [
                '{"TokenLifetimePolicy":{"Version":1,"AccessTokenLifetime":"8:00:00"}}'
              ],
              displayName: 'TokenLifetimePolicy1',
              isOrganizationDefault: true
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/tokenIssuancePolicies`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/tokenIssuancePolicies",
          "value": [
            {
              id: '457c8ef6-7a9c-4c9c-ba05-a12b7654c95a',
              deletedDateTime: null,
              definition: [
                '{ "TokenIssuancePolicy":{"TokenResponseSigningPolicy":"TokenOnly","SamlTokenVersion":"1.1","SigningAlgorithm":"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256","Version":1}}'
              ],
              displayName: 'TokenIssuancePolicy1',
              isOrganizationDefault: true
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/adminConsentRequestPolicy/$entity",
        "isEnabled": false,
        "notifyReviewers": false,
        "remindersEnabled": false,
        "requestDurationInDays": 0,
        "version": 0,
        "reviewers": []
      },
      {
        "id": "db9d4b58-3488-4da4-9994-49773c454e33",
        "displayName": "Custom app management policy",
        "description": "Custom policy that enforces app management restrictions on specific applications and service principals.",
        "isEnabled": false,
        "restrictions": {
          "passwordCredentials": [
            {
              "restrictionType": "passwordAddition",
              "maxLifetime": null,
              "restrictForAppsCreatedAfterDateTime": "2019-10-19T10:37:00Z"
            },
            {
              "restrictionType": "passwordLifetime",
              "maxLifetime": "P4DT12H30M5S",
              "restrictForAppsCreatedAfterDateTime": "2017-10-19T10:37:00Z"
            },
            {
              "restrictionType": "symmetricKeyAddition",
              "maxLifetime": null,
              "restrictForAppsCreatedAfterDateTime": "2021-10-19T10:37:00Z"
            },
            {
              "restrictionType": "symmetricKeyLifetime",
              "maxLifetime": "P4D",
              "restrictForAppsCreatedAfterDateTime": "2014-10-19T10:37:00Z"
            }
          ],
          "keyCredentials": [
            {
              "restrictionType": "asymmetricKeyLifetime",
              "maxLifetime": "P90D",
              "restrictForAppsCreatedAfterDateTime": "2014-10-19T10:37:00Z"
            }
          ]
        }
      },
      {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationFlowsPolicy/$entity",
        "id": "authenticationFlowsPolicy",
        "displayName": "Authentication flows policy",
        "description": "Authentication flows policy allows modification of settings related to authentication flows in AAD tenant, such as self-service sign up configuration.",
        "selfServiceSignUp": {
          "isEnabled": false
        }
      },
      {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#authenticationMethodsPolicy",
        "id": "authenticationMethodsPolicy",
        "displayName": "Authentication Methods Policy",
        "description": "The tenant-wide policy that controls which authentication methods are allowed in the tenant, authentication method registration requirements, and self-service password reset settings",
        "lastModifiedDateTime": "2025-03-03T09:38:22.5409946Z",
        "policyVersion": "1.5",
        "policyMigrationState": "migrationInProgress",
        "registrationEnforcement": {
          "authenticationMethodsRegistrationCampaign": {
            "snoozeDurationInDays": 1,
            "state": "default",
            "excludeTargets": [],
            "includeTargets": [
              {
                "id": "all_users",
                "targetType": "group",
                "targetedAuthenticationMethod": "microsoftAuthenticator"
              }
            ]
          }
        },
        "authenticationMethodConfigurations@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations",
        "authenticationMethodConfigurations": [
          {
            "@odata.type": "#microsoft.graph.fido2AuthenticationMethodConfiguration",
            "id": "Fido2",
            "state": "disabled",
            "isSelfServiceRegistrationAllowed": true,
            "isAttestationEnforced": false,
            "excludeTargets": [],
            "keyRestrictions": {
              "isEnforced": false,
              "enforcementType": "block",
              "aaGuids": []
            },
            "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('Fido2')/microsoft.graph.fido2AuthenticationMethodConfiguration/includeTargets",
            "includeTargets": [
              {
                "targetType": "group",
                "id": "all_users",
                "isRegistrationRequired": false
              }
            ]
          },
          {
            "@odata.type": "#microsoft.graph.microsoftAuthenticatorAuthenticationMethodConfiguration",
            "id": "MicrosoftAuthenticator",
            "state": "enabled",
            "isSoftwareOathEnabled": false,
            "excludeTargets": [],
            "featureSettings": {
              "displayAppInformationRequiredState": {
                "state": "default",
                "includeTarget": {
                  "targetType": "group",
                  "id": "all_users"
                },
                "excludeTarget": {
                  "targetType": "group",
                  "id": "00000000-0000-0000-0000-000000000000"
                }
              },
              "displayLocationInformationRequiredState": {
                "state": "default",
                "includeTarget": {
                  "targetType": "group",
                  "id": "all_users"
                },
                "excludeTarget": {
                  "targetType": "group",
                  "id": "00000000-0000-0000-0000-000000000000"
                }
              }
            },
            "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('MicrosoftAuthenticator')/microsoft.graph.microsoftAuthenticatorAuthenticationMethodConfiguration/includeTargets",
            "includeTargets": [
              {
                "targetType": "group",
                "id": "all_users",
                "isRegistrationRequired": false,
                "authenticationMode": "any"
              }
            ]
          },
          {
            "@odata.type": "#microsoft.graph.smsAuthenticationMethodConfiguration",
            "id": "Sms",
            "state": "disabled",
            "excludeTargets": [],
            "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('Sms')/microsoft.graph.smsAuthenticationMethodConfiguration/includeTargets",
            "includeTargets": [
              {
                "targetType": "group",
                "id": "all_users",
                "isRegistrationRequired": false,
                "isUsableForSignIn": true
              }
            ]
          },
          {
            "@odata.type": "#microsoft.graph.temporaryAccessPassAuthenticationMethodConfiguration",
            "id": "TemporaryAccessPass",
            "state": "disabled",
            "defaultLifetimeInMinutes": 60,
            "defaultLength": 8,
            "minimumLifetimeInMinutes": 60,
            "maximumLifetimeInMinutes": 480,
            "isUsableOnce": false,
            "excludeTargets": [],
            "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('TemporaryAccessPass')/microsoft.graph.temporaryAccessPassAuthenticationMethodConfiguration/includeTargets",
            "includeTargets": [
              {
                "targetType": "group",
                "id": "all_users",
                "isRegistrationRequired": false
              }
            ]
          },
          {
            "@odata.type": "#microsoft.graph.softwareOathAuthenticationMethodConfiguration",
            "id": "SoftwareOath",
            "state": "disabled",
            "excludeTargets": [],
            "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('SoftwareOath')/microsoft.graph.softwareOathAuthenticationMethodConfiguration/includeTargets",
            "includeTargets": [
              {
                "targetType": "group",
                "id": "all_users",
                "isRegistrationRequired": false
              }
            ]
          },
          {
            "@odata.type": "#microsoft.graph.voiceAuthenticationMethodConfiguration",
            "id": "Voice",
            "state": "disabled",
            "isOfficePhoneAllowed": false,
            "excludeTargets": [],
            "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('Voice')/microsoft.graph.voiceAuthenticationMethodConfiguration/includeTargets",
            "includeTargets": [
              {
                "targetType": "group",
                "id": "all_users",
                "isRegistrationRequired": false
              }
            ]
          },
          {
            "@odata.type": "#microsoft.graph.emailAuthenticationMethodConfiguration",
            "id": "Email",
            "state": "enabled",
            "allowExternalIdToUseEmailOtp": "default",
            "excludeTargets": [],
            "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('Email')/microsoft.graph.emailAuthenticationMethodConfiguration/includeTargets",
            "includeTargets": []
          },
          {
            "@odata.type": "#microsoft.graph.x509CertificateAuthenticationMethodConfiguration",
            "id": "X509Certificate",
            "state": "disabled",
            "excludeTargets": [],
            "certificateUserBindings": [
              {
                "x509CertificateField": "PrincipalName",
                "userProperty": "userPrincipalName",
                "priority": 1,
                "trustAffinityLevel": "low"
              },
              {
                "x509CertificateField": "RFC822Name",
                "userProperty": "userPrincipalName",
                "priority": 2,
                "trustAffinityLevel": "low"
              },
              {
                "x509CertificateField": "SubjectKeyIdentifier",
                "userProperty": "certificateUserIds",
                "priority": 3,
                "trustAffinityLevel": "high"
              }
            ],
            "authenticationModeConfiguration": {
              "x509CertificateAuthenticationDefaultMode": "x509CertificateSingleFactor",
              "x509CertificateDefaultRequiredAffinityLevel": "low",
              "rules": []
            },
            "crlValidationConfiguration": {
              "state": "disabled",
              "exemptedCertificateAuthoritiesSubjectKeyIdentifiers": []
            },
            "includeTargets@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationMethodsPolicy/authenticationMethodConfigurations('X509Certificate')/microsoft.graph.x509CertificateAuthenticationMethodConfiguration/includeTargets",
            "includeTargets": [
              {
                "targetType": "group",
                "id": "all_users",
                "isRegistrationRequired": false
              }
            ]
          }
        ]
      },
      {
        "id": "00000000-0000-0000-0000-000000000002",
        "createdDateTime": "2021-12-01T00:00:00Z",
        "modifiedDateTime": "2021-12-01T00:00:00Z",
        "displayName": "Multifactor authentication",
        "description": "Combinations of methods that satisfy strong authentication, such as a password + SMS",
        "policyType": "builtIn",
        "requirementsSatisfied": "mfa",
        "allowedCombinations": [
          "windowsHelloForBusiness",
          "fido2",
          "x509CertificateMultiFactor",
          "deviceBasedPush",
          "temporaryAccessPassOneTime",
          "temporaryAccessPassMultiUse",
          "password,microsoftAuthenticatorPush",
          "password,softwareOath",
          "password,hardwareOath",
          "password,sms",
          "password,voice",
          "federatedMultiFactor",
          "microsoftAuthenticatorPush,federatedSingleFactor",
          "softwareOath,federatedSingleFactor",
          "hardwareOath,federatedSingleFactor",
          "sms,federatedSingleFactor",
          "voice,federatedSingleFactor"
        ],
        "combinationConfigurations@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationStrengthPolicies('00000000-0000-0000-0000-000000000002')/combinationConfigurations",
        "combinationConfigurations": []
      },
      {
        "id": "00000000-0000-0000-0000-000000000003",
        "createdDateTime": "2021-12-01T00:00:00Z",
        "modifiedDateTime": "2021-12-01T00:00:00Z",
        "displayName": "Passwordless MFA",
        "description": "Passwordless methods that satisfy strong authentication, such as Passwordless sign-in with the Microsoft Authenticator",
        "policyType": "builtIn",
        "requirementsSatisfied": "mfa",
        "allowedCombinations": [
          "windowsHelloForBusiness",
          "fido2",
          "x509CertificateMultiFactor",
          "deviceBasedPush"
        ],
        "combinationConfigurations@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationStrengthPolicies('00000000-0000-0000-0000-000000000003')/combinationConfigurations",
        "combinationConfigurations": []
      },
      {
        "id": "00000000-0000-0000-0000-000000000004",
        "createdDateTime": "2021-12-01T00:00:00Z",
        "modifiedDateTime": "2021-12-01T00:00:00Z",
        "displayName": "Phishing-resistant MFA",
        "description": "Phishing-resistant, Passwordless methods for the strongest authentication, such as a FIDO2 security key",
        "policyType": "builtIn",
        "requirementsSatisfied": "mfa",
        "allowedCombinations": [
          "windowsHelloForBusiness",
          "fido2",
          "x509CertificateMultiFactor"
        ],
        "combinationConfigurations@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authenticationStrengthPolicies('00000000-0000-0000-0000-000000000004')/combinationConfigurations",
        "combinationConfigurations": []
      },
      {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authorizationPolicy/$entity",
        "@odata.id": "https://graph.microsoft.com/v2/b30f2eac-f6b4-4f87-9dcb-cdf7ae1f8923/authorizationPolicy/authorizationPolicy",
        "id": "authorizationPolicy",
        "allowInvitesFrom": "everyone",
        "allowedToSignUpEmailBasedSubscriptions": true,
        "allowedToUseSSPR": true,
        "allowEmailVerifiedUsersToJoinOrganization": true,
        "blockMsolPowerShell": null,
        "displayName": "Authorization Policy",
        "description": "Used to manage authorization related settings across the company.",
        "defaultUserRolePermissions": {
          "allowedToCreateApps": true,
          "allowedToCreateSecurityGroups": true,
          "allowedToReadOtherUsers": true,
          "permissionGrantPoliciesAssigned": [
            "ManagePermissionGrantsForSelf.microsoft-user-default-legacy"
          ]
        }
      },
      {
        "id": "e8c26533-d895-49b4-8eed-00e9f29418ef",
        "templateId": null,
        "displayName": "Default policy",
        "createdDateTime": "2024-09-27T13:14:06.7074253Z",
        "modifiedDateTime": null,
        "state": "enabledForReportingButNotEnforced",
        "sessionControls": null,
        "conditions": {
          "userRiskLevels": [],
          "signInRiskLevels": [],
          "clientAppTypes": [
            "all"
          ],
          "servicePrincipalRiskLevels": [],
          "insiderRiskLevels": null,
          "platforms": null,
          "locations": null,
          "devices": null,
          "clientApplications": null,
          "authenticationFlows": null,
          "applications": {
            "includeApplications": [
              "None"
            ],
            "excludeApplications": [],
            "includeUserActions": [],
            "includeAuthenticationContextClassReferences": [],
            "applicationFilter": null
          },
          "users": {
            "includeUsers": [
              "52f26d18-d151-434f-ae14-a4a83122b2b2"
            ],
            "excludeUsers": [],
            "includeGroups": [],
            "excludeGroups": [],
            "includeRoles": [],
            "excludeRoles": [],
            "includeGuestsOrExternalUsers": null,
            "excludeGuestsOrExternalUsers": null
          }
        },
        "grantControls": {
          "operator": "OR",
          "builtInControls": [
            "compliantDevice"
          ],
          "customAuthenticationFactors": [],
          "termsOfUse": [],
          "authenticationStrength@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/conditionalAccessPolicies('e8c26533-d895-49b4-8eed-00e9f29418ef')/grantControls/authenticationStrength/$entity",
          "authenticationStrength": null
        }
      },
      {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/crossTenantAccessPolicy/$entity",
        "id": "d4699183-964f-4e62-879a-463c03360364",
        "displayName": "CrossTenantAccessPolicy for 00000000-0000-0000-0000-000000000005",
        "allowedCloudEndpoints": []
      },
      {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/defaultAppManagementPolicy/$entity",
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Default app management tenant policy",
        "description": "Default tenant policy that enforces app management restrictions on applications and service principals. To apply policy to targeted resources, create a new policy under appManagementPolicies collection.",
        "isEnabled": false,
        "applicationRestrictions": {
          "passwordCredentials": [],
          "keyCredentials": []
        },
        "servicePrincipalRestrictions": {
          "passwordCredentials": [],
          "keyCredentials": []
        }
      },
      {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/deviceRegistrationPolicy/$entity",
        "multiFactorAuthConfiguration": "notRequired",
        "id": "deviceRegistrationPolicy",
        "displayName": "Device Registration Policy",
        "description": "Tenant-wide policy that manages initial provisioning controls using quota restrictions, additional authentication and authorization checks",
        "userDeviceQuota": 50,
        "azureADRegistration": {
          "isAdminConfigurable": false,
          "allowedToRegister": {
            "@odata.type": "#microsoft.graph.allDeviceRegistrationMembership"
          }
        },
        "azureADJoin": {
          "isAdminConfigurable": true,
          "allowedToJoin": {
            "@odata.type": "#microsoft.graph.allDeviceRegistrationMembership"
          }
        },
        "localAdminPassword": {
          "isEnabled": false
        }
      },
      {
        "id": "e3c2f23a-edd2-43a8-849f-154e70794ac5",
        "displayName": "PassthroughAuthentication rollout policy",
        "description": "PassthroughAuthentication rollout policy",
        "feature": "passthroughAuthentication",
        "isEnabled": true,
        "isAppliedToOrganization": false
      },
      {
        "id": "df85e4d9-e8c4-4033-a41c-73419a95c29c",
        "displayName": "SeamlessSso rollout policy",
        "description": "SeamlessSso rollout policy",
        "feature": "seamlessSso",
        "isEnabled": true,
        "isAppliedToOrganization": false
      },
      {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/identitySecurityDefaultsEnforcementPolicy/$entity",
        "id": "00000000-0000-0000-0000-000000000005",
        "displayName": "Security Defaults",
        "description": "Security defaults is a set of basic identity security mechanisms recommended by Microsoft. When enabled, these recommendations will be automatically enforced in your organization. Administrators and users will be better protected from common identity related attacks.",
        "isEnabled": false
      },
      {
        "id": "microsoft-all-application-permissions",
        "displayName": "All application permissions, for any client app",
        "description": "Includes all application permissions (app roles), for all APIs, for any client application.",
        "includes@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/permissionGrantPolicies('microsoft-all-application-permissions')/includes",
        "includes": [
          {
            "id": "bddda1ec-0174-44d5-84e2-47fb0ac01595",
            "permissionClassification": "all",
            "permissionType": "application",
            "resourceApplication": "any",
            "permissions": [
              "all"
            ],
            "clientApplicationIds": [
              "all"
            ],
            "clientApplicationTenantIds": [
              "all"
            ],
            "clientApplicationPublisherIds": [
              "all"
            ],
            "clientApplicationsFromVerifiedPublisherOnly": false
          }
        ],
        "excludes@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/permissionGrantPolicies('microsoft-all-application-permissions')/excludes",
        "excludes": []
      },
      {
        "id": "microsoft-user-default-recommended",
        "displayName": "Microsoft User Default Recommended Policy",
        "description": "Permissions consentable based on Microsoft's current recommendations.",
        "includes@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/permissionGrantPolicies('microsoft-user-default-recommended')/includes",
        "includes": [
          {
            "id": "939e5649-d754-4aa4-90df-8bfb027d11cc",
            "permissionClassification": "all",
            "permissionType": "delegated",
            "resourceApplication": "00000003-0000-0000-c000-000000000000",
            "permissions": [
              "7427e0e9-2fba-42fe-b0c0-848c9e6a8182",
              "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
              "37f7f235-527c-4136-accd-4a02d197296e",
              "64a6cdd6-aab1-4aaf-94b8-3cc8405e90d0",
              "14dad69e-099b-42c9-810b-d002981feec1"
            ],
            "clientApplicationIds": [
              "all"
            ],
            "clientApplicationTenantIds": [
              "all"
            ],
            "clientApplicationPublisherIds": [
              "all"
            ],
            "clientApplicationsFromVerifiedPublisherOnly": true
          }
        ],
        "excludes@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/permissionGrantPolicies('microsoft-user-default-recommended')/excludes",
        "excludes": []
      },
      {
        "id": "DirectoryRole_a457c42c-0f2e-4a25-be2a-545e840add1f_7ace6474-d11c-4a14-bc8f-3c9fdfc34930",
        "displayName": "DirectoryRole",
        "description": "DirectoryRole",
        "isOrganizationDefault": false,
        "scopeId": "/",
        "scopeType": "DirectoryRole",
        "lastModifiedDateTime": null,
        "lastModifiedBy": {
          "displayName": null,
          "id": null
        }
      },
      {
        "id": "DirectoryRole_a457c42c-0f2e-4a25-be2a-545e840add1f_c1001179-7988-4481-98b8-f641310eb7de",
        "displayName": "DirectoryRole",
        "description": "DirectoryRole",
        "isOrganizationDefault": false,
        "scopeId": "/",
        "scopeType": "DirectoryRole",
        "lastModifiedDateTime": null,
        "lastModifiedBy": {
          "displayName": null,
          "id": null
        }
      },
      {
        id: '457c8ef6-7a9c-4c9c-ba05-a12b7654c95a',
        deletedDateTime: null,
        definition: [
          '{ "TokenIssuancePolicy":{"TokenResponseSigningPolicy":"TokenOnly","SamlTokenVersion":"1.1","SigningAlgorithm":"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256","Version":1}}'
        ],
        displayName: 'TokenIssuancePolicy1',
        isOrganizationDefault: true
      },
      {
        id: 'a457c42c-0f2e-4a25-be2a-545e840add1f',
        deletedDateTime: null,
        definition: [
          '{"TokenLifetimePolicy":{"Version":1,"AccessTokenLifetime":"8:00:00"}}'
        ],
        displayName: 'TokenLifetimePolicy1',
        isOrganizationDefault: true
      }
    ]));
  });

  it('retrieves the role management policies', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/policies/roleManagementPolicies?$filter=scopeId eq '/' and scopeType eq 'DirectoryRole'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/roleManagementPolicies",
          "value": [
            {
              "id": "DirectoryRole_a457c42c-0f2e-4a25-be2a-545e840add1f_7ace6474-d11c-4a14-bc8f-3c9fdfc34930",
              "displayName": "DirectoryRole",
              "description": "DirectoryRole",
              "isOrganizationDefault": false,
              "scopeId": "/",
              "scopeType": "DirectoryRole",
              "lastModifiedDateTime": null,
              "lastModifiedBy": {
                "displayName": null,
                "id": null
              }
            },
            {
              "id": "DirectoryRole_a457c42c-0f2e-4a25-be2a-545e840add1f_c1001179-7988-4481-98b8-f641310eb7de",
              "displayName": "DirectoryRole",
              "description": "DirectoryRole",
              "isOrganizationDefault": false,
              "scopeId": "/",
              "scopeType": "DirectoryRole",
              "lastModifiedDateTime": null,
              "lastModifiedBy": {
                "displayName": null,
                "id": null
              }
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        type: "roleManagement"
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "id": "DirectoryRole_a457c42c-0f2e-4a25-be2a-545e840add1f_7ace6474-d11c-4a14-bc8f-3c9fdfc34930",
        "displayName": "DirectoryRole",
        "description": "DirectoryRole",
        "isOrganizationDefault": false,
        "scopeId": "/",
        "scopeType": "DirectoryRole",
        "lastModifiedDateTime": null,
        "lastModifiedBy": {
          "displayName": null,
          "id": null
        }
      },
      {
        "id": "DirectoryRole_a457c42c-0f2e-4a25-be2a-545e840add1f_c1001179-7988-4481-98b8-f641310eb7de",
        "displayName": "DirectoryRole",
        "description": "DirectoryRole",
        "isOrganizationDefault": false,
        "scopeId": "/",
        "scopeType": "DirectoryRole",
        "lastModifiedDateTime": null,
        "lastModifiedBy": {
          "displayName": null,
          "id": null
        }
      }
    ]));
  });

  it('correctly handles API OData error for specified policies', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred.'));

    await assert.rejects(command.action(logger, { options: { type: "foo" } } as any), new CommandError("An error has occurred."));
  });

  it('correctly handles API OData error for all policies', async () => {
    sinon.stub(request, 'get').rejects(new Error("An error has occurred."));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError("An error has occurred."));
  });

  it('accepts type to be activityBasedTimeout', async () => {
    const actual = await command.validate({
      options:
      {
        type: "activityBasedTimeout"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be authorization', async () => {
    const actual = await command.validate({
      options:
      {
        type: "authorization"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be claimsMapping', async () => {
    const actual = await command.validate({
      options:
      {
        type: "claimsMapping"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be homeRealmDiscovery', async () => {
    const actual = await command.validate({
      options:
      {
        type: "homeRealmDiscovery"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be identitySecurityDefaultsEnforcement', async () => {
    const actual = await command.validate({
      options:
      {
        type: "identitySecurityDefaultsEnforcement"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be tokenLifetime', async () => {
    const actual = await command.validate({
      options:
      {
        type: "tokenLifetime"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be tokenIssuance', async () => {
    const actual = await command.validate({
      options:
      {
        type: "tokenIssuance"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be adminConsentRequest', async () => {
    const actual = await command.validate({
      options:
      {
        type: "adminConsentRequest"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be appManagement', async () => {
    const actual = await command.validate({
      options:
      {
        type: "appManagement"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be authenticationFlows', async () => {
    const actual = await command.validate({
      options:
      {
        type: "authenticationFlows"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be authenticationMethods', async () => {
    const actual = await command.validate({
      options:
      {
        type: "authenticationMethods"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be authenticationStrength', async () => {
    const actual = await command.validate({
      options:
      {
        type: "authenticationStrength"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be conditionalAccess', async () => {
    const actual = await command.validate({
      options:
      {
        type: "conditionalAccess"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be crossTenantAccess', async () => {
    const actual = await command.validate({
      options:
      {
        type: "crossTenantAccess"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be defaultAppManagement', async () => {
    const actual = await command.validate({
      options:
      {
        type: "defaultAppManagement"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be deviceRegistration', async () => {
    const actual = await command.validate({
      options:
      {
        type: "deviceRegistration"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be permissionGrant', async () => {
    const actual = await command.validate({
      options:
      {
        type: "permissionGrant"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts type to be roleManagement', async () => {
    const actual = await command.validate({
      options:
      {
        type: "roleManagement"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('rejects invalid type', async () => {
    const type = 'foo';
    const actual = await command.validate({
      options: {
        type: type
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
