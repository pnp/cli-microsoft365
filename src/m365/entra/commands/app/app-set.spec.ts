import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-set.js';
import { settingsNames } from '../../../../settingsNames.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';

describe(commands.APP_SET, () => {

  //#region Mocked Responses  
  const appDetailsResponse: any = { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity", "id": "95cfe30d-ed44-4f9d-b73d-c66560f72e83", "deletedDateTime": null, "appId": "ff254847-12c7-44cf-921e-8883dbd622a7", "applicationTemplateId": null, "disabledByMicrosoftStatus": null, "createdDateTime": "2022-02-07T08:51:18Z", "displayName": "Angular Teams app", "description": null, "groupMembershipClaims": null, "identifierUris": ["api://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"], "isDeviceOnlyAuthSupported": null, "isFallbackPublicClient": null, "notes": null, "publisherDomain": "contoso.onmicrosoft.com", "serviceManagementReference": null, "signInAudience": "AzureADMyOrg", "tags": [], "tokenEncryptionKeyId": null, "defaultRedirectUri": null, "certification": null, "optionalClaims": null, "addIns": [], "api": { "acceptMappedClaims": null, "knownClientApplications": [], "requestedAccessTokenVersion": null, "oauth2PermissionScopes": [{ "adminConsentDescription": "Access as a user", "adminConsentDisplayName": "Access as a user", "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5", "isEnabled": true, "type": "User", "userConsentDescription": null, "userConsentDisplayName": null, "value": "access_as_user" }], "preAuthorizedApplications": [{ "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346", "delegatedPermissionIds": ["cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"] }, { "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264", "delegatedPermissionIds": ["cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"] }] }, "appRoles": [], "info": { "logoUrl": null, "marketingUrl": null, "privacyStatementUrl": null, "supportUrl": null, "termsOfServiceUrl": null }, "keyCredentials": [], "parentalControlSettings": { "countriesBlockedForMinors": [], "legalAgeGroupRule": "Allow" }, "passwordCredentials": [], "publicClient": { "redirectUris": [] }, "requiredResourceAccess": [{ "resourceAppId": "00000003-0000-0000-c000-000000000000", "resourceAccess": [{ "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d", "type": "Scope" }] }], "verifiedPublisher": { "displayName": null, "verifiedPublisherId": null, "addedDateTime": null }, "web": { "homePageUrl": null, "logoutUrl": null, "redirectUris": [], "implicitGrantSettings": { "enableAccessTokenIssuance": false, "enableIdTokenIssuance": false } }, "spa": { "redirectUris": [] } };

  const appResponse = {
    value: [
      {
        "id": "5b31c38c-2584-42f0-aa47-657fb3a84230"
      }
    ]
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      fs.existsSync,
      fs.readFileSync,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound,
      entraApp.getAppRegistrationByAppId,
      entraApp.getAppRegistrationByAppName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('allows unknown options', () => {
    assert.strictEqual(command.allowUnknownOptions(), true);
  });

  it('updates uris for the specified appId', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8') {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      })
    });
  });

  it('updates unknown options for the specified appId', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        Object.keys(opts.data).indexOf('extension_b7d8e648520f41d3b9c0fdeb91768a0a_jobGroupTracker') > -1 &&
        opts.data.extension_b7d8e648520f41d3b9c0fdeb91768a0a_jobGroupTracker === 'JobGroupN') {
        return;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        Object.keys(opts.data).indexOf('identifierUris') > -1 &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8') {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8',
        extension_b7d8e648520f41d3b9c0fdeb91768a0a_jobGroupTracker: 'JobGroupN'
      })
    });
  });

  it('updates multiple URIs for the specified appId', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' &&
        opts.data.identifierUris[1] === 'api://testapi') {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8,api://testapi'
      })
    });
  });

  it('updates uris for the specified objectId', async () => {
    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8') {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      })
    });
  });

  it('updates multiple URIs for the specified objectId', async () => {
    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' &&
        opts.data.identifierUris[1] === 'api://testapi') {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8,api://testapi'
      })
    });
  });

  it('updates uris for the specified name', async () => {
    sinon.stub(entraApp, "getAppRegistrationByAppName").resolves(appResponse.value[0]);
    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8') {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        name: 'My app',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      })
    });
  });

  it('updates multiple URIs for the specified name', async () => {
    sinon.stub(entraApp, "getAppRegistrationByAppName").resolves(appResponse.value[0]);
    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' &&
        opts.data.identifierUris[1] === 'api://testapi') {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        name: 'My app',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8,api://testapi'
      })
    });
  });

  it('skips updating uris if no uris specified', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230'
      })
    });
  });

  it('sets spa redirectUri for an app without redirectUris', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/e4528262-097a-42eb-98e1-19f073dbee45`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "id": "e4528262-097a-42eb-98e1-19f073dbee45",
          "deletedDateTime": null,
          "appId": "842e1a6f-7492-4b7d-8278-563036f5bd39",
          "applicationTemplateId": null,
          "disabledByMicrosoftStatus": null,
          "createdDateTime": "2022-02-10T08:01:06Z",
          "displayName": "Angular Teams app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [
            "api://24c4-2001-1c00-80c-d00-e5da-977c-7c52-5196.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"
          ],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "serviceManagementReference": null,
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "defaultRedirectUri": null,
          "certification": null,
          "optionalClaims": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [
              {
                "adminConsentDescription": "Access as a user",
                "adminConsentDisplayName": "Access as a user",
                "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5",
                "isEnabled": true,
                "type": "User",
                "userConsentDescription": null,
                "userConsentDisplayName": null,
                "value": "access_as_user"
              }
            ],
            "preAuthorizedApplications": [
              {
                "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              },
              {
                "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              }
            ]
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
                  "type": "Scope"
                }
              ]
            }
          ],
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          },
          "spa": {
            "redirectUris": []
          }
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/e4528262-097a-42eb-98e1-19f073dbee45' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "publicClient": {
            "redirectUris": []
          },
          "spa": {
            "redirectUris": [
              "https://24c4-2001-1c00-80c-d00-e5da-977c-7c52-5194.ngrok.io/auth"
            ]
          },
          "web": {
            "redirectUris": []
          }
        })) {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        objectId: 'e4528262-097a-42eb-98e1-19f073dbee45',
        redirectUris: 'https://24c4-2001-1c00-80c-d00-e5da-977c-7c52-5194.ngrok.io/auth',
        platform: 'spa'
      })
    });
  });

  it('sets web redirectUri for an app with existing spa redirectUris', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "id": "95cfe30d-ed44-4f9d-b73d-c66560f72e83",
          "deletedDateTime": null,
          "appId": "ff254847-12c7-44cf-921e-8883dbd622a7",
          "applicationTemplateId": null,
          "disabledByMicrosoftStatus": null,
          "createdDateTime": "2022-02-07T08:51:18Z",
          "displayName": "Angular Teams app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [
            "api://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"
          ],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "serviceManagementReference": null,
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "defaultRedirectUri": null,
          "certification": null,
          "optionalClaims": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [
              {
                "adminConsentDescription": "Access as a user",
                "adminConsentDisplayName": "Access as a user",
                "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5",
                "isEnabled": true,
                "type": "User",
                "userConsentDescription": null,
                "userConsentDisplayName": null,
                "value": "access_as_user"
              }
            ],
            "preAuthorizedApplications": [
              {
                "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              },
              {
                "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              }
            ]
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
                  "type": "Scope"
                }
              ]
            }
          ],
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          },
          "spa": {
            "redirectUris": [
              "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
              "http://localhost/auth"
            ]
          }
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "publicClient": {
            "redirectUris": []
          },
          "spa": {
            "redirectUris": [
              "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
              "http://localhost/auth"
            ]
          },
          "web": {
            "redirectUris": [
              "https://foo.com"
            ]
          }
        })) {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        redirectUris: 'https://foo.com',
        platform: 'web'
      })
    });
  });

  it('sets publicClient redirectUri for an app with existing spa and web redirectUris', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "id": "95cfe30d-ed44-4f9d-b73d-c66560f72e83",
          "deletedDateTime": null,
          "appId": "ff254847-12c7-44cf-921e-8883dbd622a7",
          "applicationTemplateId": null,
          "disabledByMicrosoftStatus": null,
          "createdDateTime": "2022-02-07T08:51:18Z",
          "displayName": "Angular Teams app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [
            "api://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"
          ],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "serviceManagementReference": null,
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "defaultRedirectUri": null,
          "certification": null,
          "optionalClaims": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [
              {
                "adminConsentDescription": "Access as a user",
                "adminConsentDisplayName": "Access as a user",
                "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5",
                "isEnabled": true,
                "type": "User",
                "userConsentDescription": null,
                "userConsentDisplayName": null,
                "value": "access_as_user"
              }
            ],
            "preAuthorizedApplications": [
              {
                "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              },
              {
                "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              }
            ]
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
                  "type": "Scope"
                }
              ]
            }
          ],
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [
              "https://foo.com"
            ],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          },
          "spa": {
            "redirectUris": [
              "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
              "http://localhost/auth"
            ]
          }
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "publicClient": {
            "redirectUris": [
              "https://foo1.com"
            ]
          },
          "spa": {
            "redirectUris": [
              "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
              "http://localhost/auth"
            ]
          },
          "web": {
            "redirectUris": [
              "https://foo.com"
            ]
          }
        })) {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        redirectUris: 'https://foo1.com',
        platform: 'publicClient'
      })
    });
  });

  it('replaces existing redirectUri with a new one', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "id": "95cfe30d-ed44-4f9d-b73d-c66560f72e83",
          "deletedDateTime": null,
          "appId": "ff254847-12c7-44cf-921e-8883dbd622a7",
          "applicationTemplateId": null,
          "disabledByMicrosoftStatus": null,
          "createdDateTime": "2022-02-07T08:51:18Z",
          "displayName": "Angular Teams app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [
            "api://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"
          ],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "serviceManagementReference": null,
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "defaultRedirectUri": null,
          "certification": null,
          "optionalClaims": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [
              {
                "adminConsentDescription": "Access as a user",
                "adminConsentDisplayName": "Access as a user",
                "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5",
                "isEnabled": true,
                "type": "User",
                "userConsentDescription": null,
                "userConsentDisplayName": null,
                "value": "access_as_user"
              }
            ],
            "preAuthorizedApplications": [
              {
                "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              },
              {
                "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              }
            ]
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
                  "type": "Scope"
                }
              ]
            }
          ],
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          },
          "spa": {
            "redirectUris": [
              "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
              "http://localhost/auth"
            ]
          }
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "publicClient": {
            "redirectUris": []
          },
          "spa": {
            "redirectUris": [
              "http://localhost/auth",
              "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5194.ngrok.io/auth"
            ]
          },
          "web": {
            "redirectUris": []
          }
        })) {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        redirectUris: 'https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5194.ngrok.io/auth',
        platform: 'spa',
        redirectUrisToRemove: 'https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth'
      })
    });
  });

  it('adds new certificate using base64 string', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return appDetailsResponse;
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83?$select=keyCredentials`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "keyCredentials": []
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "keyCredentials": [{
            "type": "AsymmetricX509Cert",
            "usage": "Verify",
            "displayName": "some certificate",
            "key": "somecertificatebase64string"
          }]
        })) {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        certificateDisplayName: 'some certificate',
        certificateBase64Encoded: 'somecertificatebase64string'
      })
    });
  });

  it('adds new certificate using base64 string (with null keyCredentials)', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return appDetailsResponse;
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83?$select=keyCredentials`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "keyCredentials": null
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "keyCredentials": [{
            "type": "AsymmetricX509Cert",
            "usage": "Verify",
            "displayName": "some certificate",
            "key": "somecertificatebase64string"
          }]
        })) {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        certificateDisplayName: 'some certificate',
        certificateBase64Encoded: 'somecertificatebase64string'
      })
    });
  });

  it('adds new certificate using certificate file', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return appDetailsResponse;
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83?$select=keyCredentials`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "keyCredentials": []
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "keyCredentials": [{
            "type": "AsymmetricX509Cert",
            "usage": "Verify",
            "displayName": "some certificate",
            "key": "somecertificatebase64string"
          }]
        })) {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns("somecertificatebase64string");

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        certificateDisplayName: 'some certificate',
        certificateFile: 'C:\\temp\\some-certificate.cer'
      })
    });
  });

  it('updates allowPublicClientFlows value for the specified appId', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.isFallbackPublicClient === true) {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
        allowPublicClientFlows: true
      })
    });
  });


  it('handles error when certificate file cannot be read', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return appDetailsResponse;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').throws(new Error("An error has occurred"));

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        certificateDisplayName: 'some certificate',
        certificateFile: 'C:\\temp\\some-certificate.cer'
      })
    }), new CommandError(`Error reading certificate file: Error: An error has occurred. Please add the certificate using base64 option '--certificateBase64Encoded'.`));
  });

  it('handles error when the app specified with objectId not found', async () => {
    sinon.stub(request, 'patch').rejects(new Error(`Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`));

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      })
    }), new CommandError(`Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('handles error when the app specified with the appId not found', async () => {
    const error = `App with appId '9b1b1e42-794b-4c71-93ac-5ed92488b67f' not found in Microsoft Entra ID`;
    sinon.stub(entraApp, 'getAppRegistrationByAppId').rejects(new Error(error));

    sinon.stub(request, 'patch').rejects('PATCH request executed');

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      })
    }), new CommandError(`App with appId '9b1b1e42-794b-4c71-93ac-5ed92488b67f' not found in Microsoft Entra ID`));
  });

  it('handles error when the app specified with name not found', async () => {
    const error = `App with name 'My app' not found in Microsoft Entra ID`;
    sinon.stub(entraApp, 'getAppRegistrationByAppName').rejects(new Error(error));
    sinon.stub(request, 'patch').rejects('PATCH request executed');

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        name: 'My app',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      })
    }), new CommandError(error));
  });

  it('handles error when multiple apps with the specified name found', async () => {
    const error = `Multiple apps with name 'My app' found in Microsoft Entra ID. Found: 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g.`;
    sinon.stub(entraApp, 'getAppRegistrationByAppName').rejects(new Error(error));
    sinon.stub(request, 'patch').rejects('PATCH request executed');

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        name: 'My app',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      })
    }), new CommandError(error));
  });

  it('handles error when retrieving information about app through name failed', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        name: 'My app',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      })
    }), new CommandError(`An error has occurred`));
  });

  it('fails validation if appId and objectId specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = commandOptionsSchema.safeParse({ appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if appId and name specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = commandOptionsSchema.safeParse({ appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if objectId and name specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = commandOptionsSchema.safeParse({ objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither appId, objectId nor name specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if redirectUris specified without platform', async () => {
    const actual = commandOptionsSchema.safeParse({ objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8', redirectUris: 'https://foo.com' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if invalid platform specified', async () => {
    const actual = commandOptionsSchema.safeParse({ objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8', redirectUris: 'https://foo.com', platform: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if certificateDisplayName is specified without certificate', async () => {
    const actual = commandOptionsSchema.safeParse({ objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8', certificateDisplayName: 'Some certificate' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both certificateBase64Encoded and certificateFile are specified', async () => {
    const actual = commandOptionsSchema.safeParse({ objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8', certificateFile: 'c:\\temp\\some-certificate.cer', certificateBase64Encoded: 'somebase64string' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if certificateFile specified with certificateDisplayName', async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);

    const actual = commandOptionsSchema.safeParse({ name: 'My Microsoft Entra app', certificateDisplayName: 'Some certificate', certificateFile: 'c:\\temp\\some-certificate.cer' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when certificate file is not found', async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => false);

    const actual = commandOptionsSchema.safeParse({ debug: true, objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83', certificateDisplayName: 'some certificate', certificateFile: 'C:\\temp\\some-certificate.cer' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if required options specified (appId)', async () => {
    const actual = commandOptionsSchema.safeParse({ appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if required options specified (objectId)', async () => {
    const actual = commandOptionsSchema.safeParse({ objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if required options specified (name)', async () => {
    const actual = commandOptionsSchema.safeParse({ name: 'My app', uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when redirectUris specified with spa', async () => {
    const actual = commandOptionsSchema.safeParse({ appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', redirectUris: 'https://foo.com', platform: 'spa' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when redirectUris specified with publicClient', async () => {
    const actual = commandOptionsSchema.safeParse({ appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', redirectUris: 'https://foo.com', platform: 'publicClient' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when redirectUris specified with web', async () => {
    const actual = commandOptionsSchema.safeParse({ appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', redirectUris: 'https://foo.com', platform: 'web' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when allowPublicClientFlows is specified as true', async () => {
    const actual = commandOptionsSchema.safeParse({ appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', allowPublicClientFlows: true });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when allowPublicClientFlows is specified as false', async () => {
    const actual = commandOptionsSchema.safeParse({ appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', allowPublicClientFlows: false });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when allowPublicClientFlows is not correct boolean value', async () => {
    const actual = commandOptionsSchema.safeParse({ appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', allowPublicClientFlows: 'foo' });
    assert.strictEqual(actual.success, false);
  });
});