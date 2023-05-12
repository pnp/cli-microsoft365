import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import Command, { CommandError } from '../../../../Command';
import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import request from '../../../../request';
import { telemetry } from '../../../../telemetry';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./serviceprincipal-permissionrequest-list');

describe(commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  const oAuth2PermissionGrantsResponse = {
    value: [
      {
        clientId: '0b79cadb-a5ea-4678-be6a-ff308846158f',
        consentType: 'AllPrincipals',
        expiryTime: '9999-12-31T23:59:59.9999999Z',
        id: 'jsCikIbnAEGeqTbCYb5sDZXCr9YICndHoJUQvLfiOQM',
        principalId: null,
        resourceId: '4e9edf33-64c9-45af-912c-c2cbd711f2df',
        scope: 'Calendars.Read User.Read',
        startTime: '0001-01-01T00:00:00Z'
      }
    ]
  };

  const spoClientExtensibilityWebApplicationPrincipalResponse = {
    value: [
      {
        id: '0b79cadb-a5ea-4678-be6a-ff308846158f',
        deletedDateTime: null,
        accountEnabled: true,
        alternativeNames: [],
        appDisplayName: 'SharePoint Online Client Extensibility Web Application Principal',
        appDescription: null,
        appId: '912a70e1-da16-417d-a789-57c122d180cd',
        applicationTemplateId: null,
        appOwnerOrganizationId: 'dfd9835f-2e19-40b1-a349-5a0dab067350',
        appRoleAssignmentRequired: false,
        createdDateTime: '2022-09-20T13:02:57Z',
        description: null,
        disabledByMicrosoftStatus: null,
        displayName: 'SharePoint Online Client Extensibility Web Application Principal',
        homepage: null,
        loginUrl: null,
        logoutUrl: null,
        notes: null,
        notificationEmailAddresses: [],
        preferredSingleSignOnMode: null,
        preferredTokenSigningKeyThumbprint: null,
        replyUrls: [
          "https://fluidpreview.office.net/spfxsinglesignon",
          "https://dev.fluidpreview.office.net/spfxsinglesignon",
          "https://contoso-admin.sharepoint.com/_forms/spfxsinglesignon.aspx",
          "https://contoso.sharepoint.com/",
          "https://contoso.sharepoint.com/_forms/spfxsinglesignon.aspx",
          "https://contoso.sharepoint.com/_forms/spfxsinglesignon.aspx?redirect"
        ],
        servicePrincipalNames: [
          "api://943d747b-2cc0-4258-ab4d-cb02c9737532/contoso.sharepoint.com",
          "api://943d747b-2cc0-4258-ab4d-cb02c9737532/microsoft.spfx3rdparty.com",
          "0b79cadb-a5ea-4678-be6a-ff308846158f"
        ],
        servicePrincipalType: 'Application',
        signInAudience: 'AzureADMyOrg',
        tags: [],
        tokenEncryptionKeyId: null,
        samlSingleSignOnSettings: null,
        addIns: [],
        appRoles: [],
        info: {
          logoUrl: null,
          marketingUrl: null,
          privacyStatementUrl: null,
          supportUrl: null,
          termsOfServiceUrl: null
        },
        keyCredentials: [],
        oauth2PermissionScopes: [
          {
            adminConsentDescription: "Allow the application to access SharePoint Online Client Extensibility Web Application Principal on behalf of the signed-in user.",
            adminConsentDisplayName: "Access SharePoint Online Client Extensibility Web Application Principal",
            id: "0b79cadb-a5ea-4678-be6a-ff308846158f",
            isEnabled: true,
            type: "User",
            userConsentDescription: "Allow the application to access SharePoint Online Client Extensibility Web Application Principal on your behalf.",
            userConsentDisplayName: "Access SharePoint Online Client Extensibility Web Application Principal",
            value: "user_impersonation"
          }
        ],
        passwordCredentials: [],
        resourceSpecificApplicationPermissions: [],
        verifiedPublisher: {
          displayName: null,
          verifiedPublisherId: null,
          addedDateTime: null
        }
      }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists pending permission requests (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/?$filter=displayName eq 'SharePoint Online Client Extensibility Web Application Principal'`) {
        return spoClientExtensibilityWebApplicationPrincipalResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/oAuth2Permissiongrants/?$filter=clientId eq '${spoClientExtensibilityWebApplicationPrincipalResponse.value[0].id}' and consentType eq 'AllPrincipals'`) {
        return oAuth2PermissionGrantsResponse;
      }

      throw 'invalid request';
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><Query Id="13" ObjectPathId="11"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionRequests" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "ed4e3a9e-5007-4000-d6f5-927416c34546"
          }, 10, {
            "IsNull": false
          }, 12, {
            "IsNull": false
          }, 13, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionRequestCollection", "_Child_Items_": [
              {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionRequest", "Id": "\/Guid(4dc4c043-25ee-40f2-81d3-b3bf63da7538)\/", "Resource": "Microsoft Graph", "ResourceId": "Microsoft Graph", "Scope": "Mail.Read"
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith([{
      Id: '4dc4c043-25ee-40f2-81d3-b3bf63da7538',
      Resource: 'Microsoft Graph',
      ResourceId: 'Microsoft Graph',
      Scope: 'Mail.Read'
    }]));
  });

  it('lists pending permission requests', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/?$filter=displayName eq 'SharePoint Online Client Extensibility Web Application Principal'`) {
        return spoClientExtensibilityWebApplicationPrincipalResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/oAuth2Permissiongrants/?$filter=clientId eq '${spoClientExtensibilityWebApplicationPrincipalResponse.value[0].id}' and consentType eq 'AllPrincipals'`) {
        return oAuth2PermissionGrantsResponse;
      }

      throw 'invalid request';
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><Query Id="13" ObjectPathId="11"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionRequests" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "ed4e3a9e-5007-4000-d6f5-927416c34546"
          }, 10, {
            "IsNull": false
          }, 12, {
            "IsNull": false
          }, 13, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionRequestCollection", "_Child_Items_": [
              {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionRequest", "Id": "\/Guid(4dc4c043-25ee-40f2-81d3-b3bf63da7538)\/", "Resource": "Microsoft Graph", "ResourceId": "Microsoft Graph", "Scope": "Mail.Read"
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith([{
      Id: '4dc4c043-25ee-40f2-81d3-b3bf63da7538',
      Resource: 'Microsoft Graph',
      ResourceId: 'Microsoft Graph',
      Scope: 'Mail.Read'
    }]));
  });

  it('lists pending permission requests when no service principal is found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/?$filter=displayName eq 'SharePoint Online Client Extensibility Web Application Principal'`) {
        return { value: [] };
      }

      throw 'invalid request';
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><Query Id="13" ObjectPathId="11"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionRequests" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "ed4e3a9e-5007-4000-d6f5-927416c34546"
          }, 10, {
            "IsNull": false
          }, 12, {
            "IsNull": false
          }, 13, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionRequestCollection", "_Child_Items_": [
              {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionRequest", "Id": "\/Guid(4dc4c043-25ee-40f2-81d3-b3bf63da7538)\/", "Resource": "Microsoft Graph", "ResourceId": "Microsoft Graph", "Scope": "Mail.Read"
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith([{
      Id: '4dc4c043-25ee-40f2-81d3-b3bf63da7538',
      Resource: 'Microsoft Graph',
      ResourceId: 'Microsoft Graph',
      Scope: 'Mail.Read'
    }]));
  });

  it('lists pending permission requests when no oAuth2Permissiongrants are found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/?$filter=displayName eq 'SharePoint Online Client Extensibility Web Application Principal'`) {
        return spoClientExtensibilityWebApplicationPrincipalResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/oAuth2Permissiongrants/?$filter=clientId eq '${spoClientExtensibilityWebApplicationPrincipalResponse.value[0].id}' and consentType eq 'AllPrincipals'`) {
        return { value: [] };
      }

      throw 'invalid request';
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><Query Id="13" ObjectPathId="11"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionRequests" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "ed4e3a9e-5007-4000-d6f5-927416c34546"
          }, 10, {
            "IsNull": false
          }, 12, {
            "IsNull": false
          }, 13, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionRequestCollection", "_Child_Items_": [
              {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionRequest", "Id": "\/Guid(4dc4c043-25ee-40f2-81d3-b3bf63da7538)\/", "Resource": "Microsoft Graph", "ResourceId": "Microsoft Graph", "Scope": "Mail.Read"
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith([{
      Id: '4dc4c043-25ee-40f2-81d3-b3bf63da7538',
      Resource: 'Microsoft Graph',
      ResourceId: 'Microsoft Graph',
      Scope: 'Mail.Read'
    }]));
  });

  it('correctly handles error when retrieving pending permission requests', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/?$filter=displayName eq 'SharePoint Online Client Extensibility Web Application Principal'`) {
        return spoClientExtensibilityWebApplicationPrincipalResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/oAuth2Permissiongrants/?$filter=clientId eq '${spoClientExtensibilityWebApplicationPrincipalResponse.value[0].id}' and consentType eq 'AllPrincipals'`) {
        return oAuth2PermissionGrantsResponse;
      }

      throw 'invalid request';
    });

    sinon.stub(request, 'post').callsFake(() => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
            "ErrorMessage": "File Not Found.", "ErrorValue": null, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26", "ErrorCode": -2147024894, "ErrorTypeName": "System.IO.FileNotFoundException"
          }, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26"
        }
      ]));
    });
    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('File Not Found.'));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/?$filter=displayName eq 'SharePoint Online Client Extensibility Web Application Principal'`) {
        return spoClientExtensibilityWebApplicationPrincipalResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/oAuth2Permissiongrants/?$filter=clientId eq '${spoClientExtensibilityWebApplicationPrincipalResponse.value[0].id}' and consentType eq 'AllPrincipals'`) {
        return oAuth2PermissionGrantsResponse;
      }

      throw 'invalid request';
    });

    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));
    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('An error has occurred'));
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });
});
