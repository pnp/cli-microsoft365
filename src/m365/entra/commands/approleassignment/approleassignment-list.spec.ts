import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './approleassignment-list.js';
import { settingsNames } from '../../../../settingsNames.js';

class ServicePrincipalAppRoleAssignments {
  private static AppRoleAssignments: any = {
    value: [
      {
        "id": "im2nOkVB0UCJyrFb25Q7_eZg4Yr51ZhDlErpioz6f8k",
        "deletedDateTime": null,
        "createdDateTime": "2020-02-11T16:42:20.2272849Z",
        "appRoleId": "df021288-bdef-4463-88db-98f22de89214",
        "principalDisplayName": "Product Catalog daemon",
        "principalId": "3aa76d8a-4145-40d1-89ca-b15bdb943bfd",
        "principalType": "ServicePrincipal",
        "resourceDisplayName": "Microsoft Graph",
        "resourceId": "b1ce2d04-5502-4142-ba53-819327b74b5b"
      },
      {
        "id": "im2nOkVB0UCJyrFb25Q7_c9ubVNI2s9PkLasaAPuNQM",
        "deletedDateTime": null,
        "createdDateTime": "2020-02-11T01:27:47.395556Z",
        "appRoleId": "9116d0c7-0632-4203-889f-a24a08442b3d",
        "principalDisplayName": "Product Catalog daemon",
        "principalId": "3aa76d8a-4145-40d1-89ca-b15bdb943bfd",
        "principalType": "ServicePrincipal",
        "resourceDisplayName": "Contoso Product Catalog service",
        "resourceId": "b3598f45-9d8c-41c9-b5f0-81eb7ea8551f"
      }
    ]
  };

  static NoAppRoleAssignments: any = { value: [] };
  static WithAppRoleAssignments: any = ServicePrincipalAppRoleAssignments.AppRoleAssignments;
}

class ServicePrincipalCollections {
  private static oneServicePrincipalWithAppRoleAssignments: any = {
    value: [
      {
        "appRoleAssignments": ServicePrincipalAppRoleAssignments.WithAppRoleAssignments.value,
        "id": "3aa76d8a-4145-40d1-89ca-b15bdb943bfd",
        "deletedDateTime": null,
        "accountEnabled": true,
        "addIns": [],
        "alternativeNames": [],
        "appDisplayName": "Product Catalog daemon",
        "appId": "36e3a540-6f25-4483-9542-9f5fa00bb633",
        "applicationTemplateId": null,
        "appOwnerTenantId": "187d6ed4-5c94-40eb-87c7-d311ec5f647a",
        "appRoleAssignmentRequired": false,
        "appRoles": [],
        "displayName": "Product Catalog daemon",
        "errorUrl": null,
        "homepage": null,
        "info": {
          "termsOfService": null,
          "support": null,
          "privacy": null,
          "marketing": null
        },
        "keyCredentials": [],
        "logoutUrl": null,
        "notificationEmailAddresses": [],
        "oauth2PermissionScopes": [],
        "passwordCredentials": [],
        "preferredSingleSignOnMode": null,
        "preferredTokenSigningKeyEndDateTime": null,
        "preferredTokenSigningKeyThumbprint": null,
        "publisherName": "Contoso",
        "replyUrls": [],
        "samlMetadataUrl": null,
        "samlSingleSignOnSettings": null,
        "servicePrincipalNames": [
          "36e3a540-6f25-4483-9542-9f5fa00bb633"
        ],
        "servicePrincipalType": "Application",
        "signInAudience": "AzureADMyOrg",
        "tags": [
          "WindowsAzureActiveDirectoryIntegratedApp"
        ],
        "tokenEncryptionKeyId": null
      }
    ]
  };

  private static oneServicePrincipalWithNoAppRoleAssignments: any = {
    value: [
      {
        "appRoleAssignments": [],
        "id": "43a9e7d8-0469-42f5-8c9d-17ac8c974ba6",
        "deletedDateTime": null,
        "accountEnabled": true,
        "addIns": [],
        "alternativeNames": [],
        "appDisplayName": "Product Catalog WebApp",
        "appId": "1c21749e-df7a-4fed-b3ab-921dce3bb124",
        "applicationTemplateId": null,
        "appOwnerTenantId": "187d6ed4-5c94-40eb-87c7-d311ec5f647a",
        "appRoleAssignmentRequired": false,
        "appRoles": [],
        "displayName": "Product Catalog WebApp",
        "errorUrl": null,
        "homepage": null,
        "info": {
          "termsOfService": null,
          "support": null,
          "privacy": null,
          "marketing": null
        },
        "keyCredentials": [],
        "logoutUrl": "https://localhost:5001/signout-oidc",
        "notificationEmailAddresses": [],
        "oauth2PermissionScopes": [],
        "passwordCredentials": [],
        "preferredSingleSignOnMode": null,
        "preferredTokenSigningKeyEndDateTime": null,
        "preferredTokenSigningKeyThumbprint": null,
        "publisherName": "Contoso",
        "replyUrls": [
          "https://localhost:5001/signin-oidc"
        ],
        "samlMetadataUrl": null,
        "samlSingleSignOnSettings": null,
        "servicePrincipalNames": [
          "1c21749e-df7a-4fed-b3ab-921dce3bb124"
        ],
        "servicePrincipalType": "Application",
        "signInAudience": "AzureADMyOrg",
        "tags": [
          "WindowsAzureActiveDirectoryIntegratedApp"
        ],
        "tokenEncryptionKeyId": null
      }
    ]
  };

  static ServicePrincipalByAppIdNotFound: any = { value: [] };
  static ServicePrincipalByAppIdNoAppRoles: any = ServicePrincipalCollections.oneServicePrincipalWithNoAppRoleAssignments;
  static ServicePrincipalByDisplayName: any = ServicePrincipalCollections.oneServicePrincipalWithAppRoleAssignments;
  static ServicePrincipalByAppId: any = ServicePrincipalCollections.oneServicePrincipalWithAppRoleAssignments;
}

class ServicePrincipalObject {
  static servicePrincipalCustomAppWithAppRole: any = {
    "id": "b3598f45-9d8c-41c9-b5f0-81eb7ea8551f",
    "deletedDateTime": null,
    "accountEnabled": true,
    "addIns": [],
    "alternativeNames": [],
    "appDisplayName": "Contoso Product Catalog service",
    "appId": "97a1ab8b-9ede-41fc-8370-7199a4c16224",
    "applicationTemplateId": null,
    "appOwnerTenantId": "187d6ed4-5c94-40eb-87c7-d311ec5f647a",
    "appRoleAssignmentRequired": false,
    "appRoles": [
      {
        "allowedMemberTypes": [
          "Application"
        ],
        "description": "Accesses the Product Catalog API as an application.",
        "displayName": "access_as_application",
        "id": "9116d0c7-0632-4203-889f-a24a08442b3d",
        "isEnabled": true,
        "value": "access_as_application"
      }
    ],
    "displayName": "Contoso Product Catalog service",
    "errorUrl": null,
    "homepage": null,
    "info": {
      "termsOfService": null,
      "support": null,
      "privacy": null,
      "marketing": null
    },
    "keyCredentials": [],
    "logoutUrl": null,
    "notificationEmailAddresses": [],
    "oauth2PermissionScopes": [
      {
        "adminConsentDescription": "Allows the app to write Product Categories",
        "adminConsentDisplayName": "Write Product Categories",
        "id": "88bd47c3-6961-481b-b8c5-d2e923a776ea",
        "isEnabled": true,
        "type": "Admin",
        "userConsentDescription": "Allows the app to write Product Categories",
        "userConsentDisplayName": "Write Product Categories",
        "value": "Category.Write"
      },
      {
        "adminConsentDescription": "Allows the app to read Product Categories",
        "adminConsentDisplayName": "Read Product Categories",
        "id": "442ce90e-98bd-4067-9915-556ac96ea376",
        "isEnabled": true,
        "type": "Admin",
        "userConsentDescription": "Allows the app to read Product Categories",
        "userConsentDisplayName": "Read Product Categories",
        "value": "Category.Read"
      },
      {
        "adminConsentDescription": "Allows users to update product information.",
        "adminConsentDisplayName": "Update Products",
        "id": "0f289128-502f-4991-b7ee-dca9ecdfec66",
        "isEnabled": true,
        "type": "User",
        "userConsentDescription": "Allows users to update product information.",
        "userConsentDisplayName": "Update Products",
        "value": "Product.Write"
      },
      {
        "adminConsentDescription": "Allows the user to read Product information",
        "adminConsentDisplayName": "Read Products",
        "id": "68ae834c-e4c4-4660-8c4e-0e4ffe044e77",
        "isEnabled": true,
        "type": "User",
        "userConsentDescription": "Allows the user to read Product information",
        "userConsentDisplayName": "Read Products",
        "value": "Product.Read"
      }
    ],
    "passwordCredentials": [],
    "preferredSingleSignOnMode": null,
    "preferredTokenSigningKeyEndDateTime": null,
    "preferredTokenSigningKeyThumbprint": null,
    "publisherName": "Contoso",
    "replyUrls": [],
    "samlMetadataUrl": null,
    "samlSingleSignOnSettings": null,
    "servicePrincipalNames": [
      "api://97a1ab8b-9ede-41fc-8370-7199a4c16224",
      "97a1ab8b-9ede-41fc-8370-7199a4c16224"
    ],
    "servicePrincipalType": "Application",
    "signInAudience": "AzureADMyOrg",
    "tags": [
      "WindowsAzureActiveDirectoryIntegratedApp"
    ],
    "tokenEncryptionKeyId": null
  };

  static servicePrincipalCustomAppWithNoAppRole: any = {
    "id": "003c6308-0075-4e45-b310-d04c72bd649f",
    "deletedDateTime": null,
    "accountEnabled": true,
    "addIns": [],
    "alternativeNames": [],
    "appDisplayName": "Contoso Product Catalog native client",
    "appId": "ea79e953-7984-4f6f-bbad-56d6d71070d1",
    "applicationTemplateId": null,
    "appOwnerTenantId": "187d6ed4-5c94-40eb-87c7-d311ec5f647a",
    "appRoleAssignmentRequired": false,
    "appRoles": [],
    "displayName": "Contoso Product Catalog native client",
    "errorUrl": null,
    "homepage": null,
    "info": {
      "termsOfService": null,
      "support": null,
      "privacy": null,
      "marketing": null
    },
    "keyCredentials": [],
    "logoutUrl": null,
    "notificationEmailAddresses": [],
    "oauth2PermissionScopes": [],
    "passwordCredentials": [],
    "preferredSingleSignOnMode": null,
    "preferredTokenSigningKeyEndDateTime": null,
    "preferredTokenSigningKeyThumbprint": null,
    "publisherName": "Contoso",
    "replyUrls": [
      "https://login.microsoftonline.com/common/oauth2/nativeclient"
    ],
    "samlMetadataUrl": null,
    "samlSingleSignOnSettings": null,
    "servicePrincipalNames": [
      "ea79e953-7984-4f6f-bbad-56d6d71070d1"
    ],
    "servicePrincipalType": "Application",
    "signInAudience": "AzureADMyOrg",
    "tags": [
      "WindowsAzureActiveDirectoryIntegratedApp"
    ],
    "tokenEncryptionKeyId": null
  };

  static servicePrincipalMicrosoftGraphWithAppRole: any = {
    "id": "b1ce2d04-5502-4142-ba53-819327b74b5b",
    "deletedDateTime": null,
    "accountEnabled": true,
    "addIns": [],
    "alternativeNames": [],
    "appDisplayName": "Microsoft Graph",
    "appId": "0000003-0000-0000-c000-000000000000",
    "appliationTemplateId": null,
    "appOwnerTenantId": "f8cdef31-a31e-4b4a-93e4-5f571e91255a",
    "appRoleAssignmenRequired": false,
    "appRoles": [
      {
        "allowedMemberTypes": [
          "Application"
        ],
        "description": "Allows the app to read user profiles without a signed in user.",
        "displayName": "Read all users' full profiles",
        "id": "df021288-bdef-4463-88db-98f22de89214",
        "isEnabled": true,
        "value": "User.Read.All"
      }
    ],
    "passwordCredentials": [],
    "preferredSingleSignOnMode": null,
    "preferredTokenSigningKeyEndDateTime": null,
    "preferredTokenSigningKeyThumbprint": null,
    "publisherName": "Microsoft Services",
    "replyUrls": [],
    "samlMetadataUrl": null,
    "samlSingleSignOnSettings": null,
    "servicePrincipalNames": [
      "00000003-0000-0000-c000-000000000000/ags.windows.net",
      "00000003-0000-0000-c000-000000000000",
      "https://canary.graph.microsoft.com",
      "https://graph.microsoft.com",
      "https://ags.windows.net",
      "https://graph.microsoft.us",
      "https://graph.microsoft.com/",
      "https://dod-graph.microsoft.us"
    ],
    "servicePrincipalType": "Application",
    "signInAudience": "AzureADMultipleOrgs",
    "tags": [],
    "tokenEncryptionKeyId": null
  };
}

class CommandActionParameters {
  static appIdWithRoleAssignments: string = "36e3a540-6f25-4483-9542-9f5fa00bb633";
  static appNameWithRoleAssignments: string = "Product Catalog daemon";
  static appIdWithNoRoleAssignments: string = "1c21749e-df7a-4fed-b3ab-921dce3bb124";
  static objectIdWithRoleAssignments: string = "3aa76d8a-4145-40d1-89ca-b15bdb943bfd";
  static objectIdNoRoleAssignments: string = "021d971f-779d-439b-8006-9f084423f344";
  static invalidAppId: string = "12345678-abcd-9876-fedc-0123456789ab";
}

class InternalRequestStub {
  static customAppId: string = "b3598f45-9d8c-41c9-b5f0-81eb7ea8551f";
  static microsoftGraphAppId: string = "b1ce2d04-5502-4142-ba53-819327b74b5b";
}

class RequestStub {
  static retrieveAppRoles = (async (opts: any) => {
    // we need to fake three calls:
    // 1. query the service principal endpoint based on input parameters
    // 2. get the service principal for the assigned resource(s)
    // 3. get the app roles of the resource

    // query for service principal
    if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?$expand=appRoleAssignments&$filter=`) > -1) {
      // by app id
      if ((opts.url as string).indexOf(`appId eq '${CommandActionParameters.appIdWithRoleAssignments}'`) > -1) {
        return ServicePrincipalCollections.ServicePrincipalByAppId;
      }
      // by display name
      if ((opts.url as string).indexOf(`displayName eq '${formatting.encodeQueryParameter(CommandActionParameters.appNameWithRoleAssignments)}'`) > -1) {
        return ServicePrincipalCollections.ServicePrincipalByDisplayName;
      }
      // by app id: no app role assignments
      if ((opts.url as string).indexOf(`appId eq '${CommandActionParameters.appIdWithNoRoleAssignments}'`) > -1) {
        return ServicePrincipalCollections.ServicePrincipalByAppIdNotFound;
      }
      // by app id: does not exist
      if ((opts.url as string).indexOf(`appId eq '${CommandActionParameters.invalidAppId}'`) > -1) {
        return ServicePrincipalCollections.ServicePrincipalByAppIdNotFound;
      }
    }

    if ((opts.url as string).indexOf(`/v1.0/servicePrincipals/${InternalRequestStub.customAppId}`) > -1) {
      return ServicePrincipalObject.servicePrincipalCustomAppWithAppRole;
    }

    if ((opts.url as string).indexOf(`/v1.0/servicePrincipals/${InternalRequestStub.microsoftGraphAppId}`) > -1) {
      return ServicePrincipalObject.servicePrincipalMicrosoftGraphWithAppRole;
    }

    // get service principal app role assignments : roles found
    if ((opts.url as string).indexOf(`/v1.0/servicePrincipals/${CommandActionParameters.objectIdWithRoleAssignments}/appRoleAssignments`) > -1) {
      return ServicePrincipalAppRoleAssignments.WithAppRoleAssignments;
    }

    // get service principal app role assignments : no roles found
    if ((opts.url as string).indexOf(`/v1.0/servicePrincipals/${CommandActionParameters.objectIdNoRoleAssignments}/appRoleAssignments`) > -1) {
      return ServicePrincipalAppRoleAssignments.NoAppRoleAssignments;
    }

    throw 'Invalid request';
  });
}

describe(commands.APPROLEASSIGNMENT_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const jsonOutput = [
    {
      "appRoleId": "df021288-bdef-4463-88db-98f22de89214",
      "resourceDisplayName": "Microsoft Graph",
      "resourceId": "b1ce2d04-5502-4142-ba53-819327b74b5b",
      "roleId": "df021288-bdef-4463-88db-98f22de89214",
      "roleName": "User.Read.All",
      "created": "2020-02-11T16:42:20.2272849Z",
      "deleted": null
    },
    {
      "appRoleId": "9116d0c7-0632-4203-889f-a24a08442b3d",
      "resourceDisplayName": "Contoso Product Catalog service",
      "resourceId": "b3598f45-9d8c-41c9-b5f0-81eb7ea8551f",
      "roleId": "9116d0c7-0632-4203-889f-a24a08442b3d",
      "roleName": "access_as_application",
      "created": "2020-02-11T01:27:47.395556Z",
      "deleted": null
    }
  ];

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APPROLEASSIGNMENT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['resourceDisplayName', 'roleName']);
  });

  it('retrieves App Role assignments for the specified appDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(RequestStub.retrieveAppRoles);

    await command.action(logger, { options: { output: 'json', appDisplayName: CommandActionParameters.appNameWithRoleAssignments } });
    assert(loggerLogSpy.calledWith(jsonOutput));
  });

  it('retrieves App Role assignments for the specified appId', async () => {
    sinon.stub(request, 'get').callsFake(RequestStub.retrieveAppRoles);

    await command.action(logger, { options: { output: 'json', appId: CommandActionParameters.appIdWithRoleAssignments } });
    assert(loggerLogSpy.calledWith(jsonOutput));
  });

  it('retrieves App Role assignments for the specified appId and outputs text', async () => {
    sinon.stub(request, 'get').callsFake(RequestStub.retrieveAppRoles);

    await command.action(logger, { options: { output: 'text', appId: CommandActionParameters.appIdWithRoleAssignments } });
    assert(loggerLogSpy.calledWith(jsonOutput));
  });

  it('retrieves App Role assignments for the specified appObjectId and outputs text', async () => {
    sinon.stub(request, 'get').callsFake(RequestStub.retrieveAppRoles);

    await command.action(logger, { options: { output: 'text', appObjectId: CommandActionParameters.objectIdWithRoleAssignments } });
    assert(loggerLogSpy.calledWith(jsonOutput));
  });

  it('correctly handles an appId that does not exist', async () => {
    sinon.stub(request, 'get').callsFake(RequestStub.retrieveAppRoles);

    await assert.rejects(command.action(logger, { options: { appId: CommandActionParameters.invalidAppId } } as any), new CommandError('app registration not found'));
  });

  it('correctly handles a service principal that does not have any app role assignments', async () => {
    sinon.stub(request, 'get').callsFake(RequestStub.retrieveAppRoles);

    await assert.rejects(command.action(logger, { options: { appObjectId: CommandActionParameters.objectIdNoRoleAssignments } } as any), new CommandError('no app role assignments found'));
  });

  it('correctly handles no app role assignments for the specified app', async () => {
    sinon.stub(request, 'get').callsFake(RequestStub.retrieveAppRoles);

    await assert.rejects(command.action(logger, { options: { appId: CommandActionParameters.appIdWithNoRoleAssignments } } as any), new CommandError('app registration not found'));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: `Resource '' does not exist or one of its queried reference-property objects are not present`
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { appObjectId: '021d971f-779d-439b-8006-9f084423f344' } } as any), new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });

  it('fails validation if neither appId nor appDisplayName are not specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { appId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appObjectId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { appObjectId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appId and appDisplayName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: CommandActionParameters.appIdWithNoRoleAssignments, appDisplayName: CommandActionParameters.appNameWithRoleAssignments } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appObjectId and appDisplayName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appDisplayName: CommandActionParameters.appNameWithRoleAssignments, appObjectId: CommandActionParameters.objectIdWithRoleAssignments } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the appId option specified', async () => {
    const actual = await command.validate({ options: { appId: CommandActionParameters.appIdWithNoRoleAssignments } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying appId', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--appId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying appDisplayName', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--appDisplayName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});

