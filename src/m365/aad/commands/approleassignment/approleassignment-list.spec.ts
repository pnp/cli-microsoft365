import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./approleassignment-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

class ServicePrincipalCollections {

  private static oneServicePrincipalWithAppRoleAssignments: any =
    {
      value: [
        {
          "odata.type": "Microsoft.DirectoryServices.ServicePrincipal",
          "appRoleAssignments@odata.navigationLinkUrl": "https://graph.windows.net/myorganization/directoryObjects/3aa76d8a-4145-40d1-89ca-b15bdb943bfd/Microsoft.DirectoryServices.ServicePrincipal/appRoleAssignments",
          "appRoleAssignments": [
            {
              "objectType": "AppRoleAssignment",
              "objectId": "im2nOkVB0UCJyrFb25Q7_eZg4Yr51ZhDlErpioz6f8k",
              "deletionTimestamp": null,
              "creationTimestamp": "2020-02-11T16:42:20.2272849Z",
              "id": "df021288-bdef-4463-88db-98f22de89214",
              "principalDisplayName": "Product Catalog daemon",
              "principalId": "3aa76d8a-4145-40d1-89ca-b15bdb943bfd",
              "principalType": "ServicePrincipal",
              "resourceDisplayName": "Microsoft Graph",
              "resourceId": "b1ce2d04-5502-4142-ba53-819327b74b5b"
            },
            {
              "objectType": "AppRoleAssignment",
              "objectId": "im2nOkVB0UCJyrFb25Q7_c9ubVNI2s9PkLasaAPuNQM",
              "deletionTimestamp": null,
              "creationTimestamp": "2020-02-11T01:27:47.395556Z",
              "id": "9116d0c7-0632-4203-889f-a24a08442b3d",
              "principalDisplayName": "Product Catalog daemon",
              "principalId": "3aa76d8a-4145-40d1-89ca-b15bdb943bfd",
              "principalType": "ServicePrincipal",
              "resourceDisplayName": "Contoso Product Catalog service",
              "resourceId": "b3598f45-9d8c-41c9-b5f0-81eb7ea8551f"
            }
          ],
          "objectType": "ServicePrincipal",
          "objectId": "3aa76d8a-4145-40d1-89ca-b15bdb943bfd",
          "deletionTimestamp": null,
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
          "informationalUrls": {
            "termsOfService": null,
            "support": null,
            "privacy": null,
            "marketing": null
          },
          "keyCredentials": [],
          "logoutUrl": null,
          "notificationEmailAddresses": [],
          "oauth2Permissions": [],
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

  private static oneServicePrincipalWithNoAppRoleAssignments: any =
    {
      value: [
        {
          "odata.type": "Microsoft.DirectoryServices.ServicePrincipal",
          "appRoleAssignments@odata.navigationLinkUrl": "https://graph.windows.net/myorganization/directoryObjects/43a9e7d8-0469-42f5-8c9d-17ac8c974ba6/Microsoft.DirectoryServices.ServicePrincipal/appRoleAssignments",
          "appRoleAssignments": [],
          "objectType": "ServicePrincipal",
          "objectId": "43a9e7d8-0469-42f5-8c9d-17ac8c974ba6",
          "deletionTimestamp": null,
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
          "informationalUrls": {
            "termsOfService": null,
            "support": null,
            "privacy": null,
            "marketing": null
          },
          "keyCredentials": [],
          "logoutUrl": "https://localhost:5001/signout-oidc",
          "notificationEmailAddresses": [],
          "oauth2Permissions": [],
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
    "objectType": "ServicePrincipal",
    "objectId": "b3598f45-9d8c-41c9-b5f0-81eb7ea8551f",
    "deletionTimestamp": null,
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
    "informationalUrls": {
      "termsOfService": null,
      "support": null,
      "privacy": null,
      "marketing": null
    },
    "keyCredentials": [],
    "logoutUrl": null,
    "notificationEmailAddresses": [],
    "oauth2Permissions": [
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

  static servicePrincipalCustomAppWithNoAppRole: any =
    {
      "odata.metadata": "https://graph.windows.net/myorganization/$metadata#directoryObjects/@Element",
      "odata.type": "Microsoft.DirectoryServices.ServicePrincipal",
      "objectType": "ServicePrincipal",
      "objectId": "003c6308-0075-4e45-b310-d04c72bd649f",
      "deletionTimestamp": null,
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
      "informationalUrls": {
        "termsOfService": null,
        "support": null,
        "privacy": null,
        "marketing": null
      },
      "keyCredentials": [],
      "logoutUrl": null,
      "notificationEmailAddresses": [],
      "oauth2Permissions": [],
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

  static servicePrincipalMicrosoftGraphWithAppRole: any =
    {
      "objectType": "ServicePrincipal",
      "objectId": "b1ce2d04-5502-4142-ba53-819327b74b5b",
      "deletionimestamp": null,
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
  static objectIdWithRoleAssigments: string = "3aa76d8a-4145-40d1-89ca-b15bdb943bfd";
  static invalidAppId: string = "12345678-abcd-9876-fedc-0123456789ab";
}

class InternalRequestStub {
  static customAppId: string = "b3598f45-9d8c-41c9-b5f0-81eb7ea8551f";
  static microsoftGraphAppId: string = "b1ce2d04-5502-4142-ba53-819327b74b5b";
}

class RequestStub {
  static retrieveAppRoles = ((opts: any) => {
    // we need to fake three calls:
    // 2. get the service principal for the assigned resource(s)
    // 3. get the app roles of the resource

    // query for service principal
    if ((opts.url as string).indexOf(`/myorganization/servicePrincipals?api-version=1.6&$expand=appRoleAssignments&$filter=`) > -1) {
      // by app id
      if ((opts.url as string).indexOf(`appId eq '${CommandActionParameters.appIdWithRoleAssignments}'`) > -1) {
        return Promise.resolve(ServicePrincipalCollections.ServicePrincipalByAppId);
      }
      // by object id
      if ((opts.url as string).indexOf(`objectId eq '${CommandActionParameters.objectIdWithRoleAssigments}'`) > -1) {
        return Promise.resolve(ServicePrincipalCollections.ServicePrincipalByAppId);
      }
      // by display name
      if ((opts.url as string).indexOf(`displayName eq '${encodeURIComponent(CommandActionParameters.appNameWithRoleAssignments)}'`) > -1) {
        return Promise.resolve(ServicePrincipalCollections.ServicePrincipalByDisplayName);
      }
      // by app id: no app role assignments
      if ((opts.url as string).indexOf(`appId eq '${CommandActionParameters.appIdWithNoRoleAssignments}'`) > -1) {
        return Promise.resolve(ServicePrincipalCollections.ServicePrincipalByAppIdNotFound);
      }
      // by app id: does not exist
      if ((opts.url as string).indexOf(`appId eq '${CommandActionParameters.invalidAppId}'`) > -1) {
        return Promise.resolve(ServicePrincipalCollections.ServicePrincipalByAppIdNotFound);
      }
    }

    if ((opts.url as string).indexOf(`/myorganization/servicePrincipals/${InternalRequestStub.customAppId}?api-version=1.6`) > -1) {
      return Promise.resolve(ServicePrincipalObject.servicePrincipalCustomAppWithAppRole);
    }

    if ((opts.url as string).indexOf(`/myorganization/servicePrincipals/${InternalRequestStub.microsoftGraphAppId}?api-version=1.6`) > -1) {
      return Promise.resolve(ServicePrincipalObject.servicePrincipalMicrosoftGraphWithAppRole);
    }

    return Promise.reject('Invalid request');
  })
}


describe(commands.APPROLEASSIGNMENT_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  let textOutput = [
    {
      "resourceDisplayName": "Microsoft Graph",
      "roleName": "User.Read.All"
    },
    {
      "resourceDisplayName": "Contoso Product Catalog service",
      "roleName": "access_as_application"
    }
  ];
  let jsonOutput = [
    {
      "appRoleId": "df021288-bdef-4463-88db-98f22de89214",
      "resourceDisplayName": "Microsoft Graph",
      "resourceId": "b1ce2d04-5502-4142-ba53-819327b74b5b",
      "roleId": "df021288-bdef-4463-88db-98f22de89214",
      "roleName": "User.Read.All"
    },
    {
      "appRoleId": "9116d0c7-0632-4203-889f-a24a08442b3d",
      "resourceDisplayName": "Contoso Product Catalog service",
      "resourceId": "b3598f45-9d8c-41c9-b5f0-81eb7ea8551f",
      "roleId": "9116d0c7-0632-4203-889f-a24a08442b3d",
      "roleName": "access_as_application"
    }
  ];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.APPROLEASSIGNMENT_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves App Role assignments for the specified displayName', (done) => {
    sinon.stub(request, 'get').callsFake(RequestStub.retrieveAppRoles);

    cmdInstance.action({ options: { output: 'json', displayName: CommandActionParameters.appNameWithRoleAssignments } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('retrieves App Role assignments for the specified appId', (done) => {
    sinon.stub(request, 'get').callsFake(RequestStub.retrieveAppRoles);

    cmdInstance.action({ options: { output: 'json', appId: CommandActionParameters.appIdWithRoleAssignments } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves App Role assignments for the specified appId and outputs text', (done) => {
    sinon.stub(request, 'get').callsFake(RequestStub.retrieveAppRoles);

    cmdInstance.action({ options: { output: 'text', appId: CommandActionParameters.appIdWithRoleAssignments } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(textOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves App Role assignments for the specified objectId and outputs text', (done) => {
    sinon.stub(request, 'get').callsFake(RequestStub.retrieveAppRoles);

    cmdInstance.action({ options: { output: 'text', objectId: CommandActionParameters.objectIdWithRoleAssigments } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(textOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles an appId that does not exist', (done) => {
    sinon.stub(request, 'get').callsFake(RequestStub.retrieveAppRoles);

    cmdInstance.action({ options: { appId: CommandActionParameters.invalidAppId } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('app registration not found')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no app role assignments for the specified app', (done) => {
    sinon.stub(request, 'get').callsFake(RequestStub.retrieveAppRoles);

    cmdInstance.action({ options: { appId: CommandActionParameters.appIdWithNoRoleAssignments } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('app registration not found')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: `Resource '' does not exist or one of its queried reference-property objects are not present`
            }
          }
        }
      });
    });

    cmdInstance.action({ options: { debug: false, appId: '36e3a540-6f25-4483-9542-9f5fa00bb633' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if neither appId nor displayName are not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appId: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the objectId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { objectId: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appId and displayName are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appId: CommandActionParameters.appIdWithNoRoleAssignments, displayName: CommandActionParameters.appNameWithRoleAssignments } });
    assert.notStrictEqual(actual, true);
  })

  it('fails validation if objectId and displayName are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: CommandActionParameters.appNameWithRoleAssignments, objectId: CommandActionParameters.objectIdWithRoleAssigments } });
    assert.notStrictEqual(actual, true);
  })

  it('passes validation when the appId option specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appId: CommandActionParameters.appIdWithNoRoleAssignments } });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying appId', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--appId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying displayName', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--displayName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});

