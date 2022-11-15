import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./folder-get');

describe(commands.FOLDER_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let stubGetResponses: any;

  const expectedFolder = {
    "ListItemAllFields": {
      "RoleAssignments": [
        {
          "Member": {
            "Id": 3,
            "IsHiddenInUI": false,
            "LoginName": "Test_Site Owners",
            "Title": "Test_Site Owners",
            "PrincipalType": 8,
            "AllowMembersEditMembership": false,
            "AllowRequestToJoinLeave": false,
            "AutoAcceptRequestToJoinLeave": false,
            "Description": null,
            "OnlyAllowMembersViewMembership": false,
            "OwnerTitle": "Test_Site Owners",
            "RequestToJoinLeaveEmailSetting": "",
            "PrincipalTypeString": "SharePointGroup"
          },
          "RoleDefinitionBindings": [
            {
              "BasePermissions": {
                "High": "2147483647",
                "Low": "4294967295"
              },
              "Description": "Has full control.",
              "Hidden": false,
              "Id": 1073741829,
              "Name": "Full Control",
              "Order": 1,
              "RoleTypeKind": 5,
              "BasePermissionsValue": [
                "ViewListItems",
                "AddListItems",
                "EditListItems",
                "DeleteListItems",
                "ApproveItems",
                "OpenItems",
                "ViewVersions",
                "DeleteVersions",
                "CancelCheckout",
                "ManagePersonalViews",
                "ManageLists",
                "ViewFormPages",
                "AnonymousSearchAccessList",
                "Open",
                "ViewPages",
                "AddAndCustomizePages",
                "ApplyThemeAndBorder",
                "ApplyStyleSheets",
                "ViewUsageData",
                "CreateSSCSite",
                "ManageSubwebs",
                "CreateGroups",
                "ManagePermissions",
                "BrowseDirectories",
                "BrowseUserInfo",
                "AddDelPrivateWebParts",
                "UpdatePersonalWebParts",
                "ManageWeb",
                "AnonymousSearchAccessWebLists",
                "UseClientIntegration",
                "UseRemoteAPIs",
                "ManageAlerts",
                "CreateAlerts",
                "EditMyUserInfo",
                "EnumeratePermissions"
              ],
              "RoleTypeKindValue": "Administrator"
            }
          ],
          "PrincipalId": 3
        },
        {
          "Member": {
            "Id": 4,
            "IsHiddenInUI": false,
            "LoginName": "Test_Site Visitors",
            "Title": "Test_Site Visitors",
            "PrincipalType": 8,
            "AllowMembersEditMembership": false,
            "AllowRequestToJoinLeave": false,
            "AutoAcceptRequestToJoinLeave": false,
            "Description": null,
            "OnlyAllowMembersViewMembership": false,
            "OwnerTitle": "Test_Site Owners",
            "RequestToJoinLeaveEmailSetting": "",
            "PrincipalTypeString": "SharePointGroup"
          },
          "RoleDefinitionBindings": [
            {
              "BasePermissions": {
                "High": "176",
                "Low": "138612833"
              },
              "Description": "Can view pages and list items and download documents.",
              "Hidden": false,
              "Id": 1073741826,
              "Name": "Read",
              "Order": 128,
              "RoleTypeKind": 2,
              "BasePermissionsValue": [
                "ViewListItems",
                "OpenItems",
                "ViewVersions",
                "ViewFormPages",
                "Open",
                "ViewPages",
                "CreateSSCSite",
                "BrowseUserInfo",
                "UseClientIntegration",
                "UseRemoteAPIs",
                "CreateAlerts"
              ],
              "RoleTypeKindValue": "Reader"
            }
          ],
          "PrincipalId": 4
        },
        {
          "Member": {
            "Id": 5,
            "IsHiddenInUI": false,
            "LoginName": "Test_Site Members",
            "Title": "Test_Site Members",
            "PrincipalType": 8,
            "AllowMembersEditMembership": true,
            "AllowRequestToJoinLeave": false,
            "AutoAcceptRequestToJoinLeave": false,
            "Description": null,
            "OnlyAllowMembersViewMembership": false,
            "OwnerTitle": "Test_Site Owners",
            "RequestToJoinLeaveEmailSetting": "",
            "PrincipalTypeString": "SharePointGroup"
          },
          "RoleDefinitionBindings": [
            {
              "BasePermissions": {
                "High": "432",
                "Low": "1011030767"
              },
              "Description": "Can add, edit and delete lists; can view, add, update and delete list items and documents.",
              "Hidden": false,
              "Id": 1073741830,
              "Name": "Edit",
              "Order": 48,
              "RoleTypeKind": 6,
              "BasePermissionsValue": [
                "ViewListItems",
                "AddListItems",
                "EditListItems",
                "DeleteListItems",
                "OpenItems",
                "ViewVersions",
                "DeleteVersions",
                "ManagePersonalViews",
                "ManageLists",
                "ViewFormPages",
                "Open",
                "ViewPages",
                "CreateSSCSite",
                "BrowseDirectories",
                "BrowseUserInfo",
                "AddDelPrivateWebParts",
                "UpdatePersonalWebParts",
                "UseClientIntegration",
                "UseRemoteAPIs",
                "CreateAlerts",
                "EditMyUserInfo"
              ],
              "RoleTypeKindValue": "Editor"
            }
          ],
          "PrincipalId": 5
        },
        {
          "Member": {
            "Id": 10,
            "IsHiddenInUI": false,
            "LoginName": "i:0#.f|membership|reshmeeauckloo@reshmeeauckloo.onmicrosoft.com",
            "Title": "Reshmee Auckloo",
            "PrincipalType": 1,
            "Email": "reshmeeauckloo@reshmeeauckloo.onmicrosoft.com",
            "Expiration": "",
            "IsEmailAuthenticationGuestUser": false,
            "IsShareByEmailGuestUser": false,
            "IsSiteAdmin": false,
            "UserId": {
              "NameId": "10032000decdad9a",
              "NameIdIssuer": "urn:federation:microsoftonline"
            },
            "UserPrincipalName": "reshmeeauckloo@reshmeeauckloo.onmicrosoft.com",
            "PrincipalTypeString": "User"
          },
          "RoleDefinitionBindings": [
            {
              "BasePermissions": {
                "High": "432",
                "Low": "1011030767"
              },
              "Description": "Can add, edit and delete lists; can view, add, update and delete list items and documents.",
              "Hidden": false,
              "Id": 1073741830,
              "Name": "Edit",
              "Order": 48,
              "RoleTypeKind": 6,
              "BasePermissionsValue": [
                "ViewListItems",
                "AddListItems",
                "EditListItems",
                "DeleteListItems",
                "OpenItems",
                "ViewVersions",
                "DeleteVersions",
                "ManagePersonalViews",
                "ManageLists",
                "ViewFormPages",
                "Open",
                "ViewPages",
                "CreateSSCSite",
                "BrowseDirectories",
                "BrowseUserInfo",
                "AddDelPrivateWebParts",
                "UpdatePersonalWebParts",
                "UseClientIntegration",
                "UseRemoteAPIs",
                "CreateAlerts",
                "EditMyUserInfo"
              ],
              "RoleTypeKindValue": "Editor"
            }
          ],
          "PrincipalId": 10
        },
        {
          "Member": {
            "Id": 11,
            "IsHiddenInUI": false,
            "LoginName": "i:0#.w|nt service\\spsearch",
            "Title": "spsearch",
            "PrincipalType": 1,
            "Email": "",
            "Expiration": "",
            "IsEmailAuthenticationGuestUser": false,
            "IsShareByEmailGuestUser": false,
            "IsSiteAdmin": false,
            "UserId": {
              "NameId": "s-1-5-80-87383287-2054257049-3601873072-440163018-3271026472",
              "NameIdIssuer": "urn:office:idp:activedirectory"
            },
            "UserPrincipalName": null,
            "PrincipalTypeString": "User"
          },
          "RoleDefinitionBindings": [
            {
              "BasePermissions": {
                "High": "2147483647",
                "Low": "4294967295"
              },
              "Description": "Has full control.",
              "Hidden": false,
              "Id": 1073741829,
              "Name": "Full Control",
              "Order": 1,
              "RoleTypeKind": 5,
              "BasePermissionsValue": [
                "ViewListItems",
                "AddListItems",
                "EditListItems",
                "DeleteListItems",
                "ApproveItems",
                "OpenItems",
                "ViewVersions",
                "DeleteVersions",
                "CancelCheckout",
                "ManagePersonalViews",
                "ManageLists",
                "ViewFormPages",
                "AnonymousSearchAccessList",
                "Open",
                "ViewPages",
                "AddAndCustomizePages",
                "ApplyThemeAndBorder",
                "ApplyStyleSheets",
                "ViewUsageData",
                "CreateSSCSite",
                "ManageSubwebs",
                "CreateGroups",
                "ManagePermissions",
                "BrowseDirectories",
                "BrowseUserInfo",
                "AddDelPrivateWebParts",
                "UpdatePersonalWebParts",
                "ManageWeb",
                "AnonymousSearchAccessWebLists",
                "UseClientIntegration",
                "UseRemoteAPIs",
                "ManageAlerts",
                "CreateAlerts",
                "EditMyUserInfo",
                "EnumeratePermissions"
              ],
              "RoleTypeKindValue": "Administrator"
            }
          ],
          "PrincipalId": 11
        }
      ],
      "HasUniqueRoleAssignments": true,
      "FileSystemObjectType": 1,
      "Id": 2,
      "ServerRedirectedEmbedUri": null,
      "ServerRedirectedEmbedUrl": "",
      "ContentTypeId": "0x012000F5E8903B1ACF184F8C31914CA58A6548",
      "Modified": "2022-09-28T00:12:12",
      "ComplianceAssetId": null,
      "Title": null,
      "ID": 2,
      "Created": "2022-09-28T00:12:12",
      "AuthorId": 10,
      "EditorId": 10,
      "OData__CopySource": null,
      "CheckoutUserId": null,
      "OData__UIVersionString": "1.0",
      "GUID": "d62d7cbb-9822-4e7e-8ce6-ba98ccd3243d"
    },
    "Exists": true,
    "IsWOPIEnabled": false,
    "ItemCount": 0,
    "Name": "FolderPermission",
    "ProgID": null,
    "ServerRelativeUrl": "/sites/Test1/Shared Documents/FolderPermission",
    "TimeCreated": "2022-09-28T07:12:12Z",
    "TimeLastModified": "2022-09-28T07:12:12Z",
    "UniqueId": "f5143d4f-b7e1-4d19-886f-a15517bd4635",
    "WelcomePage": ""
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;

    stubGetResponses = (getResp: any = null) => {
      return sinon.stub(request, 'get').callsFake((opts) => {
        if ((opts.url as string).indexOf('GetFolderByServerRelativeUrl') > -1 || (opts.url as string).indexOf('GetFolderById') > -1) {
          if (getResp) {
            return getResp;
          }
          else {
            return Promise.resolve({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "test1", "ProgID": null, "ServerRelativeUrl": "/sites/test1/Shared Documents/test1", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" });
          }
        }

        return Promise.reject('Invalid request');
      });
    };
    commandInfo = Cli.getCommandInfo(command);
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
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FOLDER_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [['url', 'id']]);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', url: '/Shared Documents' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL and url specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', url: '/Shared Documents' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('should correctly handle folder get reject request', async () => {
    stubGetResponses(new Promise((resolve, reject) => { reject('error1'); }));

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        url: '/Shared Documents'
      }
    } as any), new CommandError('error1'));
  });

  it('should show tip when folder get rejects with error code 500', async () => {
    sinon.stub(request, 'get').rejects({ statusCode: 500 });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        url: '/Shared Documents'
      }
    } as any), new CommandError('Please check the folder URL. Folder might not exist on the specified URL'));
  });

  it('should correctly handle folder get success request', async () => {
    stubGetResponses();

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        url: '/Shared Documents'
      }
    });
    assert(loggerLogSpy.lastCall.calledWith({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "test1", "ProgID": null, "ServerRelativeUrl": "/sites/test1/Shared Documents/test1", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" }));
  });

  it('should pass the correct id params to request', async () => {
    const request = stubGetResponses();

    await command.action(logger, {
      options: {
        debug: false,
        output: 'json',
        webUrl: 'https://contoso.sharepoint.com',
        id: 'b2307a39-e878-458b-bc90-03bc578531d6'
      }
    });
    const lastCall: any = request.lastCall.args[0];
    assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/_api/web/GetFolderById(\'b2307a39-e878-458b-bc90-03bc578531d6\')');
  });

  it('should pass the correct url params to request', async () => {
    const request = stubGetResponses();

    await command.action(logger, {
      options: {
        debug: false,
        output: 'json',
        webUrl: 'https://contoso.sharepoint.com',
        url: '/Shared Documents'
      }
    });
    const lastCall: any = request.lastCall.args[0];
    assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativeUrl(\'%2FShared%20Documents\')');
  });

  it('should pass the correct url params to request (sites/test1)', async () => {
    const request = stubGetResponses();

    await command.action(logger, {
      options: {
        debug: false,
        output: 'json',
        webUrl: 'https://contoso.sharepoint.com/sites/test1',
        url: 'Shared Documents/'
      }
    });
    const lastCall: any = request.lastCall.args[0];
    assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/sites/test1/_api/web/GetFolderByServerRelativeUrl(\'%2Fsites%2Ftest1%2FShared%20Documents\')');
  });

  it('retrieves details of folder if folder url and withPermissions option is passed', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve(expectedFolder);
    });

    await command.action(logger, {
      options: {
        debug: false,
        webUrl: 'https://contoso.sharepoint.com/sites/test1',
        url: 'Shared Documents/FolderPermission',
        withPermissions: true
      }
    });
    assert(loggerLogSpy.calledWith(expectedFolder));
  });

  it('should show tip when root folder is used withPermissions', async () => {
    const error = "Please ensure the specified folder URL or folder Id does not refer to a root folder. Use \'spo list get\' with withPermissions instead.";

    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativeUrl('%2FShared%20Documents')?$expand=ListItemAllFields/HasUniqueRoleAssignments,ListItemAllFields/RoleAssignments/Member,ListItemAllFields/RoleAssignments/RoleDefinitionBindings`)) {
        return Promise.resolve({ "data": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 2, "Name": "Shared Documents", "ProgID": null, "ServerRelativeUrl": "/Shared Documents", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" } });
      }
      return Promise.reject(error);
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        url: 'Shared Documents',
        withPermissions: true
      }
    } as any), new CommandError(error));
  });
  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });
});