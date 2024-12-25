import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './list-get.js';

describe(commands.LIST_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const versionPolicies = {
    VersionPolicies: {
      DefaultExpireAfterDays: 0,
      DefaultTrimMode: 0
    }
  };
  const listResponse = {
    AllowContentTypes: true,
    BaseTemplate: 109,
    BaseType: 1,
    ContentTypesEnabled: false,
    CrawlNonDefaultViews: false,
    Created: null,
    CurrentChangeToken: null,
    CustomActionElements: null,
    DefaultContentApprovalWorkflowId: "00000000-0000-0000-0000-000000000000",
    DefaultItemOpenUseListSetting: false,
    Description: "",
    Direction: "none",
    DocumentTemplateUrl: null,
    DraftVersionVisibility: 0,
    EnableAttachments: false,
    EnableFolderCreation: true,
    EnableMinorVersions: false,
    EnableModeration: false,
    EnableVersioning: false,
    EntityTypeName: "Documents",
    ExemptFromBlockDownloadOfNonViewableFiles: false,
    FileSavePostProcessingEnabled: false,
    ForceCheckout: false,
    HasExternalDataSource: false,
    Hidden: false,
    Id: "14b2b6ed-0885-4814-bfd6-594737cc3ae3",
    ImagePath: null,
    ImageUrl: null,
    IrmEnabled: false,
    IrmExpire: false,
    IrmReject: false,
    IsApplicationList: false,
    IsCatalog: false,
    IsPrivate: false,
    ItemCount: 69,
    LastItemDeletedDate: null,
    LastItemModifiedDate: null,
    LastItemUserModifiedDate: null,
    ListExperienceOptions: 0,
    ListItemEntityTypeFullName: null,
    MajorVersionLimit: 0,
    MajorWithMinorVersionsLimit: 0,
    MultipleDataList: false,
    NoCrawl: false,
    ParentWebPath: null,
    ParentWebUrl: null,
    ParserDisabled: false,
    ServerTemplateCanCreateFolders: true,
    TemplateFeatureId: null,
    Title: "Documents"
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves and prints all details of the list if the title option is passed and retrieves version policies for a document library', async () => {
    const webUrl = 'https://contoso.sharepoint.com';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists/GetByTitle('${listResponse.Title}')`) {
        return listResponse;
      }
      if (opts.url === `${webUrl}/_api/web/lists/GetByTitle('${listResponse.Title}')?$select=VersionPolicies&$expand=VersionPolicies`) {
        return versionPolicies;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        title: listResponse.Title,
        webUrl: webUrl
      }
    });

    assert(loggerLogSpy.calledWith({ ...listResponse, ...versionPolicies }));
  });

  it('retrieves and prints all details of the list if the title option is passed and does not retrieve version policies for a generic list', async () => {
    const webUrl = 'https://contoso.sharepoint.com';
    const listResponseGeneric = { ...listResponse, BaseTemplate: 100 };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists/GetByTitle('${listResponse.Title}')`) {
        return listResponseGeneric;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        title: listResponse.Title,
        webUrl: webUrl
      }
    });

    assert(loggerLogSpy.calledOnceWithExactly(listResponseGeneric));
  });

  it('retrieves details of list if title and properties option is passed (debug)', async () => {
    sinon.stub(request, 'get').resolves({
      "Id": "14b2b6ed-0885-4814-bfd6-594737cc3ae3",
      "Title": "Documents"
    });

    await command.action(logger, {
      options: {
        debug: true,
        title: 'Documents',
        webUrl: 'https://contoso.sharepoint.com',
        properties: 'Title,Id'
      }
    });

    assert(loggerLogSpy.calledWith({
      Id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
      Title: 'Documents'
    }));
  });

  it('retrieves details of list if list id and properties option is passed (debug)', async () => {
    sinon.stub(request, 'get').resolves(
      {
        "Id": "14b2b6ed-0885-4814-bfd6-594737cc3ae3",
        "Title": "Documents"
      }
    );

    await command.action(logger, {
      options: {
        debug: true,
        id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
        webUrl: 'https://contoso.sharepoint.com',
        properties: 'Title,Id'
      }
    });

    assert(loggerLogSpy.calledWith({
      Id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
      Title: 'Documents'
    }));
  });

  it('retrieves details of list if url and properties option is passed (debug)', async () => {
    sinon.stub(request, 'get').resolves(
      {
        "Id": "14b2b6ed-0885-4814-bfd6-594737cc3ae3",
        "Title": "Documents"
      }
    );

    await command.action(logger, {
      options: {
        debug: true,
        url: 'Shared Documents',
        webUrl: 'https://contoso.sharepoint.com',
        properties: 'Title,Id'
      }
    });

    assert(loggerLogSpy.calledWith({
      Id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
      Title: 'Documents'
    }));
  });

  it('retrieves details of list if list id and withPermissions option is passed', async () => {
    sinon.stub(request, 'get').resolves({
      "RoleAssignments": [
        {
          "Member": {
            "Id": 3,
            "IsHiddenInUI": false,
            "LoginName": "MySite Owners",
            "Title": "MySite Owners",
            "PrincipalType": 8,
            "AllowMembersEditMembership": false,
            "AllowRequestToJoinLeave": false,
            "AutoAcceptRequestToJoinLeave": false,
            "Description": null,
            "OnlyAllowMembersViewMembership": false,
            "OwnerTitle": "MySite Owners",
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
              "RoleTypeKind": 5
            }
          ],
          "PrincipalId": 3
        },
        {
          "Member": {
            "Id": 4,
            "IsHiddenInUI": false,
            "LoginName": "MySite Visitors",
            "Title": "MySite Visitors",
            "PrincipalType": 8,
            "AllowMembersEditMembership": false,
            "AllowRequestToJoinLeave": false,
            "AutoAcceptRequestToJoinLeave": false,
            "Description": null,
            "OnlyAllowMembersViewMembership": false,
            "OwnerTitle": "MySite Owners",
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
              "RoleTypeKind": 2
            }
          ],
          "PrincipalId": 4
        },
        {
          "Member": {
            "Id": 5,
            "IsHiddenInUI": false,
            "LoginName": "MySite Members",
            "Title": "MySite Members",
            "PrincipalType": 8,
            "AllowMembersEditMembership": true,
            "AllowRequestToJoinLeave": false,
            "AutoAcceptRequestToJoinLeave": false,
            "Description": null,
            "OnlyAllowMembersViewMembership": false,
            "OwnerTitle": "MySite Owners",
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
              "RoleTypeKind": 6
            }
          ],
          "PrincipalId": 5
        },
        {
          "Member": {
            "Id": 12,
            "IsHiddenInUI": false,
            "LoginName": "i:0#.f|membership|user@contoso.onmicrosoft.com",
            "Title": "User",
            "PrincipalType": 1,
            "Email": "user@contoso.onmicrosoft.com",
            "Expiration": "",
            "IsEmailAuthenticationGuestUser": false,
            "IsShareByEmailGuestUser": false,
            "IsSiteAdmin": false,
            "UserId": {
              "NameId": "10032000f65ded70",
              "NameIdIssuer": "urn:federation:microsoftonline"
            },
            "UserPrincipalName": "user@contoso.onmicrosoft.com",
            "PrincipalTypeString": "User"
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
              "RoleTypeKind": 2
            }
          ],
          "PrincipalId": 12
        }
      ],
      "VersionPolicies": {
        "DefaultExpireAfterDays": 0,
        "DefaultTrimMode": 0
      },
      "HasUniqueRoleAssignments": true,
      "AllowContentTypes": true,
      "BaseTemplate": 109,
      "BaseType": 1,
      "ContentTypesEnabled": false,
      "CrawlNonDefaultViews": false,
      "Created": null,
      "CurrentChangeToken": null,
      "CustomActionElements": null,
      "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000",
      "DefaultItemOpenUseListSetting": false,
      "Description": "",
      "Direction": "none",
      "DocumentTemplateUrl": null,
      "DraftVersionVisibility": 0,
      "EnableAttachments": false,
      "EnableFolderCreation": true,
      "EnableMinorVersions": false,
      "EnableModeration": false,
      "EnableVersioning": false,
      "EntityTypeName": "Documents",
      "ExemptFromBlockDownloadOfNonViewableFiles": false,
      "FileSavePostProcessingEnabled": false,
      "ForceCheckout": false,
      "HasExternalDataSource": false,
      "Hidden": false,
      "Id": "14b2b6ed-0885-4814-bfd6-594737cc3ae3",
      "ImagePath": null,
      "ImageUrl": null,
      "IrmEnabled": false,
      "IrmExpire": false,
      "IrmReject": false,
      "IsApplicationList": false,
      "IsCatalog": false,
      "IsPrivate": false,
      "ItemCount": 69,
      "LastItemDeletedDate": null,
      "LastItemModifiedDate": null,
      "LastItemUserModifiedDate": null,
      "ListExperienceOptions": 0,
      "ListItemEntityTypeFullName": null,
      "MajorVersionLimit": 0,
      "MajorWithMinorVersionsLimit": 0,
      "MultipleDataList": false,
      "NoCrawl": false,
      "ParentWebPath": null,
      "ParentWebUrl": null,
      "ParserDisabled": false,
      "ServerTemplateCanCreateFolders": true,
      "TemplateFeatureId": null,
      "Title": "Documents"
    });

    await command.action(logger, {
      options: {
        id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
        webUrl: 'https://contoso.sharepoint.com',
        withPermissions: true
      }
    });

    assert(loggerLogSpy.calledWith({
      AllowContentTypes: true,
      BaseTemplate: 109,
      BaseType: 1,
      ContentTypesEnabled: false,
      CrawlNonDefaultViews: false,
      Created: null,
      CurrentChangeToken: null,
      CustomActionElements: null,
      DefaultContentApprovalWorkflowId: '00000000-0000-0000-0000-000000000000',
      DefaultItemOpenUseListSetting: false,
      Description: '',
      Direction: 'none',
      DocumentTemplateUrl: null,
      DraftVersionVisibility: 0,
      EnableAttachments: false,
      EnableFolderCreation: true,
      EnableMinorVersions: false,
      EnableModeration: false,
      EnableVersioning: false,
      EntityTypeName: 'Documents',
      ExemptFromBlockDownloadOfNonViewableFiles: false,
      FileSavePostProcessingEnabled: false,
      ForceCheckout: false,
      HasExternalDataSource: false,
      HasUniqueRoleAssignments: true,
      Hidden: false,
      Id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
      ImagePath: null,
      ImageUrl: null,
      IrmEnabled: false,
      IrmExpire: false,
      IrmReject: false,
      IsApplicationList: false,
      IsCatalog: false,
      IsPrivate: false,
      ItemCount: 69,
      LastItemDeletedDate: null,
      LastItemModifiedDate: null,
      LastItemUserModifiedDate: null,
      ListExperienceOptions: 0,
      ListItemEntityTypeFullName: null,
      MajorVersionLimit: 0,
      MajorWithMinorVersionsLimit: 0,
      MultipleDataList: false,
      NoCrawl: false,
      ParentWebPath: null,
      ParentWebUrl: null,
      ParserDisabled: false,
      VersionPolicies: {
        DefaultExpireAfterDays: 0,
        DefaultTrimMode: 0,
        DefaultTrimModeValue: "NoExpiration"
      },
      RoleAssignments: [
        {
          Member: {
            Id: 3,
            IsHiddenInUI: false,
            LoginName: "MySite Owners",
            Title: "MySite Owners",
            PrincipalType: 8,
            AllowMembersEditMembership: false,
            AllowRequestToJoinLeave: false,
            AutoAcceptRequestToJoinLeave: false,
            Description: null,
            OnlyAllowMembersViewMembership: false,
            OwnerTitle: "MySite Owners",
            RequestToJoinLeaveEmailSetting: "",
            PrincipalTypeString: "SharePointGroup"
          },
          RoleDefinitionBindings: [
            {
              BasePermissions: {
                High: "2147483647",
                Low: "4294967295"
              },
              Description: "Has full control.",
              Hidden: false,
              Id: 1073741829,
              Name: "Full Control",
              Order: 1,
              RoleTypeKind: 5
            }
          ],
          PrincipalId: 3
        },
        {
          Member: {
            Id: 4,
            IsHiddenInUI: false,
            LoginName: "MySite Visitors",
            Title: "MySite Visitors",
            PrincipalType: 8,
            AllowMembersEditMembership: false,
            AllowRequestToJoinLeave: false,
            AutoAcceptRequestToJoinLeave: false,
            Description: null,
            OnlyAllowMembersViewMembership: false,
            OwnerTitle: "MySite Owners",
            RequestToJoinLeaveEmailSetting: "",
            PrincipalTypeString: "SharePointGroup"
          },
          RoleDefinitionBindings: [
            {
              BasePermissions: {
                High: "176",
                Low: "138612833"
              },
              Description: "Can view pages and list items and download documents.",
              Hidden: false,
              Id: 1073741826,
              Name: "Read",
              Order: 128,
              RoleTypeKind: 2
            }
          ],
          PrincipalId: 4
        },
        {
          Member: {
            Id: 5,
            IsHiddenInUI: false,
            LoginName: "MySite Members",
            Title: "MySite Members",
            PrincipalType: 8,
            AllowMembersEditMembership: true,
            AllowRequestToJoinLeave: false,
            AutoAcceptRequestToJoinLeave: false,
            Description: null,
            OnlyAllowMembersViewMembership: false,
            OwnerTitle: "MySite Owners",
            RequestToJoinLeaveEmailSetting: "",
            PrincipalTypeString: "SharePointGroup"
          },
          RoleDefinitionBindings: [
            {
              BasePermissions: {
                High: "432",
                Low: "1011030767"
              },
              Description: "Can add, edit and delete lists; can view, add, update and delete list items and documents.",
              Hidden: false,
              Id: 1073741830,
              Name: "Edit",
              Order: 48,
              RoleTypeKind: 6
            }
          ],
          PrincipalId: 5
        },
        {
          Member: {
            Id: 12,
            IsHiddenInUI: false,
            LoginName: "i:0#.f|membership|user@contoso.onmicrosoft.com",
            Title: "User",
            PrincipalType: 1,
            Email: "user@contoso.onmicrosoft.com",
            Expiration: "",
            IsEmailAuthenticationGuestUser: false,
            IsShareByEmailGuestUser: false,
            IsSiteAdmin: false,
            UserId: {
              NameId: "10032000f65ded70",
              NameIdIssuer: "urn:federation:microsoftonline"
            },
            UserPrincipalName: "user@contoso.onmicrosoft.com",
            PrincipalTypeString: "User"
          },
          RoleDefinitionBindings: [
            {
              BasePermissions: {
                High: "176",
                Low: "138612833"
              },
              Description: "Can view pages and list items and download documents.",
              Hidden: false,
              Id: 1073741826,
              Name: "Read",
              Order: 128,
              RoleTypeKind: 2
            }
          ],
          PrincipalId: 12
        }
      ],
      ServerTemplateCanCreateFolders: true,
      TemplateFeatureId: null,
      Title: 'Documents'
    }));
  });

  it('retrieves details of list if list id, properties and withPermissions option is passed', async () => {
    sinon.stub(request, 'get').resolves(
      {
        "RoleAssignments": [
          {
            "Member": {
              "Id": 3,
              "IsHiddenInUI": false,
              "LoginName": "MySite Owners",
              "Title": "MySite Owners",
              "PrincipalType": 8,
              "AllowMembersEditMembership": false,
              "AllowRequestToJoinLeave": false,
              "AutoAcceptRequestToJoinLeave": false,
              "Description": null,
              "OnlyAllowMembersViewMembership": false,
              "OwnerTitle": "MySite Owners",
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
                "RoleTypeKind": 5
              }
            ],
            "PrincipalId": 3
          },
          {
            "Member": {
              "Id": 4,
              "IsHiddenInUI": false,
              "LoginName": "MySite Visitors",
              "Title": "MySite Visitors",
              "PrincipalType": 8,
              "AllowMembersEditMembership": false,
              "AllowRequestToJoinLeave": false,
              "AutoAcceptRequestToJoinLeave": false,
              "Description": null,
              "OnlyAllowMembersViewMembership": false,
              "OwnerTitle": "MySite Owners",
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
                "RoleTypeKind": 2
              }
            ],
            "PrincipalId": 4
          },
          {
            "Member": {
              "Id": 5,
              "IsHiddenInUI": false,
              "LoginName": "MySite Members",
              "Title": "MySite Members",
              "PrincipalType": 8,
              "AllowMembersEditMembership": true,
              "AllowRequestToJoinLeave": false,
              "AutoAcceptRequestToJoinLeave": false,
              "Description": null,
              "OnlyAllowMembersViewMembership": false,
              "OwnerTitle": "MySite Owners",
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
                "RoleTypeKind": 6
              }
            ],
            "PrincipalId": 5
          },
          {
            "Member": {
              "Id": 12,
              "IsHiddenInUI": false,
              "LoginName": "i:0#.f|membership|user@contoso.onmicrosoft.com",
              "Title": "User",
              "PrincipalType": 1,
              "Email": "user@contoso.onmicrosoft.com",
              "Expiration": "",
              "IsEmailAuthenticationGuestUser": false,
              "IsShareByEmailGuestUser": false,
              "IsSiteAdmin": false,
              "UserId": {
                "NameId": "10032000f65ded70",
                "NameIdIssuer": "urn:federation:microsoftonline"
              },
              "UserPrincipalName": "user@contoso.onmicrosoft.com",
              "PrincipalTypeString": "User"
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
                "RoleTypeKind": 2
              }
            ],
            "PrincipalId": 12
          }
        ],
        "HasUniqueRoleAssignments": true,
        "Id": "14b2b6ed-0885-4814-bfd6-594737cc3ae3",
        "Title": "Documents"
      }
    );

    await command.action(logger, {
      options: {
        id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
        webUrl: 'https://contoso.sharepoint.com',
        properties: 'Title,Id',
        withPermissions: true
      }
    });

    assert(loggerLogSpy.calledWith({
      HasUniqueRoleAssignments: true,
      Id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
      RoleAssignments: [
        {
          Member: {
            Id: 3,
            IsHiddenInUI: false,
            LoginName: "MySite Owners",
            Title: "MySite Owners",
            PrincipalType: 8,
            AllowMembersEditMembership: false,
            AllowRequestToJoinLeave: false,
            AutoAcceptRequestToJoinLeave: false,
            Description: null,
            OnlyAllowMembersViewMembership: false,
            OwnerTitle: "MySite Owners",
            RequestToJoinLeaveEmailSetting: "",
            PrincipalTypeString: "SharePointGroup"
          },
          RoleDefinitionBindings: [
            {
              BasePermissions: {
                High: "2147483647",
                Low: "4294967295"
              },
              Description: "Has full control.",
              Hidden: false,
              Id: 1073741829,
              Name: "Full Control",
              Order: 1,
              RoleTypeKind: 5
            }
          ],
          PrincipalId: 3
        },
        {
          Member: {
            Id: 4,
            IsHiddenInUI: false,
            LoginName: "MySite Visitors",
            Title: "MySite Visitors",
            PrincipalType: 8,
            AllowMembersEditMembership: false,
            AllowRequestToJoinLeave: false,
            AutoAcceptRequestToJoinLeave: false,
            Description: null,
            OnlyAllowMembersViewMembership: false,
            OwnerTitle: "MySite Owners",
            RequestToJoinLeaveEmailSetting: "",
            PrincipalTypeString: "SharePointGroup"
          },
          RoleDefinitionBindings: [
            {
              BasePermissions: {
                High: "176",
                Low: "138612833"
              },
              Description: "Can view pages and list items and download documents.",
              Hidden: false,
              Id: 1073741826,
              Name: "Read",
              Order: 128,
              RoleTypeKind: 2
            }
          ],
          PrincipalId: 4
        },
        {
          Member: {
            Id: 5,
            IsHiddenInUI: false,
            LoginName: "MySite Members",
            Title: "MySite Members",
            PrincipalType: 8,
            AllowMembersEditMembership: true,
            AllowRequestToJoinLeave: false,
            AutoAcceptRequestToJoinLeave: false,
            Description: null,
            OnlyAllowMembersViewMembership: false,
            OwnerTitle: "MySite Owners",
            RequestToJoinLeaveEmailSetting: "",
            PrincipalTypeString: "SharePointGroup"
          },
          RoleDefinitionBindings: [
            {
              BasePermissions: {
                High: "432",
                Low: "1011030767"
              },
              Description: "Can add, edit and delete lists; can view, add, update and delete list items and documents.",
              Hidden: false,
              Id: 1073741830,
              Name: "Edit",
              Order: 48,
              RoleTypeKind: 6
            }
          ],
          PrincipalId: 5
        },
        {
          Member: {
            Id: 12,
            IsHiddenInUI: false,
            LoginName: "i:0#.f|membership|user@contoso.onmicrosoft.com",
            Title: "User",
            PrincipalType: 1,
            Email: "user@contoso.onmicrosoft.com",
            Expiration: "",
            IsEmailAuthenticationGuestUser: false,
            IsShareByEmailGuestUser: false,
            IsSiteAdmin: false,
            UserId: {
              NameId: "10032000f65ded70",
              NameIdIssuer: "urn:federation:microsoftonline"
            },
            UserPrincipalName: "user@contoso.onmicrosoft.com",
            PrincipalTypeString: "User"
          },
          RoleDefinitionBindings: [
            {
              BasePermissions: {
                High: "176",
                Low: "138612833"
              },
              Description: "Can view pages and list items and download documents.",
              Hidden: false,
              Id: 1073741826,
              Name: "Read",
              Order: 128,
              RoleTypeKind: 2
            }
          ],
          PrincipalId: 12
        }
      ],
      Title: 'Documents'
    }));
  });

  it('retrieves details of list if list title, properties and withPermissions option is passed', async () => {
    sinon.stub(request, 'get').resolves({
      "RoleAssignments": [
        {
          "Member": {
            "Id": 3,
            "IsHiddenInUI": false,
            "LoginName": "MySite Owners",
            "Title": "MySite Owners",
            "PrincipalType": 8,
            "AllowMembersEditMembership": false,
            "AllowRequestToJoinLeave": false,
            "AutoAcceptRequestToJoinLeave": false,
            "Description": null,
            "OnlyAllowMembersViewMembership": false,
            "OwnerTitle": "MySite Owners",
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
              "RoleTypeKind": 5
            }
          ],
          "PrincipalId": 3
        },
        {
          "Member": {
            "Id": 4,
            "IsHiddenInUI": false,
            "LoginName": "MySite Visitors",
            "Title": "MySite Visitors",
            "PrincipalType": 8,
            "AllowMembersEditMembership": false,
            "AllowRequestToJoinLeave": false,
            "AutoAcceptRequestToJoinLeave": false,
            "Description": null,
            "OnlyAllowMembersViewMembership": false,
            "OwnerTitle": "MySite Owners",
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
              "RoleTypeKind": 2
            }
          ],
          "PrincipalId": 4
        },
        {
          "Member": {
            "Id": 5,
            "IsHiddenInUI": false,
            "LoginName": "MySite Members",
            "Title": "MySite Members",
            "PrincipalType": 8,
            "AllowMembersEditMembership": true,
            "AllowRequestToJoinLeave": false,
            "AutoAcceptRequestToJoinLeave": false,
            "Description": null,
            "OnlyAllowMembersViewMembership": false,
            "OwnerTitle": "MySite Owners",
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
              "RoleTypeKind": 6
            }
          ],
          "PrincipalId": 5
        },
        {
          "Member": {
            "Id": 12,
            "IsHiddenInUI": false,
            "LoginName": "i:0#.f|membership|user@contoso.onmicrosoft.com",
            "Title": "User",
            "PrincipalType": 1,
            "Email": "user@contoso.onmicrosoft.com",
            "Expiration": "",
            "IsEmailAuthenticationGuestUser": false,
            "IsShareByEmailGuestUser": false,
            "IsSiteAdmin": false,
            "UserId": {
              "NameId": "10032000f65ded70",
              "NameIdIssuer": "urn:federation:microsoftonline"
            },
            "UserPrincipalName": "user@contoso.onmicrosoft.com",
            "PrincipalTypeString": "User"
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
              "RoleTypeKind": 2
            }
          ],
          "PrincipalId": 12
        }
      ],
      "HasUniqueRoleAssignments": true,
      "Id": "14b2b6ed-0885-4814-bfd6-594737cc3ae3",
      "Title": "Documents"
    });

    await command.action(logger, {
      options: {
        title: 'Documents',
        webUrl: 'https://contoso.sharepoint.com',
        properties: 'Title,Id',
        withPermissions: true
      }
    });

    assert(loggerLogSpy.calledWith({
      HasUniqueRoleAssignments: true,
      Id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
      RoleAssignments: [
        {
          Member: {
            Id: 3,
            IsHiddenInUI: false,
            LoginName: "MySite Owners",
            Title: "MySite Owners",
            PrincipalType: 8,
            AllowMembersEditMembership: false,
            AllowRequestToJoinLeave: false,
            AutoAcceptRequestToJoinLeave: false,
            Description: null,
            OnlyAllowMembersViewMembership: false,
            OwnerTitle: "MySite Owners",
            RequestToJoinLeaveEmailSetting: "",
            PrincipalTypeString: "SharePointGroup"
          },
          RoleDefinitionBindings: [
            {
              BasePermissions: {
                High: "2147483647",
                Low: "4294967295"
              },
              Description: "Has full control.",
              Hidden: false,
              Id: 1073741829,
              Name: "Full Control",
              Order: 1,
              RoleTypeKind: 5
            }
          ],
          PrincipalId: 3
        },
        {
          Member: {
            Id: 4,
            IsHiddenInUI: false,
            LoginName: "MySite Visitors",
            Title: "MySite Visitors",
            PrincipalType: 8,
            AllowMembersEditMembership: false,
            AllowRequestToJoinLeave: false,
            AutoAcceptRequestToJoinLeave: false,
            Description: null,
            OnlyAllowMembersViewMembership: false,
            OwnerTitle: "MySite Owners",
            RequestToJoinLeaveEmailSetting: "",
            PrincipalTypeString: "SharePointGroup"
          },
          RoleDefinitionBindings: [
            {
              BasePermissions: {
                High: "176",
                Low: "138612833"
              },
              Description: "Can view pages and list items and download documents.",
              Hidden: false,
              Id: 1073741826,
              Name: "Read",
              Order: 128,
              RoleTypeKind: 2
            }
          ],
          PrincipalId: 4
        },
        {
          Member: {
            Id: 5,
            IsHiddenInUI: false,
            LoginName: "MySite Members",
            Title: "MySite Members",
            PrincipalType: 8,
            AllowMembersEditMembership: true,
            AllowRequestToJoinLeave: false,
            AutoAcceptRequestToJoinLeave: false,
            Description: null,
            OnlyAllowMembersViewMembership: false,
            OwnerTitle: "MySite Owners",
            RequestToJoinLeaveEmailSetting: "",
            PrincipalTypeString: "SharePointGroup"
          },
          RoleDefinitionBindings: [
            {
              BasePermissions: {
                High: "432",
                Low: "1011030767"
              },
              Description: "Can add, edit and delete lists; can view, add, update and delete list items and documents.",
              Hidden: false,
              Id: 1073741830,
              Name: "Edit",
              Order: 48,
              RoleTypeKind: 6
            }
          ],
          PrincipalId: 5
        },
        {
          Member: {
            Id: 12,
            IsHiddenInUI: false,
            LoginName: "i:0#.f|membership|user@contoso.onmicrosoft.com",
            Title: "User",
            PrincipalType: 1,
            Email: "user@contoso.onmicrosoft.com",
            Expiration: "",
            IsEmailAuthenticationGuestUser: false,
            IsShareByEmailGuestUser: false,
            IsSiteAdmin: false,
            UserId: {
              NameId: "10032000f65ded70",
              NameIdIssuer: "urn:federation:microsoftonline"
            },
            UserPrincipalName: "user@contoso.onmicrosoft.com",
            PrincipalTypeString: "User"
          },
          RoleDefinitionBindings: [
            {
              BasePermissions: {
                High: "176",
                Low: "138612833"
              },
              Description: "Can view pages and list items and download documents.",
              Hidden: false,
              Id: 1073741826,
              Name: "Read",
              Order: 128,
              RoleTypeKind: 2
            }
          ],
          PrincipalId: 12
        }
      ],
      Title: 'Documents'
    }));
  });

  it('retrieves details of list with correct expand properties', async () => {
    sinon.stub(request, 'get').resolves(
      {
        "RootFolder": {
          "ServerRelativeUrl": "/Lists/TestBatchList"
        }
      }
    );

    await command.action(logger, {
      options: {
        title: 'Documents',
        webUrl: 'https://contoso.sharepoint.com',
        properties: 'RootFolder/ServerRelativeUrl'
      }
    });

    assert(loggerLogSpy.calledWith({
      RootFolder: {
        ServerRelativeUrl: "/Lists/TestBatchList"
      }
    }));
  });

  it('command correctly handles list get reject request', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        throw error;
      }

      throw 'Invalid request';
    });

    const actionTitle: string = 'Documents';

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        title: actionTitle,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });

  it('retrieves the default list in the specified site by providing a webUrl', async () => {
    const defaultSiteList = { ...listResponse, BaseTemplate: 101, ParentWebUrl: "/", ListItemEntityTypeFullName: "SP.Data.Shared_x0020_DocumentsItem" };

    sinon.stub(request, 'get').callsFake(async (opts: CliRequestOptions) => {
      if (opts.url?.includes('https://contoso.sharepoint.com/_api/web/DefaultDocumentLibrary')) {
        return defaultSiteList;
      }
      else {
        throw new Error(`Invalid request ${opts.url}`);
      }
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com'
      }
    });

    assert(loggerLogSpy.calledWithMatch({
      BaseTemplate: 101,
      ParentWebUrl: "/",
      ListItemEntityTypeFullName: "SP.Data.Shared_x0020_DocumentsItem"
    }));
  });
});