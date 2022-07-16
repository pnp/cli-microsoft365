import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./list-get');

describe(commands.LIST_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LIST_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves and prints all details of list if title option is passed', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve(
        {
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
        }
      );
    });

    command.action(logger, {
      options: {
        debug: true,
        title: 'Documents',
        webUrl: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
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
          ServerTemplateCanCreateFolders: true,
          TemplateFeatureId: null,
          Title: 'Documents'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves details of list if title and properties option is passed', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve(
        {
          "Id": "14b2b6ed-0885-4814-bfd6-594737cc3ae3",
          "Title": "Documents"
        }
      );
    });

    command.action(logger, {
      options: {
        debug: true,
        title: 'Documents',
        webUrl: 'https://contoso.sharepoint.com',
        properties:'Title,Id'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          Id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
          Title: 'Documents'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('retrieves details of list if list id and properties option is passed', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve(
        {
          "Id": "14b2b6ed-0885-4814-bfd6-594737cc3ae3",
          "Title": "Documents"
        }
      );
    });

    command.action(logger, {
      options: {
        debug: true,
        id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
        webUrl: 'https://contoso.sharepoint.com',
        properties:'Title,Id'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          Id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
          Title: 'Documents'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves details of list if list id and withPermissions option is passed', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve(
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
        }
      );
    });

    command.action(logger, {
      options: {
        debug: false,
        id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
        webUrl: 'https://contoso.sharepoint.com',
        withPermissions: true
      }
    }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('retrieves details of list if list id, properties and withPermissions option is passed', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve(
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
    });

    command.action(logger, {
      options: {
        debug: false,
        id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
        webUrl: 'https://contoso.sharepoint.com',
        properties:'Title,Id',
        withPermissions: true
      }
    }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves details of list if list title, properties and withPermissions option is passed', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve(
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
    });

    command.action(logger, {
      options: {
        debug: false,
        title: 'Documents',
        webUrl: 'https://contoso.sharepoint.com',
        properties:'Title,Id',
        withPermissions: true
      }
    }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('command correctly handles list get reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    const actionTitle: string = 'Documents';

    command.action(logger, {
      options: {
        debug: true,
        title: actionTitle,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when id option is passed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid') > -1) {
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    command.action(logger, {
      options: {
        debug: false,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }, () => {

      try {
        assert(1 === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('supports specifying URL', () => {
    const options = command.options();
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('supports specifying properties', () => {
    const options = command.options();
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--properties') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  
  it('fails validation if both id and title options are not passed', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails validation if both id and title options are passed', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', title: 'Documents' } });
    assert.notStrictEqual(actual, true);
  });
});