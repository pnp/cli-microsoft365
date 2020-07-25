import commands from '../../commands';
import Command, { CommandValidate, CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./list-sitescript-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_SITESCRIPT_GET, () => {
  let log: any[];
  let cmdInstance: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.LIST_SITESCRIPT_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('extracts the site script from the given list if title option is passed (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/Microsoft_SharePoint_Utilities_WebTemplateExtensions_SiteScriptUtility_GetSiteScriptFromList`) > -1) {
        return Promise.resolve({
          value: {
            "actions": [
              {
                "verb": "createSPList",
                "listName": "MyLibrary",
                "templateType": 101,
                "subactions": [
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{47b1b86f-9f8a-4dbe-a75e-ca5d9b0f566c}\" Type=\"URL\" Name=\"_ShortcutUrl\" DisplayName=\"Shortcut URL\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUrl\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{2662ad77-2410-4938-b01c-e5e43321bad4}\" Type=\"Guid\" Name=\"_ShortcutSiteId\" DisplayName=\"Shortcut Site Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutSiteId\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{e2a3861f-c216-47d7-820f-7cb638862ab2}\" Type=\"Guid\" Name=\"_ShortcutWebId\" DisplayName=\"Shortcut Web Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutWebId\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{e8fea999-553d-4f45-be52-d941627e9fe5}\" Type=\"Guid\" Name=\"_ShortcutUniqueId\" DisplayName=\"Shortcut Unique Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUniqueId\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field DisplayName=\"MyText\" Format=\"Dropdown\" MaxLength=\"255\" Title=\"MyText\" Type=\"Text\" ID=\"{dbd0f8fa-5131-44ed-b7a1-23bfffc50ac8}\" StaticName=\"MyText\" Name=\"MyText\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field DisplayName=\"MyDate\" FriendlyDisplayFormat=\"Disabled\" Format=\"DateTime\" Title=\"MyDate\" Type=\"DateTime\" ID=\"{f98a4e28-5fb3-4737-9a24-3ad533552bf5}\" StaticName=\"MyDate\" Name=\"MyDate\"><Default>[today]</Default></Field>"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field Decimals=\"2\" DisplayName=\"MyNumber\" Format=\"Dropdown\" Percentage=\"FALSE\" Title=\"MyNumber\" Type=\"Number\" ID=\"{496aa48c-0cf7-4990-be49-d373aa327e0c}\" StaticName=\"MyNumber\" Name=\"MyNumber\"><Default>100</Default></Field>"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{e52012a0-51eb-4c0c-8dfb-9b8a0ebedcb6}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"Combine\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Merge\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Combine\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"EncodedAbsUrl\" /><FieldRef Name=\"TemplateUrl\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkCombine\" type=\"CHECKBOX\" title=\"Merge]]\" href=\"]]></HTML><Field Name=\"EncodedAbsUrl\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkUrl\" type=\"HIDDEN\" href=\"]]></HTML><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkProgID\" type=\"HIDDEN\" href=\"]]></HTML><MapToControl><HTML>|</HTML><GetFileExtension><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /></GetFileExtension></MapToControl><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{5d36727b-bcb2-47d2-a231-1f0bc63b7439}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"RepairDocument\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Relink\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"RepairDocument\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"ID\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkRepair\" type=\"CHECKBOX\" title=\"Relink\" docid=\"]]></HTML><Field Name=\"ID\" /><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                  },
                  {
                    "verb": "addSPView",
                    "name": "All Documents",
                    "viewFields": [
                      "DocIcon",
                      "LinkFilename",
                      "MyText",
                      "MyDate",
                      "MyNumber"
                    ],
                    "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                    "rowLimit": 30,
                    "isPaged": true,
                    "makeDefault": true,
                    "formatterJSON": ""
                  }
                ]
              },
              {
                "verb": "addNavLink",
                "url": "MyLibrary/Forms/AllItems.aspx",
                "displayName": "MyLibrary",
                "isWebRelative": true
              }
            ]
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists/GetByTitle('MyLibrary')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" }, "AllowContentTypes": true, "BaseTemplate": 101, "BaseType": 1, "ContentTypesEnabled": false, "CrawlNonDefaultViews": false, "Created": "2019-01-11T10:03:19Z", "CurrentChangeToken": { "StringValue": "1;3;fb4b0cf8-c006-4802-a1ea-57e0e4852188;636827981522200000;96826061" }, "CustomActionElements": { "Items": [{ "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vdw", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdxFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdx", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdmFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdm", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenXsn", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "xsn", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XsnLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument2", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.2", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument3", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.3", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument4", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.4", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }] }, "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000", "DefaultItemOpenUseListSetting": false, "Description": "", "Direction": "none", "DisableGridEditing": false, "DocumentTemplateUrl": "/sites/team1/MyLibrary/Forms/template.dotx", "DraftVersionVisibility": 0, "EnableAttachments": false, "EnableFolderCreation": true, "EnableMinorVersions": false, "EnableModeration": false, "EnableRequestSignOff": true, "EnableVersioning": true, "EntityTypeName": "MyLibrary", "ExemptFromBlockDownloadOfNonViewableFiles": false, "FileSavePostProcessingEnabled": false, "ForceCheckout": false, "HasExternalDataSource": false, "Hidden": false, "Id": "fb4b0cf8-c006-4802-a1ea-57e0e4852188", "ImagePath": { "DecodedUrl": "/_layouts/15/images/itdl.png?rev=45" }, "ImageUrl": "/_layouts/15/images/itdl.png?rev=45", "IrmEnabled": false, "IrmExpire": false, "IrmReject": false, "IsApplicationList": false, "IsCatalog": false, "IsPrivate": false, "ItemCount": 0, "LastItemDeletedDate": "2019-01-11T10:03:19Z", "LastItemModifiedDate": "2019-01-11T10:04:15Z", "LastItemUserModifiedDate": "2019-01-11T10:03:19Z", "ListExperienceOptions": 0, "ListItemEntityTypeFullName": "SP.Data.MyLibraryItem", "MajorVersionLimit": 500, "MajorWithMinorVersionsLimit": 0, "MultipleDataList": false, "NoCrawl": false, "ParentWebPath": { "DecodedUrl": "/sites/team1" }, "ParentWebUrl": "/sites/team1", "ParserDisabled": false, "ServerTemplateCanCreateFolders": true, "TemplateFeatureId": "00bfea71-e717-4e80-aa17-d0c71b360101", "Title": "MyLibrary" }
        );
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary'
      }
    }, () => {
      try {
        const expected = {
          "actions": [
            {
              "verb": "createSPList",
              "listName": "MyLibrary",
              "templateType": 101,
              "subactions": [
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{47b1b86f-9f8a-4dbe-a75e-ca5d9b0f566c}\" Type=\"URL\" Name=\"_ShortcutUrl\" DisplayName=\"Shortcut URL\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUrl\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{2662ad77-2410-4938-b01c-e5e43321bad4}\" Type=\"Guid\" Name=\"_ShortcutSiteId\" DisplayName=\"Shortcut Site Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutSiteId\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{e2a3861f-c216-47d7-820f-7cb638862ab2}\" Type=\"Guid\" Name=\"_ShortcutWebId\" DisplayName=\"Shortcut Web Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutWebId\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{e8fea999-553d-4f45-be52-d941627e9fe5}\" Type=\"Guid\" Name=\"_ShortcutUniqueId\" DisplayName=\"Shortcut Unique Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUniqueId\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field DisplayName=\"MyText\" Format=\"Dropdown\" MaxLength=\"255\" Title=\"MyText\" Type=\"Text\" ID=\"{dbd0f8fa-5131-44ed-b7a1-23bfffc50ac8}\" StaticName=\"MyText\" Name=\"MyText\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field DisplayName=\"MyDate\" FriendlyDisplayFormat=\"Disabled\" Format=\"DateTime\" Title=\"MyDate\" Type=\"DateTime\" ID=\"{f98a4e28-5fb3-4737-9a24-3ad533552bf5}\" StaticName=\"MyDate\" Name=\"MyDate\"><Default>[today]</Default></Field>"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field Decimals=\"2\" DisplayName=\"MyNumber\" Format=\"Dropdown\" Percentage=\"FALSE\" Title=\"MyNumber\" Type=\"Number\" ID=\"{496aa48c-0cf7-4990-be49-d373aa327e0c}\" StaticName=\"MyNumber\" Name=\"MyNumber\"><Default>100</Default></Field>"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{e52012a0-51eb-4c0c-8dfb-9b8a0ebedcb6}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"Combine\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Merge\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Combine\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"EncodedAbsUrl\" /><FieldRef Name=\"TemplateUrl\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkCombine\" type=\"CHECKBOX\" title=\"Merge]]\" href=\"]]></HTML><Field Name=\"EncodedAbsUrl\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkUrl\" type=\"HIDDEN\" href=\"]]></HTML><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkProgID\" type=\"HIDDEN\" href=\"]]></HTML><MapToControl><HTML>|</HTML><GetFileExtension><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /></GetFileExtension></MapToControl><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{5d36727b-bcb2-47d2-a231-1f0bc63b7439}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"RepairDocument\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Relink\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"RepairDocument\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"ID\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkRepair\" type=\"CHECKBOX\" title=\"Relink\" docid=\"]]></HTML><Field Name=\"ID\" /><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                },
                {
                  "verb": "addSPView",
                  "name": "All Documents",
                  "viewFields": [
                    "DocIcon",
                    "LinkFilename",
                    "MyText",
                    "MyDate",
                    "MyNumber"
                  ],
                  "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                  "rowLimit": 30,
                  "isPaged": true,
                  "makeDefault": true,
                  "formatterJSON": ""
                }
              ]
            },
            {
              "verb": "addNavLink",
              "url": "MyLibrary/Forms/AllItems.aspx",
              "displayName": "MyLibrary",
              "isWebRelative": true
            }
          ]
        };
        const actual = log[log.length - 1];
        assert.strictEqual(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('extracts the site script from the given list if title option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/Microsoft_SharePoint_Utilities_WebTemplateExtensions_SiteScriptUtility_GetSiteScriptFromList`) > -1) {
        return Promise.resolve({
          value: {
            "actions": [
              {
                "verb": "createSPList",
                "listName": "MyLibrary",
                "templateType": 101,
                "subactions": [
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{47b1b86f-9f8a-4dbe-a75e-ca5d9b0f566c}\" Type=\"URL\" Name=\"_ShortcutUrl\" DisplayName=\"Shortcut URL\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUrl\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{2662ad77-2410-4938-b01c-e5e43321bad4}\" Type=\"Guid\" Name=\"_ShortcutSiteId\" DisplayName=\"Shortcut Site Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutSiteId\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{e2a3861f-c216-47d7-820f-7cb638862ab2}\" Type=\"Guid\" Name=\"_ShortcutWebId\" DisplayName=\"Shortcut Web Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutWebId\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{e8fea999-553d-4f45-be52-d941627e9fe5}\" Type=\"Guid\" Name=\"_ShortcutUniqueId\" DisplayName=\"Shortcut Unique Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUniqueId\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field DisplayName=\"MyText\" Format=\"Dropdown\" MaxLength=\"255\" Title=\"MyText\" Type=\"Text\" ID=\"{dbd0f8fa-5131-44ed-b7a1-23bfffc50ac8}\" StaticName=\"MyText\" Name=\"MyText\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field DisplayName=\"MyDate\" FriendlyDisplayFormat=\"Disabled\" Format=\"DateTime\" Title=\"MyDate\" Type=\"DateTime\" ID=\"{f98a4e28-5fb3-4737-9a24-3ad533552bf5}\" StaticName=\"MyDate\" Name=\"MyDate\"><Default>[today]</Default></Field>"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field Decimals=\"2\" DisplayName=\"MyNumber\" Format=\"Dropdown\" Percentage=\"FALSE\" Title=\"MyNumber\" Type=\"Number\" ID=\"{496aa48c-0cf7-4990-be49-d373aa327e0c}\" StaticName=\"MyNumber\" Name=\"MyNumber\"><Default>100</Default></Field>"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{e52012a0-51eb-4c0c-8dfb-9b8a0ebedcb6}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"Combine\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Merge\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Combine\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"EncodedAbsUrl\" /><FieldRef Name=\"TemplateUrl\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkCombine\" type=\"CHECKBOX\" title=\"Merge]]\" href=\"]]></HTML><Field Name=\"EncodedAbsUrl\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkUrl\" type=\"HIDDEN\" href=\"]]></HTML><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkProgID\" type=\"HIDDEN\" href=\"]]></HTML><MapToControl><HTML>|</HTML><GetFileExtension><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /></GetFileExtension></MapToControl><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{5d36727b-bcb2-47d2-a231-1f0bc63b7439}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"RepairDocument\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Relink\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"RepairDocument\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"ID\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkRepair\" type=\"CHECKBOX\" title=\"Relink\" docid=\"]]></HTML><Field Name=\"ID\" /><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                  },
                  {
                    "verb": "addSPView",
                    "name": "All Documents",
                    "viewFields": [
                      "DocIcon",
                      "LinkFilename",
                      "MyText",
                      "MyDate",
                      "MyNumber"
                    ],
                    "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                    "rowLimit": 30,
                    "isPaged": true,
                    "makeDefault": true,
                    "formatterJSON": ""
                  }
                ]
              },
              {
                "verb": "addNavLink",
                "url": "MyLibrary/Forms/AllItems.aspx",
                "displayName": "MyLibrary",
                "isWebRelative": true
              }
            ]
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists/GetByTitle('MyLibrary')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" }, "AllowContentTypes": true, "BaseTemplate": 101, "BaseType": 1, "ContentTypesEnabled": false, "CrawlNonDefaultViews": false, "Created": "2019-01-11T10:03:19Z", "CurrentChangeToken": { "StringValue": "1;3;fb4b0cf8-c006-4802-a1ea-57e0e4852188;636827981522200000;96826061" }, "CustomActionElements": { "Items": [{ "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vdw", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdxFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdx", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdmFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdm", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenXsn", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "xsn", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XsnLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument2", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.2", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument3", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.3", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument4", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.4", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }] }, "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000", "DefaultItemOpenUseListSetting": false, "Description": "", "Direction": "none", "DisableGridEditing": false, "DocumentTemplateUrl": "/sites/team1/MyLibrary/Forms/template.dotx", "DraftVersionVisibility": 0, "EnableAttachments": false, "EnableFolderCreation": true, "EnableMinorVersions": false, "EnableModeration": false, "EnableRequestSignOff": true, "EnableVersioning": true, "EntityTypeName": "MyLibrary", "ExemptFromBlockDownloadOfNonViewableFiles": false, "FileSavePostProcessingEnabled": false, "ForceCheckout": false, "HasExternalDataSource": false, "Hidden": false, "Id": "fb4b0cf8-c006-4802-a1ea-57e0e4852188", "ImagePath": { "DecodedUrl": "/_layouts/15/images/itdl.png?rev=45" }, "ImageUrl": "/_layouts/15/images/itdl.png?rev=45", "IrmEnabled": false, "IrmExpire": false, "IrmReject": false, "IsApplicationList": false, "IsCatalog": false, "IsPrivate": false, "ItemCount": 0, "LastItemDeletedDate": "2019-01-11T10:03:19Z", "LastItemModifiedDate": "2019-01-11T10:04:15Z", "LastItemUserModifiedDate": "2019-01-11T10:03:19Z", "ListExperienceOptions": 0, "ListItemEntityTypeFullName": "SP.Data.MyLibraryItem", "MajorVersionLimit": 500, "MajorWithMinorVersionsLimit": 0, "MultipleDataList": false, "NoCrawl": false, "ParentWebPath": { "DecodedUrl": "/sites/team1" }, "ParentWebUrl": "/sites/team1", "ParserDisabled": false, "ServerTemplateCanCreateFolders": true, "TemplateFeatureId": "00bfea71-e717-4e80-aa17-d0c71b360101", "Title": "MyLibrary" }
        );
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary'
      }
    }, () => {
      try {
        const expected = {
          "actions": [
            {
              "verb": "createSPList",
              "listName": "MyLibrary",
              "templateType": 101,
              "subactions": [
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{47b1b86f-9f8a-4dbe-a75e-ca5d9b0f566c}\" Type=\"URL\" Name=\"_ShortcutUrl\" DisplayName=\"Shortcut URL\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUrl\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{2662ad77-2410-4938-b01c-e5e43321bad4}\" Type=\"Guid\" Name=\"_ShortcutSiteId\" DisplayName=\"Shortcut Site Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutSiteId\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{e2a3861f-c216-47d7-820f-7cb638862ab2}\" Type=\"Guid\" Name=\"_ShortcutWebId\" DisplayName=\"Shortcut Web Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutWebId\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{e8fea999-553d-4f45-be52-d941627e9fe5}\" Type=\"Guid\" Name=\"_ShortcutUniqueId\" DisplayName=\"Shortcut Unique Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUniqueId\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field DisplayName=\"MyText\" Format=\"Dropdown\" MaxLength=\"255\" Title=\"MyText\" Type=\"Text\" ID=\"{dbd0f8fa-5131-44ed-b7a1-23bfffc50ac8}\" StaticName=\"MyText\" Name=\"MyText\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field DisplayName=\"MyDate\" FriendlyDisplayFormat=\"Disabled\" Format=\"DateTime\" Title=\"MyDate\" Type=\"DateTime\" ID=\"{f98a4e28-5fb3-4737-9a24-3ad533552bf5}\" StaticName=\"MyDate\" Name=\"MyDate\"><Default>[today]</Default></Field>"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field Decimals=\"2\" DisplayName=\"MyNumber\" Format=\"Dropdown\" Percentage=\"FALSE\" Title=\"MyNumber\" Type=\"Number\" ID=\"{496aa48c-0cf7-4990-be49-d373aa327e0c}\" StaticName=\"MyNumber\" Name=\"MyNumber\"><Default>100</Default></Field>"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{e52012a0-51eb-4c0c-8dfb-9b8a0ebedcb6}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"Combine\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Merge\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Combine\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"EncodedAbsUrl\" /><FieldRef Name=\"TemplateUrl\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkCombine\" type=\"CHECKBOX\" title=\"Merge]]\" href=\"]]></HTML><Field Name=\"EncodedAbsUrl\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkUrl\" type=\"HIDDEN\" href=\"]]></HTML><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkProgID\" type=\"HIDDEN\" href=\"]]></HTML><MapToControl><HTML>|</HTML><GetFileExtension><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /></GetFileExtension></MapToControl><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{5d36727b-bcb2-47d2-a231-1f0bc63b7439}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"RepairDocument\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Relink\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"RepairDocument\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"ID\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkRepair\" type=\"CHECKBOX\" title=\"Relink\" docid=\"]]></HTML><Field Name=\"ID\" /><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                },
                {
                  "verb": "addSPView",
                  "name": "All Documents",
                  "viewFields": [
                    "DocIcon",
                    "LinkFilename",
                    "MyText",
                    "MyDate",
                    "MyNumber"
                  ],
                  "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                  "rowLimit": 30,
                  "isPaged": true,
                  "makeDefault": true,
                  "formatterJSON": ""
                }
              ]
            },
            {
              "verb": "addNavLink",
              "url": "MyLibrary/Forms/AllItems.aspx",
              "displayName": "MyLibrary",
              "isWebRelative": true
            }
          ]
        };
        const actual = log[log.length - 1];
        assert.strictEqual(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('extracts the site script from the given list if list id option is passed (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/Microsoft_SharePoint_Utilities_WebTemplateExtensions_SiteScriptUtility_GetSiteScriptFromList`) > -1) {
        return Promise.resolve({
          value: {
            "actions": [
              {
                "verb": "createSPList",
                "listName": "MyLibrary",
                "templateType": 101,
                "subactions": [
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{47b1b86f-9f8a-4dbe-a75e-ca5d9b0f566c}\" Type=\"URL\" Name=\"_ShortcutUrl\" DisplayName=\"Shortcut URL\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUrl\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{2662ad77-2410-4938-b01c-e5e43321bad4}\" Type=\"Guid\" Name=\"_ShortcutSiteId\" DisplayName=\"Shortcut Site Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutSiteId\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{e2a3861f-c216-47d7-820f-7cb638862ab2}\" Type=\"Guid\" Name=\"_ShortcutWebId\" DisplayName=\"Shortcut Web Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutWebId\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{e8fea999-553d-4f45-be52-d941627e9fe5}\" Type=\"Guid\" Name=\"_ShortcutUniqueId\" DisplayName=\"Shortcut Unique Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUniqueId\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field DisplayName=\"MyText\" Format=\"Dropdown\" MaxLength=\"255\" Title=\"MyText\" Type=\"Text\" ID=\"{dbd0f8fa-5131-44ed-b7a1-23bfffc50ac8}\" StaticName=\"MyText\" Name=\"MyText\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field DisplayName=\"MyDate\" FriendlyDisplayFormat=\"Disabled\" Format=\"DateTime\" Title=\"MyDate\" Type=\"DateTime\" ID=\"{f98a4e28-5fb3-4737-9a24-3ad533552bf5}\" StaticName=\"MyDate\" Name=\"MyDate\"><Default>[today]</Default></Field>"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field Decimals=\"2\" DisplayName=\"MyNumber\" Format=\"Dropdown\" Percentage=\"FALSE\" Title=\"MyNumber\" Type=\"Number\" ID=\"{496aa48c-0cf7-4990-be49-d373aa327e0c}\" StaticName=\"MyNumber\" Name=\"MyNumber\"><Default>100</Default></Field>"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{e52012a0-51eb-4c0c-8dfb-9b8a0ebedcb6}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"Combine\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Merge\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Combine\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"EncodedAbsUrl\" /><FieldRef Name=\"TemplateUrl\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkCombine\" type=\"CHECKBOX\" title=\"Merge]]\" href=\"]]></HTML><Field Name=\"EncodedAbsUrl\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkUrl\" type=\"HIDDEN\" href=\"]]></HTML><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkProgID\" type=\"HIDDEN\" href=\"]]></HTML><MapToControl><HTML>|</HTML><GetFileExtension><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /></GetFileExtension></MapToControl><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{5d36727b-bcb2-47d2-a231-1f0bc63b7439}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"RepairDocument\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Relink\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"RepairDocument\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"ID\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkRepair\" type=\"CHECKBOX\" title=\"Relink\" docid=\"]]></HTML><Field Name=\"ID\" /><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                  },
                  {
                    "verb": "addSPView",
                    "name": "All Documents",
                    "viewFields": [
                      "DocIcon",
                      "LinkFilename",
                      "MyText",
                      "MyDate",
                      "MyNumber"
                    ],
                    "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                    "rowLimit": 30,
                    "isPaged": true,
                    "makeDefault": true,
                    "formatterJSON": ""
                  }
                ]
              },
              {
                "verb": "addNavLink",
                "url": "MyLibrary/Forms/AllItems.aspx",
                "displayName": "MyLibrary",
                "isWebRelative": true
              }
            ]
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'fb4b0cf8-c006-4802-a1ea-57e0e4852188')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" }, "AllowContentTypes": true, "BaseTemplate": 101, "BaseType": 1, "ContentTypesEnabled": false, "CrawlNonDefaultViews": false, "Created": "2019-01-11T10:03:19Z", "CurrentChangeToken": { "StringValue": "1;3;fb4b0cf8-c006-4802-a1ea-57e0e4852188;636827981522200000;96826061" }, "CustomActionElements": { "Items": [{ "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vdw", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdxFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdx", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdmFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdm", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenXsn", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "xsn", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XsnLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument2", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.2", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument3", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.3", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument4", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.4", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }] }, "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000", "DefaultItemOpenUseListSetting": false, "Description": "", "Direction": "none", "DisableGridEditing": false, "DocumentTemplateUrl": "/sites/team1/MyLibrary/Forms/template.dotx", "DraftVersionVisibility": 0, "EnableAttachments": false, "EnableFolderCreation": true, "EnableMinorVersions": false, "EnableModeration": false, "EnableRequestSignOff": true, "EnableVersioning": true, "EntityTypeName": "MyLibrary", "ExemptFromBlockDownloadOfNonViewableFiles": false, "FileSavePostProcessingEnabled": false, "ForceCheckout": false, "HasExternalDataSource": false, "Hidden": false, "Id": "fb4b0cf8-c006-4802-a1ea-57e0e4852188", "ImagePath": { "DecodedUrl": "/_layouts/15/images/itdl.png?rev=45" }, "ImageUrl": "/_layouts/15/images/itdl.png?rev=45", "IrmEnabled": false, "IrmExpire": false, "IrmReject": false, "IsApplicationList": false, "IsCatalog": false, "IsPrivate": false, "ItemCount": 0, "LastItemDeletedDate": "2019-01-11T10:03:19Z", "LastItemModifiedDate": "2019-01-11T10:04:15Z", "LastItemUserModifiedDate": "2019-01-11T10:03:19Z", "ListExperienceOptions": 0, "ListItemEntityTypeFullName": "SP.Data.MyLibraryItem", "MajorVersionLimit": 500, "MajorWithMinorVersionsLimit": 0, "MultipleDataList": false, "NoCrawl": false, "ParentWebPath": { "DecodedUrl": "/sites/team1" }, "ParentWebUrl": "/sites/team1", "ParserDisabled": false, "ServerTemplateCanCreateFolders": true, "TemplateFeatureId": "00bfea71-e717-4e80-aa17-d0c71b360101", "Title": "MyLibrary" }
        );
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: 'fb4b0cf8-c006-4802-a1ea-57e0e4852188'
      }
    }, () => {
      try {
        const expected = {
          "actions": [
            {
              "verb": "createSPList",
              "listName": "MyLibrary",
              "templateType": 101,
              "subactions": [
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{47b1b86f-9f8a-4dbe-a75e-ca5d9b0f566c}\" Type=\"URL\" Name=\"_ShortcutUrl\" DisplayName=\"Shortcut URL\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUrl\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{2662ad77-2410-4938-b01c-e5e43321bad4}\" Type=\"Guid\" Name=\"_ShortcutSiteId\" DisplayName=\"Shortcut Site Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutSiteId\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{e2a3861f-c216-47d7-820f-7cb638862ab2}\" Type=\"Guid\" Name=\"_ShortcutWebId\" DisplayName=\"Shortcut Web Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutWebId\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{e8fea999-553d-4f45-be52-d941627e9fe5}\" Type=\"Guid\" Name=\"_ShortcutUniqueId\" DisplayName=\"Shortcut Unique Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUniqueId\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field DisplayName=\"MyText\" Format=\"Dropdown\" MaxLength=\"255\" Title=\"MyText\" Type=\"Text\" ID=\"{dbd0f8fa-5131-44ed-b7a1-23bfffc50ac8}\" StaticName=\"MyText\" Name=\"MyText\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field DisplayName=\"MyDate\" FriendlyDisplayFormat=\"Disabled\" Format=\"DateTime\" Title=\"MyDate\" Type=\"DateTime\" ID=\"{f98a4e28-5fb3-4737-9a24-3ad533552bf5}\" StaticName=\"MyDate\" Name=\"MyDate\"><Default>[today]</Default></Field>"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field Decimals=\"2\" DisplayName=\"MyNumber\" Format=\"Dropdown\" Percentage=\"FALSE\" Title=\"MyNumber\" Type=\"Number\" ID=\"{496aa48c-0cf7-4990-be49-d373aa327e0c}\" StaticName=\"MyNumber\" Name=\"MyNumber\"><Default>100</Default></Field>"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{e52012a0-51eb-4c0c-8dfb-9b8a0ebedcb6}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"Combine\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Merge\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Combine\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"EncodedAbsUrl\" /><FieldRef Name=\"TemplateUrl\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkCombine\" type=\"CHECKBOX\" title=\"Merge]]\" href=\"]]></HTML><Field Name=\"EncodedAbsUrl\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkUrl\" type=\"HIDDEN\" href=\"]]></HTML><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkProgID\" type=\"HIDDEN\" href=\"]]></HTML><MapToControl><HTML>|</HTML><GetFileExtension><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /></GetFileExtension></MapToControl><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{5d36727b-bcb2-47d2-a231-1f0bc63b7439}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"RepairDocument\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Relink\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"RepairDocument\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"ID\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkRepair\" type=\"CHECKBOX\" title=\"Relink\" docid=\"]]></HTML><Field Name=\"ID\" /><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                },
                {
                  "verb": "addSPView",
                  "name": "All Documents",
                  "viewFields": [
                    "DocIcon",
                    "LinkFilename",
                    "MyText",
                    "MyDate",
                    "MyNumber"
                  ],
                  "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                  "rowLimit": 30,
                  "isPaged": true,
                  "makeDefault": true,
                  "formatterJSON": ""
                }
              ]
            },
            {
              "verb": "addNavLink",
              "url": "MyLibrary/Forms/AllItems.aspx",
              "displayName": "MyLibrary",
              "isWebRelative": true
            }
          ]
        };
        const actual = log[log.length - 1];
        assert.strictEqual(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('extracts the site script from the given list if list id option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/Microsoft_SharePoint_Utilities_WebTemplateExtensions_SiteScriptUtility_GetSiteScriptFromList`) > -1) {
        return Promise.resolve({
          value: {
            "actions": [
              {
                "verb": "createSPList",
                "listName": "MyLibrary",
                "templateType": 101,
                "subactions": [
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{47b1b86f-9f8a-4dbe-a75e-ca5d9b0f566c}\" Type=\"URL\" Name=\"_ShortcutUrl\" DisplayName=\"Shortcut URL\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUrl\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{2662ad77-2410-4938-b01c-e5e43321bad4}\" Type=\"Guid\" Name=\"_ShortcutSiteId\" DisplayName=\"Shortcut Site Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutSiteId\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{e2a3861f-c216-47d7-820f-7cb638862ab2}\" Type=\"Guid\" Name=\"_ShortcutWebId\" DisplayName=\"Shortcut Web Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutWebId\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{e8fea999-553d-4f45-be52-d941627e9fe5}\" Type=\"Guid\" Name=\"_ShortcutUniqueId\" DisplayName=\"Shortcut Unique Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUniqueId\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field DisplayName=\"MyText\" Format=\"Dropdown\" MaxLength=\"255\" Title=\"MyText\" Type=\"Text\" ID=\"{dbd0f8fa-5131-44ed-b7a1-23bfffc50ac8}\" StaticName=\"MyText\" Name=\"MyText\" />"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field DisplayName=\"MyDate\" FriendlyDisplayFormat=\"Disabled\" Format=\"DateTime\" Title=\"MyDate\" Type=\"DateTime\" ID=\"{f98a4e28-5fb3-4737-9a24-3ad533552bf5}\" StaticName=\"MyDate\" Name=\"MyDate\"><Default>[today]</Default></Field>"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field Decimals=\"2\" DisplayName=\"MyNumber\" Format=\"Dropdown\" Percentage=\"FALSE\" Title=\"MyNumber\" Type=\"Number\" ID=\"{496aa48c-0cf7-4990-be49-d373aa327e0c}\" StaticName=\"MyNumber\" Name=\"MyNumber\"><Default>100</Default></Field>"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{e52012a0-51eb-4c0c-8dfb-9b8a0ebedcb6}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"Combine\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Merge\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Combine\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"EncodedAbsUrl\" /><FieldRef Name=\"TemplateUrl\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkCombine\" type=\"CHECKBOX\" title=\"Merge]]\" href=\"]]></HTML><Field Name=\"EncodedAbsUrl\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkUrl\" type=\"HIDDEN\" href=\"]]></HTML><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkProgID\" type=\"HIDDEN\" href=\"]]></HTML><MapToControl><HTML>|</HTML><GetFileExtension><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /></GetFileExtension></MapToControl><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                  },
                  {
                    "verb": "addSPFieldXml",
                    "schemaXml": "<Field ID=\"{5d36727b-bcb2-47d2-a231-1f0bc63b7439}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"RepairDocument\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Relink\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"RepairDocument\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"ID\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkRepair\" type=\"CHECKBOX\" title=\"Relink\" docid=\"]]></HTML><Field Name=\"ID\" /><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                  },
                  {
                    "verb": "addSPView",
                    "name": "All Documents",
                    "viewFields": [
                      "DocIcon",
                      "LinkFilename",
                      "MyText",
                      "MyDate",
                      "MyNumber"
                    ],
                    "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                    "rowLimit": 30,
                    "isPaged": true,
                    "makeDefault": true,
                    "formatterJSON": ""
                  }
                ]
              },
              {
                "verb": "addNavLink",
                "url": "MyLibrary/Forms/AllItems.aspx",
                "displayName": "MyLibrary",
                "isWebRelative": true
              }
            ]
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'fb4b0cf8-c006-4802-a1ea-57e0e4852188')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" }, "AllowContentTypes": true, "BaseTemplate": 101, "BaseType": 1, "ContentTypesEnabled": false, "CrawlNonDefaultViews": false, "Created": "2019-01-11T10:03:19Z", "CurrentChangeToken": { "StringValue": "1;3;fb4b0cf8-c006-4802-a1ea-57e0e4852188;636827981522200000;96826061" }, "CustomActionElements": { "Items": [{ "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vdw", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdxFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdx", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdmFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdm", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenXsn", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "xsn", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XsnLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument2", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.2", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument3", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.3", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument4", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.4", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }] }, "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000", "DefaultItemOpenUseListSetting": false, "Description": "", "Direction": "none", "DisableGridEditing": false, "DocumentTemplateUrl": "/sites/team1/MyLibrary/Forms/template.dotx", "DraftVersionVisibility": 0, "EnableAttachments": false, "EnableFolderCreation": true, "EnableMinorVersions": false, "EnableModeration": false, "EnableRequestSignOff": true, "EnableVersioning": true, "EntityTypeName": "MyLibrary", "ExemptFromBlockDownloadOfNonViewableFiles": false, "FileSavePostProcessingEnabled": false, "ForceCheckout": false, "HasExternalDataSource": false, "Hidden": false, "Id": "fb4b0cf8-c006-4802-a1ea-57e0e4852188", "ImagePath": { "DecodedUrl": "/_layouts/15/images/itdl.png?rev=45" }, "ImageUrl": "/_layouts/15/images/itdl.png?rev=45", "IrmEnabled": false, "IrmExpire": false, "IrmReject": false, "IsApplicationList": false, "IsCatalog": false, "IsPrivate": false, "ItemCount": 0, "LastItemDeletedDate": "2019-01-11T10:03:19Z", "LastItemModifiedDate": "2019-01-11T10:04:15Z", "LastItemUserModifiedDate": "2019-01-11T10:03:19Z", "ListExperienceOptions": 0, "ListItemEntityTypeFullName": "SP.Data.MyLibraryItem", "MajorVersionLimit": 500, "MajorWithMinorVersionsLimit": 0, "MultipleDataList": false, "NoCrawl": false, "ParentWebPath": { "DecodedUrl": "/sites/team1" }, "ParentWebUrl": "/sites/team1", "ParserDisabled": false, "ServerTemplateCanCreateFolders": true, "TemplateFeatureId": "00bfea71-e717-4e80-aa17-d0c71b360101", "Title": "MyLibrary" }
        );
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: 'fb4b0cf8-c006-4802-a1ea-57e0e4852188'
      }
    }, () => {
      try {
        const expected = {
          "actions": [
            {
              "verb": "createSPList",
              "listName": "MyLibrary",
              "templateType": 101,
              "subactions": [
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{47b1b86f-9f8a-4dbe-a75e-ca5d9b0f566c}\" Type=\"URL\" Name=\"_ShortcutUrl\" DisplayName=\"Shortcut URL\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUrl\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{2662ad77-2410-4938-b01c-e5e43321bad4}\" Type=\"Guid\" Name=\"_ShortcutSiteId\" DisplayName=\"Shortcut Site Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutSiteId\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{e2a3861f-c216-47d7-820f-7cb638862ab2}\" Type=\"Guid\" Name=\"_ShortcutWebId\" DisplayName=\"Shortcut Web Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutWebId\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{e8fea999-553d-4f45-be52-d941627e9fe5}\" Type=\"Guid\" Name=\"_ShortcutUniqueId\" DisplayName=\"Shortcut Unique Id\" DisplaceOnUpgrade=\"TRUE\" Indexed=\"FALSE\" Required=\"FALSE\" Hidden=\"TRUE\" ReadOnlyField=\"TRUE\" ShowInEditForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"_ShortcutUniqueId\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field DisplayName=\"MyText\" Format=\"Dropdown\" MaxLength=\"255\" Title=\"MyText\" Type=\"Text\" ID=\"{dbd0f8fa-5131-44ed-b7a1-23bfffc50ac8}\" StaticName=\"MyText\" Name=\"MyText\" />"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field DisplayName=\"MyDate\" FriendlyDisplayFormat=\"Disabled\" Format=\"DateTime\" Title=\"MyDate\" Type=\"DateTime\" ID=\"{f98a4e28-5fb3-4737-9a24-3ad533552bf5}\" StaticName=\"MyDate\" Name=\"MyDate\"><Default>[today]</Default></Field>"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field Decimals=\"2\" DisplayName=\"MyNumber\" Format=\"Dropdown\" Percentage=\"FALSE\" Title=\"MyNumber\" Type=\"Number\" ID=\"{496aa48c-0cf7-4990-be49-d373aa327e0c}\" StaticName=\"MyNumber\" Name=\"MyNumber\"><Default>100</Default></Field>"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{e52012a0-51eb-4c0c-8dfb-9b8a0ebedcb6}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"Combine\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Merge\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Combine\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"EncodedAbsUrl\" /><FieldRef Name=\"TemplateUrl\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkCombine\" type=\"CHECKBOX\" title=\"Merge]]\" href=\"]]></HTML><Field Name=\"EncodedAbsUrl\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkUrl\" type=\"HIDDEN\" href=\"]]></HTML><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /><HTML><![CDATA[\">]]></HTML><HTML><![CDATA[<input id=\"chkProgID\" type=\"HIDDEN\" href=\"]]></HTML><MapToControl><HTML>|</HTML><GetFileExtension><Column Name=\"TemplateUrl\" HTMLEncode=\"TRUE\" /></GetFileExtension></MapToControl><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                },
                {
                  "verb": "addSPFieldXml",
                  "schemaXml": "<Field ID=\"{5d36727b-bcb2-47d2-a231-1f0bc63b7439}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"RepairDocument\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"Relink\" Filterable=\"FALSE\" Sortable=\"FALSE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"RepairDocument\"><FieldRefs><FieldRef Name=\"FSObjType\" Key=\"Primary\" /><FieldRef Name=\"ID\" /></FieldRefs><DisplayPattern><IfEqual><Expr1><Field Name=\"FSObjType\" /></Expr1><Expr2>0</Expr2><Then><HTML><![CDATA[<input id=\"chkRepair\" type=\"CHECKBOX\" title=\"Relink\" docid=\"]]></HTML><Field Name=\"ID\" /><HTML><![CDATA[\">]]></HTML></Then></IfEqual></DisplayPattern></Field>"
                },
                {
                  "verb": "addSPView",
                  "name": "All Documents",
                  "viewFields": [
                    "DocIcon",
                    "LinkFilename",
                    "MyText",
                    "MyDate",
                    "MyNumber"
                  ],
                  "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                  "rowLimit": 30,
                  "isPaged": true,
                  "makeDefault": true,
                  "formatterJSON": ""
                }
              ]
            },
            {
              "verb": "addNavLink",
              "url": "MyLibrary/Forms/AllItems.aspx",
              "displayName": "MyLibrary",
              "isWebRelative": true
            }
          ]
        };
        const actual = log[log.length - 1];
        assert.strictEqual(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when trying to extract site script if the server did not return generated site script (using listId)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/Microsoft_SharePoint_Utilities_WebTemplateExtensions_SiteScriptUtility_GetSiteScriptFromList`) > -1) {
        return Promise.resolve({
          value: ''
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'fb4b0cf8-c006-4802-a1ea-57e0e4852188')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" }, "AllowContentTypes": true, "BaseTemplate": 101, "BaseType": 1, "ContentTypesEnabled": false, "CrawlNonDefaultViews": false, "Created": "2019-01-11T10:03:19Z", "CurrentChangeToken": { "StringValue": "1;3;fb4b0cf8-c006-4802-a1ea-57e0e4852188;636827981522200000;96826061" }, "CustomActionElements": { "Items": [{ "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vdw", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdxFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdx", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdmFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdm", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenXsn", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "xsn", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XsnLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument2", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.2", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument3", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.3", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument4", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.4", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }] }, "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000", "DefaultItemOpenUseListSetting": false, "Description": "", "Direction": "none", "DisableGridEditing": false, "DocumentTemplateUrl": "/sites/team1/MyLibrary/Forms/template.dotx", "DraftVersionVisibility": 0, "EnableAttachments": false, "EnableFolderCreation": true, "EnableMinorVersions": false, "EnableModeration": false, "EnableRequestSignOff": true, "EnableVersioning": true, "EntityTypeName": "MyLibrary", "ExemptFromBlockDownloadOfNonViewableFiles": false, "FileSavePostProcessingEnabled": false, "ForceCheckout": false, "HasExternalDataSource": false, "Hidden": false, "Id": "fb4b0cf8-c006-4802-a1ea-57e0e4852188", "ImagePath": { "DecodedUrl": "/_layouts/15/images/itdl.png?rev=45" }, "ImageUrl": "/_layouts/15/images/itdl.png?rev=45", "IrmEnabled": false, "IrmExpire": false, "IrmReject": false, "IsApplicationList": false, "IsCatalog": false, "IsPrivate": false, "ItemCount": 0, "LastItemDeletedDate": "2019-01-11T10:03:19Z", "LastItemModifiedDate": "2019-01-11T10:04:15Z", "LastItemUserModifiedDate": "2019-01-11T10:03:19Z", "ListExperienceOptions": 0, "ListItemEntityTypeFullName": "SP.Data.MyLibraryItem", "MajorVersionLimit": 500, "MajorWithMinorVersionsLimit": 0, "MultipleDataList": false, "NoCrawl": false, "ParentWebPath": { "DecodedUrl": "/sites/team1" }, "ParentWebUrl": "/sites/team1", "ParserDisabled": false, "ServerTemplateCanCreateFolders": true, "TemplateFeatureId": "00bfea71-e717-4e80-aa17-d0c71b360101", "Title": "MyLibrary" }
        );
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: 'fb4b0cf8-c006-4802-a1ea-57e0e4852188',
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred, the site script could not be extracted from list 'fb4b0cf8-c006-4802-a1ea-57e0e4852188'")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when trying to extract site script if the server did not return generated site script (using listTitle)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/Microsoft_SharePoint_Utilities_WebTemplateExtensions_SiteScriptUtility_GetSiteScriptFromList`) > -1) {
        return Promise.resolve({
          value: ''
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists/GetByTitle('MyLibrary')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" }, "AllowContentTypes": true, "BaseTemplate": 101, "BaseType": 1, "ContentTypesEnabled": false, "CrawlNonDefaultViews": false, "Created": "2019-01-11T10:03:19Z", "CurrentChangeToken": { "StringValue": "1;3;fb4b0cf8-c006-4802-a1ea-57e0e4852188;636827981522200000;96826061" }, "CustomActionElements": { "Items": [{ "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vdw", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdxFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdx", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdmFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdm", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenXsn", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "xsn", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XsnLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument2", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.2", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument3", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.3", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument4", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.4", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }] }, "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000", "DefaultItemOpenUseListSetting": false, "Description": "", "Direction": "none", "DisableGridEditing": false, "DocumentTemplateUrl": "/sites/team1/MyLibrary/Forms/template.dotx", "DraftVersionVisibility": 0, "EnableAttachments": false, "EnableFolderCreation": true, "EnableMinorVersions": false, "EnableModeration": false, "EnableRequestSignOff": true, "EnableVersioning": true, "EntityTypeName": "MyLibrary", "ExemptFromBlockDownloadOfNonViewableFiles": false, "FileSavePostProcessingEnabled": false, "ForceCheckout": false, "HasExternalDataSource": false, "Hidden": false, "Id": "fb4b0cf8-c006-4802-a1ea-57e0e4852188", "ImagePath": { "DecodedUrl": "/_layouts/15/images/itdl.png?rev=45" }, "ImageUrl": "/_layouts/15/images/itdl.png?rev=45", "IrmEnabled": false, "IrmExpire": false, "IrmReject": false, "IsApplicationList": false, "IsCatalog": false, "IsPrivate": false, "ItemCount": 0, "LastItemDeletedDate": "2019-01-11T10:03:19Z", "LastItemModifiedDate": "2019-01-11T10:04:15Z", "LastItemUserModifiedDate": "2019-01-11T10:03:19Z", "ListExperienceOptions": 0, "ListItemEntityTypeFullName": "SP.Data.MyLibraryItem", "MajorVersionLimit": 500, "MajorWithMinorVersionsLimit": 0, "MultipleDataList": false, "NoCrawl": false, "ParentWebPath": { "DecodedUrl": "/sites/team1" }, "ParentWebUrl": "/sites/team1", "ParserDisabled": false, "ServerTemplateCanCreateFolders": true, "TemplateFeatureId": "00bfea71-e717-4e80-aa17-d0c71b360101", "Title": "MyLibrary" }
        );
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary',
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred, the site script could not be extracted from list 'MyLibrary'")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when trying to extract site script from a list that doesn\'t exist', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/Microsoft_SharePoint_Utilities_WebTemplateExtensions_SiteScriptUtility_GetSiteScriptFromList`) > -1) {
        return Promise.resolve({
          value: ''
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')`) > -1) {
        return Promise.reject(new Error("404 - \"404 FILE NOT FOUND\""));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('404 - "404 FILE NOT FOUND"')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when listId option is passed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
        debug: false,
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

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails validation if the listid option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'XXXXX' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listid option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
    assert(actual);
  });

  it('fails validation if both listId and listTitle options are passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85', listTitle: 'Documents' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both listId and listTitle options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });
});