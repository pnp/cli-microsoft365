import commands from '../../commands';
import Command, { CommandValidate, CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./list-label-get');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.LIST_LABEL_GET, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth,
      (command as any).getRequestDigest
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.LIST_LABEL_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.LIST_LABEL_GET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team1', title: 'Documents' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the compliance label from the given list if title option is passed (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
        return Promise.resolve({
            "@odata.context": "https://contoso.sharepoint.com/sites/team1/_api/$metadata#SP.CompliancePolicy.ComplianceTag",
            "AcceptMessagesOnlyFromSendersOrMembers": false,
            "AccessType": null,
            "AllowAccessFromUnmanagedDevice": null,
            "AutoDelete": false,
            "BlockDelete": false,
            "BlockEdit": false,
            "ContainsSiteLabel": false,
            "DisplayName": "",
            "EncryptionRMSTemplateId": null,
            "HasRetentionAction": false,
            "IsEventTag": false,
            "Notes": null,
            "RequireSenderAuthenticationEnabled": false,
            "ReviewerEmail": null,
            "SharingCapabilities": null,
            "SuperLock": false,
            "TagDuration": 0,
            "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
            "TagName": "Sensitive",
            "TagRetentionBasedOn": null                     
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists/GetByTitle('MyLibrary')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" }, "AllowContentTypes": true, "BaseTemplate": 101, "BaseType": 1, "ContentTypesEnabled": false, "CrawlNonDefaultViews": false, "Created": "2019-01-11T10:03:19Z", "CurrentChangeToken": { "StringValue": "1;3;fb4b0cf8-c006-4802-a1ea-57e0e4852188;636827981522200000;96826061" }, "CustomActionElements": { "Items": [{ "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vdw", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdxFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdx", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdmFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdm", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenXsn", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "xsn", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XsnLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument2", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.2", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument3", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.3", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument4", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.4", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }] }, "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000", "DefaultItemOpenUseListSetting": false, "Description": "", "Direction": "none", "DisableGridEditing": false, "DocumentTemplateUrl": "/sites/team1/MyLibrary/Forms/template.dotx", "DraftVersionVisibility": 0, "EnableAttachments": false, "EnableFolderCreation": true, "EnableMinorVersions": false, "EnableModeration": false, "EnableRequestSignOff": true, "EnableVersioning": true, "EntityTypeName": "MyLibrary", "ExemptFromBlockDownloadOfNonViewableFiles": false, "FileSavePostProcessingEnabled": false, "ForceCheckout": false, "HasExternalDataSource": false, "Hidden": false, "Id": "fb4b0cf8-c006-4802-a1ea-57e0e4852188", "ImagePath": { "DecodedUrl": "/_layouts/15/images/itdl.png?rev=45" }, "ImageUrl": "/_layouts/15/images/itdl.png?rev=45", "IrmEnabled": false, "IrmExpire": false, "IrmReject": false, "IsApplicationList": false, "IsCatalog": false, "IsPrivate": false, "ItemCount": 0, "LastItemDeletedDate": "2019-01-11T10:03:19Z", "LastItemModifiedDate": "2019-01-11T10:04:15Z", "LastItemUserModifiedDate": "2019-01-11T10:03:19Z", "ListExperienceOptions": 0, "ListItemEntityTypeFullName": "SP.Data.MyLibraryItem", "MajorVersionLimit": 500, "MajorWithMinorVersionsLimit": 0, "MultipleDataList": false, "NoCrawl": false, "ParentWebPath": { "DecodedUrl": "/sites/team1" }, "ParentWebUrl": "/sites/team1", "ParserDisabled": false, "ServerTemplateCanCreateFolders": true, "TemplateFeatureId": "00bfea71-e717-4e80-aa17-d0c71b360101", "Title": "MyLibrary" }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary'
      }
    }, () => {
      try {
        const expected =   {
          "@odata.context": "https://contoso.sharepoint.com/sites/team1/_api/$metadata#SP.CompliancePolicy.ComplianceTag",
          "AcceptMessagesOnlyFromSendersOrMembers": false,
          "AccessType": null,
          "AllowAccessFromUnmanagedDevice": null,
          "AutoDelete": false,
          "BlockDelete": false,
          "BlockEdit": false,
          "ContainsSiteLabel": false,
          "DisplayName": "",
          "EncryptionRMSTemplateId": null,
          "HasRetentionAction": false,
          "IsEventTag": false,
          "Notes": null,
          "RequireSenderAuthenticationEnabled": false,
          "ReviewerEmail": null,
          "SharingCapabilities": null,
          "SuperLock": false,
          "TagDuration": 0,
          "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
          "TagName": "Sensitive",
          "TagRetentionBasedOn": null
        
        } ;   
        const actual = log[log.length - 1];
        assert.equal(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the compliance label from the given list if title option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/`) > -1) {
        return Promise.resolve({
            "@odata.context": "https://contoso.sharepoint.com/sites/team1/_api/$metadata#SP.CompliancePolicy.ComplianceTag",
            "AcceptMessagesOnlyFromSendersOrMembers": false,
            "AccessType": null,
            "AllowAccessFromUnmanagedDevice": null,
            "AutoDelete": false,
            "BlockDelete": false,
            "BlockEdit": false,
            "ContainsSiteLabel": false,
            "DisplayName": "",
            "EncryptionRMSTemplateId": null,
            "HasRetentionAction": false,
            "IsEventTag": false,
            "Notes": null,
            "RequireSenderAuthenticationEnabled": false,
            "ReviewerEmail": null,
            "SharingCapabilities": null,
            "SuperLock": false,
            "TagDuration": 0,
            "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
            "TagName": "Sensitive",
            "TagRetentionBasedOn": null          
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists/GetByTitle('MyLibrary')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" }, "AllowContentTypes": true, "BaseTemplate": 101, "BaseType": 1, "ContentTypesEnabled": false, "CrawlNonDefaultViews": false, "Created": "2019-01-11T10:03:19Z", "CurrentChangeToken": { "StringValue": "1;3;fb4b0cf8-c006-4802-a1ea-57e0e4852188;636827981522200000;96826061" }, "CustomActionElements": { "Items": [{ "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vdw", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdxFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdx", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdmFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdm", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenXsn", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "xsn", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XsnLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument2", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.2", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument3", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.3", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument4", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.4", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }] }, "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000", "DefaultItemOpenUseListSetting": false, "Description": "", "Direction": "none", "DisableGridEditing": false, "DocumentTemplateUrl": "/sites/team1/MyLibrary/Forms/template.dotx", "DraftVersionVisibility": 0, "EnableAttachments": false, "EnableFolderCreation": true, "EnableMinorVersions": false, "EnableModeration": false, "EnableRequestSignOff": true, "EnableVersioning": true, "EntityTypeName": "MyLibrary", "ExemptFromBlockDownloadOfNonViewableFiles": false, "FileSavePostProcessingEnabled": false, "ForceCheckout": false, "HasExternalDataSource": false, "Hidden": false, "Id": "fb4b0cf8-c006-4802-a1ea-57e0e4852188", "ImagePath": { "DecodedUrl": "/_layouts/15/images/itdl.png?rev=45" }, "ImageUrl": "/_layouts/15/images/itdl.png?rev=45", "IrmEnabled": false, "IrmExpire": false, "IrmReject": false, "IsApplicationList": false, "IsCatalog": false, "IsPrivate": false, "ItemCount": 0, "LastItemDeletedDate": "2019-01-11T10:03:19Z", "LastItemModifiedDate": "2019-01-11T10:04:15Z", "LastItemUserModifiedDate": "2019-01-11T10:03:19Z", "ListExperienceOptions": 0, "ListItemEntityTypeFullName": "SP.Data.MyLibraryItem", "MajorVersionLimit": 500, "MajorWithMinorVersionsLimit": 0, "MultipleDataList": false, "NoCrawl": false, "ParentWebPath": { "DecodedUrl": "/sites/team1" }, "ParentWebUrl": "/sites/team1", "ParserDisabled": false, "ServerTemplateCanCreateFolders": true, "TemplateFeatureId": "00bfea71-e717-4e80-aa17-d0c71b360101", "Title": "MyLibrary" }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary'
      }
    }, () => {
      try {
        const expected =  {
          "@odata.context": "https://contoso.sharepoint.com/sites/team1/_api/$metadata#SP.CompliancePolicy.ComplianceTag",
          "AcceptMessagesOnlyFromSendersOrMembers": false,
          "AccessType": null,
          "AllowAccessFromUnmanagedDevice": null,
          "AutoDelete": false,
          "BlockDelete": false,
          "BlockEdit": false,
          "ContainsSiteLabel": false,
          "DisplayName": "",
          "EncryptionRMSTemplateId": null,
          "HasRetentionAction": false,
          "IsEventTag": false,
          "Notes": null,
          "RequireSenderAuthenticationEnabled": false,
          "ReviewerEmail": null,
          "SharingCapabilities": null,
          "SuperLock": false,
          "TagDuration": 0,
          "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
          "TagName": "Sensitive",
          "TagRetentionBasedOn": null
        
        } ;
        const actual = log[log.length - 1];
        assert.equal(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the compliance label from the given list if list id option is passed (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
        return Promise.resolve({
            "@odata.context": "https://contoso.sharepoint.com/sites/team1/_api/$metadata#SP.CompliancePolicy.ComplianceTag",
            "AcceptMessagesOnlyFromSendersOrMembers": false,
            "AccessType": null,
            "AllowAccessFromUnmanagedDevice": null,
            "AutoDelete": false,
            "BlockDelete": false,
            "BlockEdit": false,
            "ContainsSiteLabel": false,
            "DisplayName": "",
            "EncryptionRMSTemplateId": null,
            "HasRetentionAction": false,
            "IsEventTag": false,
            "Notes": null,
            "RequireSenderAuthenticationEnabled": false,
            "ReviewerEmail": null,
            "SharingCapabilities": null,
            "SuperLock": false,
            "TagDuration": 0,
            "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
            "TagName": "Sensitive",
            "TagRetentionBasedOn": null        
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'fb4b0cf8-c006-4802-a1ea-57e0e4852188')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" }, "AllowContentTypes": true, "BaseTemplate": 101, "BaseType": 1, "ContentTypesEnabled": false, "CrawlNonDefaultViews": false, "Created": "2019-01-11T10:03:19Z", "CurrentChangeToken": { "StringValue": "1;3;fb4b0cf8-c006-4802-a1ea-57e0e4852188;636827981522200000;96826061" }, "CustomActionElements": { "Items": [{ "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vdw", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdxFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdx", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdmFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdm", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenXsn", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "xsn", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XsnLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument2", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.2", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument3", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.3", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument4", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.4", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }] }, "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000", "DefaultItemOpenUseListSetting": false, "Description": "", "Direction": "none", "DisableGridEditing": false, "DocumentTemplateUrl": "/sites/team1/MyLibrary/Forms/template.dotx", "DraftVersionVisibility": 0, "EnableAttachments": false, "EnableFolderCreation": true, "EnableMinorVersions": false, "EnableModeration": false, "EnableRequestSignOff": true, "EnableVersioning": true, "EntityTypeName": "MyLibrary", "ExemptFromBlockDownloadOfNonViewableFiles": false, "FileSavePostProcessingEnabled": false, "ForceCheckout": false, "HasExternalDataSource": false, "Hidden": false, "Id": "fb4b0cf8-c006-4802-a1ea-57e0e4852188", "ImagePath": { "DecodedUrl": "/_layouts/15/images/itdl.png?rev=45" }, "ImageUrl": "/_layouts/15/images/itdl.png?rev=45", "IrmEnabled": false, "IrmExpire": false, "IrmReject": false, "IsApplicationList": false, "IsCatalog": false, "IsPrivate": false, "ItemCount": 0, "LastItemDeletedDate": "2019-01-11T10:03:19Z", "LastItemModifiedDate": "2019-01-11T10:04:15Z", "LastItemUserModifiedDate": "2019-01-11T10:03:19Z", "ListExperienceOptions": 0, "ListItemEntityTypeFullName": "SP.Data.MyLibraryItem", "MajorVersionLimit": 500, "MajorWithMinorVersionsLimit": 0, "MultipleDataList": false, "NoCrawl": false, "ParentWebPath": { "DecodedUrl": "/sites/team1" }, "ParentWebUrl": "/sites/team1", "ParserDisabled": false, "ServerTemplateCanCreateFolders": true, "TemplateFeatureId": "00bfea71-e717-4e80-aa17-d0c71b360101", "Title": "MyLibrary" }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: 'fb4b0cf8-c006-4802-a1ea-57e0e4852188'
      }
    }, () => {
      try {
        const expected = {
          "@odata.context": "https://contoso.sharepoint.com/sites/team1/_api/$metadata#SP.CompliancePolicy.ComplianceTag",
          "AcceptMessagesOnlyFromSendersOrMembers": false,
          "AccessType": null,
          "AllowAccessFromUnmanagedDevice": null,
          "AutoDelete": false,
          "BlockDelete": false,
          "BlockEdit": false,
          "ContainsSiteLabel": false,
          "DisplayName": "",
          "EncryptionRMSTemplateId": null,
          "HasRetentionAction": false,
          "IsEventTag": false,
          "Notes": null,
          "RequireSenderAuthenticationEnabled": false,
          "ReviewerEmail": null,
          "SharingCapabilities": null,
          "SuperLock": false,
          "TagDuration": 0,
          "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
          "TagName": "Sensitive",
          "TagRetentionBasedOn": null
        
        };    
        const actual = log[log.length - 1];
        assert.equal(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the compliance label from the given list if list id option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
        return Promise.resolve({
            "@odata.context": "https://contoso.sharepoint.com/sites/team1/_api/$metadata#SP.CompliancePolicy.ComplianceTag",
            "AcceptMessagesOnlyFromSendersOrMembers": false,
            "AccessType": null,
            "AllowAccessFromUnmanagedDevice": null,
            "AutoDelete": false,
            "BlockDelete": false,
            "BlockEdit": false,
            "ContainsSiteLabel": false,
            "DisplayName": "",
            "EncryptionRMSTemplateId": null,
            "HasRetentionAction": false,
            "IsEventTag": false,
            "Notes": null,
            "RequireSenderAuthenticationEnabled": false,
            "ReviewerEmail": null,
            "SharingCapabilities": null,
            "SuperLock": false,
            "TagDuration": 0,
            "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
            "TagName": "Sensitive",
            "TagRetentionBasedOn": null
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'fb4b0cf8-c006-4802-a1ea-57e0e4852188')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" }, "AllowContentTypes": true, "BaseTemplate": 101, "BaseType": 1, "ContentTypesEnabled": false, "CrawlNonDefaultViews": false, "Created": "2019-01-11T10:03:19Z", "CurrentChangeToken": { "StringValue": "1;3;fb4b0cf8-c006-4802-a1ea-57e0e4852188;636827981522200000;96826061" }, "CustomActionElements": { "Items": [{ "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vdw", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdxFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdx", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdmFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdm", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenXsn", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "xsn", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XsnLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument2", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.2", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument3", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.3", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument4", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.4", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }] }, "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000", "DefaultItemOpenUseListSetting": false, "Description": "", "Direction": "none", "DisableGridEditing": false, "DocumentTemplateUrl": "/sites/team1/MyLibrary/Forms/template.dotx", "DraftVersionVisibility": 0, "EnableAttachments": false, "EnableFolderCreation": true, "EnableMinorVersions": false, "EnableModeration": false, "EnableRequestSignOff": true, "EnableVersioning": true, "EntityTypeName": "MyLibrary", "ExemptFromBlockDownloadOfNonViewableFiles": false, "FileSavePostProcessingEnabled": false, "ForceCheckout": false, "HasExternalDataSource": false, "Hidden": false, "Id": "fb4b0cf8-c006-4802-a1ea-57e0e4852188", "ImagePath": { "DecodedUrl": "/_layouts/15/images/itdl.png?rev=45" }, "ImageUrl": "/_layouts/15/images/itdl.png?rev=45", "IrmEnabled": false, "IrmExpire": false, "IrmReject": false, "IsApplicationList": false, "IsCatalog": false, "IsPrivate": false, "ItemCount": 0, "LastItemDeletedDate": "2019-01-11T10:03:19Z", "LastItemModifiedDate": "2019-01-11T10:04:15Z", "LastItemUserModifiedDate": "2019-01-11T10:03:19Z", "ListExperienceOptions": 0, "ListItemEntityTypeFullName": "SP.Data.MyLibraryItem", "MajorVersionLimit": 500, "MajorWithMinorVersionsLimit": 0, "MultipleDataList": false, "NoCrawl": false, "ParentWebPath": { "DecodedUrl": "/sites/team1" }, "ParentWebUrl": "/sites/team1", "ParserDisabled": false, "ServerTemplateCanCreateFolders": true, "TemplateFeatureId": "00bfea71-e717-4e80-aa17-d0c71b360101", "Title": "MyLibrary" }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: 'fb4b0cf8-c006-4802-a1ea-57e0e4852188'
      }
    }, () => {
      try {
        const expected =  {
          "@odata.context": "https://contoso.sharepoint.com/sites/team1/_api/$metadata#SP.CompliancePolicy.ComplianceTag",
          "AcceptMessagesOnlyFromSendersOrMembers": false,
          "AccessType": null,
          "AllowAccessFromUnmanagedDevice": null,
          "AutoDelete": false,
          "BlockDelete": false,
          "BlockEdit": false,
          "ContainsSiteLabel": false,
          "DisplayName": "",
          "EncryptionRMSTemplateId": null,
          "HasRetentionAction": false,
          "IsEventTag": false,
          "Notes": null,
          "RequireSenderAuthenticationEnabled": false,
          "ReviewerEmail": null,
          "SharingCapabilities": null,
          "SuperLock": false,
          "TagDuration": 0,
          "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
          "TagName": "Sensitive",
          "TagRetentionBasedOn": null        
        }; 
        const actual = log[log.length - 1];
        assert.equal(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when trying to get the complaince label for the list (using listId)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'fb4b0cf8-c006-4802-a1ea-57e0e4852188')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" }, "AllowContentTypes": true, "BaseTemplate": 101, "BaseType": 1, "ContentTypesEnabled": false, "CrawlNonDefaultViews": false, "Created": "2019-01-11T10:03:19Z", "CurrentChangeToken": { "StringValue": "1;3;fb4b0cf8-c006-4802-a1ea-57e0e4852188;636827981522200000;96826061" }, "CustomActionElements": { "Items": [{ "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vdw", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdxFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdx", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdmFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdm", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenXsn", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "xsn", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XsnLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument2", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.2", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument3", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.3", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument4", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.4", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }] }, "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000", "DefaultItemOpenUseListSetting": false, "Description": "", "Direction": "none", "DisableGridEditing": false, "DocumentTemplateUrl": "/sites/team1/MyLibrary/Forms/template.dotx", "DraftVersionVisibility": 0, "EnableAttachments": false, "EnableFolderCreation": true, "EnableMinorVersions": false, "EnableModeration": false, "EnableRequestSignOff": true, "EnableVersioning": true, "EntityTypeName": "MyLibrary", "ExemptFromBlockDownloadOfNonViewableFiles": false, "FileSavePostProcessingEnabled": false, "ForceCheckout": false, "HasExternalDataSource": false, "Hidden": false, "Id": "fb4b0cf8-c006-4802-a1ea-57e0e4852188", "ImagePath": { "DecodedUrl": "/_layouts/15/images/itdl.png?rev=45" }, "ImageUrl": "/_layouts/15/images/itdl.png?rev=45", "IrmEnabled": false, "IrmExpire": false, "IrmReject": false, "IsApplicationList": false, "IsCatalog": false, "IsPrivate": false, "ItemCount": 0, "LastItemDeletedDate": "2019-01-11T10:03:19Z", "LastItemModifiedDate": "2019-01-11T10:04:15Z", "LastItemUserModifiedDate": "2019-01-11T10:03:19Z", "ListExperienceOptions": 0, "ListItemEntityTypeFullName": "SP.Data.MyLibraryItem", "MajorVersionLimit": 500, "MajorWithMinorVersionsLimit": 0, "MultipleDataList": false, "NoCrawl": false, "ParentWebPath": { "DecodedUrl": "/sites/team1" }, "ParentWebUrl": "/sites/team1", "ParserDisabled": false, "ServerTemplateCanCreateFolders": true, "TemplateFeatureId": "00bfea71-e717-4e80-aa17-d0c71b360101", "Title": "MyLibrary" }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: 'fb4b0cf8-c006-4802-a1ea-57e0e4852188',
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred, compliance label could not be retrieved from list 'fb4b0cf8-c006-4802-a1ea-57e0e4852188'")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when trying to get compliance label for the list (using listTitle)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists/GetByTitle('MyLibrary')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" }, "AllowContentTypes": true, "BaseTemplate": 101, "BaseType": 1, "ContentTypesEnabled": false, "CrawlNonDefaultViews": false, "Created": "2019-01-11T10:03:19Z", "CurrentChangeToken": { "StringValue": "1;3;fb4b0cf8-c006-4802-a1ea-57e0e4852188;636827981522200000;96826061" }, "CustomActionElements": { "Items": [{ "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vdw", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdxFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdx", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "vwaViewAsWebAccessVsdmFromEcb", "EnabledScript": null, "ImageUrl": null, "Location": "EditControlBlock", "RegistrationId": "vsdm", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "View in Web Browser", "UrlAction": "~site/_layouts/15/VisioWebAccess/VisioWebAccess.aspx?listguid={ListId}&itemid={ItemId}&DefaultItemOpen=1" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenXsn", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "xsn", "RegistrationType": 4, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XsnLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument2", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.2", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument3", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.3", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }, { "ClientSideComponentId": "00000000-0000-0000-0000-000000000000", "ClientSideComponentProperties": "", "CommandUIExtension": null, "Id": "FormServerEcbItemOpenInfoPathDocument4", "EnabledScript": null, "ImageUrl": "/_layouts/15/images/icxddoc.gif?rev=45", "Location": "EditControlBlock", "RegistrationId": "InfoPath.Document.4", "RegistrationType": 3, "RequireSiteAdministrator": false, "Rights": { "High": "0", "Low": "1" }, "Title": "Edit in Browser", "UrlAction": "~site/_layouts/15/formserver.aspx?XmlLocation={ItemUrl}&OpenIn=Browser&Source={Source}" }] }, "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000", "DefaultItemOpenUseListSetting": false, "Description": "", "Direction": "none", "DisableGridEditing": false, "DocumentTemplateUrl": "/sites/team1/MyLibrary/Forms/template.dotx", "DraftVersionVisibility": 0, "EnableAttachments": false, "EnableFolderCreation": true, "EnableMinorVersions": false, "EnableModeration": false, "EnableRequestSignOff": true, "EnableVersioning": true, "EntityTypeName": "MyLibrary", "ExemptFromBlockDownloadOfNonViewableFiles": false, "FileSavePostProcessingEnabled": false, "ForceCheckout": false, "HasExternalDataSource": false, "Hidden": false, "Id": "fb4b0cf8-c006-4802-a1ea-57e0e4852188", "ImagePath": { "DecodedUrl": "/_layouts/15/images/itdl.png?rev=45" }, "ImageUrl": "/_layouts/15/images/itdl.png?rev=45", "IrmEnabled": false, "IrmExpire": false, "IrmReject": false, "IsApplicationList": false, "IsCatalog": false, "IsPrivate": false, "ItemCount": 0, "LastItemDeletedDate": "2019-01-11T10:03:19Z", "LastItemModifiedDate": "2019-01-11T10:04:15Z", "LastItemUserModifiedDate": "2019-01-11T10:03:19Z", "ListExperienceOptions": 0, "ListItemEntityTypeFullName": "SP.Data.MyLibraryItem", "MajorVersionLimit": 500, "MajorWithMinorVersionsLimit": 0, "MultipleDataList": false, "NoCrawl": false, "ParentWebPath": { "DecodedUrl": "/sites/team1" }, "ParentWebUrl": "/sites/team1", "ParserDisabled": false, "ServerTemplateCanCreateFolders": true, "TemplateFeatureId": "00bfea71-e717-4e80-aa17-d0c71b360101", "Title": "MyLibrary" }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary',
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred, compliance label could not be retrieved from list 'MyLibrary'")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when trying to get compliance lable from  a list that doesn\'t exist', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
        return Promise.resolve([]);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')`) > -1) {
        return Promise.reject(new Error("404 - \"404 FILE NOT FOUND\""));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('404 - "404 FILE NOT FOUND"')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when listId option is passed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/web/lists(guid') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

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

  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails validation if the listid option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'XXXXX' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the listid option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
    assert(actual);
  });

  it('fails validation if both listId and listTitle options are passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85', listTitle: 'Documents' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if both listId and listTitle options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.LIST_LABEL_GET));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

});