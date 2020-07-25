import commands from '../../commands';
import Command, { CommandValidate, CommandError, CommandOption, CommandTypes } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./list-contenttype-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_CONTENTTYPE_ADD, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

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
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
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
    assert.strictEqual(command.name.startsWith(commands.LIST_CONTENTTYPE_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds content type to the specific list if listTitle option is passed (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/ContentTypes/AddAvailableContentType`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0 &&
          JSON.stringify(opts.body) === JSON.stringify({
            contentTypeId: '0x0120'
          })) {
          return Promise.resolve({
            "Description": "Create a new folder.",
            "DisplayFormTemplateName": "ListForm",
            "DisplayFormUrl": "",
            "DocumentTemplate": "",
            "DocumentTemplateUrl": "",
            "EditFormTemplateName": "ListForm",
            "EditFormUrl": "",
            "Group": "Folder Content Types",
            "Hidden": false,
            "Id": {
              "StringValue": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
            },
            "JSLink": "",
            "MobileDisplayFormUrl": "",
            "MobileEditFormUrl": "",
            "MobileNewFormUrl": "",
            "Name": "Folder",
            "NewFormTemplateName": "ListForm",
            "NewFormUrl": "",
            "ReadOnly": false,
            "SchemaXml": "<ContentType ID=\"0x0120000EAD53EDAD7C6647B0D976EEC953F99E\" Name=\"Folder\" Group=\"Folder Content Types\" Description=\"Create a new folder.\" Sealed=\"TRUE\" Version=\"0\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" FeatureId=\"{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}\"><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" ShowInNewForm=\"FALSE\" ShowInFileDlg=\"FALSE\" DisplayName=\"Title\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" ColName=\"nvarchar8\" Required=\"FALSE\" Hidden=\"TRUE\"/><Field ID=\"{8553196d-ec8d-4564-9861-3dbe931050c8}\" ShowInFileDlg=\"FALSE\" ShowInVersionHistory=\"FALSE\" Type=\"File\" Name=\"FileLeafRef\" DisplayName=\"Name\" AuthoringInfo=\"(for use in forms)\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"LeafName\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" Required=\"TRUE\" NoCustomize=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FileLeafRef\" FromBaseType=\"TRUE\" Hidden=\"FALSE\"/><Field ID=\"{b824e17e-a1b3-426e-aecf-f0184d900485}\" Name=\"ItemChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Item Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"ItemChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ItemChildCount\" FromBaseType=\"TRUE\"/><Field ID=\"{960ff01f-2b6d-4f1b-9c3f-e19ad8927341}\" Name=\"FolderChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Folder Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"FolderChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FolderChildCount\" FromBaseType=\"TRUE\"/></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>",
            "Scope": "/Shared Documents",
            "Sealed": true,
            "StringId": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        listTitle: 'Documents',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        contentTypeId: '0x0120'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": "Create a new folder.",
          "DisplayFormTemplateName": "ListForm",
          "DisplayFormUrl": "",
          "DocumentTemplate": "",
          "DocumentTemplateUrl": "",
          "EditFormTemplateName": "ListForm",
          "EditFormUrl": "",
          "Group": "Folder Content Types",
          "Hidden": false,
          "Id": {
            "StringValue": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
          },
          "JSLink": "",
          "MobileDisplayFormUrl": "",
          "MobileEditFormUrl": "",
          "MobileNewFormUrl": "",
          "Name": "Folder",
          "NewFormTemplateName": "ListForm",
          "NewFormUrl": "",
          "ReadOnly": false,
          "SchemaXml": "<ContentType ID=\"0x0120000EAD53EDAD7C6647B0D976EEC953F99E\" Name=\"Folder\" Group=\"Folder Content Types\" Description=\"Create a new folder.\" Sealed=\"TRUE\" Version=\"0\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" FeatureId=\"{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}\"><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" ShowInNewForm=\"FALSE\" ShowInFileDlg=\"FALSE\" DisplayName=\"Title\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" ColName=\"nvarchar8\" Required=\"FALSE\" Hidden=\"TRUE\"/><Field ID=\"{8553196d-ec8d-4564-9861-3dbe931050c8}\" ShowInFileDlg=\"FALSE\" ShowInVersionHistory=\"FALSE\" Type=\"File\" Name=\"FileLeafRef\" DisplayName=\"Name\" AuthoringInfo=\"(for use in forms)\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"LeafName\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" Required=\"TRUE\" NoCustomize=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FileLeafRef\" FromBaseType=\"TRUE\" Hidden=\"FALSE\"/><Field ID=\"{b824e17e-a1b3-426e-aecf-f0184d900485}\" Name=\"ItemChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Item Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"ItemChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ItemChildCount\" FromBaseType=\"TRUE\"/><Field ID=\"{960ff01f-2b6d-4f1b-9c3f-e19ad8927341}\" Name=\"FolderChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Folder Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"FolderChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FolderChildCount\" FromBaseType=\"TRUE\"/></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>",
          "Scope": "/Shared Documents",
          "Sealed": true,
          "StringId": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds content type to the specific list if listTitle option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/ContentTypes/AddAvailableContentType`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0 &&
          JSON.stringify(opts.body) === JSON.stringify({
            contentTypeId: '0x0120'
          })) {
          return Promise.resolve({
            "Description": "Create a new folder.",
            "DisplayFormTemplateName": "ListForm",
            "DisplayFormUrl": "",
            "DocumentTemplate": "",
            "DocumentTemplateUrl": "",
            "EditFormTemplateName": "ListForm",
            "EditFormUrl": "",
            "Group": "Folder Content Types",
            "Hidden": false,
            "Id": {
              "StringValue": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
            },
            "JSLink": "",
            "MobileDisplayFormUrl": "",
            "MobileEditFormUrl": "",
            "MobileNewFormUrl": "",
            "Name": "Folder",
            "NewFormTemplateName": "ListForm",
            "NewFormUrl": "",
            "ReadOnly": false,
            "SchemaXml": "<ContentType ID=\"0x0120000EAD53EDAD7C6647B0D976EEC953F99E\" Name=\"Folder\" Group=\"Folder Content Types\" Description=\"Create a new folder.\" Sealed=\"TRUE\" Version=\"0\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" FeatureId=\"{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}\"><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" ShowInNewForm=\"FALSE\" ShowInFileDlg=\"FALSE\" DisplayName=\"Title\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" ColName=\"nvarchar8\" Required=\"FALSE\" Hidden=\"TRUE\"/><Field ID=\"{8553196d-ec8d-4564-9861-3dbe931050c8}\" ShowInFileDlg=\"FALSE\" ShowInVersionHistory=\"FALSE\" Type=\"File\" Name=\"FileLeafRef\" DisplayName=\"Name\" AuthoringInfo=\"(for use in forms)\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"LeafName\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" Required=\"TRUE\" NoCustomize=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FileLeafRef\" FromBaseType=\"TRUE\" Hidden=\"FALSE\"/><Field ID=\"{b824e17e-a1b3-426e-aecf-f0184d900485}\" Name=\"ItemChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Item Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"ItemChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ItemChildCount\" FromBaseType=\"TRUE\"/><Field ID=\"{960ff01f-2b6d-4f1b-9c3f-e19ad8927341}\" Name=\"FolderChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Folder Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"FolderChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FolderChildCount\" FromBaseType=\"TRUE\"/></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>",
            "Scope": "/Shared Documents",
            "Sealed": true,
            "StringId": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        listTitle: 'Documents',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        contentTypeId: '0x0120'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": "Create a new folder.",
          "DisplayFormTemplateName": "ListForm",
          "DisplayFormUrl": "",
          "DocumentTemplate": "",
          "DocumentTemplateUrl": "",
          "EditFormTemplateName": "ListForm",
          "EditFormUrl": "",
          "Group": "Folder Content Types",
          "Hidden": false,
          "Id": {
            "StringValue": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
          },
          "JSLink": "",
          "MobileDisplayFormUrl": "",
          "MobileEditFormUrl": "",
          "MobileNewFormUrl": "",
          "Name": "Folder",
          "NewFormTemplateName": "ListForm",
          "NewFormUrl": "",
          "ReadOnly": false,
          "SchemaXml": "<ContentType ID=\"0x0120000EAD53EDAD7C6647B0D976EEC953F99E\" Name=\"Folder\" Group=\"Folder Content Types\" Description=\"Create a new folder.\" Sealed=\"TRUE\" Version=\"0\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" FeatureId=\"{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}\"><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" ShowInNewForm=\"FALSE\" ShowInFileDlg=\"FALSE\" DisplayName=\"Title\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" ColName=\"nvarchar8\" Required=\"FALSE\" Hidden=\"TRUE\"/><Field ID=\"{8553196d-ec8d-4564-9861-3dbe931050c8}\" ShowInFileDlg=\"FALSE\" ShowInVersionHistory=\"FALSE\" Type=\"File\" Name=\"FileLeafRef\" DisplayName=\"Name\" AuthoringInfo=\"(for use in forms)\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"LeafName\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" Required=\"TRUE\" NoCustomize=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FileLeafRef\" FromBaseType=\"TRUE\" Hidden=\"FALSE\"/><Field ID=\"{b824e17e-a1b3-426e-aecf-f0184d900485}\" Name=\"ItemChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Item Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"ItemChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ItemChildCount\" FromBaseType=\"TRUE\"/><Field ID=\"{960ff01f-2b6d-4f1b-9c3f-e19ad8927341}\" Name=\"FolderChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Folder Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"FolderChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FolderChildCount\" FromBaseType=\"TRUE\"/></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>",
          "Scope": "/Shared Documents",
          "Sealed": true,
          "StringId": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds content type to the specific list if listId option is passed (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/ContentTypes/AddAvailableContentType`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0 &&
          JSON.stringify(opts.body) === JSON.stringify({
            contentTypeId: '0x0120'
          })) {
          return Promise.resolve({
            "Description": "Create a new folder.",
            "DisplayFormTemplateName": "ListForm",
            "DisplayFormUrl": "",
            "DocumentTemplate": "",
            "DocumentTemplateUrl": "",
            "EditFormTemplateName": "ListForm",
            "EditFormUrl": "",
            "Group": "Folder Content Types",
            "Hidden": false,
            "Id": {
              "StringValue": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
            },
            "JSLink": "",
            "MobileDisplayFormUrl": "",
            "MobileEditFormUrl": "",
            "MobileNewFormUrl": "",
            "Name": "Folder",
            "NewFormTemplateName": "ListForm",
            "NewFormUrl": "",
            "ReadOnly": false,
            "SchemaXml": "<ContentType ID=\"0x0120000EAD53EDAD7C6647B0D976EEC953F99E\" Name=\"Folder\" Group=\"Folder Content Types\" Description=\"Create a new folder.\" Sealed=\"TRUE\" Version=\"0\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" FeatureId=\"{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}\"><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" ShowInNewForm=\"FALSE\" ShowInFileDlg=\"FALSE\" DisplayName=\"Title\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" ColName=\"nvarchar8\" Required=\"FALSE\" Hidden=\"TRUE\"/><Field ID=\"{8553196d-ec8d-4564-9861-3dbe931050c8}\" ShowInFileDlg=\"FALSE\" ShowInVersionHistory=\"FALSE\" Type=\"File\" Name=\"FileLeafRef\" DisplayName=\"Name\" AuthoringInfo=\"(for use in forms)\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"LeafName\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" Required=\"TRUE\" NoCustomize=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FileLeafRef\" FromBaseType=\"TRUE\" Hidden=\"FALSE\"/><Field ID=\"{b824e17e-a1b3-426e-aecf-f0184d900485}\" Name=\"ItemChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Item Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"ItemChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ItemChildCount\" FromBaseType=\"TRUE\"/><Field ID=\"{960ff01f-2b6d-4f1b-9c3f-e19ad8927341}\" Name=\"FolderChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Folder Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"FolderChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FolderChildCount\" FromBaseType=\"TRUE\"/></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>",
            "Scope": "/Shared Documents",
            "Sealed": true,
            "StringId": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        contentTypeId: '0x0120'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": "Create a new folder.",
          "DisplayFormTemplateName": "ListForm",
          "DisplayFormUrl": "",
          "DocumentTemplate": "",
          "DocumentTemplateUrl": "",
          "EditFormTemplateName": "ListForm",
          "EditFormUrl": "",
          "Group": "Folder Content Types",
          "Hidden": false,
          "Id": {
            "StringValue": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
          },
          "JSLink": "",
          "MobileDisplayFormUrl": "",
          "MobileEditFormUrl": "",
          "MobileNewFormUrl": "",
          "Name": "Folder",
          "NewFormTemplateName": "ListForm",
          "NewFormUrl": "",
          "ReadOnly": false,
          "SchemaXml": "<ContentType ID=\"0x0120000EAD53EDAD7C6647B0D976EEC953F99E\" Name=\"Folder\" Group=\"Folder Content Types\" Description=\"Create a new folder.\" Sealed=\"TRUE\" Version=\"0\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" FeatureId=\"{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}\"><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" ShowInNewForm=\"FALSE\" ShowInFileDlg=\"FALSE\" DisplayName=\"Title\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" ColName=\"nvarchar8\" Required=\"FALSE\" Hidden=\"TRUE\"/><Field ID=\"{8553196d-ec8d-4564-9861-3dbe931050c8}\" ShowInFileDlg=\"FALSE\" ShowInVersionHistory=\"FALSE\" Type=\"File\" Name=\"FileLeafRef\" DisplayName=\"Name\" AuthoringInfo=\"(for use in forms)\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"LeafName\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" Required=\"TRUE\" NoCustomize=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FileLeafRef\" FromBaseType=\"TRUE\" Hidden=\"FALSE\"/><Field ID=\"{b824e17e-a1b3-426e-aecf-f0184d900485}\" Name=\"ItemChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Item Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"ItemChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ItemChildCount\" FromBaseType=\"TRUE\"/><Field ID=\"{960ff01f-2b6d-4f1b-9c3f-e19ad8927341}\" Name=\"FolderChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Folder Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"FolderChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FolderChildCount\" FromBaseType=\"TRUE\"/></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>",
          "Scope": "/Shared Documents",
          "Sealed": true,
          "StringId": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds content type to the specific list if listId option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/ContentTypes/AddAvailableContentType`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0 &&
          JSON.stringify(opts.body) === JSON.stringify({
            contentTypeId: '0x0120'
          })) {
          return Promise.resolve({
            "Description": "Create a new folder.",
            "DisplayFormTemplateName": "ListForm",
            "DisplayFormUrl": "",
            "DocumentTemplate": "",
            "DocumentTemplateUrl": "",
            "EditFormTemplateName": "ListForm",
            "EditFormUrl": "",
            "Group": "Folder Content Types",
            "Hidden": false,
            "Id": {
              "StringValue": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
            },
            "JSLink": "",
            "MobileDisplayFormUrl": "",
            "MobileEditFormUrl": "",
            "MobileNewFormUrl": "",
            "Name": "Folder",
            "NewFormTemplateName": "ListForm",
            "NewFormUrl": "",
            "ReadOnly": false,
            "SchemaXml": "<ContentType ID=\"0x0120000EAD53EDAD7C6647B0D976EEC953F99E\" Name=\"Folder\" Group=\"Folder Content Types\" Description=\"Create a new folder.\" Sealed=\"TRUE\" Version=\"0\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" FeatureId=\"{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}\"><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" ShowInNewForm=\"FALSE\" ShowInFileDlg=\"FALSE\" DisplayName=\"Title\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" ColName=\"nvarchar8\" Required=\"FALSE\" Hidden=\"TRUE\"/><Field ID=\"{8553196d-ec8d-4564-9861-3dbe931050c8}\" ShowInFileDlg=\"FALSE\" ShowInVersionHistory=\"FALSE\" Type=\"File\" Name=\"FileLeafRef\" DisplayName=\"Name\" AuthoringInfo=\"(for use in forms)\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"LeafName\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" Required=\"TRUE\" NoCustomize=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FileLeafRef\" FromBaseType=\"TRUE\" Hidden=\"FALSE\"/><Field ID=\"{b824e17e-a1b3-426e-aecf-f0184d900485}\" Name=\"ItemChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Item Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"ItemChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ItemChildCount\" FromBaseType=\"TRUE\"/><Field ID=\"{960ff01f-2b6d-4f1b-9c3f-e19ad8927341}\" Name=\"FolderChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Folder Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"FolderChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FolderChildCount\" FromBaseType=\"TRUE\"/></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>",
            "Scope": "/Shared Documents",
            "Sealed": true,
            "StringId": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        contentTypeId: '0x0120'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": "Create a new folder.",
          "DisplayFormTemplateName": "ListForm",
          "DisplayFormUrl": "",
          "DocumentTemplate": "",
          "DocumentTemplateUrl": "",
          "EditFormTemplateName": "ListForm",
          "EditFormUrl": "",
          "Group": "Folder Content Types",
          "Hidden": false,
          "Id": {
            "StringValue": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
          },
          "JSLink": "",
          "MobileDisplayFormUrl": "",
          "MobileEditFormUrl": "",
          "MobileNewFormUrl": "",
          "Name": "Folder",
          "NewFormTemplateName": "ListForm",
          "NewFormUrl": "",
          "ReadOnly": false,
          "SchemaXml": "<ContentType ID=\"0x0120000EAD53EDAD7C6647B0D976EEC953F99E\" Name=\"Folder\" Group=\"Folder Content Types\" Description=\"Create a new folder.\" Sealed=\"TRUE\" Version=\"0\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" FeatureId=\"{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}\"><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" ShowInNewForm=\"FALSE\" ShowInFileDlg=\"FALSE\" DisplayName=\"Title\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" ColName=\"nvarchar8\" Required=\"FALSE\" Hidden=\"TRUE\"/><Field ID=\"{8553196d-ec8d-4564-9861-3dbe931050c8}\" ShowInFileDlg=\"FALSE\" ShowInVersionHistory=\"FALSE\" Type=\"File\" Name=\"FileLeafRef\" DisplayName=\"Name\" AuthoringInfo=\"(for use in forms)\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"LeafName\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" Required=\"TRUE\" NoCustomize=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FileLeafRef\" FromBaseType=\"TRUE\" Hidden=\"FALSE\"/><Field ID=\"{b824e17e-a1b3-426e-aecf-f0184d900485}\" Name=\"ItemChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Item Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"ItemChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ItemChildCount\" FromBaseType=\"TRUE\"/><Field ID=\"{960ff01f-2b6d-4f1b-9c3f-e19ad8927341}\" Name=\"FolderChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Folder Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"FolderChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FolderChildCount\" FromBaseType=\"TRUE\"/></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>",
          "Scope": "/Shared Documents",
          "Sealed": true,
          "StringId": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
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
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    const actionTitle: string = 'Documents';

    cmdInstance.action({
      options: {
        debug: true,
        listTitle: actionTitle,
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0120'
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

  it('uses correct API url when listTitle option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        return Promise.resolve({
          "Description": "Create a new folder.",
          "DisplayFormTemplateName": "ListForm",
          "DisplayFormUrl": "",
          "DocumentTemplate": "",
          "DocumentTemplateUrl": "",
          "EditFormTemplateName": "ListForm",
          "EditFormUrl": "",
          "Group": "Folder Content Types",
          "Hidden": false,
          "Id": {
            "StringValue": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
          },
          "JSLink": "",
          "MobileDisplayFormUrl": "",
          "MobileEditFormUrl": "",
          "MobileNewFormUrl": "",
          "Name": "Folder",
          "NewFormTemplateName": "ListForm",
          "NewFormUrl": "",
          "ReadOnly": false,
          "SchemaXml": "<ContentType ID=\"0x0120000EAD53EDAD7C6647B0D976EEC953F99E\" Name=\"Folder\" Group=\"Folder Content Types\" Description=\"Create a new folder.\" Sealed=\"TRUE\" Version=\"0\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" FeatureId=\"{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}\"><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" ShowInNewForm=\"FALSE\" ShowInFileDlg=\"FALSE\" DisplayName=\"Title\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" ColName=\"nvarchar8\" Required=\"FALSE\" Hidden=\"TRUE\"/><Field ID=\"{8553196d-ec8d-4564-9861-3dbe931050c8}\" ShowInFileDlg=\"FALSE\" ShowInVersionHistory=\"FALSE\" Type=\"File\" Name=\"FileLeafRef\" DisplayName=\"Name\" AuthoringInfo=\"(for use in forms)\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"LeafName\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" Required=\"TRUE\" NoCustomize=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FileLeafRef\" FromBaseType=\"TRUE\" Hidden=\"FALSE\"/><Field ID=\"{b824e17e-a1b3-426e-aecf-f0184d900485}\" Name=\"ItemChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Item Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"ItemChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ItemChildCount\" FromBaseType=\"TRUE\"/><Field ID=\"{960ff01f-2b6d-4f1b-9c3f-e19ad8927341}\" Name=\"FolderChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Folder Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"FolderChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FolderChildCount\" FromBaseType=\"TRUE\"/></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>",
          "Scope": "/Shared Documents",
          "Sealed": true,
          "StringId": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
        });
      }

      return Promise.reject('Invalid request');
    });

    const actionTitle: string = 'Documents';

    cmdInstance.action({
      options: {
        debug: false,
        listTitle: actionTitle,
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0120'
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

  it('uses correct API url when listId option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid') > -1) {
        return Promise.resolve({
          "Description": "Create a new folder.",
          "DisplayFormTemplateName": "ListForm",
          "DisplayFormUrl": "",
          "DocumentTemplate": "",
          "DocumentTemplateUrl": "",
          "EditFormTemplateName": "ListForm",
          "EditFormUrl": "",
          "Group": "Folder Content Types",
          "Hidden": false,
          "Id": {
            "StringValue": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
          },
          "JSLink": "",
          "MobileDisplayFormUrl": "",
          "MobileEditFormUrl": "",
          "MobileNewFormUrl": "",
          "Name": "Folder",
          "NewFormTemplateName": "ListForm",
          "NewFormUrl": "",
          "ReadOnly": false,
          "SchemaXml": "<ContentType ID=\"0x0120000EAD53EDAD7C6647B0D976EEC953F99E\" Name=\"Folder\" Group=\"Folder Content Types\" Description=\"Create a new folder.\" Sealed=\"TRUE\" Version=\"0\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" FeatureId=\"{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}\"><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" ShowInNewForm=\"FALSE\" ShowInFileDlg=\"FALSE\" DisplayName=\"Title\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" ColName=\"nvarchar8\" Required=\"FALSE\" Hidden=\"TRUE\"/><Field ID=\"{8553196d-ec8d-4564-9861-3dbe931050c8}\" ShowInFileDlg=\"FALSE\" ShowInVersionHistory=\"FALSE\" Type=\"File\" Name=\"FileLeafRef\" DisplayName=\"Name\" AuthoringInfo=\"(for use in forms)\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"LeafName\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" Required=\"TRUE\" NoCustomize=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FileLeafRef\" FromBaseType=\"TRUE\" Hidden=\"FALSE\"/><Field ID=\"{b824e17e-a1b3-426e-aecf-f0184d900485}\" Name=\"ItemChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Item Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"ItemChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ItemChildCount\" FromBaseType=\"TRUE\"/><Field ID=\"{960ff01f-2b6d-4f1b-9c3f-e19ad8927341}\" Name=\"FolderChildCount\" DisplaceOnUpgrade=\"TRUE\" ReadOnly=\"TRUE\" ShowInFileDlg=\"FALSE\" Type=\"Lookup\" DisplayName=\"Folder Child Count\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"FolderChildCount\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FolderChildCount\" FromBaseType=\"TRUE\"/></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>",
          "Scope": "/Shared Documents",
          "Sealed": true,
          "StringId": "0x0120000EAD53EDAD7C6647B0D976EEC953F99E"
        });
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    cmdInstance.action({
      options: {
        debug: false,
        listId: actionId,
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0120'
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

  it('fails validation if both listId and listTitle options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', contentTypeId: '0x0120' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', contentTypeId: '0x0120' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', contentTypeId: '0x0120' } });
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', contentTypeId: '0x0120' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', contentTypeId: '0x0120' } });
    assert(actual);
  });

  it('passes validation if the listTitle option is passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', contentTypeId: '0x0120' } });
    assert(actual);
  });

  it('fails validation if both listId and listTitle options are passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listTitle: 'Documents', contentTypeId: '0x0120' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the contentTypeId option is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listTitle: 'Documents' } });
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

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notStrictEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('configures contentTypeId as string option', () => {
    const types = (command.types() as CommandTypes);
    ['c', 'contentTypeId'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });
});