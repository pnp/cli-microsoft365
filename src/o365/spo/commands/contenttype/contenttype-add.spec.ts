import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError, CommandTypes } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./contenttype-add');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.CONTENTTYPE_ADD, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(command as any, 'getRequestDigestForSite').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
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
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth,
      (command as any).getRequestDigestForSite
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.CONTENTTYPE_ADD), true);
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
        assert.equal(telemetry.name, commands.CONTENTTYPE_ADD);
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
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates site content type', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/contenttypes`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          Name: 'PnP Tile',
          Id: { StringValue: '0x0100FF0B2E33A3718B46A3909298D240FD93' },
          Group: 'PnP Content Types'
        })) {
        return Promise.resolve({"Description":"Create a new list item.","DisplayFormTemplateName":"ListForm","DisplayFormUrl":"","DocumentTemplate":"","DocumentTemplateUrl":"","EditFormTemplateName":"ListForm","EditFormUrl":"","Group":"PnP Content Types","Hidden":false,"Id":{"StringValue":"0x010098998426EC27DF43841EF165675F4BF6"},"JSLink":"","MobileDisplayFormUrl":"","MobileEditFormUrl":"","MobileNewFormUrl":"","Name":"PnP Tile","NewFormTemplateName":"ListForm","NewFormUrl":"","ReadOnly":false,"SchemaXml":"<ContentType ID=\"0x010098998426EC27DF43841EF165675F4BF6\" Name=\"PnP Tile\" Group=\"PnP Content Types\" Description=\"Create a new list item.\" Version=\"1\"><Folder TargetName=\"_cts/PnP Tile\" /><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Name=\"ContentType\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" Type=\"Computed\" DisplayName=\"Content Type\" Sealed=\"TRUE\" Sortable=\"FALSE\" RenderXMLUsingPattern=\"TRUE\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" Customization=\"\"><FieldRefs><FieldRef ID=\"{03e45e84-1992-4d42-9116-26f756012634}\" Name=\"ContentTypeId\" /></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\" /></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Name=\"Title\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" Group=\"_Hidden\" Type=\"Text\" DisplayName=\"Title\" Required=\"TRUE\" FromBaseType=\"TRUE\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" Customization=\"\" ShowInNewForm=\"TRUE\" ShowInEditForm=\"TRUE\"></Field></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>","Scope":"/sites/portal","Sealed":false,"StringId":"0x010098998426EC27DF43841EF165675F4BF6"});
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Tile', id: '0x0100FF0B2E33A3718B46A3909298D240FD93', group: 'PnP Content Types' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({"Description":"Create a new list item.","DisplayFormTemplateName":"ListForm","DisplayFormUrl":"","DocumentTemplate":"","DocumentTemplateUrl":"","EditFormTemplateName":"ListForm","EditFormUrl":"","Group":"PnP Content Types","Hidden":false,"Id":{"StringValue":"0x010098998426EC27DF43841EF165675F4BF6"},"JSLink":"","MobileDisplayFormUrl":"","MobileEditFormUrl":"","MobileNewFormUrl":"","Name":"PnP Tile","NewFormTemplateName":"ListForm","NewFormUrl":"","ReadOnly":false,"SchemaXml":"<ContentType ID=\"0x010098998426EC27DF43841EF165675F4BF6\" Name=\"PnP Tile\" Group=\"PnP Content Types\" Description=\"Create a new list item.\" Version=\"1\"><Folder TargetName=\"_cts/PnP Tile\" /><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Name=\"ContentType\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" Type=\"Computed\" DisplayName=\"Content Type\" Sealed=\"TRUE\" Sortable=\"FALSE\" RenderXMLUsingPattern=\"TRUE\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" Customization=\"\"><FieldRefs><FieldRef ID=\"{03e45e84-1992-4d42-9116-26f756012634}\" Name=\"ContentTypeId\" /></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\" /></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Name=\"Title\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" Group=\"_Hidden\" Type=\"Text\" DisplayName=\"Title\" Required=\"TRUE\" FromBaseType=\"TRUE\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" Customization=\"\" ShowInNewForm=\"TRUE\" ShowInEditForm=\"TRUE\"></Field></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>","Scope":"/sites/portal","Sealed":false,"StringId":"0x010098998426EC27DF43841EF165675F4BF6"}));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates list content type (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists/getByTitle('Documents')/contenttypes`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          Name: 'PnP Tile',
          Id: { StringValue: '0x0100FF0B2E33A3718B46A3909298D240FD93' },
          Description: 'Create a new menu tile'
        })) {
        return Promise.resolve({"Description":"Create a new menu tile","DisplayFormTemplateName":"ListForm","DisplayFormUrl":"","DocumentTemplate":"template.dotx","DocumentTemplateUrl":"/sites/portal/Shared Documents/Forms/PnP Tile/template.dotx","EditFormTemplateName":"ListForm","EditFormUrl":"","Group":"List Content Types","Hidden":false,"Id":{"StringValue":"0x01007F34F00FE277BA438CCAA9B9FDA73CEC"},"JSLink":"","MobileDisplayFormUrl":"","MobileEditFormUrl":"","MobileNewFormUrl":"","Name":"PnP Tile","NewFormTemplateName":"ListForm","NewFormUrl":"","ReadOnly":false,"SchemaXml":"<ContentType ID=\"0x01007F34F00FE277BA438CCAA9B9FDA73CEC\" Name=\"PnP Tile\" Group=\"List Content Types\" Description=\"Defines menu tile\" Version=\"2\"><Folder TargetName=\"Forms/PnP Tile\"/><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" ShowInNewForm=\"TRUE\" ShowInFileDlg=\"FALSE\" DisplayName=\"Title\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" ColName=\"nvarchar8\" Required=\"TRUE\" ShowInEditForm=\"TRUE\"/></Fields><DocumentTemplate TargetName=\"Forms/PnP Tile/template.dotx\"/><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>","Scope":"/sites/portal/Shared Documents","Sealed":false,"StringId":"0x01007F34F00FE277BA438CCAA9B9FDA73CEC"});
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Tile', listTitle: 'Documents', id: '0x0100FF0B2E33A3718B46A3909298D240FD93', description: 'Create a new menu tile' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({"Description":"Create a new menu tile","DisplayFormTemplateName":"ListForm","DisplayFormUrl":"","DocumentTemplate":"template.dotx","DocumentTemplateUrl":"/sites/portal/Shared Documents/Forms/PnP Tile/template.dotx","EditFormTemplateName":"ListForm","EditFormUrl":"","Group":"List Content Types","Hidden":false,"Id":{"StringValue":"0x01007F34F00FE277BA438CCAA9B9FDA73CEC"},"JSLink":"","MobileDisplayFormUrl":"","MobileEditFormUrl":"","MobileNewFormUrl":"","Name":"PnP Tile","NewFormTemplateName":"ListForm","NewFormUrl":"","ReadOnly":false,"SchemaXml":"<ContentType ID=\"0x01007F34F00FE277BA438CCAA9B9FDA73CEC\" Name=\"PnP Tile\" Group=\"List Content Types\" Description=\"Defines menu tile\" Version=\"2\"><Folder TargetName=\"Forms/PnP Tile\"/><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" ShowInNewForm=\"TRUE\" ShowInFileDlg=\"FALSE\" DisplayName=\"Title\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" ColName=\"nvarchar8\" Required=\"TRUE\" ShowInEditForm=\"TRUE\"/></Fields><DocumentTemplate TargetName=\"Forms/PnP Tile/template.dotx\"/><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>","Scope":"/sites/portal/Shared Documents","Sealed":false,"StringId":"0x01007F34F00FE277BA438CCAA9B9FDA73CEC"}));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when creating content type', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'PnP Tile', id: '0x010098998426EC27DF43841EF165675F4BF6' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('fails validation if site URL is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'PnP Tile', id: '0x0100FF0B2E33A3718B46A3909298D240FD93' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'site.com', name: 'PnP Tile', id: '0x0100FF0B2E33A3718B46A3909298D240FD93' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the content type name is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', id: '0x0100FF0B2E33A3718B46A3909298D240FD93' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the content type id is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Tile' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Tile', id: '0x0100FF0B2E33A3718B46A3909298D240FD93' } });
    assert.equal(actual, true);
  });

  it('configures command types', () => {
    assert.notEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('configures id as string option', () => {
    const types = (command.types() as CommandTypes);
    ['i', 'id'].forEach(o => {
      assert.notEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
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
    assert(find.calledWith(commands.CONTENTTYPE_ADD));
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

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Tile', id: '0x0100FF0B2E33A3718B46A3909298D240FD93' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});