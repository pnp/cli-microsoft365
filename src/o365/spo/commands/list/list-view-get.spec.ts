import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./list-view-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_VIEW_GET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
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
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.LIST_VIEW_GET), true);
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
        assert.equal(telemetry.name, commands.LIST_VIEW_GET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified list doesn\'t exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-2130575322, Microsoft.SharePoint.SPException",
            "message": {
              "lang": "en-US",
              "value": "List does not exist.\n\nThe page you selected contains a list that does not exist. It may have been deleted by another user."
            }
          }
        }
      })
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', listTitle: 'List', viewTitle: 'All items' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("List does not exist.\n\nThe page you selected contains a list that does not exist. It may have been deleted by another user.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified view doesn\'t exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-2147024809, System.ArgumentException",
            "message": {
              "lang": "en-US",
              "value": "The specified view is invalid."
            }
          }
        }
      })
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', listTitle: 'List', viewTitle: 'All Items' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("The specified view is invalid.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should successfully get the list view with specified its Id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {

      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('List%201')/views/getById('ba84217c-8561-4234-aa95-265081e74be9')`) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ "Aggregations": null, "AggregationsStatus": null, "BaseViewId": "1", "ColumnWidth": null, "ContentTypeId": { "StringValue": "0x" }, "CustomFormatter": null, "DefaultView": true, "DefaultViewForContentType": false, "EditorModified": false, "Formats": null, "Hidden": false, "HtmlSchemaXml": "<View Name=\"{BA84217C-8561-4234-AA95-265081E74BE9}\" DefaultView=\"TRUE\" MobileView=\"TRUE\" MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Items\" Url=\"/Lists/l2/AllItems.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=45\"><Toolbar Type=\"Standard\" /><XslLink Default=\"TRUE\">main.xsl</XslLink><JSLink>clienttemplates.js</JSLink><RowLimit Paged=\"TRUE\">30</RowLimit><ViewFields><FieldRef Name=\"LinkTitle\" /></ViewFields><Query><OrderBy><FieldRef Name=\"ID\" /></OrderBy></Query><ParameterBindings><ParameterBinding Name=\"NoAnnouncements\" Location=\"Resource(wss,noXinviewofY_LIST)\" /><ParameterBinding Name=\"NoAnnouncementsHowTo\" Location=\"Resource(wss,noXinviewofY_DEFAULT)\" /></ParameterBindings></View>", "Id": "ba84217c-8561-4234-aa95-265081e74be9", "ImageUrl": "/_layouts/15/images/generic.png?rev=45", "IncludeRootFolder": false, "ViewJoins": null, "JSLink": "clienttemplates.js", "ListViewXml": "<View Name=\"{BA84217C-8561-4234-AA95-265081E74BE9}\" DefaultView=\"TRUE\" MobileView=\"TRUE\" MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Items\" Url=\"/Lists/l2/AllItems.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=45\" ><Query><OrderBy><FieldRef Name=\"ID\" /></OrderBy></Query><ViewFields><FieldRef Name=\"LinkTitle\" /></ViewFields><RowLimit Paged=\"TRUE\">30</RowLimit><JSLink>clienttemplates.js</JSLink><XslLink Default=\"TRUE\">main.xsl</XslLink><Toolbar Type=\"Standard\"/></View>", "Method": null, "MobileDefaultView": true, "MobileView": true, "ModerationType": null, "NewDocumentTemplates": null, "OrderedView": false, "Paged": true, "PersonalView": false, "ViewProjectedFields": null, "ViewQuery": "<OrderBy><FieldRef Name=\"ID\" /></OrderBy>", "ReadOnlyView": false, "RequiresClientIntegration": false, "RowLimit": 30, "Scope": 0, "ServerRelativePath": { "DecodedUrl": "/Lists/l2/AllItems.aspx" }, "ServerRelativeUrl": "/Lists/l2/AllItems.aspx", "StyleId": null, "TabularView": true, "Threaded": false, "Title": "All Items", "Toolbar": "", "ToolbarTemplateName": null, "ViewType": "HTML", "ViewData": null, "VisualizationInfo": null });
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewId: 'ba84217c-8561-4234-aa95-265081e74be9' } }, () => {
      try {
        assert.equal(cmdInstanceLogSpy.lastCall.args[0].Id, 'ba84217c-8561-4234-aa95-265081e74be9');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should successfully get the list view with specified its name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === "https://contoso.sharepoint.com/_api/web/GetList('%2Flists%2FList1')/views/getByTitle('All%20Items')") {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ "Aggregations": null, "AggregationsStatus": null, "BaseViewId": "1", "ColumnWidth": null, "ContentTypeId": { "StringValue": "0x" }, "CustomFormatter": null, "DefaultView": true, "DefaultViewForContentType": false, "EditorModified": false, "Formats": null, "Hidden": false, "HtmlSchemaXml": "<View Name=\"{BA84217C-8561-4234-AA95-265081E74BE9}\" DefaultView=\"TRUE\" MobileView=\"TRUE\" MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Items\" Url=\"/Lists/l2/AllItems.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=45\"><Toolbar Type=\"Standard\" /><XslLink Default=\"TRUE\">main.xsl</XslLink><JSLink>clienttemplates.js</JSLink><RowLimit Paged=\"TRUE\">30</RowLimit><ViewFields><FieldRef Name=\"LinkTitle\" /></ViewFields><Query><OrderBy><FieldRef Name=\"ID\" /></OrderBy></Query><ParameterBindings><ParameterBinding Name=\"NoAnnouncements\" Location=\"Resource(wss,noXinviewofY_LIST)\" /><ParameterBinding Name=\"NoAnnouncementsHowTo\" Location=\"Resource(wss,noXinviewofY_DEFAULT)\" /></ParameterBindings></View>", "Id": "ba84217c-8561-4234-aa95-265081e74be9", "ImageUrl": "/_layouts/15/images/generic.png?rev=45", "IncludeRootFolder": false, "ViewJoins": null, "JSLink": "clienttemplates.js", "ListViewXml": "<View Name=\"{BA84217C-8561-4234-AA95-265081E74BE9}\" DefaultView=\"TRUE\" MobileView=\"TRUE\" MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Items\" Url=\"/Lists/l2/AllItems.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=45\" ><Query><OrderBy><FieldRef Name=\"ID\" /></OrderBy></Query><ViewFields><FieldRef Name=\"LinkTitle\" /></ViewFields><RowLimit Paged=\"TRUE\">30</RowLimit><JSLink>clienttemplates.js</JSLink><XslLink Default=\"TRUE\">main.xsl</XslLink><Toolbar Type=\"Standard\"/></View>", "Method": null, "MobileDefaultView": true, "MobileView": true, "ModerationType": null, "NewDocumentTemplates": null, "OrderedView": false, "Paged": true, "PersonalView": false, "ViewProjectedFields": null, "ViewQuery": "<OrderBy><FieldRef Name=\"ID\" /></OrderBy>", "ReadOnlyView": false, "RequiresClientIntegration": false, "RowLimit": 30, "Scope": 0, "ServerRelativePath": { "DecodedUrl": "/Lists/l2/AllItems.aspx" }, "ServerRelativeUrl": "/Lists/l2/AllItems.aspx", "StyleId": null, "TabularView": true, "Threaded": false, "Title": "All Items", "Toolbar": "", "ToolbarTemplateName": null, "ViewType": "HTML", "ViewData": null, "VisualizationInfo": null });
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', listUrl: 'lists/List1', viewTitle: 'All Items' } }, () => {
      try {
        assert.equal(cmdInstanceLogSpy.lastCall.args[0].Title, 'All Items');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should successfully get the list view with specified its name and list id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === "https://contoso.sharepoint.com/_api/web/lists(guid'dac05e4a-5f6c-41dd-bba3-2be1104c711e')/views/getById('ba84217c-8561-4234-aa95-265081e74be9')") {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ "Aggregations": null, "AggregationsStatus": null, "BaseViewId": "1", "ColumnWidth": null, "ContentTypeId": { "StringValue": "0x" }, "CustomFormatter": null, "DefaultView": true, "DefaultViewForContentType": false, "EditorModified": false, "Formats": null, "Hidden": false, "HtmlSchemaXml": "<View Name=\"{BA84217C-8561-4234-AA95-265081E74BE9}\" DefaultView=\"TRUE\" MobileView=\"TRUE\" MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Items\" Url=\"/Lists/l2/AllItems.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=45\"><Toolbar Type=\"Standard\" /><XslLink Default=\"TRUE\">main.xsl</XslLink><JSLink>clienttemplates.js</JSLink><RowLimit Paged=\"TRUE\">30</RowLimit><ViewFields><FieldRef Name=\"LinkTitle\" /></ViewFields><Query><OrderBy><FieldRef Name=\"ID\" /></OrderBy></Query><ParameterBindings><ParameterBinding Name=\"NoAnnouncements\" Location=\"Resource(wss,noXinviewofY_LIST)\" /><ParameterBinding Name=\"NoAnnouncementsHowTo\" Location=\"Resource(wss,noXinviewofY_DEFAULT)\" /></ParameterBindings></View>", "Id": "ba84217c-8561-4234-aa95-265081e74be9", "ImageUrl": "/_layouts/15/images/generic.png?rev=45", "IncludeRootFolder": false, "ViewJoins": null, "JSLink": "clienttemplates.js", "ListViewXml": "<View Name=\"{BA84217C-8561-4234-AA95-265081E74BE9}\" DefaultView=\"TRUE\" MobileView=\"TRUE\" MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Items\" Url=\"/Lists/l2/AllItems.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=45\" ><Query><OrderBy><FieldRef Name=\"ID\" /></OrderBy></Query><ViewFields><FieldRef Name=\"LinkTitle\" /></ViewFields><RowLimit Paged=\"TRUE\">30</RowLimit><JSLink>clienttemplates.js</JSLink><XslLink Default=\"TRUE\">main.xsl</XslLink><Toolbar Type=\"Standard\"/></View>", "Method": null, "MobileDefaultView": true, "MobileView": true, "ModerationType": null, "NewDocumentTemplates": null, "OrderedView": false, "Paged": true, "PersonalView": false, "ViewProjectedFields": null, "ViewQuery": "<OrderBy><FieldRef Name=\"ID\" /></OrderBy>", "ReadOnlyView": false, "RequiresClientIntegration": false, "RowLimit": 30, "Scope": 0, "ServerRelativePath": { "DecodedUrl": "/Lists/l2/AllItems.aspx" }, "ServerRelativeUrl": "/Lists/l2/AllItems.aspx", "StyleId": null, "TabularView": true, "Threaded": false, "Title": "All Items", "Toolbar": "", "ToolbarTemplateName": null, "ViewType": "HTML", "ViewData": null, "VisualizationInfo": null });
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com', listId: 'dac05e4a-5f6c-41dd-bba3-2be1104c711e', viewId: 'ba84217c-8561-4234-aa95-265081e74be9' } }, () => {
      try {
        assert.equal(cmdInstanceLogSpy.lastCall.args[0].Title, 'All Items');
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('fails validation if webUrl is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { listTitle: 'List 1', viewTitle: 'All items' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'invalid', listTitle: 'List 1', viewTitle: 'All items' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if neither listId nor listTitle nor listUrl specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', viewTitle: 'All items' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if listId is not a GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'invalid', viewTitle: 'All items' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if neither viewId nor viewTitle specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if both viewId and viewTitle specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', viewTitle: 'All items' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if viewId is not a GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewId: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when viewTitle and listTitle specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewTitle: 'All items' } });
    assert.equal(actual, true);
  });

  it('passes validation when viewId and listId specified and valid GUIDs', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', viewId: '330f29c5-5c4c-465f-9f4b-7903020ae1cf' } });
    assert.equal(actual, true);
  });

  it('passes validation when viewId and listUrl specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listUrl: 'lists/list1', viewId: '330f29c5-5c4c-465f-9f4b-7903020ae1cf' } });
    assert.equal(actual, true);
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
    assert(find.calledWith(commands.LIST_VIEW_GET));
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
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com' } }, (err?: any) => {
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