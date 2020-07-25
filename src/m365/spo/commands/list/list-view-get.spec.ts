import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./list-view-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_VIEW_GET, () => {
  let log: string[];
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
    assert.strictEqual(command.name.startsWith(commands.LIST_VIEW_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', listTitle: 'List', viewTitle: 'All items' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("List does not exist.\n\nThe page you selected contains a list that does not exist. It may have been deleted by another user.")));
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', listTitle: 'List', viewTitle: 'All Items' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("The specified view is invalid.")));
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
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ "Aggregations": null, "AggregationsStatus": null, "BaseViewId": "1", "ColumnWidth": null, "ContentTypeId": { "StringValue": "0x" }, "CustomFormatter": null, "DefaultView": true, "DefaultViewForContentType": false, "EditorModified": false, "Formats": null, "Hidden": false, "HtmlSchemaXml": "<View Name=\"{BA84217C-8561-4234-AA95-265081E74BE9}\" DefaultView=\"TRUE\" MobileView=\"TRUE\" MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Items\" Url=\"/Lists/l2/AllItems.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=45\"><Toolbar Type=\"Standard\" /><XslLink Default=\"TRUE\">main.xsl</XslLink><JSLink>clienttemplates.js</JSLink><RowLimit Paged=\"TRUE\">30</RowLimit><ViewFields><FieldRef Name=\"LinkTitle\" /></ViewFields><Query><OrderBy><FieldRef Name=\"ID\" /></OrderBy></Query><ParameterBindings><ParameterBinding Name=\"NoAnnouncements\" Location=\"Resource(wss,noXinviewofY_LIST)\" /><ParameterBinding Name=\"NoAnnouncementsHowTo\" Location=\"Resource(wss,noXinviewofY_DEFAULT)\" /></ParameterBindings></View>", "Id": "ba84217c-8561-4234-aa95-265081e74be9", "ImageUrl": "/_layouts/15/images/generic.png?rev=45", "IncludeRootFolder": false, "ViewJoins": null, "JSLink": "clienttemplates.js", "ListViewXml": "<View Name=\"{BA84217C-8561-4234-AA95-265081E74BE9}\" DefaultView=\"TRUE\" MobileView=\"TRUE\" MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Items\" Url=\"/Lists/l2/AllItems.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=45\" ><Query><OrderBy><FieldRef Name=\"ID\" /></OrderBy></Query><ViewFields><FieldRef Name=\"LinkTitle\" /></ViewFields><RowLimit Paged=\"TRUE\">30</RowLimit><JSLink>clienttemplates.js</JSLink><XslLink Default=\"TRUE\">main.xsl</XslLink><Toolbar Type=\"Standard\"/></View>", "Method": null, "MobileDefaultView": true, "MobileView": true, "ModerationType": null, "NewDocumentTemplates": null, "OrderedView": false, "Paged": true, "PersonalView": false, "ViewProjectedFields": null, "ViewQuery": "<OrderBy><FieldRef Name=\"ID\" /></OrderBy>", "ReadOnlyView": false, "RequiresClientIntegration": false, "RowLimit": 30, "Scope": 0, "ServerRelativePath": { "DecodedUrl": "/Lists/l2/AllItems.aspx" }, "ServerRelativeUrl": "/Lists/l2/AllItems.aspx", "StyleId": null, "TabularView": true, "Threaded": false, "Title": "All Items", "Toolbar": "", "ToolbarTemplateName": null, "ViewType": "HTML", "ViewData": null, "VisualizationInfo": null });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewId: 'ba84217c-8561-4234-aa95-265081e74be9' } }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].Id, 'ba84217c-8561-4234-aa95-265081e74be9');
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
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ "Aggregations": null, "AggregationsStatus": null, "BaseViewId": "1", "ColumnWidth": null, "ContentTypeId": { "StringValue": "0x" }, "CustomFormatter": null, "DefaultView": true, "DefaultViewForContentType": false, "EditorModified": false, "Formats": null, "Hidden": false, "HtmlSchemaXml": "<View Name=\"{BA84217C-8561-4234-AA95-265081E74BE9}\" DefaultView=\"TRUE\" MobileView=\"TRUE\" MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Items\" Url=\"/Lists/l2/AllItems.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=45\"><Toolbar Type=\"Standard\" /><XslLink Default=\"TRUE\">main.xsl</XslLink><JSLink>clienttemplates.js</JSLink><RowLimit Paged=\"TRUE\">30</RowLimit><ViewFields><FieldRef Name=\"LinkTitle\" /></ViewFields><Query><OrderBy><FieldRef Name=\"ID\" /></OrderBy></Query><ParameterBindings><ParameterBinding Name=\"NoAnnouncements\" Location=\"Resource(wss,noXinviewofY_LIST)\" /><ParameterBinding Name=\"NoAnnouncementsHowTo\" Location=\"Resource(wss,noXinviewofY_DEFAULT)\" /></ParameterBindings></View>", "Id": "ba84217c-8561-4234-aa95-265081e74be9", "ImageUrl": "/_layouts/15/images/generic.png?rev=45", "IncludeRootFolder": false, "ViewJoins": null, "JSLink": "clienttemplates.js", "ListViewXml": "<View Name=\"{BA84217C-8561-4234-AA95-265081E74BE9}\" DefaultView=\"TRUE\" MobileView=\"TRUE\" MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Items\" Url=\"/Lists/l2/AllItems.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=45\" ><Query><OrderBy><FieldRef Name=\"ID\" /></OrderBy></Query><ViewFields><FieldRef Name=\"LinkTitle\" /></ViewFields><RowLimit Paged=\"TRUE\">30</RowLimit><JSLink>clienttemplates.js</JSLink><XslLink Default=\"TRUE\">main.xsl</XslLink><Toolbar Type=\"Standard\"/></View>", "Method": null, "MobileDefaultView": true, "MobileView": true, "ModerationType": null, "NewDocumentTemplates": null, "OrderedView": false, "Paged": true, "PersonalView": false, "ViewProjectedFields": null, "ViewQuery": "<OrderBy><FieldRef Name=\"ID\" /></OrderBy>", "ReadOnlyView": false, "RequiresClientIntegration": false, "RowLimit": 30, "Scope": 0, "ServerRelativePath": { "DecodedUrl": "/Lists/l2/AllItems.aspx" }, "ServerRelativeUrl": "/Lists/l2/AllItems.aspx", "StyleId": null, "TabularView": true, "Threaded": false, "Title": "All Items", "Toolbar": "", "ToolbarTemplateName": null, "ViewType": "HTML", "ViewData": null, "VisualizationInfo": null });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', listUrl: 'lists/List1', viewTitle: 'All Items' } }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].Title, 'All Items');
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
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ "Aggregations": null, "AggregationsStatus": null, "BaseViewId": "1", "ColumnWidth": null, "ContentTypeId": { "StringValue": "0x" }, "CustomFormatter": null, "DefaultView": true, "DefaultViewForContentType": false, "EditorModified": false, "Formats": null, "Hidden": false, "HtmlSchemaXml": "<View Name=\"{BA84217C-8561-4234-AA95-265081E74BE9}\" DefaultView=\"TRUE\" MobileView=\"TRUE\" MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Items\" Url=\"/Lists/l2/AllItems.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=45\"><Toolbar Type=\"Standard\" /><XslLink Default=\"TRUE\">main.xsl</XslLink><JSLink>clienttemplates.js</JSLink><RowLimit Paged=\"TRUE\">30</RowLimit><ViewFields><FieldRef Name=\"LinkTitle\" /></ViewFields><Query><OrderBy><FieldRef Name=\"ID\" /></OrderBy></Query><ParameterBindings><ParameterBinding Name=\"NoAnnouncements\" Location=\"Resource(wss,noXinviewofY_LIST)\" /><ParameterBinding Name=\"NoAnnouncementsHowTo\" Location=\"Resource(wss,noXinviewofY_DEFAULT)\" /></ParameterBindings></View>", "Id": "ba84217c-8561-4234-aa95-265081e74be9", "ImageUrl": "/_layouts/15/images/generic.png?rev=45", "IncludeRootFolder": false, "ViewJoins": null, "JSLink": "clienttemplates.js", "ListViewXml": "<View Name=\"{BA84217C-8561-4234-AA95-265081E74BE9}\" DefaultView=\"TRUE\" MobileView=\"TRUE\" MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Items\" Url=\"/Lists/l2/AllItems.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=45\" ><Query><OrderBy><FieldRef Name=\"ID\" /></OrderBy></Query><ViewFields><FieldRef Name=\"LinkTitle\" /></ViewFields><RowLimit Paged=\"TRUE\">30</RowLimit><JSLink>clienttemplates.js</JSLink><XslLink Default=\"TRUE\">main.xsl</XslLink><Toolbar Type=\"Standard\"/></View>", "Method": null, "MobileDefaultView": true, "MobileView": true, "ModerationType": null, "NewDocumentTemplates": null, "OrderedView": false, "Paged": true, "PersonalView": false, "ViewProjectedFields": null, "ViewQuery": "<OrderBy><FieldRef Name=\"ID\" /></OrderBy>", "ReadOnlyView": false, "RequiresClientIntegration": false, "RowLimit": 30, "Scope": 0, "ServerRelativePath": { "DecodedUrl": "/Lists/l2/AllItems.aspx" }, "ServerRelativeUrl": "/Lists/l2/AllItems.aspx", "StyleId": null, "TabularView": true, "Threaded": false, "Title": "All Items", "Toolbar": "", "ToolbarTemplateName": null, "ViewType": "HTML", "ViewData": null, "VisualizationInfo": null });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com', listId: 'dac05e4a-5f6c-41dd-bba3-2be1104c711e', viewId: 'ba84217c-8561-4234-aa95-265081e74be9' } }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].Title, 'All Items');
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

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'invalid', listTitle: 'List 1', viewTitle: 'All items' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither listId nor listTitle nor listUrl specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', viewTitle: 'All items' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId is not a GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'invalid', viewTitle: 'All items' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither viewId nor viewTitle specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both viewId and viewTitle specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', viewTitle: 'All items' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if viewId is not a GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewId: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when viewTitle and listTitle specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewTitle: 'All items' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when viewId and listId specified and valid GUIDs', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', viewId: '330f29c5-5c4c-465f-9f4b-7903020ae1cf' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when viewId and listUrl specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listUrl: 'lists/list1', viewId: '330f29c5-5c4c-465f-9f4b-7903020ae1cf' } });
    assert.strictEqual(actual, true);
  });
});