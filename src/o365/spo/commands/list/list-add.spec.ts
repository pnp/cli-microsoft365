import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./list-add');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.LIST_ADD, () => {
  let vorpal: Vorpal;
  let log: any[];
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
      request.post
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
    assert.equal(command.name.startsWith(commands.LIST_ADD), true);
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
        assert.equal(telemetry.name, commands.LIST_ADD);
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
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', title: 'Documents' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified title for list', (done) => {
    const expected = 'List 1';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.Title;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: expected, baseTemplate: 'GenericList', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified baseTemplate for list', (done) => {
    const expected = 100;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.BaseTemplate;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'List 1', baseTemplate: 'GenericList', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified description for list', (done) => {
    const expected = 'List 1 description';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.Description;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'List 1', description: expected, baseTemplate: 'GenericList', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified templateFeatureId for list', (done) => {
    const expected = '00bfea71-de22-43b2-a848-c05709900100';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.TemplateFeatureId;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'List 1', baseTemplate: 'GenericList', templateFeatureId: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified schemaXml for list', (done) => {
    const expected = `<List Title=\'List 1' ID='BE9CE88C-EF3A-4A61-9A8E-F8C038442227'></List>`;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.SchemaXml;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'List 1', baseTemplate: 'GenericList', schemaXml: expected, templateFeatureId: '00bfea71-de22-43b2-a848-c05709900100', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified allowDeletion for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.AllowDeletion;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'List 1', baseTemplate: 'GenericList', allowDeletion: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified allowEveryoneViewItems for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.AllowEveryoneViewItems;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'List 1', baseTemplate: 'GenericList', allowEveryoneViewItems: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified allowMultiResponses for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.AllowMultiResponses;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'List 1', baseTemplate: 'GenericList', allowMultiResponses: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified contentTypesEnabled for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.ContentTypesEnabled;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'List 1', baseTemplate: 'GenericList', contentTypesEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified crawlNonDefaultViews for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.CrawlNonDefaultViews;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'List 1', baseTemplate: 'GenericList', crawlNonDefaultViews: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified defaultContentApprovalWorkflowId for list', (done) => {
    const expected = '00bfea71-de22-43b2-a848-c05709900100';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.DefaultContentApprovalWorkflowId;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'List 1', baseTemplate: 'GenericList', defaultContentApprovalWorkflowId: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified defaultDisplayFormUrl for list', (done) => {
    const expected = '/sites/project-x/List%201/view.aspx';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.DefaultDisplayFormUrl;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'List 1', baseTemplate: 'GenericList', defaultDisplayFormUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified defaultEditFormUrl for list', (done) => {
    const expected = '/sites/project-x/List%201/edit.aspx';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.DefaultEditFormUrl;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'List 1', baseTemplate: 'GenericList', defaultEditFormUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified direction for list', (done) => {
    const expected = 'LTR';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.Direction;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'List 1', baseTemplate: 'GenericList', direction: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
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

  it('supports specifying URL', () => {
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('offers autocomplete for the baseTemplate option', () => {
    const options = (command.options() as CommandOption[]);
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--baseTemplate') > -1) {
        assert(options[i].autocomplete);
        return;
      }
    }
    assert(false);
  });

  it('offers autocomplete for the direction option', () => {
    const options = (command.options() as CommandOption[]);
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--direction') > -1) {
        assert(options[i].autocomplete);
        return;
      }
    }
    assert(false);
  });

  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { } });
    assert.notEqual(actual, true);
  });

  it('fails validation if both title and baseTemplate options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList' } });
    assert(actual);
  });

  it('fails validation if the templateFeatureId option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', templateFeatureId: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the templateFeatureId option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', templateFeatureId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails validation if the defaultContentApprovalWorkflowId option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', defaultContentApprovalWorkflowId: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the defaultContentApprovalWorkflowId option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', defaultContentApprovalWorkflowId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails if non existing direction specified', () => {
    const directionValue = 'abc';
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', direction: directionValue } });
    assert.equal(actual, `direction value ${directionValue} is not a valid value. Allowed values are NONE|LTR|RTL`);
  });

  it('has correct direction specified', () => {
    const directionValue = 'LTR';
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'List 1', baseTemplate: 'GenericList', direction: directionValue } });
    assert(actual === true);
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
    assert(find.calledWith(commands.LIST_ADD));
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
    cmdInstance.action({
      options: {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: "https://contoso.sharepoint.com",
        debug: false
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});