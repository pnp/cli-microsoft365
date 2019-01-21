import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./feature-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.FEATURE_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;
  let stubAuth: any = () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/common/oauth2/token') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });
  }

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
      auth.ensureAccessToken,
      auth.restoreAuth,
      request.get
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.FEATURE_LIST), true);
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
        assert.equal(telemetry.name, commands.FEATURE_LIST);
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
    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/abc' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves available features from site collection', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve({
          value: [
            {
              DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              DisplayName: "TenantSitesList"
            },
            {
              DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
              DisplayName: "Ratings"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        verbose: false,
        url: 'https://contoso.sharepoint.com',
        scope: 'Site'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
            DisplayName: "TenantSitesList"
          },
          {
            DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
            DisplayName: "Ratings"
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post,
          request.get
        ]);
      }
    });
  });

  it('retrieves available features from site', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve({
          value: [
            {
              DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              DisplayName: "TenantSitesList"
            },
            {
              DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
              DisplayName: "Ratings"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        verbose: false,
        url: 'https://contoso.sharepoint.com',
        scope: 'Web'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
            DisplayName: "TenantSitesList"
          },
          {
            DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
            DisplayName: "Ratings"
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post,
          request.get
        ]);
      }
    });
  });

  it('retrieves available features from site (default) when no scope is entered', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.reject('Invalid request');
      }

      if (opts.url.indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve({
          value: [
            {
              DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              DisplayName: "TenantSitesList"
            },
            {
              DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
              DisplayName: "Ratings"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        verbose: false,
        url: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
            DisplayName: "TenantSitesList"
          },
          {
            DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
            DisplayName: "Ratings"
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post,
          request.get
        ]);
      }
    });
  });

  it('returns all properties for output JSON', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve({
          value: [
            {
              "odata.type": "SP.Feature",
              "odata.id": "https://contoso.sharepoint.com/_api/Site/Features/GetById(guid'3019c9b4-e371-438d-98f6-0a08c34d06eb')",
              "odata.editLink": "Site/Features/GetById(guid'3019c9b4-e371-438d-98f6-0a08c34d06eb')",
              "DefinitionId": "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              "DisplayName": "TenantSitesList"
            },
            {
              "odata.type": "SP.Feature",
              "odata.id": "https://contoso.sharepoint.com/_api/Site/Features/GetById(guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3')",
              "odata.editLink": "Site/Features/GetById(guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3')",
              "DefinitionId": "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
              "DisplayName": "Ratings"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const options: Object = {
      debug: true,
      url: 'https://contoso.sharepoint.com',
      scope: 'Site',
      output: 'json'
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(
          [
            {
              "odata.type": "SP.Feature",
              "odata.id": "https://contoso.sharepoint.com/_api/Site/Features/GetById(guid'3019c9b4-e371-438d-98f6-0a08c34d06eb')",
              "odata.editLink": "Site/Features/GetById(guid'3019c9b4-e371-438d-98f6-0a08c34d06eb')",
              "DefinitionId": "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              "DisplayName": "TenantSitesList"
            },
            {
              "odata.type": "SP.Feature",
              "odata.id": "https://contoso.sharepoint.com/_api/Site/Features/GetById(guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3')",
              "odata.editLink": "Site/Features/GetById(guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3')",
              "DefinitionId": "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
              "DisplayName": "Ratings"
            }
          ]));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore(request.get);
      }
    });
  });

  it('correctly handles no features in site collection', (done) => {
    stubAuth();
    sinon.stub(request, 'get').callsFake((opts) => {

      if (opts.url.indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve(JSON.stringify({ value: [] }));
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const options: Object = {
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: 'Site',
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.equal(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore(request.get);
      }
    });
  });

  it('correctly handles no features in site', (done) => {
    stubAuth();
    sinon.stub(request, 'get').callsFake((opts) => {

      if (opts.url.indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve(JSON.stringify({ value: [] }));
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const options: Object = {
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: 'Web',
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.equal(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore(request.get);
      }
    });
  });

  it('correctly handles no features in site collection (verbose)', (done) => {
    stubAuth();
    sinon.stub(request, 'get').callsFake((opts) => {

      if (opts.url.indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve(JSON.stringify({ value: [] }));
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const options: Object = {
      verbose: true,
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: 'Site',
    }
    cmdInstance.action({ options: options }, () => {
      let correctLogStatement = false;
      log.forEach(l => {
        if (!l || typeof l !== 'string') {
          return;
        }

        if (l.indexOf('No activated Features found') > -1) {
          correctLogStatement = true;
        }
      })
      try {
        assert(correctLogStatement);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore(request.get);
      }
    });
  });

  it('correctly handles no features in site (verbose)', (done) => {
    stubAuth();
    sinon.stub(request, 'get').callsFake((opts) => {

      if (opts.url.indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve(JSON.stringify({ value: [] }));
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const options: Object = {
      verbose: true,
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: 'Web',
    }
    cmdInstance.action({ options: options }, () => {
      let correctLogStatement = false;
      log.forEach(l => {
        if (!l || typeof l !== 'string') {
          return;
        }

        if (l.indexOf('No activated Features found') > -1) {
          correctLogStatement = true;
        }
      })
      try {
        assert(correctLogStatement);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore(request.get);
      }
    });
  });

  it('correctly handles web feature reject request', (done) => {
    stubAuth();

    const err = 'Invalid web Features reject request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: false,
        url: 'https://contoso.sharepoint.com',
        scope: 'Web'
      }
    }, (error?: any) => {
      try {
        assert.equal(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post,
          request.get
        ]);
      }
    });
  });

  it('correctly handles site Features reject request', (done) => {
    stubAuth();

    const err = 'Invalid site Features reject request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: false,
        verbose: true,
        url: 'https://contoso.sharepoint.com',
        scope: 'Site'
      }
    }, (error?: any) => {
      try {
        assert.equal(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore(request.get);
        Utils.restore(cmdInstance.log);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsVerboseOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsVerboseOption = true;
      }
    });
    assert(containsVerboseOption);
  });

  it('supports specifying scope', () => {
    const options = (command.options() as CommandOption[]);
    let containsScopeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[scope]') > -1) {
        containsScopeOption = true;
      }
    });
    assert(containsScopeOption);
  });

  it('doesn\'t fail if the parent doesn\'t define options', () => {
    sinon.stub(Command.prototype, 'options').callsFake(() => { return undefined; });
    const options = (command.options() as CommandOption[]);
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.equal(actual, "Required parameter url missing");
  });

  it('retrieves all Web features', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve({
          value: [
            {
              DefinitionId: "00bfea71-5932-4f9c-ad71-1557e5751100",
              DisplayName: "WebPageLibrary"
            }]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/abc', scope: 'Web' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            DefinitionId: '00bfea71-5932-4f9c-ad71-1557e5751100',
            DisplayName: 'WebPageLibrary'
          }]
        ));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('retrieves all site features', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {

      if (opts.url.indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return Promise.resolve({
          value: [
            {
              DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              DisplayName: "TenantSitesList"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/abc', scope: 'Site' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            DefinitionId: '3019c9b4-e371-438d-98f6-0a08c34d06eb',
            DisplayName: 'TenantSitesList'
          }]
        ));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        url: 'foo'
      }
    });
    assert.notEqual(actual, true);
  });

  it('passes validation when the url options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        url: "https://contoso.sharepoint.com"
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation when the url and scope options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        url: "https://contoso.sharepoint.com",
        scope: "Site"
      }
    });
    assert.equal(actual, true);
  });

  it('accepts scope to be Site', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        url: "https://contoso.sharepoint.com",
        scope: 'Site'
      }
    });
    assert.equal(actual, true);
  });

  it('accepts scope to be Web', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        url: "https://contoso.sharepoint.com",
        scope: 'Web'
      }
    });
    assert.equal(actual, true);
  });

  it('rejects invalid string scope', () => {
    const scope = 'foo';
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: "https://contoso.sharepoint.com",
        scope: scope
      }
    });
    assert.equal(actual, `${scope} is not a valid Feature scope. Allowed values are Site|Web`);
  });

  it('rejects invalid scope value specified as number', () => {
    const scope = 123;
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: "https://contoso.sharepoint.com",
        scope: scope
      }
    });
    assert.equal(actual, `${scope} is not a valid Feature scope. Allowed values are Site|Web`);
  });

  it('doesn\'t fail validation if the optional scope option not specified', () => {
    const actual = (command.validate() as CommandValidate)(
      {
        options:
        {
          url: "https://contoso.sharepoint.com"
        }
      });
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
    assert(find.calledWith(commands.FEATURE_LIST));
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
        url: "https://contoso.sharepoint.com",
        debug: false
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(auth.getAccessToken);
      }
    });
  });
});