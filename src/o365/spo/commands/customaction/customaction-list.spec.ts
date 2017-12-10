import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./customaction-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.CUSTOMACTION_LIST, () => {
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

      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
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
    assert.equal(command.name.startsWith(commands.CUSTOMACTION_LIST), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {}, appCatalogUrl: 'https://contoso-admin.sharepoint.com' }, () => {
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
    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      try {
        assert.equal(telemetry.name, commands.CUSTOMACTION_LIST);
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
    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/abc' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('getCustomActions called once when scope is Web', (done) => {
    stubAuth();

    let getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const getCustomActionsSpy = sinon.spy((command as any), 'getCustomActions');
    const options: Object = {
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: 'Web'
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(getRequestSpy.calledOnce == true);
        assert(getCustomActionsSpy.calledWith({
          debug: false,
          url: 'https://contoso.sharepoint.com',
          scope: 'Web'
        }, 'ABC', cmdInstance));
        assert(getCustomActionsSpy.calledOnce == true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore(request.get);
        Utils.restore((command as any)['getCustomActions']);
      }
    });
  });

  it('getCustomActions called once when scope is Site', (done) => {
    stubAuth();

    let getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const getCustomActionsSpy = sinon.spy((command as any), 'getCustomActions');
    const options: Object = {
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: 'Site'
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(getRequestSpy.calledOnce == true);
        assert(getCustomActionsSpy.calledWith({
          debug: false,
          url: 'https://contoso.sharepoint.com',
          scope: 'Site'
        }, 'ABC', cmdInstance));
        assert(getCustomActionsSpy.calledOnce == true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore(request.get);
        Utils.restore((command as any)['getCustomActions']);
      }
    });
  });

  it('getCustomActions called twice when scope is All', (done) => {
    stubAuth();

    let getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      if (opts.url.indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const getCustomActionsSpy = sinon.spy((command as any), 'getCustomActions');

    cmdInstance.action({
      options: {
        debug: true,
        url: 'https://contoso.sharepoint.com'
      }
    }, () => {

      try {
        assert(getRequestSpy.calledTwice == true);
        assert(getCustomActionsSpy.calledTwice == true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore(request.get);
        Utils.restore((command as any)['getCustomActions']);
      }
    });
  });

  it('searchAllScopes called when scope is All', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const searchAllScopesSpy = sinon.spy((command as any), 'searchAllScopes');
    const options: Object = {
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: "All"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(searchAllScopesSpy.calledWith(sinon.match(
          {
            url: 'https://contoso.sharepoint.com'
          }), 'ABC', cmdInstance));
        assert(searchAllScopesSpy.calledOnce == true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore(request.get);
        Utils.restore((command as any)['searchAllScopes']);
      }
    });
  });

  it('searchAllScopes correctly handles no custom actions when All scope specified', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      if (opts.url.indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
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
        scope: 'All'
      }
    }, () => {

      try {
        assert(cmdInstanceLogSpy.notCalled);
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

  it('correctly handles no custom actions when All scope specified (verbose)', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      if (opts.url.indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
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
        scope: 'All'
      }
    }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(`Custom actions not found`));
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

  it('correctly handles web custom action reject request', (done) => {
    stubAuth();

    const err = 'Invalid web custom action reject request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions') > -1) {
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
        scope: 'All'
      }
    }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError(err)));
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

  it('correctly handles site custom action reject request', (done) => {
    stubAuth();

    const err = 'Invalid site custom action reject request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      if (opts.url.indexOf('/_api/Site/UserCustomActions') > -1) {
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
        scope: 'All'
      }
    }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError(err)));
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
    assert.equal(actual, "Missing required option url");
  });

  it('retrieves all available user custom actions', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({
          value: [
            {
              Id: "9f7ade35-0f8d-4c8a-82e1-5008ab42df55",
              Location: "Microsoft.SharePoint.StandardMenu",
              Name: "customaction1",
              Scope: 3
            }]
        });
      }

      if (opts.url.indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({
          value: [
            {
              Id: "9f7ade35-0f8d-4c8a-82e1-5008ab42df56",
              Location: "Microsoft.SharePoint.StandardMenu",
              Name: "customaction2",
              Scope: 2
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/abc' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Name: 'customaction2',
            Location: 'Microsoft.SharePoint.StandardMenu',
            Scope: 'Site',
            Id: '9f7ade35-0f8d-4c8a-82e1-5008ab42df56'
          },
          {
            Name: 'customaction1',
            Location: 'Microsoft.SharePoint.StandardMenu',
            Scope: 'Web',
            Id: '9f7ade35-0f8d-4c8a-82e1-5008ab42df55'
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
    assert(actual);
  });

  it('passes validation when the url and scope options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          url: "https://contoso.sharepoint.com",
          scope: "Site"
        }
    });
    assert(actual);
  });

  it('humanize scope shows correct value when scope odata is 2', () => {
    const actual = (command as any)["humanizeScope"](2);
    assert(actual === "Site");
  });

  it('humanize scope shows correct value when scope odata is 3', () => {
    const actual = (command as any)["humanizeScope"](3);
    assert(actual === "Web");
  });

  it('humanize scope shows the scope odata value when is different than 2 and 3', () => {
    const actual = (command as any)["humanizeScope"](1);
    assert(actual === "1");
  });

  it('accepts scope to be All', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          url: "https://contoso.sharepoint.com",
          scope: 'All'
        }
    });
    assert(actual);
  });

  it('accepts scope to be Site', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          url: "https://contoso.sharepoint.com",
          scope: 'Site'
        }
    });
    assert(actual);
  });

  it('accepts scope to be Web', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          url: "https://contoso.sharepoint.com",
          scope: 'Web'
        }
    });
    assert(actual);
  });

  it('rejects invalid string scope', () => {
    const scope = 'foo';
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: "https://contoso.sharepoint.com",
        scope: scope
      }
    });
    assert.equal(actual, `${scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
  });

  it('rejects invalid scope value specified as number', () => {
    const scope = 123;
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: "https://contoso.sharepoint.com",
        scope: scope
      }
    });
    assert.equal(actual, `${scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
  });

  it('doesn\'t fail validation if the optional scope option not specified', () => {
    const actual = (command.validate() as CommandValidate)(
      {
        options:
          {
            url: "https://contoso.sharepoint.com"
          }
      });
    assert(actual);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => {},
      prompt: () => {},
      helpInformation: () => {}
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => {});
    assert(find.calledWith(commands.CUSTOMACTION_LIST));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => {},
      helpInformation: () => {}
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => {});
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
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
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