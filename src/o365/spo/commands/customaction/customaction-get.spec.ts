import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./customaction-get');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.CUSTOMACTION_GET, () => {
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
      auth.getAccessToken,
      auth.restoreAuth,
      request.get
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.CUSTOMACTION_GET), true);
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
    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/appcatalog', id: "abc" } }, () => {
      try {
        assert.equal(telemetry.name, commands.CUSTOMACTION_GET);
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
    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/appcatalog', id: "abc" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves and prints all details user custom actions', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve(
          {
            "ClientSideComponentId": "015e0fcf-fe9d-4037-95af-0a4776cdfbb4",
            "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}",
            "CommandUIExtension": null,
            "Description": null,
            "Group": null,
            "Id": "d26af83a-6421-4bb3-9f5c-8174ba645c80",
            "ImageUrl": null,
            "Location": "ClientSideExtension.ApplicationCustomizer",
            "Name": "{d26af83a-6421-4bb3-9f5c-8174ba645c80}",
            "RegistrationId": null,
            "RegistrationType": 0,
            "Rights": { "High": 0, "Low": 0 },
            "Scope": "1",
            "ScriptBlock": null,
            "ScriptSrc": null,
            "Sequence": 65536,
            "Title": "Places",
            "Url": null,
            "VersionOfUserCustomAction": "1.0.1.0"
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {
      debug: false,
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com'
    } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ 
          ClientSideComponentId: '015e0fcf-fe9d-4037-95af-0a4776cdfbb4',
          ClientSideComponentProperties: '{"testMessage":"Test message"}',
          CommandUIExtension: null,
          Description: null,
          Group: null,
          Id: 'd26af83a-6421-4bb3-9f5c-8174ba645c80',
          ImageUrl: null,
          Location: 'ClientSideExtension.ApplicationCustomizer',
          Name: '{d26af83a-6421-4bb3-9f5c-8174ba645c80}',
          RegistrationId: null,
          RegistrationType: 0,
          Rights: '{"High":0,"Low":0}',
          Scope: '1',
          ScriptBlock: null,
          ScriptSrc: null,
          Sequence: 65536,
          Title: 'Places',
          Url: null,
          VersionOfUserCustomAction: '1.0.1.0' 
        }));
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

  it('getCustomAction called once when scope is Web', (done) => {
    stubAuth();

    let getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const getCustomActionSpy = sinon.spy((command as any), 'getCustomAction');
    const options: Object = {
      debug: false,
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      scope: 'Web'
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(getRequestSpy.calledOnce == true, 'getRequestSpy.calledOnce');
        assert(getCustomActionSpy.calledWith({
          debug: false,
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          url: 'https://contoso.sharepoint.com',
          scope: 'Web'
        }, 'ABC', cmdInstance), 'getCustomActionSpy.calledWith');
        assert(getCustomActionSpy.calledOnce == true, 'getCustomActionSpy.calledOnce');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore(request.get);
        Utils.restore((command as any)['getCustomAction']);
      }
    });
  });

  it('getCustomAction called once when scope is Site', (done) => {
    stubAuth();

    let getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const getCustomActionSpy = sinon.spy((command as any), 'getCustomAction');
    const options: Object = {
      debug: false,
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      scope: 'Site'
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(getRequestSpy.calledOnce == true, 'getRequestSpy.calledOnce');
        assert(getCustomActionSpy.calledWith(
          {
            debug: false,
            id: 'b2307a39-e878-458b-bc90-03bc578531d6',
            url: 'https://contoso.sharepoint.com',
            scope: 'Site'
          }, 'ABC', cmdInstance), 'getCustomActionSpy.calledWith');
        assert(getCustomActionSpy.calledOnce == true, 'getCustomActionSpy.calledOnce');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore(request.get);
        Utils.restore((command as any)['getCustomAction']);
      }
    });
  });

  it('getCustomAction called once when scope is All, but item found on web level', (done) => {
    stubAuth();

    let getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const getCustomActionSpy = sinon.spy((command as any), 'getCustomAction');

    cmdInstance.action({
      options: {
        debug: false,
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        url: 'https://contoso.sharepoint.com',
        scope: 'All'
      }
    }, () => {

      try {
        assert(getRequestSpy.calledOnce == true);
        assert(getCustomActionSpy.calledOnce == true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore(request.get);
        Utils.restore((command as any)['getCustomAction']);
      }
    });
  });

  it('getCustomAction called twice when scope is All, but item not found on web level', (done) => {
    stubAuth();

    let getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if (opts.url.indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const getCustomActionSpy = sinon.spy((command as any), 'getCustomAction');

    cmdInstance.action({
      options: {
        debug: true,
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        url: 'https://contoso.sharepoint.com'
      }
    }, () => {

      try {
        assert(getRequestSpy.calledTwice == true);
        assert(getCustomActionSpy.calledTwice == true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore(request.get);
        Utils.restore((command as any)['getCustomAction']);
      }
    });
  });

  it('searchAllScopes called when scope is All', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions(') > -1) {
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
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      scope: "All"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(searchAllScopesSpy.calledWith(sinon.match(
          {
            id: 'b2307a39-e878-458b-bc90-03bc578531d6',
            url: 'https://contoso.sharepoint.com'
          }), 'ABC', cmdInstance), 'searchAllScopesSpy.calledWith');
        assert(searchAllScopesSpy.calledOnce == true, 'searchAllScopesSpy.calledOnce');
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

  it('searchAllScopes correctly handles custom action odata.null when All scope specified', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if (opts.url.indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    cmdInstance.action({
      options: {
        debug: false,
        verbose: false,
        id: actionId,
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

  it('searchAllScopes correctly handles custom action odata.null when All scope specified (verbose)', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if (opts.url.indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    cmdInstance.action({
      options: {
        debug: false,
        verbose: true,
        id: actionId,
        url: 'https://contoso.sharepoint.com',
        scope: 'All'
      }
    }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(`Custom action with id ${actionId} not found`));
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

  it('searchAllScopes correctly handles web custom action reject request', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }

      return Promise.reject('Invalid request');
    });

    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    cmdInstance.action({
      options: {
        debug: false,
        id: actionId,
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

  it('searchAllScopes correctly handles site custom action reject request', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }

      return Promise.reject('Invalid request');
    });

    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if (opts.url.indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    cmdInstance.action({
      options: {
        debug: false,
        verbose: true,
        id: actionId,
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

  it('fails validation if the id option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: "https://contoso.sharepoint.com" } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: "BC448D63-484F-49C5-AB8C-96B14AA68D50" } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
          url: 'foo'
        }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if the id option is not a valid guid', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          id: "foo",
          url: 'https://contoso.sharepoint.com'
        }
    });
    assert.notEqual(actual, true);
  });

  it('passes validation when the id and url options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
          url: "https://contoso.sharepoint.com"
        }
    });
    assert.equal(actual, true);
  });

  it('passes validation when the id, url and scope options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
          url: "https://contoso.sharepoint.com",
          scope: "Site"
        }
    });
    assert.equal(actual, true);
  });

  it('passes validation when the id and url option specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
          url: "https://contoso.sharepoint.com"
        }
    });
    assert.equal(actual, true);
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
          id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
          url: "https://contoso.sharepoint.com",
          scope: 'All'
        }
    });
    assert.equal(actual, true);
  });

  it('accepts scope to be Site', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
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
          id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
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
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
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
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
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
            id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
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
    assert(find.calledWith(commands.CUSTOMACTION_GET));
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
    });
  });
});