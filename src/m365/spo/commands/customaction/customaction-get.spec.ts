import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./customaction-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.CUSTOMACTION_GET, () => {
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
    assert.strictEqual(command.name.startsWith(commands.CUSTOMACTION_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves and prints all details user custom actions', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
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
    });
  });

  it('getCustomAction called once when scope is Web', (done) => {
    const getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const getCustomActionSpy = sinon.spy((command as any), 'getCustomAction');
    const options: Object = {
      debug: false,
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      scope: 'Web'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(getRequestSpy.calledOnce, 'getRequestSpy.calledOnce');
        assert(getCustomActionSpy.calledWith({
          debug: false,
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          url: 'https://contoso.sharepoint.com',
          scope: 'Web'
        }), 'getCustomActionSpy.calledWith');
        assert(getCustomActionSpy.calledOnce == true, 'getCustomActionSpy.calledOnce');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['getCustomAction']);
      }
    });
  });

  it('getCustomAction called once when scope is Site', (done) => {
    const getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const getCustomActionSpy = sinon.spy((command as any), 'getCustomAction');
    const options: Object = {
      debug: false,
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      scope: 'Site'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(getRequestSpy.calledOnce, 'getRequestSpy.calledOnce');
        assert(getCustomActionSpy.calledWith(
          {
            debug: false,
            id: 'b2307a39-e878-458b-bc90-03bc578531d6',
            url: 'https://contoso.sharepoint.com',
            scope: 'Site'
          }), 'getCustomActionSpy.calledWith');
        assert(getCustomActionSpy.calledOnce, 'getCustomActionSpy.calledOnce');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['getCustomAction']);
      }
    });
  });

  it('getCustomAction called once when scope is All, but item found on web level', (done) => {
    const getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

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
        assert(getRequestSpy.calledOnce);
        assert(getCustomActionSpy.calledOnce);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['getCustomAction']);
      }
    });
  });

  it('getCustomAction called twice when scope is All, but item not found on web level', (done) => {
    let getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const getCustomActionSpy = sinon.spy((command as any), 'getCustomAction');

    cmdInstance.action({
      options: {
        debug: true,
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        url: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
        assert(getRequestSpy.calledTwice);
        assert(getCustomActionSpy.calledTwice);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['getCustomAction']);
      }
    });
  });

  it('searchAllScopes called when scope is All', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

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
          })), 'searchAllScopesSpy.calledWith');
        assert(searchAllScopesSpy.calledOnce, 'searchAllScopesSpy.calledOnce');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['searchAllScopes']);
      }
    });
  });

  it('searchAllScopes correctly handles custom action odata.null when All scope specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

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
    });
  });

  it('searchAllScopes correctly handles custom action odata.null when All scope specified (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

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
    });
  });

  it('searchAllScopes correctly handles web custom action reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    cmdInstance.action({
      options: {
        debug: false,
        id: actionId,
        url: 'https://contoso.sharepoint.com',
        scope: 'All'
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

  it('searchAllScopes correctly handles site custom action reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    cmdInstance.action({
      options: {
        debug: false,
        verbose: true,
        id: actionId,
        url: 'https://contoso.sharepoint.com',
        scope: 'All'
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
    sinon.stub(Command.prototype, 'options').callsFake(() => { return []; });
    const options = (command.options() as CommandOption[]);
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
          url: 'foo'
        }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid guid', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          id: "foo",
          url: 'https://contoso.sharepoint.com'
        }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id and url options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
          url: "https://contoso.sharepoint.com"
        }
    });
    assert.strictEqual(actual, true);
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
    assert.strictEqual(actual, true);
  });

  it('passes validation when the id and url option specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
          url: "https://contoso.sharepoint.com"
        }
    });
    assert.strictEqual(actual, true);
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
    assert.strictEqual(actual, true);
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
    assert.strictEqual(actual, true);
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
    assert.strictEqual(actual, true);
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
    assert.strictEqual(actual, `${scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
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
    assert.strictEqual(actual, `${scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
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
    assert.strictEqual(actual, true);
  });
});