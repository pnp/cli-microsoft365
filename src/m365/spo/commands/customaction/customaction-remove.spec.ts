import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./customaction-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.CUSTOMACTION_REMOVE, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let promptOptions: any;
  const defaultPostCallsStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve(undefined);
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve(undefined);
      }

      return Promise.reject('Invalid request');
    });
  }

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
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        promptOptions = options;
        cb({ continue: false });
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    promptOptions = undefined;
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
    assert.strictEqual(command.name.startsWith(commands.CUSTOMACTION_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should user custom action removed successfully without prompting with confirmation argument', (done) => {
    defaultPostCallsStub();

    cmdInstance.action({ options: {
      verbose: false,
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      confirm: true
    } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should user custom action removed successfully (verbose) without prompting with confirmation argument', (done) => {
    defaultPostCallsStub();

    cmdInstance.action({ options: {
      verbose: true,
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      confirm: true
    } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(sinon.match('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should prompt before removing custom action when confirmation argument not passed', (done) => {
    cmdInstance.action({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', url: 'https://contoso.sharepoint.com'} }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should abort custom action remove when prompt not confirmed', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', url: 'https://contoso.sharepoint.com'}}, () => {
      try {
        assert(postCallsSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should remove custom action when prompt confirmed', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();
    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', url: 'https://contoso.sharepoint.com' }}, () => {
      try {
        assert(postCallsSpy.calledOnce);
        assert(removeScopedCustomActionSpy.calledWith(sinon.match(
          { 
            id: 'b2307a39-e878-458b-bc90-03bc578531d6',
            url: 'https://contoso.sharepoint.com'
          })));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['removeScopedCustomAction']);
      }
    });
  });

  it('should removeScopedCustomAction be called once when scope is Web', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');
    const options: Object = {
      debug: false,
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      scope: 'Web',
      confirm: true
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(postCallsSpy.calledOnce);
        assert(removeScopedCustomActionSpy.calledWith({
          debug: false,
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          url: 'https://contoso.sharepoint.com',
          scope: 'Web',
          confirm: true
        }), 'removeScopedCustomActionSpy data error');
        assert(removeScopedCustomActionSpy.calledOnce, 'removeScopedCustomActionSpy calledOnce error');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['removeScopedCustomAction']);
      }
    });
  });

  it('should removeScopedCustomAction be called once when scope is Site', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');
    const options: Object = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      scope: 'Site',
      confirm: true
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(postCallsSpy.calledOnce);
        assert(removeScopedCustomActionSpy.calledWith(
          {
            id: 'b2307a39-e878-458b-bc90-03bc578531d6',
            url: 'https://contoso.sharepoint.com',
            scope: 'Site',
            confirm: true
          }), 'removeScopedCustomActionSpy data error');
        assert(removeScopedCustomActionSpy.calledOnce, 'removeScopedCustomActionSpy calledOnce error');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['removeScopedCustomAction']);
      }
    });
  });

  it('should removeScopedCustomAction be called once when scope is All, but item found on web level', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');

    cmdInstance.action({
      options: {
        confirm: true,
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        url: 'https://contoso.sharepoint.com',
        scope: 'All'
      }
    }, () => {
      try {
        assert(postCallsSpy.calledOnce);
        assert(removeScopedCustomActionSpy.calledOnce);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['removeScopedCustomAction']);
      }
    });
  });

  it('should removeScopedCustomAction be called twice when scope is All, but item not found on web level', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve(undefined);
      }

      return Promise.reject('Invalid request');
    });

    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');

    cmdInstance.action({
      options: {
        debug: true,
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        url: 'https://contoso.sharepoint.com',
        confirm: true
      }
    }, () => {
      try {
        assert(removeScopedCustomActionSpy.calledTwice);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['removeScopedCustomAction']);
      }
    });
  });

  it('should searchAllScopes be called when scope is All', (done) => {
    defaultPostCallsStub();

    const searchAllScopesSpy = sinon.spy((command as any), 'searchAllScopes');
    const options: Object = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      confirm: true
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(searchAllScopesSpy.calledWith(sinon.match(
          {
            id: 'b2307a39-e878-458b-bc90-03bc578531d6',
            url: 'https://contoso.sharepoint.com',
            confirm: true
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

  it('should searchAllScopes correctly handles custom action odata.null when All scope specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    cmdInstance.action({
      options: {
        verbose: false,
        id: actionId,
        url: 'https://contoso.sharepoint.com',
        scope: 'All',
        confirm: true
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

  it('should searchAllScopes correctly handles custom action odata.null when All scope specified (verbose)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    cmdInstance.action({
      options: {
        verbose: true,
        id: actionId,
        url: 'https://contoso.sharepoint.com',
        scope: 'All',
        confirm: true
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

  it('should correctly handle custom action reject request (web)', (done) => {
    const err = 'abc error';

    sinon.stub(request, 'post').callsFake((opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    cmdInstance.action({
      options: {
        id: actionId,
        url: 'https://contoso.sharepoint.com',
        scope: 'All',
        confirm: true
      }
    }, (error: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle custom action reject request (site)', (done) => {
    const err = 'abc error';

    sinon.stub(request, 'post').callsFake((opts) => {
      // should return null to proceed with site when scope is All
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
        id: actionId,
        url: 'https://contoso.sharepoint.com',
        scope: 'All',
        confirm: true
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

  it('should fail validation if the id option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: "https://contoso.sharepoint.com" } });
    assert.notStrictEqual(actual, true);
  });

  it('should fail validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: "BC448D63-484F-49C5-AB8C-96B14AA68D50" } });
    assert.notStrictEqual(actual, true);
  });

  it('should fail validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
          url: 'foo'
        }
    });
    assert.notStrictEqual(actual, true);
  });

  it('should fail validation if the id option is not a valid guid', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          id: "foo",
          url: 'https://contoso.sharepoint.com'
        }
    });
    assert.notStrictEqual(actual, true);
  });

  it('should pass validation when the id and url options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
          url: "https://contoso.sharepoint.com"
        }
    });
    assert.strictEqual(actual, true);
  });

  it('should pass validation when the id, url and scope options specified', () => {
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

  it('should pass validation when the id and url option specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
          url: "https://contoso.sharepoint.com"
        }
    });
    assert.strictEqual(actual, true);
  });

  it('should accept scope to be All', () => {
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

  it('should accept scope to be Site', () => {
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

  it('should accept scope to be Web', () => {
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

  it('should reject invalid string scope', () => {
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

  it('should reject invalid scope value specified as number', () => {
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