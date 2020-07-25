import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./customaction-clear');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.CUSTOMACTION_CLEAR, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let promptOptions: any;
  const defaultPostCallsStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      // fakes clear custom actions success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions/clear') > -1) {
        return Promise.resolve(undefined);
      }

      // fakes clear custom actions success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions/clear') > -1) {
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
    assert.strictEqual(command.name.startsWith(commands.CUSTOMACTION_CLEAR), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should user custom actions cleared successfully without prompting with confirmation argument', (done) => {
    defaultPostCallsStub();

    cmdInstance.action({ options: {
      verbose: false,
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

  it('should user custom actions cleared successfully (verbose) without prompting with confirmation argument', (done) => {
    defaultPostCallsStub();

    cmdInstance.action({ options: {
      verbose: true,
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

  it('should prompt before clearing custom actions when confirmation argument not passed', (done) => {
    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com'} }, () => {
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

  it('should abort custom actions clear when prompt not confirmed', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: {  url: 'https://contoso.sharepoint.com'}}, () => {
      try {
        assert(postCallsSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should clear custom actions when prompt confirmed', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();
    const clearScopedCustomActionsSpy = sinon.spy((command as any), 'clearScopedCustomActions');

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: {  url: 'https://contoso.sharepoint.com' }}, () => {
      try {
        assert(postCallsSpy.calledTwice);
        assert(clearScopedCustomActionsSpy.calledWith(sinon.match(
          { 
            url: 'https://contoso.sharepoint.com'
          })) === true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['clearScopedCustomActions']);
      }
    });
  });

  it('should clearScopedCustomActions be called once when scope is Web', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const clearScopedCustomActionsSpy = sinon.spy((command as any), 'clearScopedCustomActions');
    const options: Object = {
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: 'Web',
      confirm: true
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(postCallsSpy.calledOnce);
        assert(clearScopedCustomActionsSpy.calledWith({
          debug: false,
          url: 'https://contoso.sharepoint.com',
          scope: 'Web',
          confirm: true
        }), 'clearScopedCustomActionsSpy data error');
        assert(clearScopedCustomActionsSpy.calledOnce, 'clearScopedCustomActionsSpy calledOnce error');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['clearScopedCustomActions']);
      }
    });
  });

  it('should clearScopedCustomActions be called once when scope is Site', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const clearScopedCustomActionsSpy = sinon.spy((command as any), 'clearScopedCustomActions');
    const options: Object = {
      url: 'https://contoso.sharepoint.com',
      scope: 'Site',
      confirm: true
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(postCallsSpy.calledOnce === true);
        assert(clearScopedCustomActionsSpy.calledWith(
          {
            url: 'https://contoso.sharepoint.com',
            scope: 'Site',
            confirm: true
          }), 'clearScopedCustomActionsSpy data error');
        assert(clearScopedCustomActionsSpy.calledOnce, 'clearScopedCustomActionsSpy calledOnce error');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['clearScopedCustomActions']);
      }
    });
  });

  it('should clearScopedCustomActions be called twice when scope is All', (done) => {
    defaultPostCallsStub();

    const clearScopedCustomActionsSpy = sinon.spy((command as any), 'clearScopedCustomActions');

    cmdInstance.action({
      options: {
        debug: true,
        url: 'https://contoso.sharepoint.com',
        confirm: true
      }
    }, () => {
      try {
        assert(clearScopedCustomActionsSpy.calledTwice);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['clearScopedCustomActions']);
      }
    });
  });

  it('should clearScopedCustomActions be calledTwice when scope is All', (done) => {
    defaultPostCallsStub();

    const clearScopedCustomActionsSpy: sinon.SinonSpy = sinon.spy((command as any), 'clearScopedCustomActions');
    const options: Object = {
      url: 'https://contoso.sharepoint.com',
      confirm: true
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(clearScopedCustomActionsSpy.calledTwice, 'clearScopedCustomActionsSpy.calledTwice');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any)['clearScopedCustomActions']);
      }
    });
  });

  it('should the post calls be have the correct endpoint urls when scope is All', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const options: Object = {
      url: 'https://contoso.sharepoint.com',
      confirm: true
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(postCallsSpy.calledWith(sinon.match({
          url: 'https://contoso.sharepoint.com/_api/Web/UserCustomActions/clear'
        })));
        assert(postCallsSpy.calledWith(sinon.match({
          url: 'https://contoso.sharepoint.com/_api/Site/UserCustomActions/clear'
        })));
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
      // fakes clear custom actions success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions/clear') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
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

  it('should correctly handle custom action reject request (site)', (done) => {
    const err = 'abc error';

    sinon.stub(request, 'post').callsFake((opts) => {
      // should return null to proceed with site when scope is All
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions/clear') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions/clear') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
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

  it('should fail validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { scope: "Web" } });
    assert.notStrictEqual(actual, true);
  });

  it('should fail validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          url: 'foo'
        }
    });
    assert.notStrictEqual(actual, true);
  });

  it('should pass validation when the url options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          url: "https://contoso.sharepoint.com"
        }
    });
    assert.strictEqual(actual, true);
  });

  it('should pass validation when the url and scope options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
          url: "https://contoso.sharepoint.com",
          scope: "Site"
        }
    });
    assert.strictEqual(actual, true);
  });

  it('should accept scope to be All', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
        {
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
        url: "https://contoso.sharepoint.com",
        scope: scope
      }
    });
    assert.strictEqual(actual, `${scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
  });
});