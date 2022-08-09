import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./customaction-clear');

describe(commands.CUSTOMACTION_CLEAR, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
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
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
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

    command.action(logger, { options: {
      verbose: false,
      url: 'https://contoso.sharepoint.com',
      confirm: true
    } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should prompt before clearing custom actions when confirmation argument not passed', (done) => {
    command.action(logger, { options: { url: 'https://contoso.sharepoint.com'} }, () => {
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
    
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, { options: {  url: 'https://contoso.sharepoint.com'}} as any, () => {
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

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: {  url: 'https://contoso.sharepoint.com' }} as any, () => {
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
        sinonUtil.restore((command as any)['clearScopedCustomActions']);
      }
    });
  });

  it('should clearScopedCustomActions be called once when scope is Web', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const clearScopedCustomActionsSpy = sinon.spy((command as any), 'clearScopedCustomActions');
    const options = {
      debug: false,
      url: 'https://contoso.sharepoint.com',
      scope: 'Web',
      confirm: true
    };

    command.action(logger, { options: options } as any, () => {
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
        sinonUtil.restore((command as any)['clearScopedCustomActions']);
      }
    });
  });

  it('should clearScopedCustomActions be called once when scope is Site', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const clearScopedCustomActionsSpy = sinon.spy((command as any), 'clearScopedCustomActions');
    const options = {
      url: 'https://contoso.sharepoint.com',
      scope: 'Site',
      confirm: true
    };

    command.action(logger, { options: options } as any, () => {
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
        sinonUtil.restore((command as any)['clearScopedCustomActions']);
      }
    });
  });

  it('should clearScopedCustomActions be called twice when scope is All', (done) => {
    defaultPostCallsStub();

    const clearScopedCustomActionsSpy = sinon.spy((command as any), 'clearScopedCustomActions');

    command.action(logger, {
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
        sinonUtil.restore((command as any)['clearScopedCustomActions']);
      }
    });
  });

  it('should clearScopedCustomActions be calledTwice when scope is All', (done) => {
    defaultPostCallsStub();

    const clearScopedCustomActionsSpy: sinon.SinonSpy = sinon.spy((command as any), 'clearScopedCustomActions');
    const options = {
      url: 'https://contoso.sharepoint.com',
      confirm: true
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(clearScopedCustomActionsSpy.calledTwice, 'clearScopedCustomActionsSpy.calledTwice');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore((command as any)['clearScopedCustomActions']);
      }
    });
  });

  it('should the post calls be have the correct endpoint urls when scope is All', (done) => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const options = {
      url: 'https://contoso.sharepoint.com',
      confirm: true
    };

    command.action(logger, { options: options } as any, () => {
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

    command.action(logger, {
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

    command.action(logger, {
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
    const options = command.options;
    let containsVerboseOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsVerboseOption = true;
      }
    });
    assert(containsVerboseOption);
  });

  it('supports specifying scope', () => {
    const options = command.options;
    let containsScopeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[scope]') > -1) {
        containsScopeOption = true;
      }
    });
    assert(containsScopeOption);
  });

  it('should fail validation if the url option not specified', async () => {
    const actual = await command.validate({ options: { scope: "Web" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should fail validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options:
        {
          url: 'foo'
        }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should pass validation when the url options specified', async () => {
    const actual = await command.validate({
      options:
        {
          url: "https://contoso.sharepoint.com"
        }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should pass validation when the url and scope options specified', async () => {
    const actual = await command.validate({
      options:
        {
          url: "https://contoso.sharepoint.com",
          scope: "Site"
        }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should accept scope to be All', async () => {
    const actual = await command.validate({
      options:
        {
          url: "https://contoso.sharepoint.com",
          scope: 'All'
        }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should accept scope to be Site', async () => {
    const actual = await command.validate({
      options:
        {
          url: "https://contoso.sharepoint.com",
          scope: 'Site'
        }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should accept scope to be Web', async () => {
    const actual = await command.validate({
      options:
        {
          url: "https://contoso.sharepoint.com",
          scope: 'Web'
        }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should reject invalid string scope', async () => {
    const scope = 'foo';
    const actual = await command.validate({
      options: {
        url: "https://contoso.sharepoint.com",
        scope: scope
      }
    }, commandInfo);
    assert.strictEqual(actual, `${scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
  });

  it('should reject invalid scope value specified as number', async () => {
    const scope = 123;
    const actual = await command.validate({
      options: {
        url: "https://contoso.sharepoint.com",
        scope: scope
      }
    }, commandInfo);
    assert.strictEqual(actual, `${scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
  });
});