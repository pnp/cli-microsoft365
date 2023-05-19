import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
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
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
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
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CUSTOMACTION_CLEAR), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should user custom actions cleared successfully without prompting with confirmation argument', async () => {
    defaultPostCallsStub();

    await command.action(logger, {
      options: {
        verbose: false,
        webUrl: 'https://contoso.sharepoint.com',
        confirm: true
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('should prompt before clearing custom actions when confirmation argument not passed', async () => {
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('should abort custom actions clear when prompt not confirmed', async () => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com' } } as any);
    assert(postCallsSpy.notCalled);
  });

  it('should clear custom actions when prompt confirmed', async () => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();
    const clearScopedCustomActionsSpy = sinon.spy((command as any), 'clearScopedCustomActions');

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    try {
      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com' } } as any);
      assert(postCallsSpy.calledTwice);
      assert(clearScopedCustomActionsSpy.calledWith(sinon.match(
        {
          webUrl: 'https://contoso.sharepoint.com'
        })) === true);
    }
    finally {
      sinonUtil.restore((command as any)['clearScopedCustomActions']);
    }
  });

  it('should clearScopedCustomActions be called once when scope is Web', async () => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const clearScopedCustomActionsSpy = sinon.spy((command as any), 'clearScopedCustomActions');
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Web',
      confirm: true
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(postCallsSpy.calledOnce);
      assert(clearScopedCustomActionsSpy.calledWith({
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Web',
        confirm: true
      }), 'clearScopedCustomActionsSpy data error');
      assert(clearScopedCustomActionsSpy.calledOnce, 'clearScopedCustomActionsSpy calledOnce error');
    }
    finally {
      sinonUtil.restore((command as any)['clearScopedCustomActions']);
    }
  });

  it('should clearScopedCustomActions be called once when scope is Site', async () => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const clearScopedCustomActionsSpy = sinon.spy((command as any), 'clearScopedCustomActions');
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Site',
      confirm: true
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(postCallsSpy.calledOnce === true);
      assert(clearScopedCustomActionsSpy.calledWith(
        {
          webUrl: 'https://contoso.sharepoint.com',
          scope: 'Site',
          confirm: true
        }), 'clearScopedCustomActionsSpy data error');
      assert(clearScopedCustomActionsSpy.calledOnce, 'clearScopedCustomActionsSpy calledOnce error');
    }
    finally {
      sinonUtil.restore((command as any)['clearScopedCustomActions']);
    }
  });

  it('should clearScopedCustomActions be called twice when scope is All', async () => {
    defaultPostCallsStub();

    const clearScopedCustomActionsSpy = sinon.spy((command as any), 'clearScopedCustomActions');

    try {
      await command.action(logger, {
        options: {
          debug: true,
          webUrl: 'https://contoso.sharepoint.com',
          confirm: true
        }
      });
      assert(clearScopedCustomActionsSpy.calledTwice);
    }
    finally {
      sinonUtil.restore((command as any)['clearScopedCustomActions']);
    }
  });

  it('should clearScopedCustomActions be calledTwice when scope is All', async () => {
    defaultPostCallsStub();

    const clearScopedCustomActionsSpy: sinon.SinonSpy = sinon.spy((command as any), 'clearScopedCustomActions');
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      confirm: true
    };


    try {
      await command.action(logger, { options: options } as any);
      assert(clearScopedCustomActionsSpy.calledTwice, 'clearScopedCustomActionsSpy.calledTwice');
    }
    finally {
      sinonUtil.restore((command as any)['clearScopedCustomActions']);
    }
  });

  it('should the post calls be have the correct endpoint urls when scope is All', async () => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      confirm: true
    };

    await command.action(logger, { options: options } as any);
    assert(postCallsSpy.calledWith(sinon.match({
      url: 'https://contoso.sharepoint.com/_api/Web/UserCustomActions/clear'
    })));
    assert(postCallsSpy.calledWith(sinon.match({
      url: 'https://contoso.sharepoint.com/_api/Site/UserCustomActions/clear'
    })));
  });

  it('should correctly handle custom action reject request (web)', async () => {
    const err = 'abc error';

    sinon.stub(request, 'post').callsFake((opts) => {
      // fakes clear custom actions success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions/clear') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All',
        confirm: true
      }
    }), new CommandError(err));
  });

  it('should correctly handle custom action reject request (site)', async () => {
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

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All',
        confirm: true
      }
    }), new CommandError(err));
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
        webUrl: 'foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should pass validation when the url options specified', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: "https://contoso.sharepoint.com"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should pass validation when the url and scope options specified', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: "https://contoso.sharepoint.com",
        scope: "Site"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should accept scope to be All', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: "https://contoso.sharepoint.com",
        scope: 'All'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should accept scope to be Site', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: "https://contoso.sharepoint.com",
        scope: 'Site'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should accept scope to be Web', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: "https://contoso.sharepoint.com",
        scope: 'Web'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should reject invalid string scope', async () => {
    const scope = 'foo';
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com",
        scope: scope
      }
    }, commandInfo);
    assert.strictEqual(actual, `${scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
  });

  it('should reject invalid scope value specified as number', async () => {
    const scope = 123;
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com",
        scope: scope
      }
    }, commandInfo);
    assert.strictEqual(actual, `${scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
  });
});
