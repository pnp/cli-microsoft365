import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './customaction-clear.js';

describe(commands.CUSTOMACTION_CLEAR, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptOptions: any;
  const defaultPostCallsStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake(async (opts) => {
      // fakes clear custom actions success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions/clear') > -1) {
        return undefined;
      }

      // fakes clear custom actions success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions/clear') > -1) {
        return undefined;
      }

      throw 'Invalid request';
    });
  };

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);
    promptOptions = undefined;
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.promptForConfirmation,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CUSTOMACTION_CLEAR);
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
        force: true
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
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com' } } as any);
    assert(postCallsSpy.notCalled);
  });

  it('should clear custom actions when prompt confirmed', async () => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();
    const clearScopedCustomActionsSpy = sinon.spy((command as any), 'clearScopedCustomActions');

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

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
      force: true
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(postCallsSpy.calledOnce);
      assert(clearScopedCustomActionsSpy.calledWith({
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Web',
        force: true
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
      force: true
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(postCallsSpy.calledOnce === true);
      assert(clearScopedCustomActionsSpy.calledWith(
        {
          webUrl: 'https://contoso.sharepoint.com',
          scope: 'Site',
          force: true
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
          force: true
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
      force: true
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
      force: true
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

    sinon.stub(request, 'post').callsFake(async (opts) => {
      // fakes clear custom actions success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions/clear') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All',
        force: true
      }
    }), new CommandError(err));
  });

  it('should correctly handle custom action reject request (site)', async () => {
    const err = 'abc error';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      // should return null to proceed with site when scope is All
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions/clear') > -1) {
        return { "odata.null": true };
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions/clear') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All',
        force: true
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
