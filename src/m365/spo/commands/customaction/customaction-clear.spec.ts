import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './customaction-clear.js';

describe(commands.CUSTOMACTION_CLEAR, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let promptIssued: boolean = false;
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
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
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
      options: commandOptionsSchema.parse({
        verbose: false,
        webUrl: 'https://contoso.sharepoint.com',
        force: true
      })
    });
    assert(loggerLogSpy.notCalled);
  });

  it('should prompt before clearing custom actions when confirmation argument not passed', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com' }) });

    assert(promptIssued);
  });

  it('should abort custom actions clear when prompt not confirmed', async () => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com' }) });
    assert(postCallsSpy.notCalled);
  });

  it('should clear custom actions when prompt confirmed', async () => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();
    const clearScopedCustomActionsSpy = sinon.spy((command as any), 'clearScopedCustomActions');

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    try {
      await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com' }) });
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
      await command.action(logger, { options: commandOptionsSchema.parse(options) });
      assert(postCallsSpy.calledOnce);
      assert(clearScopedCustomActionsSpy.calledWithMatch({
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
      await command.action(logger, { options: commandOptionsSchema.parse(options) });
      assert(postCallsSpy.calledOnce === true);
      assert(clearScopedCustomActionsSpy.calledWithMatch(
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
        options: commandOptionsSchema.parse({
          debug: true,
          webUrl: 'https://contoso.sharepoint.com',
          force: true
        })
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
      await command.action(logger, { options: commandOptionsSchema.parse(options) });
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

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
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
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All',
        force: true
      })
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
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All',
        force: true
      })
    }), new CommandError(err));
  });

  it('should fail validation if the url option not specified', () => {
    const actual = commandOptionsSchema.safeParse({ scope: "Web" });
    assert.strictEqual(actual.success, false);
  });

  it('should fail validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'foo'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, 'foo is not a valid SharePoint Online site URL.');
  });

  it('should pass validation when the url options specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com"
    });
    assert.strictEqual(actual.success, true);
  });

  it('should pass validation when the url and scope options specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: "Site"
    });
    assert.strictEqual(actual.success, true);
  });

  it('should accept scope to be All', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: 'All'
    });
    assert.strictEqual(actual.success, true);
  });

  it('should accept scope to be Site', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: 'Site'
    });
    assert.strictEqual(actual.success, true);
  });

  it('should accept scope to be Web', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: 'Web'
    });
    assert.strictEqual(actual.success, true);
  });

  it('should reject invalid string scope', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: 'foo'
    });
    assert.strictEqual(actual.success, false);
  });

  it('should reject invalid scope value specified as number', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: 123
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});
