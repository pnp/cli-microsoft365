import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./theme-remove');

describe(commands.THEME_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    promptOptions = undefined;
    sinon.stub(Cli, 'prompt').callsFake(async (options) => {
      promptOptions = options;
      return { continue: false };
    });
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
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.THEME_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should prompt before removing theme when confirmation argument not passed', async () => {
    await command.action(logger, { options: { name: 'Contoso' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('removes theme successfully without prompting with confirmation argument', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
        return 'Correct Url';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: 'Contoso',
        confirm: true
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_api/thememanager/DeleteTenantTheme');
    assert.strictEqual(postStub.lastCall.args[0].headers['accept'], 'application/json;odata=nometadata');
    assert.strictEqual(postStub.lastCall.args[0].data.name, 'Contoso');
    assert.strictEqual(loggerLogSpy.notCalled, true);
  });

  it('removes theme successfully without prompting with confirmation argument (debug)', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
        return 'Correct Url';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso',
        confirm: true
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_api/thememanager/DeleteTenantTheme');
    assert.strictEqual(postStub.lastCall.args[0].headers['accept'], 'application/json;odata=nometadata');
    assert.strictEqual(postStub.lastCall.args[0].data.name, 'Contoso');
  });

  it('removes theme successfully when prompt confirmed', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
        return 'Correct Url';
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_api/thememanager/DeleteTenantTheme');
    assert.strictEqual(postStub.lastCall.args[0].headers['accept'], 'application/json;odata=nometadata');
    assert.strictEqual(postStub.lastCall.args[0].data.name, 'Contoso');
  });

  it('handles error when removing theme', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
        throw 'An error has occurred';
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso',
        confirm: true
      }
    } as any), new CommandError('An error has occurred'));
  });
});
