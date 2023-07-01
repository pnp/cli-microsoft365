import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './theme-remove.js';

describe(commands.THEME_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    assert.strictEqual(command.name.startsWith(commands.THEME_REMOVE), true);
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
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        name: 'Contoso',
        force: true
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_api/thememanager/DeleteTenantTheme');
    assert.strictEqual(postStub.lastCall.args[0].headers['accept'], 'application/json;odata=nometadata');
    assert.strictEqual(postStub.lastCall.args[0].data.name, 'Contoso');
    assert.strictEqual(loggerLogSpy.notCalled, true);
  });

  it('removes theme successfully without prompting with confirmation argument (debug)', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso',
        force: true
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_api/thememanager/DeleteTenantTheme');
    assert.strictEqual(postStub.lastCall.args[0].headers['accept'], 'application/json;odata=nometadata');
    assert.strictEqual(postStub.lastCall.args[0].data.name, 'Contoso');
  });

  it('removes theme successfully when prompt confirmed', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
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
    sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso',
        force: true
      }
    } as any), new CommandError('An error has occurred'));
  });
});
