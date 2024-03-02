import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
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
  let promptIssued: boolean = false;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
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
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.THEME_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should prompt before removing theme when confirmation argument not passed', async () => {
    await command.action(logger, { options: { name: 'Contoso' } });

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
        force: true
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
        force: true
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

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

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

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso',
        force: true
      }
    } as any), new CommandError('An error has occurred'));
  });
});
