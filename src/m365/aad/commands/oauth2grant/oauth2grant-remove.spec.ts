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
const command: Command = require('./oauth2grant-remove');

describe(commands.OAUTH2GRANT_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.OAUTH2GRANT_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes OAuth2 permission grant when prompt confirmed (debug)', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/oauth2PermissionGrants/YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, { options: { debug: true, grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek' } });
    assert(loggerLogToStderrSpy.called);
    assert(deleteRequestStub.called);
  });

  it('removes OAuth2 permission grant when prompt confirmed', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/oauth2PermissionGrants/YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, { options: { grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek' } });
    assert(loggerLogSpy.notCalled);
    assert(deleteRequestStub.called);
  });

  it('removes OAuth2 permission grant when confirm specified', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/oauth2PermissionGrants/YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek', confirm: true } });
    assert(loggerLogSpy.notCalled);
    assert(deleteRequestStub.called);
  });

  it('prompts before removing OAuth2 permission grant when confirm option not passed', async () => {
    await command.action(logger, { options: { grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('prompts before removing OAuth2 permission grant when confirm option not passed (debug)', async () => {
    await command.action(logger, { options: { debug: true, grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing OAuth2 permission grant when prompt not confirmed', async () => {
    const deleteSpy = sinon.spy(request, 'delete');

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));

    await command.action(logger, { options: { grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek' } });
    assert(deleteSpy.notCalled);
  });

  it('aborts removing OAuth2 permission grant when prompt not confirmed (debug)', async () => {
    const deleteSpy = sinon.spy(request, 'delete');

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));

    await command.action(logger, { options: { debug: true, grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek' } });
    assert(deleteSpy.notCalled);
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: 'An error has occurred'
            }
          }
        }
      });
    });

    await assert.rejects(command.action(logger, { options: { confirm: true, grantId: 'YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek' } } as any),
      new CommandError('An error has occurred'));
  });

  it('supports specifying grantId', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--grantId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying confirmation flag', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--confirm') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
