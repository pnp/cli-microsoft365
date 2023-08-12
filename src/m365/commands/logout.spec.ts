import assert from 'assert';
import sinon from 'sinon';
import auth from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import { CommandError } from '../../Command.js';
import { telemetry } from '../../telemetry.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { sinonUtil } from '../../utils/sinonUtil.js';
import commands from './commands.js';
import command from './logout.js';

describe(commands.LOGOUT, () => {
  let log: string[];
  let logger: Logger;
  let authClearConnectionInfoStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    authClearConnectionInfoStub = sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LOGOUT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('logs out from Microsoft 365 when logged in', async () => {
    auth.service.connected = true;
    await command.action(logger, { options: { debug: true } });
    assert(!auth.service.connected);
  });

  it('logs out from Microsoft 365 when not logged in', async () => {
    auth.service.connected = false;
    await command.action(logger, { options: { debug: true } });
    assert(!auth.service.connected);
  });

  it('clears persisted connection info when logging out', async () => {
    auth.service.connected = true;
    await command.action(logger, { options: { debug: true } });
    assert(authClearConnectionInfoStub.called);
  });

  it('correctly handles error while clearing persisted connection info', async () => {
    sinonUtil.restore(auth.clearConnectionInfo);
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    const logoutSpy = sinon.spy(auth.service, 'logout');
    auth.service.connected = true;

    try {
      await command.action(logger, { options: {} });
      assert(logoutSpy.called);
    }
    finally {
      sinonUtil.restore([
        auth.clearConnectionInfo,
        auth.service.logout
      ]);
    }
  });

  it('correctly handles error while clearing persisted connection info (debug)', async () => {
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    const logoutSpy = sinon.spy(auth.service, 'logout');
    auth.service.connected = true;

    try {
      await command.action(logger, { options: { debug: true } });
      assert(logoutSpy.called);
    }
    finally {
      sinonUtil.restore([
        auth.clearConnectionInfo,
        auth.service.logout
      ]);
    }
  });

  it('correctly handles error when restoring auth information', async () => {
    sinonUtil.restore(auth.restoreAuth);
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));

    try {
      await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
    }
    finally {
      sinonUtil.restore([
        auth.clearConnectionInfo
      ]);
    }
  });
});
