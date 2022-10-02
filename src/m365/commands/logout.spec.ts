import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../appInsights';
import auth from '../../Auth';
import { Logger } from '../../cli/Logger';
import Command, { CommandError } from '../../Command';
import { sinonUtil } from '../../utils/sinonUtil';
import commands from './commands';
const command: Command = require('./logout');

describe(commands.LOGOUT, () => {
  let log: string[];
  let logger: Logger;
  let authClearConnectionInfoStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    authClearConnectionInfoStub = sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
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
      await command.action(logger, { options: { debug: false } });
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