import commands from './commands';
import Command, { CommandError } from '../../Command';
import * as sinon from 'sinon';
import appInsights from '../../appInsights';
import auth from '../../Auth';
const command: Command = require('./logout');
import * as assert from 'assert';
import Utils from '../../Utils';

describe(commands.LOGOUT, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let authClearConnectionInfoStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    authClearConnectionInfoStub = sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
  });

  beforeEach(() => {
    vorpal = require('../../vorpal-init');
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: 'logout'
      },
      log: (msg: string) => {
        log.push(msg);
      }
    };
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.LOGOUT), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('logs out from Microsoft 365 when logged in', (done) => {
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(!auth.service.connected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs out from Microsoft 365 when not logged in', (done) => {
    auth.service.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(!auth.service.connected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('clears persisted connection info when logging out', (done) => {
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(authClearConnectionInfoStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.LOGOUT));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles error while clearing persisted connection info', (done) => {
    Utils.restore(auth.clearConnectionInfo);
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    const logoutSpy = sinon.spy(auth.service, 'logout');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(logoutSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          auth.clearConnectionInfo,
          auth.service.logout
        ]);
      }
    });
  });

  it('correctly handles error while clearing persisted connection info (debug)', (done) => {
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    const logoutSpy = sinon.spy(auth.service, 'logout');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(logoutSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          auth.clearConnectionInfo,
          auth.service.logout
        ]);
      }
    });
  });

  it('correctly handles error when restoring auth information', (done) => {
    Utils.restore(auth.restoreAuth);
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          auth.clearConnectionInfo
        ]);
      }
    });
  });
});