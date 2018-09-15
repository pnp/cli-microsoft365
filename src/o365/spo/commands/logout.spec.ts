import commands from '../commands';
import Command from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth, { Site } from '../SpoAuth';
const command: Command = require('./logout');
import * as assert from 'assert';
import Utils from '../../../Utils';

describe(commands.LOGOUT, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;
  let authClearSiteConnectionInfoStub: sinon.SinonStub;

  before(() => {
    authClearSiteConnectionInfoStub = sinon.stub(auth, 'clearSiteConnectionInfo').callsFake(() => Promise.resolve());
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../vorpal-init');
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: 'spo logout'
      },
      log: (msg: string) => {
        log.push(msg);
      }
    };
    auth.site = new Site();
    sinon.stub(auth.site, 'logout').callsFake(() => { });
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore(vorpal.find);
  });

  after(() => {
    Utils.restore(appInsights.trackEvent);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.LOGOUT), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.LOGOUT);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs out from SharePoint when logged in', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(!auth.site.connected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs out from SharePoint when not logged in', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(!auth.site.connected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('clears persisted connection info when logging out', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(authClearSiteConnectionInfoStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notEqual(typeof alias, 'undefined');
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => {},
      prompt: () => {},
      helpInformation: () => {}
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => {});
    assert(find.calledWith(commands.LOGOUT));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => {},
      helpInformation: () => {}
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => {});
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
    Utils.restore(auth.clearSiteConnectionInfo);
    sinon.stub(auth, 'clearSiteConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    auth.site = new Site();
    const logoutSpy = sinon.spy(auth.site, 'logout');
    auth.site.connected = true;
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
        Utils.restore(auth.clearSiteConnectionInfo);
      }
    });
  });

  it('correctly handles error while clearing persisted connection info (debug)', (done) => {
    sinon.stub(auth, 'clearSiteConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    auth.site = new Site();
    const logoutSpy = sinon.spy(auth.site, 'logout');
    auth.site.connected = true;
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
        Utils.restore(auth.clearSiteConnectionInfo);
      }
    });
  });
});