import commands from '../commands';
import Command from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth, { Site } from '../SpoAuth';
const command: Command = require('./disconnect');
import * as assert from 'assert';
import Utils from '../../../Utils';

describe(commands.DISCONNECT, () => {
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
      log: (msg: string) => {
        log.push(msg);
      }
    };
    auth.site = new Site();
    sinon.stub(auth.site, 'disconnect').callsFake(() => { });
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore(vorpal.find);
  });

  after(() => {
    Utils.restore(appInsights.trackEvent);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.DISCONNECT), true);
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
        assert.equal(telemetry.name, commands.DISCONNECT);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('disconnects from SharePoint when connected', (done) => {
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

  it('disconnects from SharePoint when not connected', (done) => {
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

  it('clears persisted connection info when disconnecting', (done) => {
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => {},
      prompt: () => {},
      helpInformation: () => {}
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => {});
    assert(find.calledWith(commands.DISCONNECT));
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
    const disconnectSpy = sinon.spy(auth.site, 'disconnect');
    auth.site.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(disconnectSpy.called);
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
    const disconnectSpy = sinon.spy(auth.site, 'disconnect');
    auth.site.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(disconnectSpy.called);
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