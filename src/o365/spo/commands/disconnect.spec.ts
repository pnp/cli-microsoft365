import commands from '../commands';
import Command, { CommandHelp } from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth, { Site } from '../SpoAuth';
const disconnectCommand: Command = require('./disconnect');
import * as assert from 'assert';
import Utils from '../../../Utils';

describe(commands.DISCONNECT, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
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
    assert.equal(disconnectCommand.name.startsWith(commands.DISCONNECT), true);
  });

  it('has a description', () => {
    assert.notEqual(disconnectCommand.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = disconnectCommand.action();
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
    cmdInstance.action = disconnectCommand.action();
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
    cmdInstance.action = disconnectCommand.action();
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
    cmdInstance.action = disconnectCommand.action();
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

  it('has help referring to the right command', () => {
    const _helpLog: string[] = [];
    const helpLog = (msg: string) => { _helpLog.push(msg); }
    const cmd: any = {
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    (disconnectCommand.help() as CommandHelp)({}, helpLog);
    assert(find.calledWith(commands.DISCONNECT));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const log = (msg: string) => { _log.push(msg); }
    const cmd: any = {
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    (disconnectCommand.help() as CommandHelp)({}, log);
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    assert(containsExamples);
    Utils.restore(vorpal.find);
  });
});