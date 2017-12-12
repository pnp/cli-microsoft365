import commands from '../commands';
import Command, { CommandHelp, CommandError } from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth, { Site } from '../SpoAuth';
const statusCommand: Command = require('./status');
import * as assert from 'assert';
import Utils from '../../../Utils';

describe(commands.STATUS, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
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
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore(vorpal.find);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(statusCommand.name.startsWith(commands.STATUS), true);
  });

  it('has a description', () => {
    assert.notEqual(statusCommand.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = statusCommand.action();
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
    cmdInstance.action = statusCommand.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.STATUS);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows disconnected status when not connected', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = statusCommand.action();
    cmdInstance.action({ options: {} }, () => {
      let reportsDisconnected: boolean = false;
      log.forEach(l => {
        if (l && l.indexOf('Not connected') === 0) {
          reportsDisconnected = true;
        }
      });
      try {
        assert(reportsDisconnected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows disconnected status when not connected (verbose)', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = statusCommand.action();
    cmdInstance.action({ options: { verbose: true } }, () => {
      let reportsDisconnected: boolean = false;
      log.forEach(l => {
        if (l && l.indexOf('Not connected to SharePoint Online') === 0) {
          reportsDisconnected = true;
        }
      });
      try {
        assert(reportsDisconnected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows correct site URL when connected', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = statusCommand.action();
    cmdInstance.action({ options: {} }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l && l.indexOf(auth.site.url) === 0) {
          reportsCorrectValue = true;
        }
      });
      try {
        assert(reportsCorrectValue);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly reports tenant admin site', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = statusCommand.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l && l.indexOf('Is tenant admin:') > -1 && l.indexOf('true') > -1) {
          reportsCorrectValue = true;
        }
      });
      try {
        assert(reportsCorrectValue);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly reports regular site', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = statusCommand.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l && l.indexOf('Is tenant admin:') > -1 && l.indexOf('false') > -1) {
          reportsCorrectValue = true;
        }
      });
      try {
        assert(reportsCorrectValue);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly reports AAD resource', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com/sites/team';
    auth.service.resource = 'https://contoso.sharepoint.com';
    cmdInstance.action = statusCommand.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l && l.indexOf('AAD resource:') > -1 && l.indexOf(auth.service.resource) > -1) {
          reportsCorrectValue = true;
        }
      });
      try {
        assert(reportsCorrectValue);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly reports access token', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com/sites/team';
    auth.service.accessToken = 'abc';
    cmdInstance.action = statusCommand.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l && l.indexOf('Access token:') > -1 && l.indexOf('abc') > -1) {
          reportsCorrectValue = true;
        }
      });
      try {
        assert(reportsCorrectValue);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly reports refresh token', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com/sites/team';
    auth.service.refreshToken = 'abc';
    cmdInstance.action = statusCommand.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l && l.indexOf('Refresh token:') > -1 && l.indexOf('abc') > -1) {
          reportsCorrectValue = true;
        }
      });
      try {
        assert(reportsCorrectValue);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly reports expiration time', (done) => {
    const date: Date = new Date();
    const expiresAtDate: Date = new Date(0);
    expiresAtDate.setUTCSeconds(date.getUTCSeconds());

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com/sites/team';
    auth.service.expiresAt = date.getUTCSeconds();
    cmdInstance.action = statusCommand.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l && l.indexOf('Expires at:') > -1 && l.indexOf(expiresAtDate.toString()) > -1) {
          reportsCorrectValue = true;
        }
      });
      try {
        assert(reportsCorrectValue);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when restoring auth', (done) => {
    Utils.restore(auth.restoreAuth);
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    const cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    cmdInstance.action = statusCommand.action();
    cmdInstance.action({options:{}}, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('An error has occurred')));
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
    (statusCommand.help() as CommandHelp)({}, helpLog);
    assert(find.calledWith(commands.STATUS));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const log = (msg: string) => { _log.push(msg); }
    const cmd: any = {
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    (statusCommand.help() as CommandHelp)({}, log);
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});