import commands from '../commands';
import Command, { CommandError } from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth from '../AadAuth';
const command: Command = require('./status');
import * as assert from 'assert';
import Utils from '../../../Utils';
import { Service } from '../../../Auth';

describe(commands.STATUS, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;
  let cmdInstanceLogSpy: any;

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
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.service = new Service('https://graph.windows.net');
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
    assert.equal(command.name.startsWith(commands.STATUS), true);
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
        assert.equal(telemetry.name, commands.STATUS);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows disconnected status when not connected', (done) => {
    auth.service = new Service('https://graph.windows.net');
    auth.service.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('Not connected'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows disconnected status when not connected (verbose)', (done) => {
    auth.service = new Service('https://graph.windows.net');
    auth.service.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('Not connected to AAD Graph'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows connected status when connected', (done) => {
    auth.service = new Service('https://graph.windows.net');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('https://graph.windows.net'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows connected status when connected (verbose)', (done) => {
    auth.service = new Service('https://graph.windows.net');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('Connected to https://graph.windows.net'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly reports AAD resource', (done) => {
    auth.service = new Service('https://graph.windows.net');
    auth.service.connected = true;
    auth.service.resource = 'https://graph.windows.net';
    cmdInstance.action = command.action();
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
    auth.service = new Service('https://graph.windows.net');
    auth.service.connected = true;
    auth.service.accessToken = 'abc';
    cmdInstance.action = command.action();
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
    auth.service = new Service('https://graph.windows.net');
    auth.service.connected = true;
    auth.service.refreshToken = 'abc';
    cmdInstance.action = command.action();
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

    auth.service = new Service('https://graph.windows.net');
    auth.service.connected = true;
    auth.service.expiresAt = date.getUTCSeconds();
    cmdInstance.action = command.action();
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
    cmdInstance.action = command.action();
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
    const cmd: any = {
      log: (msg: string) => {},
      prompt: () => {},
      helpInformation: () => {}
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => {});
    assert(find.calledWith(commands.STATUS));
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
});