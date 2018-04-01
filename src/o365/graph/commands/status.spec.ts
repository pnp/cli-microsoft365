import commands from '../commands';
import Command, { CommandError } from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth from '../GraphAuth';
const command: Command = require('./status');
import * as assert from 'assert';
import Utils from '../../../Utils';
import { Service } from '../../../Auth';

describe(commands.STATUS, () => {
  let vorpal: Vorpal;
  let log: any[];
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
      log: (msg: any) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.service = new Service('https://graph.microsoft.com');
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
    auth.service = new Service('https://graph.microsoft.com');
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
    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('Not connected to Microsoft Graph'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows connected status when connected', (done) => {
    auth.service = new Service('https://graph.microsoft.com');
    auth.service.accessToken = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ing0NTh4eU9wbHNNMkg3TlhrMlN4MTd4MXVwYyIsImtpZCI6Ing0NTh4eU9wbHNNMkg3TlhrN1N4MTd4MXVwYyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLndpbmRvd3MubmV0IiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvY2FlZTMyZTYtNDA1ZC00MjRhLTljZjEtMjA3MWQwNDdmMjk4LyIsImlhdCI6MTUxNTAwNDc4NCwibmJmIjoxNTE1MDA0Nzg0LCJleHAiOjE1MTUwMDg2ODQsImFjciI6IjEiLCJhaW8iOiJBQVdIMi84R0FBQUFPN3c0TDBXaHZLZ1kvTXAxTGJMWFdhd2NpOEpXUUpITmpKUGNiT2RBM1BvPSIsImFtciI6WyJwd2QiXSwiYXBwaWQiOiIwNGIwNzc5NS04ZGRiLTQ2MWEtYmJlZS0wMmY5ZTFiZjdiNDYiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6IkRvZSIsImdpdmVuX25hbWUiOiJKb2huIiwiaXBhZGRyIjoiOC44LjguOCIsIm5hbWUiOiJKb2huIERvZSIsIm9pZCI6ImYzZTU5NDkxLWZjMWEtNDdjYy1hMWYwLTk1ZWQ0NTk4MzcxNyIsInB1aWQiOiIxMDk0N0ZGRUE2OEJDQ0NFIiwic2NwIjoiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIiwic3ViIjoiemZicmtUV1VQdEdWUUg1aGZRckpvVGp3TTBrUDRsY3NnLTJqeUFJb0JuOCIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJOQSIsInRpZCI6ImNhZWUzM2U2LTQwNWQtNDU0YS05Y2YxLTMwNzFkMjQxYTI5OCIsInVuaXF1ZV9uYW1lIjoiYWRtaW5AY29udG9zby5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJhZG1pbkBjb250b3NvLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6ImFUZVdpelVmUTBheFBLMVRUVXhsQUEiLCJ2ZXIiOiIxLjAifQ==.abc';
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ 
          connectedTo: 'https://graph.microsoft.com',
          connectedAs: 'admin@contoso.onmicrosoft.com'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly reports MS Graph resource', (done) => {
    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l && l.aadResource === 'https://graph.microsoft.com') {
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
    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    auth.service.accessToken = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l && l.accessToken === 'abc') {
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
    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    auth.service.refreshToken = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l && l.refreshToken === 'abc') {
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

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    auth.service.expiresOn = date.toISOString();
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l &&
          l.expiresAt &&
          l.expiresAt === date.toISOString()) {
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