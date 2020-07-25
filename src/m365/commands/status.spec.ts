import commands from './commands';
import Command, { CommandError } from '../../Command';
import * as sinon from 'sinon';
import appInsights from '../../appInsights';
import auth from '../../Auth';
const command: Command = require('./status');
import * as assert from 'assert';
import Utils from '../../Utils';

describe(commands.STATUS, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: any) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.STATUS), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('shows logged out status when not logged in', (done) => {
    auth.service.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('Logged out'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows logged out status when not logged in (verbose)', (done) => {
    auth.service.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('Logged out from Microsoft 365'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows logged in status when logged in', (done) => {
    auth.service.accessTokens['https://graph.microsoft.com'] = {
      expiresOn: 'abc',
      value: 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ing0NTh4eU9wbHNNMkg3TlhrMlN4MTd4MXVwYyIsImtpZCI6Ing0NTh4eU9wbHNNMkg3TlhrN1N4MTd4MXVwYyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLndpbmRvd3MubmV0IiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvY2FlZTMyZTYtNDA1ZC00MjRhLTljZjEtMjA3MWQwNDdmMjk4LyIsImlhdCI6MTUxNTAwNDc4NCwibmJmIjoxNTE1MDA0Nzg0LCJleHAiOjE1MTUwMDg2ODQsImFjciI6IjEiLCJhaW8iOiJBQVdIMi84R0FBQUFPN3c0TDBXaHZLZ1kvTXAxTGJMWFdhd2NpOEpXUUpITmpKUGNiT2RBM1BvPSIsImFtciI6WyJwd2QiXSwiYXBwaWQiOiIwNGIwNzc5NS04ZGRiLTQ2MWEtYmJlZS0wMmY5ZTFiZjdiNDYiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6IkRvZSIsImdpdmVuX25hbWUiOiJKb2huIiwiaXBhZGRyIjoiOC44LjguOCIsIm5hbWUiOiJKb2huIERvZSIsIm9pZCI6ImYzZTU5NDkxLWZjMWEtNDdjYy1hMWYwLTk1ZWQ0NTk4MzcxNyIsInB1aWQiOiIxMDk0N0ZGRUE2OEJDQ0NFIiwic2NwIjoiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIiwic3ViIjoiemZicmtUV1VQdEdWUUg1aGZRckpvVGp3TTBrUDRsY3NnLTJqeUFJb0JuOCIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJOQSIsInRpZCI6ImNhZWUzM2U2LTQwNWQtNDU0YS05Y2YxLTMwNzFkMjQxYTI5OCIsInVuaXF1ZV9uYW1lIjoiYWRtaW5AY29udG9zby5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJhZG1pbkBjb250b3NvLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6ImFUZVdpelVmUTBheFBLMVRUVXhsQUEiLCJ2ZXIiOiIxLjAifQ==.abc'
    };
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          connectedAs: 'admin@contoso.onmicrosoft.com'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly reports access token', (done) => {
    auth.service.connected = true;
    auth.service.accessTokens = {
      'https://graph.microsoft.com': {
        expiresOn: '123',
        value: 'abc'
      }
    };
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (JSON.stringify(l) === JSON.stringify({
          connectedAs: '',
          authType: 'DeviceCode',
          accessTokens: '{\n  "https://graph.microsoft.com": {\n    "expiresOn": "123",\n    "value": "abc"\n  }\n}'
        })) {
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

  it('correctly handles error when restoring auth', (done) => {
    Utils.restore(auth.restoreAuth);
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});