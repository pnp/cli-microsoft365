import commands from '../commands';
import Command, { CommandError } from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth, { Site } from '../SpoAuth';
const command: Command = require('./status');
import * as assert from 'assert';
import Utils from '../../../Utils';

describe(commands.STATUS, () => {
  let vorpal: Vorpal;
  let log: any[];
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
      log: (msg: any) => {
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
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
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
    cmdInstance.action = command.action();
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
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l && l.connectedTo === 'https://contoso.sharepoint.com') {
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

  it('correctly reports current user', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    auth.service.accessToken = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ing0NTh4eU9wbHNNMkg3TlhrMlN4MTd4MXVwYyIsImtpZCI6Ing0NTh4eU9wbHNNMkg3TlhrN1N4MTd4MXVwYyJ9.eyJhdWQiOiJodHRwczovL2NvbnRvc28uc2hhcmVwb2ludC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC82Nzc1M2Y2My1iYzE0LTQwMTItODY5ZS1mYTAxYTMzZmUwMjMvIiwiaWF0IjoxNTE1MzMwNDU2LCJuYmYiOjE1MTUzMzA0NTYsImV4cCI6MTUxNTMzNDM1NiwiYWNyIjoiMSIsImFpbyI6IlkyTmdZR1pUU1VUZWVYQURpdnhzYmEzYis1Vmw0dC8zM1hpUjgxdERXNlRRcm9jM3V4VUEiLCJhbXIiOlsicHdkIl0sImFwcF9kaXNwbGF5bmFtZSI6Ik1pY3Jvc29mdCBTaGFyZVBvaW50IE9ubGluZSBNYW5hZ2VtZW50IFNoZWxsIiwiYXBwaWQiOiI5YmMzYWI0OS1iNjVkLTQxMGEtODVhZC1kZTgxOWZlYmZkZGMiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6IkRvZSIsImdpdmVuX25hbWUiOiJKb2UiLCJpcGFkZHIiOiI4LjguOC44IiwibmFtZSI6IkpvZSBEb2UiLCJvaWQiOiI5NDliMTZjMS1hMDMyLTQ1M2UtYThhZS04OWE1MmJmYzFkOGEiLCJwdWlkIjoiMTAwMzdGRkVBNjdBQkNDRSIsInNjcCI6IkFsbFByb2ZpbGVzLk1hbmFnZSBTaXRlcy5GdWxsQ29udHJvbC5BbGwgVXNlci5SZWFkLkFsbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6InVpbnFkZXdYRXBHdXZMRUFQeDVTQVNhZlVfeXhTNFhseDF5Z3FQMEFvOTgiLCJ0aWQiOiI2Nzc1M2Y2My1iYzE0LTQwMTItODY5ZS1mODA4YTQzZmUwMjMiLCJ1bmlxdWVfbmFtZSI6ImFkbWluQGNvbnRvc28ub25taWNyb3NvZnQuY29tIiwidXBuIjoiYWRtaW5AY29udG9zby5vbm1pY3Jvc29mdC5jb20iLCJ1dGkiOiJlOFRRcjJKQXlVTFA3dkVoS2JNeUFBIiwidmVyIjoiMS4wIn0=.abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l && l.connectedAs === 'admin@contoso.onmicrosoft.com') {
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
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l && l.isTenantAdmin === true) {
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
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l && l.isTenantAdmin === false) {
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
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      let reportsCorrectValue: boolean = false;
      log.forEach(l => {
        if (l && l.aadResource === 'https://contoso.sharepoint.com') {
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
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com/sites/team';
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com/sites/team';
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
    const cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
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
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.STATUS));
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
});