import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./accesstoken-get');

describe(commands.ACCESSTOKEN_GET, () => {
  let log: any[];
  let loggerLogSpy: sinon.SinonSpy;
  let logger: Logger;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    auth.service.connected = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    Utils.restore([
      auth.ensureAccessToken
    ]);
    auth.service.accessTokens = {};
    auth.service.spoUrl = undefined;
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ACCESSTOKEN_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves access token for the specified resource', (done) => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.service.accessTokens['https://graph.microsoft.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    command.action(logger, { options: { debug: false, resource: 'https://graph.microsoft.com' } }, () => {
      try {
        assert(loggerLogSpy.calledWith('ABC'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves access token for SharePoint when sharepoint specified as the resource and SPO URL previously retrieved', (done) => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    auth.service.accessTokens['https://contoso.sharepoint.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    command.action(logger, { options: { debug: false, resource: 'sharepoint' } }, () => {
      try {
        assert(loggerLogSpy.calledWith('ABC'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving access token', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.reject('An error has occurred'));

    command.action(logger, { options: { debug: false, resource: 'https://graph.microsoft.com' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when sharepoint specified as resource and SPO URL not available', (done) => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.service.accessTokens['https://contoso.sharepoint.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    command.action(logger, { options: { debug: false, resource: 'sharepoint' } }, (err?: any) => {
      try {
        assert.notStrictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});