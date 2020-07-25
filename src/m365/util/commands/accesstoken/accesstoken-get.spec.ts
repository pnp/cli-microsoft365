import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
const command: Command = require('./accesstoken-get');
import * as assert from 'assert';
import Utils from '../../../../Utils';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';

describe(commands.UTIL_ACCESSTOKEN_GET, () => {
  let log: any[];
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let cmdInstance: any;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    auth.service.connected = true;
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
    Utils.restore([
      auth.ensureAccessToken
    ]);
    auth.service.accessTokens = {};
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.UTIL_ACCESSTOKEN_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves access token for the specified resource', (done) => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.service.accessTokens['https://graph.microsoft.com'] = {
      expiresOn: d.toString(),
      value: 'ABC'
    };

    cmdInstance.action({ options: { debug: false, resource: 'https://graph.microsoft.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('ABC'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving access token', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.reject('An error has occurred'));

    cmdInstance.action({ options: { debug: false, resource: 'https://graph.microsoft.com' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});