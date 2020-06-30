import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
const command: Command = require('./accesstoken-get');
import * as assert from 'assert';
import Utils from '../../../../Utils';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';

describe(commands.UTIL_ACCESSTOKEN_GET, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let cmdInstance: any;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
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
      vorpal.find,
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
    assert.equal(command.name.startsWith(commands.UTIL_ACCESSTOKEN_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
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
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if resource is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if resource is undefined', () => {
    const actual = (command.validate() as CommandValidate)({ options: { resource: undefined } });
    assert.notEqual(actual, true);
  });

  it('fails validation if resource is blank', () => {
    const actual = (command.validate() as CommandValidate)({ options: { resource: '' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when resource is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { resource: 'https://graph.microsoft.com' } });
    assert.equal(actual, true);
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.UTIL_ACCESSTOKEN_GET));
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