import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./user-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.YAMMER_USER_GET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.YAMMER_USER_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls user by e-mail', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/by_email.json?email=pl%40nubo.eu') {
        return Promise.resolve(
          [{ "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" }]
        );
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { email: "pl@nubo.eu" } }, (err?: any) => {
      try {
        assert.equal(cmdInstanceLogSpy.lastCall.args[0][0].id, 1496550646)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls user by userId', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/1496550646.json') {
        return Promise.resolve(
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" }
        );
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { userId: 1496550646 } }, (err?: any) => {
      try {
        assert.equal(cmdInstanceLogSpy.lastCall.args[0].id, 1496550646)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the current user and json', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/current.json') {
        return Promise.resolve(
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" }
        );
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { output: 'json' } }, (err?: any) => {
      try {
        assert.equal(cmdInstanceLogSpy.lastCall.args[0].id, 1496550646)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles 404 error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "statusCode": 404
      });
    });

    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("Not found (404)")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes validation without parameters', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.equal(actual, true);
  });

  it('passes validation with parameters', () => {
    const actual = (command.validate() as CommandValidate)({ options: { userId: 1496550646 } });
    assert.equal(actual, true);
  });

  it('passes validation with parameters', () => {
    const actual = (command.validate() as CommandValidate)({ options: { email: "pl@nubo.eu" } });
    assert.equal(actual, true);
  });

  it('does not pass with userId and e-mail', () => {
    const actual = (command.validate() as CommandValidate)({ options: { userId: 1496550646, email: "pl@nubo.eu" } });
    assert.equal(actual, "You are only allowed to search by ID or e-mail but not both");
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
    assert(find.calledWith(commands.YAMMER_USER_GET));
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