import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./user-get');

describe(commands.YAMMER_USER_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
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
    assert.strictEqual(command.name.startsWith(commands.YAMMER_USER_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'full_name', 'email', 'job_title', 'state', 'url']);
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
    command.action(logger, { options: { email: "pl@nubo.eu" } } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
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
    command.action(logger, { options: { userId: 1496550646 } } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 1496550646);
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
    command.action(logger, { options: { output: 'json' } } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 1496550646);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles 404 error', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        "statusCode": 404
      });
    });

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Not found (404)")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes validation without parameters', () => {
    const actual = command.validate({ options: {} });
    assert.strictEqual(actual, true);
  });

  it('passes validation with parameters', () => {
    const actual = command.validate({ options: { userId: 1496550646 } });
    assert.strictEqual(actual, true);
  });

  it('passes validation with parameters', () => {
    const actual = command.validate({ options: { email: "pl@nubo.eu" } });
    assert.strictEqual(actual, true);
  });

  it('does not pass with userId and e-mail', () => {
    const actual = command.validate({ options: { userId: 1496550646, email: "pl@nubo.eu" } });
    assert.strictEqual(actual, "You are only allowed to search by ID or e-mail but not both");
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