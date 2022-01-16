import * as assert from 'assert';
import * as sinon from 'sinon';
import { Logger } from '../../../../cli';
import { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command = require('./user-hibp');

describe(commands.USER_HIBP, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
      request.get
    ]);
  });


  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.USER_HIBP), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if userName and apiKey is not specified', () => {
    const actual = command.validate({ options: {} });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the userName is not a valid UPN', () => {
    const actual = command.validate({ options: { userName: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if userName and apiKey is specified', () => {
    const actual = command.validate({ options: { userName: "account-exists@hibp-integration-tests.com", apiKey: "2975xc539c304xf797f665x43f8x557x" } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if domain is specified', () => {
    const actual = command.validate({ options: { userName: "account-exists@hibp-integration-tests.com", apiKey: "2975xc539c304xf797f665x43f8x557x", domain: "domain.com" } });
    assert.strictEqual(actual, true);
  });

  it('checks user is pwned using userName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://haveibeenpwned.com/api/v3/breachedaccount/${encodeURIComponent('account-exists@hibp-integration-tests.com')}`) {
        // this is the actual truncated response as the API would return
        return Promise.resolve([{ "Name": "Adobe" }]);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { userName: 'account-exists@hibp-integration-tests.com', apiKey: '2975xc539c304xf797f665x43f8x557x' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([{ "Name": "Adobe" }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('checks user is pwned using userName (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://haveibeenpwned.com/api/v3/breachedaccount/${encodeURIComponent('account-exists@hibp-integration-tests.com')}`) {
        // this is the actual truncated response as the API would return
        return Promise.resolve([{ "Name": "Adobe" }]);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, userName: 'account-exists@hibp-integration-tests.com', apiKey: '2975xc539c304xf797f665x43f8x557x' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([{ "Name": "Adobe" }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('checks user is pwned using userName and multiple breaches', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://haveibeenpwned.com/api/v3/breachedaccount/${encodeURIComponent('account-exists@hibp-integration-tests.com')}`) {
        // this is the actual truncated response as the API would return
        return Promise.resolve([{ "Name": "Adobe" }, { "Name": "Gawker" }, { "Name": "Stratfor" }]);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, userName: 'account-exists@hibp-integration-tests.com', apiKey: "2975xc539c304xf797f665x43f8x557x" } }, () => {
      try {
        assert(loggerLogSpy.calledWith([{ "Name": "Adobe" }, { "Name": "Gawker" }, { "Name": "Stratfor" }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('checks user is pwned using userName and domain', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://haveibeenpwned.com/api/v3/breachedaccount/${encodeURIComponent('account-exists@hibp-integration-tests.com')}?domain=adobe.com`) {
        // this is the actual truncated response as the API would return
        return Promise.resolve([{ "Name": "Adobe" }]);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, userName: 'account-exists@hibp-integration-tests.com', domain: "adobe.com", apiKey: "2975xc539c304xf797f665x43f8x557x" } }, () => {
      try {
        assert(loggerLogSpy.calledWith([{ "Name": "Adobe" }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('checks user is pwned using userName and domain with a domain that does not exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://haveibeenpwned.com/api/v3/breachedaccount/${encodeURIComponent('account-exists@hibp-integration-tests.com')}?domain=adobe.xxx`) {
        // this is the actual truncated response as the API would return
        return Promise.reject({
          "response": {
            "status": 404
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, userName: 'account-exists@hibp-integration-tests.com', domain: "adobe.xxx", apiKey: "2975xc539c304xf797f665x43f8x557x" } }, () => {
      try {
        assert(loggerLogSpy.calledWith("No pwnage found"));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('checks user is pwned using userName (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://haveibeenpwned.com/api/v3/breachedaccount/${encodeURIComponent('account-exists@hibp-integration-tests.com')}`) {
        return Promise.resolve([{ "Name": "Adobe" }]);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, userName: 'account-exists@hibp-integration-tests.com', apiKey: "2975xc539c304xf797f665x43f8x557x" } }, () => {
      try {
        assert(loggerLogSpy.calledWith([{ "Name": "Adobe" }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no pwnage found (debug)', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        "response": {
          "status": 404
        }
      });
    });

    command.action(logger, { options: { debug: true, userName: 'account-notexists@hibp-integration-tests.com', apiKey: "2975xc539c304xf797f665x43f8x557x" } }, () => {
      try {
        assert(loggerLogSpy.calledWith("No pwnage found"));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no pwnage found (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        "response": {
          "status": 404
        }
      });
    });

    command.action(logger, { options: { verbose: true, userName: 'account-notexists@hibp-integration-tests.com', apiKey: "2975xc539c304xf797f665x43f8x557x" } }, () => {
      try {
        assert(loggerLogSpy.calledWith("No pwnage found"));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles unauthorized request', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        "statusCode": 401,
        "message": "Access denied due to improperly formed hibp-api-key."
      });
    });

    command.action(logger, { options: { debug: false, userName: 'account-notexists@hibp-integration-tests.com' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(err.message)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the userName is not a valid UPN.', () => {
    const actual = command.validate({
      options: {
        userName: "no-an-email"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach((o: { option: string; }) => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
