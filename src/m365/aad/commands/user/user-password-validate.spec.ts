import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./user-password-validate');

describe(commands.USER_PASSWORD_VALIDATE, () => {
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
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.USER_PASSWORD_VALIDATE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('password is too short', (done) => {
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/beta/users/validatePassword' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "password": "cli365"
        })) {
        return Promise.resolve({
          "isValid": false,
          "validationResults": [
            {
              "ruleName": "password_too_short",
              "validationPassed": false,
              "message": "Password is too short, length: 6."
            }
          ]
        }
        );
      }
      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, { options: { debug: false, password: 'cli365' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "isValid": false,
          "validationResults": [
            {
              "ruleName": "password_too_short",
              "validationPassed": false,
              "message": "Password is too short, length: 6."
            }
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('password complexity is not met', (done) => {
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/beta/users/validatePassword' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "password": "cli365password"
        })) {
        return Promise.resolve({
          "isValid": false,
          "validationResults": [
            {
              "ruleName": "password_not_meet_complexity",
              "validationPassed": false,
              "message": "Password does not meet complexity requirements."
            }
          ]
        }
        );
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, { options: { debug: false, password: 'cli365password' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "isValid": false,
          "validationResults": [
            {
              "ruleName": "password_not_meet_complexity",
              "validationPassed": false,
              "message": "Password does not meet complexity requirements."
            }
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('password is too weak', (done) => {
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/beta/users/validatePassword' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "password": "MyP@ssW0rd"
        })) {
        return Promise.resolve({
          "isValid": false,
          "validationResults": [
            {
              "ruleName": "password_banned",
              "validationPassed": false,
              "message": "Password is too weak and can not be used."
            }
          ]
        }
        );
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, { options: { debug: false, password: 'MyP@ssW0rd' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "isValid": false,
          "validationResults": [
            {
              "ruleName": "password_banned",
              "validationPassed": false,
              "message": "Password is too weak and can not be used."
            }
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('password meets all requirements', (done) => {
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/beta/users/validatePassword' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "password": "cli365P@ssW0rd"
        })) {
        return Promise.resolve({
          "isValid": true,
          "validationResults": [
            {
              "ruleName": "AllChecks",
              "validationPassed": true,
              "message": "Password meets all validation requirements."
            }
          ]
        }
        );
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, { options: { debug: false, password: 'cli365P@ssW0rd' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "isValid": true,
          "validationResults": [
            {
              "ruleName": "AllChecks",
              "validationPassed": true,
              "message": "Password meets all validation requirements."
            }
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "Error",
          "message": "An error has occurred",
          "innerError": {
            "request-id": "9b0df954-93b5-4de9-8b99-43c204a9acf8",
            "date": "2021-12-08T18:56:48"
          }
        }
      });
    });

    command.action(logger, { options: { debug: false, password: 'cli365P@ssW0rd079654' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`An error has occurred`)));
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