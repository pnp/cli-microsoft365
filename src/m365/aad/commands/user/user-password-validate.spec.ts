import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./user-password-validate');

describe(commands.USER_PASSWORD_VALIDATE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_PASSWORD_VALIDATE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('password is too short', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/beta/users/validatePassword' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "password": "cli365"
        })) {
        return {
          "isValid": false,
          "validationResults": [
            {
              "ruleName": "password_too_short",
              "validationPassed": false,
              "message": "Password is too short, length: 6."
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: { password: 'cli365' } });
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
  });

  it('password complexity is not met', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/beta/users/validatePassword' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "password": "cli365password"
        })) {
        return {
          "isValid": false,
          "validationResults": [
            {
              "ruleName": "password_not_meet_complexity",
              "validationPassed": false,
              "message": "Password does not meet complexity requirements."
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: { password: 'cli365password' } });
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
  });

  it('password is too weak', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/beta/users/validatePassword' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "password": "MyP@ssW0rd"
        })) {
        return {
          "isValid": false,
          "validationResults": [
            {
              "ruleName": "password_banned",
              "validationPassed": false,
              "message": "Password is too weak and can not be used."
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: { password: 'MyP@ssW0rd' } });
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
  });

  it('password meets all requirements', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/beta/users/validatePassword' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "password": "cli365P@ssW0rd"
        })) {
        return {
          "isValid": true,
          "validationResults": [
            {
              "ruleName": "AllChecks",
              "validationPassed": true,
              "message": "Password meets all validation requirements."
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: { password: 'cli365P@ssW0rd' } });
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
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'post').rejects({
      "error": {
        "code": "Error",
        "message": "An error has occurred",
        "innerError": {
          "request-id": "9b0df954-93b5-4de9-8b99-43c204a9acf8",
          "date": "2021-12-08T18:56:48"
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { password: 'cli365P@ssW0rd079654' } } as any),
      new CommandError(`An error has occurred`));
  });
});
