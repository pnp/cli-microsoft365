import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./user-license-add');

describe(commands.USER_LICENSE_ADD, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validIds = '45715bb8-13f9-4bf6-927f-ef96c102d394,0118A350-71FC-4EC3-8F0C-6A1CB8867561';
  const validIdsSingle = '45715bb8-13f9-4bf6-927f-ef96c102d394';
  const validUserId = 'eb77fbcf-6fe8-458b-985d-1747284793bc';
  const validUserName = 'John@contos.onmicrosoft.com';
  const userLicenseResponse = {
    "businessPhones": [],
    "displayName": "John Doe",
    "givenName": null,
    "jobTitle": null,
    "mail": "John@contoso.onmicrosoft.com",
    "mobilePhone": null,
    "officeLocation": null,
    "preferredLanguage": null,
    "surname": null,
    "userPrincipalName": "John@contoso.onmicrosoft.com",
    "id": "eb77fbcf-6fe8-458b-985d-1747284793bc"
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_LICENSE_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if ids is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        ids: 'Invalid GUID', userId: validUserId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        ids: validIds, userId: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (ids)', async () => {
    const actual = await command.validate({ options: { ids: validIds, userId: validUserId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('adds licenses to a user by userId', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/users/${validUserId}/assignLicense`)) {
        return userLicenseResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { verbose: true, userId: validUserId, ids: validIds } });
    assert(loggerLogSpy.calledWith(userLicenseResponse));
  });

  it('adds single license to a user by userId', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/users/${validUserId}/assignLicense`)) {
        return userLicenseResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { verbose: true, userId: validUserId, ids: validIdsSingle } });
    assert(loggerLogSpy.calledWith(userLicenseResponse));
  });

  it('adds licenses to a user by userName', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/users/${validUserName}/assignLicense`)) {
        return userLicenseResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { verbose: true, userName: validUserName, ids: validIds } });
    assert(loggerLogSpy.calledWith(userLicenseResponse));
  });

  it('fails when one license is not a valid company license', async () => {
    const error = {
      error: {
        message: 'License 0118a350-71fc-4ec3-8f0c-6a1cb8867561 does not correspond to a valid company License.'
      }
    };

    sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/users/${validUserId}/assignLicense`)) {
        throw error;
      }

      throw `Invalid request ${opts.url}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true, userId: validUserId, ids: validIdsSingle
      }
    }), new CommandError(error.error.message));
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        message: 'The license cannot be added.'
      }
    };
    sinon.stub(request, 'post').callsFake(async () => { throw error; });

    await assert.rejects(command.action(logger, {
      options: {
        userName: validUserName, ids: validIds
      }
    }), new CommandError(error.error.message));
  });
});
