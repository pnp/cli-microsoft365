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
import { accessToken } from '../../../../utils/accessToken';
import { session } from '../../../../utils/session';
const command: Command = require('./user-license-list');

describe(commands.USER_LICENSE_LIST, () => {
  const userId = '59f80e08-24b1-41f8-8586-16765fd830d3';
  const userName = 'john.doe@contoso.com';
  const licenseResponse: any = {
    "value": [
      {
        "id": "x4s03usaBkSMs5fbAhyttK6cK8RP6rdKlxeBV2I1zKw",
        "skuId": "c42b9cae-ea4f-4ab7-9717-81576235ccac",
        "skuPartNumber": "DEVELOPERPACK_E5",
        "servicePlans": [
          {
            "servicePlanId": "b76fb638-6ba6-402a-b9f9-83d28acb3d86",
            "servicePlanName": "VIVA_LEARNING_SEEDED",
            "provisioningStatus": "PendingProvisioning",
            "appliesTo": "User"
          },
          {
            "servicePlanId": "7547a3fe-08ee-4ccb-b430-5077c5041653",
            "servicePlanName": "YAMMER_ENTERPRISE",
            "provisioningStatus": "Success",
            "appliesTo": "User"
          },
          {
            "servicePlanId": "eec0eb4f-6444-4f95-aba0-50c24d67f998",
            "servicePlanName": "AAD_PREMIUM_P2",
            "provisioningStatus": "Disabled",
            "appliesTo": "User"
          }
        ]
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
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
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      accessToken.isAppOnlyAccessToken
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_LICENSE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'skuId', 'skuPartNumber']);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName is not a valid UPN', async () => {
    const actual = await command.validate({ options: { userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with a userId defined', async () => {
    const actual = await command.validate({ options: { userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if no options specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('throws an error when using application permissions and no option is specified', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, {
      options: {}
    }), new CommandError(`Specify at least 'userId' or 'userName' when using application permissions.`));
  });

  it('retrieves license details of the current logged in user', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/licenseDetails') {
        return licenseResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith(licenseResponse.value));
  });

  it('retrieves license details of a specific user by its ID', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/licenseDetails`) {
        return licenseResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: userId } });
    assert(loggerLogSpy.calledWith(licenseResponse.value));
  });

  it('retrieves license details of a specific user by its UPN', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userName}/licenseDetails`) {
        return licenseResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName } });
    assert(loggerLogSpy.calledWith(licenseResponse.value));
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        message: `Resource '' does not exist or one of its queried reference-property objects are not present.`
      }
    };
    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, {
      options: { userName: userName }
    }), new CommandError(error.error.message));
  });
});