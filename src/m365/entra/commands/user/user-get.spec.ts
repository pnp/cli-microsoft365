import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './user-get.js';
import { settingsNames } from '../../../../settingsNames.js';
import aadCommands from '../../aadCommands.js';
import { aadUser } from '../../../../utils/aadUser.js';
import { formatting } from '../../../../utils/formatting.js';

describe(commands.USER_GET, () => {
  const userId = "68be84bf-a585-4776-80b3-30aa5207aa21";
  const userName = "AarifS@contoso.onmicrosoft.com";
  const resultValue = { "id": "68be84bf-a585-4776-80b3-30aa5207aa21", "businessPhones": ["+1 425 555 0100"], "displayName": "Aarif Sherzai", "givenName": "Aarif", "jobTitle": "Administrative", "mail": null, "mobilePhone": "+1 425 555 0100", "officeLocation": null, "preferredLanguage": null, "surname": "Sherzai", "userPrincipalName": "AarifS@contoso.onmicrosoft.com" };
  const externalUserName = "john.doe_microsoft.com#EXT#@contoso.com";
  const externalUserResponse = { "id": "eb77fbcf-6fe8-458b-985d-1747284793bc", "businessPhones": ["+420 605 123 456"], "displayName": "John Doe", "givenName": "John", "jobTitle": "External consultant", "mail": null, "mobilePhone": "+420 605 123 456", "officeLocation": null, "preferredLanguage": null, "surname": "Doe", "userPrincipalName": "john.doe_microsoft.com#EXT#@contoso.com" };
  const userNameWithDollar = "$john.doe@contoso.com";
  const userNameWithDollarResponse = { "id": "eb77fbcf-6fe8-458b-985d-1747284793bc", "businessPhones": ["+420 605 123 456"], "displayName": "John Doe", "givenName": "John", "jobTitle": "Consultant", "mail": null, "mobilePhone": "+420 605 123 456", "officeLocation": null, "preferredLanguage": null, "surname": "Doe", "userPrincipalName": "$john.doe@contoso.com" };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    if (!auth.service.accessTokens[auth.defaultResource]) {
      auth.service.accessTokens[auth.defaultResource] = {
        expiresOn: '123',
        accessToken: 'abc'
      };
    }
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      accessToken.getUserIdFromAccessToken,
      accessToken.getUserNameFromAccessToken,
      cli.getSettingWithDefaultValue,
      aadUser.getUserIdByEmail
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [aadCommands.USER_GET]);
  });

  it('retrieves user using id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}`) {
        return resultValue;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: userId } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves user using @userid token', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}`) {
        return resultValue;
      }

      throw 'Invalid request';
    });

    sinon.stub(accessToken, 'getUserIdFromAccessToken').callsFake(() => { return userId; });

    await command.action(logger, { options: { id: '@meid' } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves user using id (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}`) {
        return resultValue;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: userId } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves user using user name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userName) }`) {
        return resultValue;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves external user using user name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(externalUserName)}`) {
        return externalUserResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: externalUserName } });
    assert(loggerLogSpy.calledWith(externalUserResponse));
  });

  it('retrieves user using user name which starts with $', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${formatting.encodeQueryParameter(userNameWithDollar)}')`) {
        return userNameWithDollarResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userNameWithDollar } });
    assert(loggerLogSpy.calledWith(userNameWithDollarResponse));
  });

  it('retrieves user using user name and with their direct manager', async () => {
    const resultValueWithManger: any = { ...resultValue };
    resultValueWithManger.manager = {
      "displayName": "John Doe",
      "userPrincipalName": "john.doe@contoso.onmicrosoft.com",
      "id": "eb77fbcf-6fe8-458b-985d-1747284793bc",
      "mail": "john.doe@contoso.onmicrosoft.com"
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userName) }?$expand=manager($select=displayName,userPrincipalName,id,mail)`) {
        return resultValueWithManger;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName, withManager: true } });
    assert(loggerLogSpy.calledWith(resultValueWithManger));
  });

  it('retrieves user using @meusername token', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userName) }`) {
        return resultValue;
      }

      throw 'Invalid request';
    });

    sinon.stub(accessToken, 'getUserNameFromAccessToken').callsFake(() => { return userName; });

    await command.action(logger, { options: { userName: '@meusername' } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves user using email', async () => {
    sinon.stub(aadUser, 'getUserIdByEmail').withArgs(userName).resolves(userId);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}`) {
        return resultValue;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { email: userName } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves only the specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userName) }?$select=id,mail`) {
        return { "id": "userId", "mail": null };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName, properties: 'id,mail' } });
    assert(loggerLogSpy.calledWith({ "id": "userId", "mail": null }));
  });

  it('fails to get user when user with provided email does not exists', async () => {
    sinon.stub(aadUser, 'getUserIdByEmail').withArgs(userName).throws(Error(`The specified user with email ${userName} does not exist`));

    await assert.rejects(command.action(logger, { options: { email: userName } }),
      new CommandError(`The specified user with email ${userName} does not exist`));
  });

  it('correctly handles error when user provided by id was not found', async () => {
    sinon.stub(request, 'get').rejects({
      "error": {
        "code": "Request_ResourceNotFound",
        "message": "Resource '68be84bf-a585-4776-80b3-30aa5207aa22' does not exist or one of its queried reference-property objects are not present.",
        "innerError": {
          "request-id": "9b0df954-93b5-4de9-8b99-43c204a8aaf8",
          "date": "2018-04-24T18:56:48"
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { id: '68be84bf-a585-4776-80b3-30aa5207aa22' } } as any),
      new CommandError(`Resource '68be84bf-a585-4776-80b3-30aa5207aa22' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('correctly handles error when user provided by userName was not found', async () => {
    sinon.stub(request, 'get').rejects({
      "error": {
        "code": "Request_ResourceNotFound",
        "message": `Resource '${userName}' does not exist or one of its queried reference-property objects are not present.`,
        "innerError": {
          "request-id": "9b0df954-93b5-4de9-8b99-43c204a8aaf8",
          "date": "2018-04-24T18:56:48"
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { userName: userName } } as any),
      new CommandError(`Resource '${userName}' does not exist or one of its queried reference-property objects are not present.`));
  });
  
  it('fails validation if id or email or userName options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id, email, and userName options are passed (multiple options)', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { id: "1caf7dcd-7e83-4c3a-94f7-932a1299c844", email: "john.doe@contoso.onmicrosoft.com", userName: "i:0#.f|membership|john.doe@contoso.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and email options are passed (multiple options)', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { id: "1caf7dcd-7e83-4c3a-94f7-932a1299c844", email: "john.doe@contoso.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and userName options are passed (multiple options)', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { id: "1caf7dcd-7e83-4c3a-94f7-932a1299c844", userName: "john.doe@contoso.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both email and userName options are passed (multiple options)', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { email: "jonh.deo@contoso.com", userName: "john.doe@contoso.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when userName has an invalid value', async () => {
    const actual = await command.validate({ options: { userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '68be84bf-a585-4776-80b3-30aa5207aa22' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the userName is specified', async () => {
    const actual = await command.validate({ options: { userName: 'john.doe@contoso.onmicrosoft.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the email is specified', async () => {
    const actual = await command.validate({ options: { email: 'john.doe@contoso.onmicrosoft.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
