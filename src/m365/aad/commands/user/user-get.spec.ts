import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { accessToken } from '../../../../utils/accessToken';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./user-get');

describe(commands.USER_GET, () => {
  const userId = "68be84bf-a585-4776-80b3-30aa5207aa21";
  const userName = "AarifS@contoso.onmicrosoft.com";
  const resultValue = { "id": "68be84bf-a585-4776-80b3-30aa5207aa21", "businessPhones": ["+1 425 555 0100"], "displayName": "Aarif Sherzai", "givenName": "Aarif", "jobTitle": "Administrative", "mail": null, "mobilePhone": "+1 425 555 0100", "officeLocation": null, "preferredLanguage": null, "surname": "Sherzai", "userPrincipalName": "AarifS@contoso.onmicrosoft.com" };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
    if (!auth.service.accessTokens[auth.defaultResource]) {
      auth.service.accessTokens[auth.defaultResource] = {
        expiresOn: '123',
        accessToken: 'abc'
      };
    }
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      accessToken.getUserIdFromAccessToken,
      accessToken.getUserNameFromAccessToken
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.USER_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves user using id', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=id eq '${userId}'`) > -1) {
        return Promise.resolve({ value: [resultValue] });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { id: userId } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves user using @userid token', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=id eq '${userId}'`) > -1) {
        return Promise.resolve({ value: [resultValue] });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(accessToken, 'getUserIdFromAccessToken').callsFake(() => { return userId; });

    await command.action(logger, { options: { id: '@meid' } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves user using id (debug)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=id eq '${userId}'`) > -1) {
        return Promise.resolve({ value: [resultValue] });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, id: userId } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves user using user name', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(userName)}'`) > -1) {
        return Promise.resolve({ value: [resultValue] });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { userName: userName } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves user using @meusername token', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(userName)}'`) > -1) {
        return Promise.resolve({ value: [resultValue] });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(accessToken, 'getUserNameFromAccessToken').callsFake(() => { return userName; });

    await command.action(logger, { options: { userName: '@meusername' } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves user using email', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=mail eq '${formatting.encodeQueryParameter(userName)}'`) > -1) {
        return Promise.resolve({ value: [resultValue] });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { email: userName } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves only the specified properties', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(userName)}'&$select=id,mail`) {
        return Promise.resolve({ value: [{ "id": "userId", "mail": null }] });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { userName: userName, properties: 'id,mail' } });
    assert(loggerLogSpy.calledWith({ "id": "userId", "mail": null }));
  });

  it('correctly handles user not found', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "Request_ResourceNotFound",
          "message": "Resource '68be84bf-a585-4776-80b3-30aa5207aa22' does not exist or one of its queried reference-property objects are not present.",
          "innerError": {
            "request-id": "9b0df954-93b5-4de9-8b99-43c204a8aaf8",
            "date": "2018-04-24T18:56:48"
          }
        }
      });
    });

    await assert.rejects(command.action(logger, { options: { id: '68be84bf-a585-4776-80b3-30aa5207aa22' } } as any),
      new CommandError(`Resource '68be84bf-a585-4776-80b3-30aa5207aa22' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('fails to get user when user with provided id does not exists', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=id eq '${userId}'`) > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`The specified user with id ${userId} does not exist`);
    });

    await assert.rejects(command.action(logger, { options: { id: userId } }),
      new CommandError(`The specified user with id ${userId} does not exist`));
  });

  it('fails to get user when user with provided user name does not exists', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(userName)}'`) > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`The specified user with user name ${userName} does not exist`);
    });

    await assert.rejects(command.action(logger, { options: { userName: userName } }),
      new CommandError(`The specified user with user name ${userName} does not exist`));
  });

  it('fails to get user when user with provided email does not exists', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=mail eq '${formatting.encodeQueryParameter(userName)}'`) > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`The specified user with email ${userName} does not exist`);
    });

    await assert.rejects(command.action(logger, { options: { email: userName } }),
      new CommandError(`The specified user with email ${userName} does not exist`));
  });

  it('handles error when multiple users with the specified email found', async () => {
    sinon.stub(request, 'get').callsFake(opts => {
      if ((opts.url as string).indexOf('https://graph.microsoft.com/v1.0/users?$filter') > -1) {
        return Promise.resolve({
          value: [
            resultValue,
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', userPrincipalName: 'DebraB@contoso.onmicrosoft.com' }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        email: userName
      }
    }), new CommandError(`Multiple users with email ${userName} found. Please disambiguate (user names): ${userName}, DebraB@contoso.onmicrosoft.com or (ids): ${userId}, 9b1b1e42-794b-4c71-93ac-5ed92488b67f`));
  });

  it('fails validation if id or email or userName options are not passed', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id, email, and userName options are passed (multiple options)', async () => {
    const actual = await command.validate({ options: { id: "1caf7dcd-7e83-4c3a-94f7-932a1299c844", email: "john.doe@contoso.onmicrosoft.com", userName: "i:0#.f|membership|john.doe@contoso.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and email options are passed (multiple options)', async () => {
    const actual = await command.validate({ options: { id: "1caf7dcd-7e83-4c3a-94f7-932a1299c844", email: "john.doe@contoso.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and userName options are passed (multiple options)', async () => {
    const actual = await command.validate({ options: { id: "1caf7dcd-7e83-4c3a-94f7-932a1299c844", userName: "john.doe@contoso.onmicrosoft.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both email and userName options are passed (multiple options)', async () => {
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
