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
const command: Command = require('./user-add');

describe(commands.USER_GET, () => {
  const graphBaseUrl = 'https://graph.microsoft.com/v1.0/users';
  const userName = 'john@contoso.com';
  const displayName = 'John';
  const accountEnabled = true;
  const mailNickName = 'john';
  const password = 'R@ndom1!';
  const firstName = 'John';
  const lastName = 'Doe';
  const usageLocation = 'BE';
  const officeLocation = 'Vosselaar';
  const jobTitle = 'Developer';
  const companyName = 'Microsoft';
  const department = 'IT';
  const preferredLanguage = 'NL-be';
  const managerUserId = 'f4099688-dd3f-4a55-a9f5-ddd7417c227a';
  const managerUserName = 'doe@contoso.com';
  const largeString = 'f4gsz5cD0DmR7VpVXhsKlAwIryzpC847Z4qciQ1CDveZCNuCkWtUd9I8QXVLjurVS';

  const userResponseWithoutPassword = {
    id: "f5caff1f-e9b6-4dba-a65e-d0c908c0e91b",
    businessPhones: [],
    displayName: displayName,
    givenName: firstName,
    jobTitle: jobTitle,
    mail: null,
    mobilePhone: null,
    officeLocation: officeLocation,
    preferredLanguage: preferredLanguage,
    surname: lastName,
    userPrincipalName: userName
  };

  const userResponseWithPassword = {
    ...userResponseWithoutPassword,
    password: password
  };

  const graphError = {
    error: {
      code: "Request_BadRequest",
      message: "Another object with the same value for property userPrincipalName already exists.",
      details: [
        {
          code: "ObjectConflict",
          message: "Another object with the same value for property userPrincipalName already exists.",
          target: "userPrincipalName"
        }
      ],
      innerError: {
        date: "2023-02-16T17:22:25",
        'request-id': "2726a9e1-2909-4277-ba89-144558eb9431",
        'client-request-id': "2726a9e1-2909-4277-ba89-144558eb9431"
      }
    }
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.put
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
    assert.strictEqual(command.name, commands.USER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates Azure AD user using a preset password but requiring the user to change it the next login', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === graphBaseUrl) {
        return userResponseWithoutPassword;
      }

      throw 'Invalid request';
    });


    await command.action(logger, { options: { verbose: true, userName: userName, displayName: displayName, password: password, forceChangePasswordNextSignIn: true } });
    assert(loggerLogSpy.calledWith(userResponseWithPassword));
  });

  it('creates Azure AD user and set its manager by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === graphBaseUrl) {
        return userResponseWithoutPassword;
      }

      throw 'Invalid request';
    });

    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `${graphBaseUrl}/${userResponseWithPassword.id}/manager/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName, displayName: displayName, managerUserId: managerUserId } });
    assert.strictEqual(putStub.lastCall.args[0].data['@odata.id'], `${graphBaseUrl}/${managerUserId}`);
  });

  it('creates Azure AD user and set its manager by user principal name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === graphBaseUrl) {
        return userResponseWithoutPassword;
      }

      throw 'Invalid request';
    });

    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `${graphBaseUrl}/${userResponseWithPassword.id}/manager/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName, displayName: displayName, managerUserName: managerUserName } });
    assert.strictEqual(putStub.lastCall.args[0].data['@odata.id'], `${graphBaseUrl}/${managerUserName}`);
  });

  it('correctly handles Graph error when userPrincipalName already exists in the tenant', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === graphBaseUrl) {
        throw graphError;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { userName: userName, displayName: displayName } }),
      new CommandError(graphError.error.message));
  });

  it('fails validation if userName is not a valid userPrincipalName', async () => {
    const actual = await command.validate({ options: { displayName: displayName, userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation usageLocation is not a valid usageLocation', async () => {
    const actual = await command.validate({ options: { displayName: displayName, userName: userName, usageLocation: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation preferredLanguage is not a valid preferredLanguage', async () => {
    const actual = await command.validate({ options: { displayName: displayName, userName: userName, preferredLanguage: 'z' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both managerUserId and managerUserName are specified', async () => {
    const actual = await command.validate({ options: { displayName: displayName, userName: userName, managerUserId: managerUserId, managerUserName: managerUserName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if managerUserName is not a valid userPrincipalName', async () => {
    const actual = await command.validate({ options: { displayName: displayName, userName: userName, managerUserName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if managerUserId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { displayName: displayName, userName: userName, managerUserId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if firstName has more than 64 characters', async () => {
    const actual = await command.validate({ options: { displayName: displayName, userName: userName, firstName: largeString } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if lastName has more than 64 characters', async () => {
    const actual = await command.validate({ options: { displayName: displayName, userName: userName, lastName: largeString } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if jobTitle has more than 128 characters', async () => {
    const actual = await command.validate({ options: { displayName: displayName, userName: userName, jobTitle: largeString + largeString } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if companyName has more than 64 characters', async () => {
    const actual = await command.validate({ options: { displayName: displayName, userName: userName, companyName: largeString } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if department has more than 64 characters', async () => {
    const actual = await command.validate({ options: { displayName: displayName, userName: userName, department: largeString } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if only userName and displayName are specified', async () => {
    const actual = await command.validate({ options: { displayName: displayName, userName: userName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if all options (excluding managerUserName and forceChangePasswordNextSignInWithMfa) are specified', async () => {
    const actual = await command.validate({ options: { displayName: displayName, userName: userName, accountEnabled: accountEnabled, mailNickname: mailNickName, password: password, firstName: firstName, lastName: lastName, forceChangePasswordNextSignIn: true, usageLocation: usageLocation, officeLocation: officeLocation, jobTitle: jobTitle, companyName: companyName, department: department, preferredLanguage: preferredLanguage, managerUserId: managerUserId } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
