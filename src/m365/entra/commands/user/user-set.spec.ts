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
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './user-set.js';
import { settingsNames } from '../../../../settingsNames.js';
import aadCommands from '../../aadCommands.js';

describe(commands.USER_SET, () => {
  const currentPassword = '9%9OLUg6p@Ra';
  const newPassword = 'iO$99OVj386i';
  const id = '1caf7dcd-7e83-4c3a-94f7-932a1299c844';
  const userName = 'steve@contoso.onmicrosoft.com';
  const displayName = 'John';
  const firstName = 'John';
  const lastName = 'Doe';
  const usageLocation = 'BE';
  const officeLocation = 'New York';
  const jobTitle = 'Developer';
  const department = 'IT';
  const preferredLanguage = 'NL-be';
  const managerUserId = 'f4099688-dd3f-4a55-a9f5-ddd7417c227a';
  const managerUserName = 'doe@contoso.com';
  const largeString = 'f4gsz5cD0DmR7VpVXhsKlAwIryzpC847Z4qciQ1CDveZCNuCkWtUd9I8QXVLjurVS';

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    if (!auth.connection.accessTokens[auth.defaultResource]) {
      auth.connection.accessTokens[auth.defaultResource] = {
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch,
      request.post,
      request.put,
      accessToken.getUserNameFromAccessToken,
      accessToken.getUserIdFromAccessToken,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_SET);
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
    assert.deepStrictEqual(alias, [aadCommands.USER_SET]);
  });

  it('fails validation if neither the id nor the userName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both the id and the userName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { id: id, userName: userName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if currentPassword is set without newPassword', async () => {
    const actual = await command.validate({ options: { id: id, currentPassword: currentPassword } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if newPassword is set without currentPassword', async () => {
    const actual = await command.validate({ options: { id: id, newPassword: newPassword } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if resetPassword is set without a password', async () => {
    const actual = await command.validate({ options: { id: id, resetPassword: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if resetPassword and password is set and currentPassword is also set', async () => {
    const actual = await command.validate({ options: { id: id, resetPassword: true, password: newPassword, currentPassword: currentPassword } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when userName has an invalid value', async () => {
    const actual = await command.validate({ options: { userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation usageLocation is not a valid usageLocation', async () => {
    const actual = await command.validate({ options: { displayName: displayName, id: id, usageLocation: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation preferredLanguage is not a valid preferredLanguage', async () => {
    const actual = await command.validate({ options: { displayName: displayName, id: id, preferredLanguage: 'z' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both managerUserId and managerUserName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { displayName: displayName, id: id, managerUserId: managerUserId, managerUserName: managerUserName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if managerUserName is not a valid userName', async () => {
    const actual = await command.validate({ options: { displayName: displayName, id: id, managerUserName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if managerUserId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { displayName: displayName, id: id, managerUserId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if firstName has more than 64 characters', async () => {
    const actual = await command.validate({ options: { displayName: displayName, id: id, firstName: largeString } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if lastName has more than 64 characters', async () => {
    const actual = await command.validate({ options: { displayName: displayName, id: id, lastName: largeString } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if jobTitle has more than 128 characters', async () => {
    const actual = await command.validate({ options: { displayName: displayName, id: id, jobTitle: largeString + largeString } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if companyName has more than 64 characters', async () => {
    const actual = await command.validate({ options: { displayName: displayName, id: id, companyName: largeString } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if department has more than 64 characters', async () => {
    const actual = await command.validate({ options: { displayName: displayName, id: id, department: largeString } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if forceChangePasswordNextSignIn is set without resetPassword', async () => {
    const actual = await command.validate({ options: { id: id, forceChangePasswordNextSignIn: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if forceChangePasswordNextSignInWithMfa is set without resetPassword', async () => {
    const actual = await command.validate({ options: { id: id, forceChangePasswordNextSignInWithMfa: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: id } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('allows unknown properties', () => {
    const allowUnknownOptions = command.allowUnknownOptions();
    assert.strictEqual(allowUnknownOptions, true);
  });

  it('throws error when id is not equal to current signed in id in Cli when passing both the options currentPassword and newPassword', async () => {
    sinon.stub(accessToken, 'getUserIdFromAccessToken').returns('7c47b08e-e7b3-427a-9eba-b679815148e9');
    await assert.rejects(command.action(logger, { options: { verbose: true, id: id, newPassword: newPassword, currentPassword: currentPassword } } as any),
      new CommandError(`You can only change your own password. Please use --id @meId to reference to your own userId`));
  });

  it('throws error when userName is not equal to current signed in userName in Cli when passing both the options currentPassword and newPassword', async () => {
    sinon.stub(accessToken, 'getUserNameFromAccessToken').returns('john@contoso.com');
    await assert.rejects(command.action(logger, { options: { verbose: true, userName: userName, newPassword: newPassword, currentPassword: currentPassword } } as any),
      new CommandError(`You can only change your own password. Please use --userName @meUserName to reference to your own user principal name`));
  });

  it('correctly handles user or property not found', async () => {
    sinon.stub(request, 'patch').rejects({
      "error": {
        "code": "Request_ResourceNotFound",
        "message": "Resource '1caf7dcd-7e83-4c3a-94f7-932a1299c844' does not exist or one of its queried reference-property objects are not present.",
        "innerError": {
          "request-id": "9b0df954-93b5-4de9-8b99-43c204a8aaf8",
          "date": "2018-04-24T18:56:48"
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, id: id, NonExistingProperty: 'Value' } } as any),
      new CommandError(`Resource '1caf7dcd-7e83-4c3a-94f7-932a1299c844' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('correctly updates information about the specified user', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/users/`) > -1) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        id: id,
        Department: 'Sales & Marketing',
        companyName: 'Contoso',
        displayName: displayName,
        firstName: firstName,
        lastName: lastName,
        usageLocation: usageLocation,
        officeLocation: officeLocation,
        jobTitle: jobTitle,
        department: department,
        preferredLanguage: preferredLanguage
      }
    } as any);
    assert(loggerLogSpy.notCalled);
  });

  it('correctly updates user with an empty value', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${id}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: id,
        companyName: ''
      }
    } as any);

    assert.strictEqual(patchStub.lastCall.args[0].data.companyName, null);
  });

  it('correctly resets password for a specified user by id', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${id}`
        && opts.data.passwordProfile !== undefined
        && opts.data.passwordProfile.password === newPassword
        && opts.data.passwordProfile.forceChangePasswordNextSignIn === true
        && opts.data.passwordProfile.forceChangePasswordNextSignInWithMfa === true) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        id: id,
        resetPassword: true,
        newPassword: newPassword,
        forceChangePasswordNextSignIn: true,
        forceChangePasswordNextSignInWithMfa: true
      }
    } as any);
    assert(loggerLogSpy.notCalled);
  });

  it('correctly resets password for a specified user by userName', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userName)}`
        && opts.data.passwordProfile !== undefined
        && opts.data.passwordProfile.password === newPassword
        && opts.data.passwordProfile.forceChangePasswordNextSignIn === false) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        userName: userName,
        resetPassword: true,
        newPassword: newPassword
      }
    } as any);
    assert(loggerLogSpy.notCalled);
  });

  it('correctly changes password for current user retrieved by userName', async () => {
    sinon.stub(accessToken, 'getUserNameFromAccessToken').returns(userName);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userName)}/changePassword`
        && opts.data !== undefined
        && opts.data.currentPassword === currentPassword
        && opts.data.newPassword === newPassword) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        userName: userName,
        currentPassword: currentPassword,
        newPassword: newPassword
      }
    } as any);
    assert(loggerLogSpy.notCalled);
  });

  it('correctly changes password for current user retrieved by id', async () => {
    sinon.stub(accessToken, 'getUserIdFromAccessToken').returns(id);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${id}/changePassword`
        && opts.data !== undefined
        && opts.data.currentPassword === currentPassword
        && opts.data.newPassword === newPassword) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        id: id,
        currentPassword: currentPassword,
        newPassword: newPassword
      }
    } as any);
    assert(loggerLogSpy.notCalled);
  });

  it('correctly enables the specified user', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/users/`) > -1) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        userName: userName,
        accountEnabled: true
      }
    } as any);
    assert(loggerLogSpy.notCalled);
  });

  it('updates Microsoft Entra user and set its manager by id', async () => {
    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userName}/manager/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName, managerUserId: managerUserId } });
    assert.deepEqual(putStub.lastCall.args[0].data, {
      '@odata.id': `https://graph.microsoft.com/v1.0/users/${managerUserId}`
    });
  });

  it('updates Microsoft Entra user and set its manager by user principal name', async () => {
    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userName}/manager/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, userName: userName, managerUserName: managerUserName } });
    assert.deepEqual(putStub.lastCall.args[0].data, {
      '@odata.id': `https://graph.microsoft.com/v1.0/users/${managerUserName}`
    });
  });

  it('updates Microsoft Entra user and removes manager', async () => {
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userName}/manager/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, userName: userName, removeManager: true } });
    assert(deleteStub.called);
  });
});
