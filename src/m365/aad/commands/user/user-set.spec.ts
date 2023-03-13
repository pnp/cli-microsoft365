import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { telemetry } from '../../../../telemetry';
import { accessToken } from '../../../../utils/accessToken';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./user-set');

describe(commands.USER_SET, () => {
  const currentPassword = '9%9OLUg6p@Ra';
  const newPassword = 'iO$99OVj386i';
  const objectId = '1caf7dcd-7e83-4c3a-94f7-932a1299c844';
  const userPrincipalName = 'steve@contoso.onmicrosoft.com';
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
      request.patch,
      request.post,
      accessToken.getUserNameFromAccessToken,
      accessToken.getUserIdFromAccessToken
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
    assert.strictEqual(command.name.startsWith(commands.USER_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if neither the objectId nor the userPrincipalName are specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both the objectId and the userPrincipalName are specified', async () => {
    const actual = await command.validate({ options: { objectId: objectId, userPrincipalName: userPrincipalName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the objectId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { objectId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if currentPassword is set without newPassword', async () => {
    const actual = await command.validate({ options: { objectId: objectId, currentPassword: currentPassword } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if newPassword is set without currentPassword', async () => {
    const actual = await command.validate({ options: { objectId: objectId, newPassword: newPassword } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if resetPassword is set without a password', async () => {
    const actual = await command.validate({ options: { objectId: objectId, resetPassword: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if resetPassword and password is set and currentPassword is also set', async () => {
    const actual = await command.validate({ options: { objectId: objectId, resetPassword: true, password: newPassword, currentPassword: currentPassword } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when userPrincipalName has an invalid value', async () => {
    const actual = await command.validate({ options: { userPrincipalName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the objectId is a valid GUID', async () => {
    const actual = await command.validate({ options: { objectId: objectId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('allows unknown properties', () => {
    const allowUnknownOptions = command.allowUnknownOptions();
    assert.strictEqual(allowUnknownOptions, true);
  });

  it('throws error when objectId is not equal to current signed in objectId in Cli when passing both the options currentPassword and newPassword', async () => {
    sinon.stub(accessToken, 'getUserIdFromAccessToken').callsFake(() => { return '7c47b08e-e7b3-427a-9eba-b679815148e9'; });
    await assert.rejects(command.action(logger, { options: { verbose: true, objectId: objectId, newPassword: newPassword, currentPassword: currentPassword } } as any),
      new CommandError(`You can only change your own password. Please use --objectId @meId to reference to your own userId`));
  });

  it('throws error when userPrincipalName is not equal to current signed in userPrincipalName in Cli when passing both the options currentPassword and newPassword', async () => {
    sinon.stub(accessToken, 'getUserNameFromAccessToken').callsFake(() => { return 'john@contoso.com'; });
    await assert.rejects(command.action(logger, { options: { verbose: true, userPrincipalName: userPrincipalName, newPassword: newPassword, currentPassword: currentPassword } } as any),
      new CommandError(`You can only change your own password. Please use --userPrincipalName @meUserName to reference to your own user principal name`));
  });

  it('correctly handles user or property not found', async () => {
    sinon.stub(request, 'patch').callsFake(async () => {
      throw {
        "error": {
          "code": "Request_ResourceNotFound",
          "message": "Resource '1caf7dcd-7e83-4c3a-94f7-932a1299c844' does not exist or one of its queried reference-property objects are not present.",
          "innerError": {
            "request-id": "9b0df954-93b5-4de9-8b99-43c204a8aaf8",
            "date": "2018-04-24T18:56:48"
          }
        }
      };
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, objectId: objectId, NonExistingProperty: 'Value' } } as any),
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
        objectId: objectId,
        Department: 'Sales & Marketing',
        CompanyName: 'Contoso'
      }
    } as any);
    assert(loggerLogSpy.notCalled);
  });

  it('correctly resets password for a specified user by objectId', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${objectId}`
        && opts.data.passwordProfile !== undefined
        && opts.data.passwordProfile.password === newPassword
        && opts.data.passwordProfile.forceChangePasswordNextSignIn === true) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        objectId: objectId,
        resetPassword: true,
        newPassword: newPassword,
        forceChangePasswordNextSignIn: true
      }
    } as any);
    assert(loggerLogSpy.notCalled);
  });

  it('correctly resets password for a specified user by userPrincipalName', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userPrincipalName)}`
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
        userPrincipalName: userPrincipalName,
        resetPassword: true,
        newPassword: newPassword
      }
    } as any);
    assert(loggerLogSpy.notCalled);
  });

  it('correctly changes password for current user retrieved by userPrincipalName', async () => {
    sinon.stub(accessToken, 'getUserNameFromAccessToken').callsFake(() => { return userPrincipalName; });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userPrincipalName)}/changePassword`
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
        userPrincipalName: userPrincipalName,
        currentPassword: currentPassword,
        newPassword: newPassword
      }
    } as any);
    assert(loggerLogSpy.notCalled);
  });

  it('correctly changes password for current user retrieved by objectId', async () => {
    sinon.stub(accessToken, 'getUserIdFromAccessToken').callsFake(() => { return objectId; });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${objectId}/changePassword`
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
        objectId: objectId,
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
        userPrincipalName: userPrincipalName,
        accountEnabled: true
      }
    } as any);
    assert(loggerLogSpy.notCalled);
  });
});
