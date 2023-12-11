import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import command from './administrativeunit-member-add.js';
import { aadAdministrativeUnit } from '../../../../utils/aadAdministrativeUnit.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { aadUser } from '../../../../utils/aadUser.js';
import { aadDevice } from '../../../../utils/aadDevice.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.ADMINISTRATIVEUNIT_MEMBER_ADD, () => {
  const administrativeUnitId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const administrativeUnitName = 'European Division';
  const userId = '23b415fb-baea-4995-a26e-c74073beadff';
  const userName = 'adele.vence@contoso.com';
  const groupId = '593af7e2-d27e-42b8-ad17-abe5e57dab61';
  const groupName = 'Marketing';
  const deviceId = '60c99a96-70af-4d68-a8dc-5c51b345c6ce';
  const deviceName = 'AdeleVence-PC';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      aadAdministrativeUnit.getAdministrativeUnitByDisplayName,
      aadUser.getUserIdByUpn,
      aadGroup.getGroupIdByDisplayName,
      aadDevice.getDeviceByDisplayName,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound,
      cli.promptForSelection
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ADMINISTRATIVEUNIT_MEMBER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when administrativeUnitId is a valid GUID', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: administrativeUnitId, userId: '00000000-0000-0000-0000-000000000000' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if administrativeUnitId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: 'invalid', userId: '00000000-0000-0000-0000-000000000000' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when groupId is a valid GUID', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', groupId: groupId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', groupId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when deviceId is a valid GUID', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', deviceId: deviceId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if deviceId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', deviceId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both administrativeUnitId and administrativeUnitName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: administrativeUnitId, administrativeUnitName: administrativeUnitName, userId: '00000000-0000-0000-0000-000000000000' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both administrativeUnitId and administrativeUnitName options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { userId: userId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both userId and userName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', userId: userId, userName: userName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both userId and groupId options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', userId: userId, groupId: groupId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both userId and groupName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', userId: userId, groupName: groupName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both userId and deviceId options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', userId: userId, deviceId: deviceId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both userId and deviceName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', userId: userId, deviceName: deviceName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both userName and groupId options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', userName: userName, groupId: groupId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both userName and groupName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', userName: userName, groupName: groupName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both userName and deviceId options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', userName: userName, deviceId: deviceId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both userName and deviceName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', userName: userName, deviceName: deviceName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both groupId and groupName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', groupId: groupId, groupName: groupName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both groupId and deviceId options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', groupId: groupId, deviceId: deviceId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both groupId and deviceName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', groupId: groupId, deviceName: deviceName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both groupName and deviceId options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', groupName: groupName, deviceId: deviceId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both groupName and deviceName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', groupName: groupName, deviceName: deviceName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both deviceId and deviceName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: '00000000-0000-0000-0000-000000000000', deviceId: deviceId, deviceName: deviceName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all the required options are specified', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: administrativeUnitId, userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('adds a user specified by its id to an administrative unit specified by its id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitId: administrativeUnitId, userId: userId } });
    assert(postStub.called);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "@odata.id": `https://graph.microsoft.com/v1.0/users/${userId}` });
  });

  it('adds a user specified by its name to an administrative unit specified by its name', async () => {
    sinon.stub(aadUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(aadAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitName: administrativeUnitName, userName: userName } });
    assert(postStub.called);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "@odata.id": `https://graph.microsoft.com/v1.0/users/${userId}` });
  });

  it('adds a group specified by its id to an administrative unit specified by its id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitId: administrativeUnitId, groupId: groupId } });
    assert(postStub.called);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "@odata.id": `https://graph.microsoft.com/v1.0/groups/${groupId}` });
  });

  it('adds a group specified by its name to an administrative unit specified by its name', async () => {
    sinon.stub(aadGroup, 'getGroupIdByDisplayName').withArgs(groupName).resolves(groupId);
    sinon.stub(aadAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitName: administrativeUnitName, groupName: groupName } });
    assert(postStub.called);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "@odata.id": `https://graph.microsoft.com/v1.0/groups/${groupId}` });
  });

  it('adds a device specified by its id to an administrative unit specified by its id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitId: administrativeUnitId, deviceId: deviceId } });
    assert(postStub.called);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "@odata.id": `https://graph.microsoft.com/v1.0/devices/${deviceId}` });
  });

  it('adds a device specified by its name to an administrative unit specified by its name', async () => {
    sinon.stub(aadDevice, 'getDeviceByDisplayName').withArgs(deviceName).resolves({ id: deviceId, displayName: deviceName });
    sinon.stub(aadAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitName: administrativeUnitName, deviceName: deviceName } });
    assert(postStub.called);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { "@odata.id": `https://graph.microsoft.com/v1.0/devices/${deviceId}` });
  });

  it('throws an error when administrative unit by id cannot be found', async () => {
    const error = {
      error: {
        code: 'Request_ResourceNotFound',
        message: `Resource '${administrativeUnitId}' does not exist or one of its queried reference-property objects are not present.`,
        innerError: {
          date: '2023-10-27T12:24:36',
          'request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b',
          'client-request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b'
        }
      }
    };
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/$ref`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { administrativeUnitId: administrativeUnitId, userId: userId } }), new CommandError(error.error.message));
  });
});