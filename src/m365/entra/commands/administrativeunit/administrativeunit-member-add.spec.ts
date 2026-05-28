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
import command, { options } from './administrativeunit-member-add.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraDevice } from '../../../../utils/entraDevice.js';

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
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
      entraAdministrativeUnit.getAdministrativeUnitByDisplayName,
      entraUser.getUserIdByUpn,
      entraGroup.getGroupIdByDisplayName,
      entraDevice.getDeviceByDisplayName,
      cli.handleMultipleResultsFound,
      cli.promptForSelection
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ADMINISTRATIVEUNIT_MEMBER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when administrativeUnitId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: administrativeUnitId, userId: '00000000-0000-0000-0000-000000000000' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: 'invalid', userId: '00000000-0000-0000-0000-000000000000' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when userId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', userId: userId });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', userId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when groupId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', groupId: groupId });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if groupId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', groupId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when deviceId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', deviceId: deviceId });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if deviceId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', deviceId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both administrativeUnitId and administrativeUnitName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: administrativeUnitId, administrativeUnitName: administrativeUnitName, userId: '00000000-0000-0000-0000-000000000000' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both administrativeUnitId and administrativeUnitName options are not passed', () => {
    const actual = commandOptionsSchema.safeParse({ userId: userId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both userId and userName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', userId: userId, userName: userName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both userId and groupId options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', userId: userId, groupId: groupId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both userId and groupName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', userId: userId, groupName: groupName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both userId and deviceId options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', userId: userId, deviceId: deviceId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both userId and deviceName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', userId: userId, deviceName: deviceName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both userName and groupId options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', userName: userName, groupId: groupId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both userName and groupName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', userName: userName, groupName: groupName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both userName and deviceId options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', userName: userName, deviceId: deviceId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both userName and deviceName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', userName: userName, deviceName: deviceName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both groupId and groupName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', groupId: groupId, groupName: groupName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both groupId and deviceId options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', groupId: groupId, deviceId: deviceId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both groupId and deviceName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', groupId: groupId, deviceName: deviceName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both groupName and deviceId options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', groupName: groupName, deviceId: deviceId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both groupName and deviceName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', groupName: groupName, deviceName: deviceName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both deviceId and deviceName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', deviceId: deviceId, deviceName: deviceName });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if all the required options are specified', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: administrativeUnitId, userId: userId });
    assert.strictEqual(actual.success, true);
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

  it('adds a user specified by its name to an administrative unit specified by its name (verbose)', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitName: administrativeUnitName, userName: userName, verbose: true } });
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

  it('adds a group specified by its name to an administrative unit specified by its name (verbose)', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs(groupName).resolves(groupId);
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitName: administrativeUnitName, groupName: groupName, verbose: true } });
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

  it('adds a device specified by its name to an administrative unit specified by its name (verbose)', async () => {
    sinon.stub(entraDevice, 'getDeviceByDisplayName').withArgs(deviceName).resolves({ id: deviceId, displayName: deviceName });
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitName: administrativeUnitName, deviceName: deviceName, verbose: true } });
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