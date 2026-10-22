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
import command, { options } from './administrativeunit-member-remove.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraDevice } from '../../../../utils/entraDevice.js';

describe(commands.ADMINISTRATIVEUNIT_MEMBER_REMOVE, () => {
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
  let promptIssued: boolean;

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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      entraAdministrativeUnit.getAdministrativeUnitByDisplayName,
      entraUser.getUserIdByUpn,
      entraGroup.getGroupIdByDisplayName,
      entraDevice.getDeviceByDisplayName,
      cli.handleMultipleResultsFound,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ADMINISTRATIVEUNIT_MEMBER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when administrativeUnitId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: administrativeUnitId, id: '00000000-0000-0000-0000-000000000000' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: 'invalid', id: '00000000-0000-0000-0000-000000000000' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when id is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', id: userId });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', id: 'invalid' });
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

  it('fails validation if id, userId, userName, groupId, groupName, deviceId and deviceName options are not passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: administrativeUnitId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both id and userId options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', id: userId, userId: userId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both id and userName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', id: userId, userName: userName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both id and groupId options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', id: userId, groupId: groupId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both id and groupName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', id: userId, groupName: groupName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both id and deviceId options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', id: userId, deviceId: deviceId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both id and deviceName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ administrativeUnitId: '00000000-0000-0000-0000-000000000000', id: userId, deviceName: deviceName });
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

  it('removes the member specified by id from administrative unit specified by id without prompting for confirmation', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${userId}/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitId: administrativeUnitId, id: userId, force: true } });
    assert(deleteRequestStub.called);
  });

  it('removes the member specified by name from administrative unit specified by displayName while prompting for confirmation', async () => {
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').resolves({ id: administrativeUnitId, displayName: administrativeUnitName });
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${userId}/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { administrativeUnitName: administrativeUnitName, userName: userName } });
    assert(deleteRequestStub.called);
  });

  it('removes a member specified by its id from an administrative unit specified by its id', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${userId}/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitId: administrativeUnitId, id: userId, force: true } });
    assert(deleteRequestStub.called);
  });

  it('removes a user specified by its id from an administrative unit specified by its id', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${userId}/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitId: administrativeUnitId, userId: userId, force: true } });
    assert(deleteRequestStub.called);
  });

  it('removes a user specified by its name from an administrative unit specified by its name (verbose)', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${userId}/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitName: administrativeUnitName, userName: userName, force: true, verbose: true } });
    assert(deleteRequestStub.called);
  });

  it('removes a group specified by its id from an administrative unit specified by its id', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${groupId}/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitId: administrativeUnitId, groupId: groupId, force: true } });
    assert(deleteRequestStub.called);
  });

  it('removes a group specified by its name from an administrative unit specified by its name (verbose)', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs(groupName).resolves(groupId);
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${groupId}/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitName: administrativeUnitName, groupName: groupName, force: true, verbose: true } });
    assert(deleteRequestStub.called);
  });

  it('removes a device specified by its id from an administrative unit specified by its id', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${deviceId}/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitId: administrativeUnitId, deviceId: deviceId, force: true } });
    assert(deleteRequestStub.called);
  });

  it('removes a device specified by its name from an administrative unit specified by its name (verbose)', async () => {
    sinon.stub(entraDevice, 'getDeviceByDisplayName').withArgs(deviceName).resolves({ id: deviceId, displayName: deviceName });
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${deviceId}/$ref`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitName: administrativeUnitName, deviceName: deviceName, force: true, verbose: true } });
    assert(deleteRequestStub.called);
  });

  it('throws an error when administrative unit specified by id cannot be found', async () => {
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
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${userId}/$ref`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { administrativeUnitId: administrativeUnitId, userId: userId, force: true } }),
      new CommandError(error.error.message));
  });

  it('prompts before removing a member from an administrative unit when confirm option not passed', async () => {
    await command.action(logger, { options: { administrativeUnitId: administrativeUnitId, userId: userId } });

    assert(promptIssued);
  });
});