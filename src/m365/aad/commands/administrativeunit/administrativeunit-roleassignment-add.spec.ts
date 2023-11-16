import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import commands from '../../commands.js';
import command from './administrativeunit-roleassignment-add.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { aadAdministrativeUnit } from '../../../../utils/aadAdministrativeUnit.js';
import { aadUser } from '../../../../utils/aadUser.js';
import { roleAssignment } from '../../../../utils/roleAssignment.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.ADMINISTRATIVEUNIT_ROLEASSIGNMENT_ADD, () => {
  const roleDefinitionId = 'fe930be7-5e62-47db-91af-98c3a49a38b1';
  const roleDefinitionName = 'User Administrator';
  const userId = '2056d2f6-3257-4253-8cfc-b73393e414e5';
  const userName = 'AdeleVance@contoso.com';
  const administrativeUnitId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const administrativeUnitName = 'Marketing Department';
  const unifiedRoleAssignment = {
    "id": "BH21sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHIWb7-1",
    "roleDefinitionId": "fe930be7-5e62-47db-91af-98c3a49a38b1",
    "principalId": "2056d2f6-3257-4253-8cfc-b73393e414e5",
    "directoryScopeId": "/administrativeUnits/fc33aa61-cf0e-46b6-9506-f633347202ab"
  };

  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    (command as any).pollingInterval = 0;
  });

  afterEach(() => {
    sinonUtil.restore([
      aadAdministrativeUnit.getAdministrativeUnitByDisplayName,
      aadUser.getUserIdByUpn,
      roleAssignment.createRoleAssignmentWithAdministrativeUnitScope,
      roleDefinition.getRoleDefinitionByDisplayName,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ADMINISTRATIVEUNIT_ROLEASSIGNMENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation if administrative unit id, role definition id and user id are passed', async () => {
    const actual = await command.validate({
      options: {
        administrativeUnitId: administrativeUnitId,
        roleDefinitionId: roleDefinitionId,
        userId: userId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if administrative unit name, role definition name and user name are passed', async () => {
    const actual = await command.validate({
      options: {
        administrativeUnitName: administrativeUnitName,
        roleDefinitionName: roleDefinitionName,
        userName: userName
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if both user id and user name are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        administrativeUnitId: administrativeUnitId,
        roleDefinitionId: roleDefinitionId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both user id and user name are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        administrativeUnitId: administrativeUnitId,
        roleDefinitionId: roleDefinitionId,
        userId: userId,
        userName: userName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both role definition id and role definition name are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        administrativeUnitId: administrativeUnitId,
        userId: userId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both role definition id and role definition name are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        administrativeUnitId: administrativeUnitId,
        roleDefinitionId: roleDefinitionId,
        roleDefinitionName: roleDefinitionName,
        userId: userId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both administrative unit id and administrative unit name are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        roleDefinitionId: roleDefinitionId,
        userId: userId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both administrative unit id and administrative unit name are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        administrativeUnitId: administrativeUnitId,
        administrativeUnitName: administrativeUnitName,
        roleDefinitionId: roleDefinitionId,
        userId: userId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if administrative unit id is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        administrativeUnitId: '123',
        roleDefinitionId: roleDefinitionId,
        userId: userId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if role definition id is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        administrativeUnitId: administrativeUnitId,
        roleDefinitionId: '123',
        userId: userId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if user id is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        administrativeUnitId: administrativeUnitId,
        roleDefinitionId: roleDefinitionId,
        userId: '123'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('correctly assign a role specified by id to and administrative unit specified by id and to a user specified by id', async () => {
    sinon.stub(roleAssignment, 'createRoleAssignmentWithAdministrativeUnitScope').withArgs(roleDefinitionId, userId, administrativeUnitId).resolves(unifiedRoleAssignment);

    await command.action(logger, {
      options: {
        administrativeUnitId: administrativeUnitId,
        roleDefinitionId: roleDefinitionId,
        userId: userId
      }
    });
    //assert.strictEqual(loggerLogSpy.lastCall,'');
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignment));
  });

  it('correctly assign a role specified by name to and administrative unit specified by name and to a user specified by name', async () => {
    sinon.stub(aadAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });
    sinon.stub(aadUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(roleDefinition, 'getRoleDefinitionByDisplayName').withArgs(roleDefinitionName).resolves({ id: roleDefinitionId, displayName: roleDefinitionName });
    sinon.stub(roleAssignment, 'createRoleAssignmentWithAdministrativeUnitScope').withArgs(roleDefinitionId, userId, administrativeUnitId).resolves(unifiedRoleAssignment);

    await command.action(logger, {
      options: {
        administrativeUnitName: administrativeUnitName,
        roleDefinitionName: roleDefinitionName,
        userName: userName
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignment));
  });

  it('correctly handles error', async () => {
    sinon.stub(roleAssignment, 'createRoleAssignmentWithAdministrativeUnitScope').throws(Error('Invalid request'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Invalid request'));
  });
});