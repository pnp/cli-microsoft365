import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import commands from '../../commands.js';
import command from './administrativeunit-roleassignment-add.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { roleAssignment } from '../../../../utils/roleAssignment.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { settingsNames } from '../../../../settingsNames.js';
import request from '../../../../request.js';

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

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
      entraAdministrativeUnit.getAdministrativeUnitByDisplayName,
      entraUser.getUserIdByUpn,
      roleAssignment.createRoleAssignmentWithAdministrativeUnitScope,
      roleDefinition.getRoleDefinitionByDisplayName,
      cli.getSettingWithDefaultValue,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
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

  it('correctly assigns a role specified by id to and administrative unit specified by id and to a user specified by id', async () => {
    sinon.stub(roleAssignment, 'createRoleAssignmentWithAdministrativeUnitScope').withArgs(roleDefinitionId, userId, administrativeUnitId).resolves(unifiedRoleAssignment);

    await command.action(logger, {
      options: {
        administrativeUnitId: administrativeUnitId,
        roleDefinitionId: roleDefinitionId,
        userId: userId
      }
    });

    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignment));
  });

  it('correctly assigns a role specified by name to and administrative unit specified by name and to a user specified by name (verbose)', async () => {
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(roleDefinition, 'getRoleDefinitionByDisplayName').withArgs(roleDefinitionName).resolves({ id: roleDefinitionId, displayName: roleDefinitionName });
    sinon.stub(roleAssignment, 'createRoleAssignmentWithAdministrativeUnitScope').withArgs(roleDefinitionId, userId, administrativeUnitId).resolves(unifiedRoleAssignment);

    await command.action(logger, {
      options: {
        administrativeUnitName: administrativeUnitName,
        roleDefinitionName: roleDefinitionName,
        userName: userName,
        verbose: true
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignment));
  });

  it('correctly handles error', async () => {
    sinon.stub(roleAssignment, 'createRoleAssignmentWithAdministrativeUnitScope').throws(Error('Invalid request'));

    await assert.rejects(command.action(logger, {
      options: {
        administrativeUnitId: administrativeUnitId,
        roleDefinitionId: roleDefinitionId,
        userId: userId
      }
    }), new CommandError('Invalid request'));
  });

  it('fails if an administrative unit specified by name was not found', async () => {
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).throws(Error("The specified administrative unit 'Marketing Department' does not exist."));
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(roleDefinition, 'getRoleDefinitionByDisplayName').withArgs(roleDefinitionName).resolves({ id: roleDefinitionId, displayName: roleDefinitionName });
    sinon.stub(roleAssignment, 'createRoleAssignmentWithAdministrativeUnitScope').withArgs(roleDefinitionId, userId, administrativeUnitId).resolves(unifiedRoleAssignment);

    await assert.rejects(command.action(logger, {
      options: {
        administrativeUnitName: administrativeUnitName,
        roleDefinitionName: roleDefinitionName,
        userName: userName
      }
    }), new CommandError("The specified administrative unit 'Marketing Department' does not exist."));
  });

  it('fails if a role definition specified by name was not found', async () => {
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(roleDefinition, 'getRoleDefinitionByDisplayName').withArgs(roleDefinitionName).throws(Error("The specified role definition 'User Administrator' does not exist."));
    sinon.stub(roleAssignment, 'createRoleAssignmentWithAdministrativeUnitScope').withArgs(roleDefinitionId, userId, administrativeUnitId).resolves(unifiedRoleAssignment);

    await assert.rejects(command.action(logger, {
      options: {
        administrativeUnitName: administrativeUnitName,
        roleDefinitionName: roleDefinitionName,
        userName: userName
      }
    }), new CommandError("The specified role definition 'User Administrator' does not exist."));
  });

  it('fails if a user specified by UPN was not found', async () => {
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).throws(Error("The specified user with user name AdeleVance@contoso.com does not exist."));
    sinon.stub(roleDefinition, 'getRoleDefinitionByDisplayName').withArgs(roleDefinitionName).resolves({ id: roleDefinitionId, displayName: roleDefinitionName });
    sinon.stub(roleAssignment, 'createRoleAssignmentWithAdministrativeUnitScope').withArgs(roleDefinitionId, userId, administrativeUnitId).resolves(unifiedRoleAssignment);

    await assert.rejects(command.action(logger, {
      options: {
        administrativeUnitName: administrativeUnitName,
        roleDefinitionName: roleDefinitionName,
        userName: userName
      }
    }), new CommandError("The specified user with user name AdeleVance@contoso.com does not exist."));
  });

  it('correctly handles API OData error when creating role assignment with an administrative unit scope failed', async () => {
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(roleDefinition, 'getRoleDefinitionByDisplayName').withArgs(roleDefinitionName).resolves({ id: roleDefinitionId, displayName: roleDefinitionName });
    sinon.stub(request, 'post').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, {
      options: {
        administrativeUnitName: administrativeUnitName,
        roleDefinitionName: roleDefinitionName,
        userName: userName
      }
    }), new CommandError("Invalid request"));
  });
});