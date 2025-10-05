import { AdministrativeUnit, UnifiedRoleAssignment, UnifiedRoleDefinition } from '@microsoft/microsoft-graph-types';
import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { customAppScope } from '../../../../utils/customAppScope.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraServicePrincipal } from '../../../../utils/entraServicePrincipal.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { pid } from '../../../../utils/pid.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './approleassignment-add.js';

describe(commands.APPROLEASSIGNMENT_ADD, () => {
  const principalId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const principalName = 'ContosoApp';
  const roleDefinitionId = 'abc3aa61-cf0e-46b6-9506-f63334729876';
  const roleDefinitionName = 'Application Mail.Read';
  const userId = '7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091';
  const userName = 'john.doe@contoso.com';
  const groupId = '1a70e568-d286-4ad1-b036-734ff8667915';
  const groupName = 'Contoso Group';
  const administrativeUnitId = '31537af4-6d77-4bb9-a681-d2394888ea26';
  const administrativeUnitName = 'Contoso Administrative Unit';
  const customAppScopeId = '0d2ba203-f407-4de0-9f6a-7e411484e4da';
  const customAppScopeName = 'Users from Marketing Department';
  const unifiedRoleDefinition: UnifiedRoleDefinition = {
    id: roleDefinitionId,
    displayName: roleDefinitionName
  };
  const administrativeUnit: AdministrativeUnit = {
    id: administrativeUnitId,
    displayName: administrativeUnitName
  };
  const customApplicationScope = {
    "id": "0d2ba203-f407-4de0-9f6a-7e411484e4da",
    "type": "RecipientScope",
    "displayName": "Users from Marketing Department",
    "customAttributes": {
      "Exclusive": false,
      "RecipientFilter": "Department -eq 'Marketing'"
    }
  };
  const unifiedRoleAssignmentScopeTenant: UnifiedRoleAssignment = {
    id: 'YUb1sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHVi2I-1',
    principalId: '/ServicePrincipals/fc33aa61-cf0e-46b6-9506-f633347202ab',
    roleDefinitionId: 'abc3aa61-cf0e-46b6-9506-f63334729876',
    directoryScopeId: '/'
  };
  const unifiedRoleAssignmentScopeUser: UnifiedRoleAssignment = {
    id: 'YUb1sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHVi2I-2',
    principalId: '/ServicePrincipals/fc33aa61-cf0e-46b6-9506-f633347202ab',
    roleDefinitionId: 'abc3aa61-cf0e-46b6-9506-f63334729876',
    directoryScopeId: '/users/7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091'
  };
  const unifiedRoleAssignmentScopeGroup: UnifiedRoleAssignment = {
    id: 'YUb1sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHVi2I-3',
    principalId: '/ServicePrincipals/fc33aa61-cf0e-46b6-9506-f633347202ab',
    roleDefinitionId: 'abc3aa61-cf0e-46b6-9506-f63334729876',
    directoryScopeId: '/groups/1a70e568-d286-4ad1-b036-734ff8667915'
  };
  const unifiedRoleAssignmentScopeAdministrativeUnit: UnifiedRoleAssignment = {
    id: 'YUb1sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHVi2I-4',
    principalId: '/ServicePrincipals/fc33aa61-cf0e-46b6-9506-f633347202ab',
    roleDefinitionId: 'abc3aa61-cf0e-46b6-9506-f63334729876',
    directoryScopeId: '/administrativeUnits/31537af4-6d77-4bb9-a681-d2394888ea26'
  };
  const unifiedRoleAssignmentScopeCustom: UnifiedRoleAssignment = {
    id: 'YUb1sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHVi2I-5',
    principalId: '/ServicePrincipals/fc33aa61-cf0e-46b6-9506-f633347202ab',
    roleDefinitionId: 'abc3aa61-cf0e-46b6-9506-f63334729876',
    appScopeId: '0d2ba203-f407-4de0-9f6a-7e411484e4da'
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APPROLEASSIGNMENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if roleDefinitionId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: 'foo',
      principalId: principalId,
      scope: 'tenant'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if principalId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: 'foo',
      scope: 'tenant'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userId: 'foo',
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but userName is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userName: 'foo',
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to group, but groupId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupId: 'foo',
      scope: 'group'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to administrativeUnit, but administrativeUnitId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitId: 'foo',
      scope: 'administrativeUnit'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to custom, but customAppScopeId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeId: 'foo',
      scope: 'custom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId is specified, but scope user is missing', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userId: userId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName is specified, but scope user is missing', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userName: userName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but neither userId nor userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but userId and userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userId: userId,
      userName: userName,
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but userId and groupId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userId: userId,
      groupId: groupId,
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but userId and groupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userId: userId,
      groupName: groupName,
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but userId and administrativeUnitId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userId: userId,
      administrativeUnitId: administrativeUnitId,
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but userId and administrativeUnitName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userId: userId,
      administrativeUnitName: administrativeUnitName,
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but userId and customAppScopeId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userId: userId,
      customAppScopeId: customAppScopeId,
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but userId and customAppScopeName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userId: userId,
      customAppScopeName: customAppScopeName,
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but userName and groupId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userName: userName,
      groupId: groupId,
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but userName and groupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userName: userName,
      groupName: groupName,
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but userName and administrativeUnitId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userName: userName,
      administrativeUnitId: administrativeUnitId,
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but userName and administrativeUnitName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userName: userName,
      administrativeUnitName: administrativeUnitName,
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but userName and customAppScopeId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userName: userName,
      customAppScopeId: customAppScopeId,
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to user, but userName and customAppScopeName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userName: userName,
      customAppScopeName: customAppScopeName,
      scope: 'user'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to group, but neither groupId nor groupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scope: 'group'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to group, but groupId and groupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupId: groupId,
      groupName: groupName,
      scope: 'group'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupId is specified, but scope group is missing', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupId: groupId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupName is specified, but scope group is missing', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupName: groupName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to group, but groupId and userId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupId: groupId,
      userId: userId,
      scope: 'group'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to group, but groupId and userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupId: groupId,
      userName: userName,
      scope: 'group'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to group, but groupId and administrativeUnitId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupId: groupId,
      administrativeUnitId: administrativeUnitId,
      scope: 'group'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to group, but groupId and administrativeUnitName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupId: groupId,
      administrativeUnitName: administrativeUnitName,
      scope: 'group'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to group, but groupId and customAppScopeId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupId: groupId,
      customAppScopeId: customAppScopeId,
      scope: 'group'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to group, but groupId and customAppScopeName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupId: groupId,
      customAppScopeName: customAppScopeName,
      scope: 'group'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to group, but groupName and userId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupName: groupName,
      userId: userId,
      scope: 'group'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to group, but groupName and userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupName: groupName,
      userName: userName,
      scope: 'group'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to group, but groupName and administrativeUnitId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupName: groupName,
      administrativeUnitId: administrativeUnitId,
      scope: 'group'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to group, but groupName and administrativeUnitName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupName: groupName,
      administrativeUnitName: administrativeUnitName,
      scope: 'group'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to group, but groupName and customAppScopeId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupName: groupName,
      customAppScopeId: customAppScopeId,
      scope: 'group'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to group, but groupName and customAppScopeName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupName: groupName,
      customAppScopeName: customAppScopeName,
      scope: 'group'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to administrativeUnit, but neither administrativeUnitId nor administrativeUnitName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scope: 'administrativeUnit'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to administrativeUnit, but administrativeUnitId and administrativeUnitName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitId: administrativeUnitId,
      administrativeUnitName: administrativeUnitName,
      scope: 'administrativeUnit'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to administrativeUnit, but administrativeUnitId and userId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitId: administrativeUnitId,
      userId: userId,
      scope: 'administrativeUnit'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to administrativeUnit, but administrativeUnitId and userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitId: administrativeUnitId,
      userName: userName,
      scope: 'administrativeUnit'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to administrativeUnit, but administrativeUnitId and groupId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitId: administrativeUnitId,
      groupId: groupId,
      scope: 'administrativeUnit'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to administrativeUnit, but administrativeUnitId and groupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitId: administrativeUnitId,
      groupName: groupName,
      scope: 'administrativeUnit'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to administrativeUnit, but administrativeUnitId and customAppScopeId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitId: administrativeUnitId,
      customAppScopeId: customAppScopeId,
      scope: 'administrativeUnit'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to administrativeUnit, but administrativeUnitId and customAppScopeName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitId: administrativeUnitId,
      customAppScopeName: customAppScopeName,
      scope: 'administrativeUnit'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to administrativeUnit, but administrativeUnitName and userId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitName: administrativeUnitName,
      userId: userId,
      scope: 'administrativeUnit'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to administrativeUnit, but administrativeUnitName and userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitName: administrativeUnitName,
      userName: userName,
      scope: 'administrativeUnit'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to administrativeUnit, but administrativeUnitName and groupId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitName: administrativeUnitName,
      groupId: groupId,
      scope: 'administrativeUnit'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to administrativeUnit, but administrativeUnitName and groupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitName: administrativeUnitName,
      groupName: groupName,
      scope: 'administrativeUnit'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to administrativeUnit, but administrativeUnitName and customAppScopeId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitName: administrativeUnitName,
      customAppScopeId: customAppScopeId,
      scope: 'administrativeUnit'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to administrativeUnit, but administrativeUnitName and customAppScopeName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitName: administrativeUnitName,
      customAppScopeName: customAppScopeName,
      scope: 'administrativeUnit'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to custom, but neither customAppScopeId nor customAppScopeName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scope: 'custom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to custom, but customAppScopeId and userId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeId: customAppScopeId,
      userId: userId,
      scope: 'custom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to custom, but customAppScopeId and userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeId: customAppScopeId,
      userName: userName,
      scope: 'custom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to custom, but customAppScopeId and groupId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeId: customAppScopeId,
      groupId: groupId,
      scope: 'custom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to custom, but customAppScopeId and groupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeId: customAppScopeId,
      groupName: groupName,
      scope: 'custom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to custom, but customAppScopeId and administrativeUnitId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeId: customAppScopeId,
      administrativeUnitId: administrativeUnitId,
      scope: 'custom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to custom, but customAppScopeId and administrativeUnitName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeId: customAppScopeId,
      administrativeUnitName: administrativeUnitName,
      scope: 'custom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to custom, but customAppScopeName and userId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeName: customAppScopeName,
      userId: userId,
      scope: 'custom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to custom, but customAppScopeName and userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeName: customAppScopeName,
      userName: userName,
      scope: 'custom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to custom, but customAppScopeName and groupId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeName: customAppScopeName,
      groupId: groupId,
      scope: 'custom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to custom, but customAppScopeName and groupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeName: customAppScopeName,
      groupName: groupName,
      scope: 'custom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to custom, but customAppScopeName and administrativeUnitId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeName: customAppScopeName,
      administrativeUnitId: administrativeUnitId,
      scope: 'custom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scope is set to custom, but customAppScopeName and administrativeUnitName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeName: customAppScopeName,
      administrativeUnitName: administrativeUnitName,
      scope: 'custom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitId is specified, but scope administrativeUnit is missing', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitId: administrativeUnitId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitName is specified, but scope administrativeUnit is missing', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitName: administrativeUnitName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitId and scope tenant is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitId: administrativeUnitId,
      scope: 'tenant'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitName and scope tenant is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitName: administrativeUnitName,
      scope: 'tenant'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if customAppScopeId and customAppScopeName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeId: customAppScopeId,
      customAppScopeName: customAppScopeName,
      scope: 'custom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if customAppScopeId is specified, but scope custom is missing', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeId: customAppScopeId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if customAppScopeName is specified, but scope customAppScopeName is missing', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeName: customAppScopeName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if roleDefinitionId and roleDefinitionName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      roleDefinitionName: roleDefinitionName,
      principalId: principalId,
      scope: 'tenant'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither roleDefinitionId nor roleDefinitionName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      principalId: principalId,
      scope: 'tenant'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if principalId and principalName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      principalName: principalName,
      scope: 'tenant'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither principalId nor principalName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      scope: 'tenant'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if no scope is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      principalId: principalId,
      roleDefinitionId: roleDefinitionId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/",
          "appScopeId": null
        })) {
        return unifiedRoleAssignmentScopeTenant;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scope: 'tenant',
      verbose: true
    });
    await command.action(logger, {
      options: result.data!
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeTenant));
  });

  it('correctly creates the role assignment for the role specified by name and the service principal specified by id', async () => {
    sinon.stub(roleDefinition, 'getExchangeRoleDefinitionByDisplayName').withArgs(roleDefinitionName).resolves(unifiedRoleDefinition);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/",
          "appScopeId": null
        })) {
        return unifiedRoleAssignmentScopeTenant;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionName: roleDefinitionName,
      principalId: principalId,
      scope: 'tenant',
      verbose: true
    });
    await command.action(logger, {
      options: result.data!
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeTenant));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by name', async () => {
    sinon.stub(entraServicePrincipal, 'getServicePrincipalByAppName').withArgs(principalName, 'id').resolves({ id: principalId });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/",
          "appScopeId": null
        })) {
        return unifiedRoleAssignmentScopeTenant;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalName: principalName,
      scope: 'tenant',
      verbose: true
    });
    await command.action(logger, {
      options: result.data!
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeTenant));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by id with the scope restricted to the user specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/users/7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091",
          "appScopeId": null
        })) {
        return unifiedRoleAssignmentScopeUser;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userId: userId,
      scope: 'user',
      verbose: true
    });
    await command.action(logger, {
      options: result.data!
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeUser));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by id with the scope restricted to the user specified by name', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/users/7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091",
          "appScopeId": null
        })) {
        return unifiedRoleAssignmentScopeUser;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      userName: userName,
      scope: 'user',
      verbose: true
    });
    await command.action(logger, {
      options: result.data!
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeUser));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by id with the scope restricted to the group specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/groups/1a70e568-d286-4ad1-b036-734ff8667915",
          "appScopeId": null
        })) {
        return unifiedRoleAssignmentScopeGroup;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupId: groupId,
      scope: 'group',
      verbose: true
    });
    await command.action(logger, {
      options: result.data!
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeGroup));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by id with the scope restricted to the group specified by name', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs(groupName).resolves(groupId);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/groups/1a70e568-d286-4ad1-b036-734ff8667915",
          "appScopeId": null
        })) {
        return unifiedRoleAssignmentScopeGroup;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      groupName: groupName,
      scope: 'group',
      verbose: true
    });
    await command.action(logger, {
      options: result.data!
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeGroup));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by id with the scope restricted to the administrative unit specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/administrativeUnits/31537af4-6d77-4bb9-a681-d2394888ea26",
          "appScopeId": null
        })) {
        return unifiedRoleAssignmentScopeAdministrativeUnit;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitId: administrativeUnitId,
      scope: 'administrativeUnit',
      verbose: true
    });
    await command.action(logger, {
      options: result.data!
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeAdministrativeUnit));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by id with the scope restricted to the administrative unit specified by name', async () => {
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves(administrativeUnit);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/administrativeUnits/31537af4-6d77-4bb9-a681-d2394888ea26",
          "appScopeId": null
        })) {
        return unifiedRoleAssignmentScopeAdministrativeUnit;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      administrativeUnitName: administrativeUnitName,
      scope: 'administrativeUnit',
      verbose: true
    });
    await command.action(logger, {
      options: result.data!
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeAdministrativeUnit));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by id with the scope restricted to the custom application scope specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": null,
          "appScopeId": `${customAppScopeId}`
        })) {
        return unifiedRoleAssignmentScopeCustom;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeId: customAppScopeId,
      scope: 'custom',
      verbose: true
    });
    await command.action(logger, {
      options: result.data!
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeCustom));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by id with the scope restricted to the custom application scope specified by name', async () => {
    sinon.stub(customAppScope, 'getCustomAppScopeByDisplayName').withArgs(customAppScopeName).resolves(customApplicationScope);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": null,
          "appScopeId": `${customAppScopeId}`
        })) {
        return unifiedRoleAssignmentScopeCustom;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      customAppScopeName: customAppScopeName,
      scope: 'custom',
      verbose: true
    });
    await command.action(logger, {
      options: result.data!
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeCustom));
  });

  it('correctly handles API OData error', async () => {
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
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scope: 'tenant'
    });
    await assert.rejects(command.action(logger, {
      options: result.data!
    }), new CommandError('Invalid request'));
  });
});
