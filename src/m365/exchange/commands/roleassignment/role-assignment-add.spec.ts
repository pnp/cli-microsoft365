import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { cli } from '../../../../cli/cli.js';
import command from './role-assignment-add.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import request from '../../../../request.js';
import { CommandError } from '../../../../Command.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { AdministrativeUnit, UnifiedRoleAssignment, UnifiedRoleDefinition } from '@microsoft/microsoft-graph-types';
import { entraServicePrincipal } from '../../../../utils/entraServicePrincipal.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

describe(commands.ROLE_ASSIGNMENT_ADD, () => {
  const principalId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const principalName = 'ContosoApp';
  const roleDefinitionId = 'abc3aa61-cf0e-46b6-9506-f63334729876';
  const roleDefinitionName = 'Application Mail.Read';
  const scopeUserId = '7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091';
  const scopeUserName = 'john.doe@contoso.com';
  const scopeGroupId = '1a70e568-d286-4ad1-b036-734ff8667915';
  const scopeGroupName = 'Contoso Group';
  const scopeAdministrativeUnitId = '31537af4-6d77-4bb9-a681-d2394888ea26';
  const scopeAdministrativeUnitName = 'Contoso Administrative Unit';
  const unifiedRoleDefinition: UnifiedRoleDefinition = {
    id: roleDefinitionId,
    displayName: roleDefinitionName
  };
  const administrativeUnit: AdministrativeUnit = {
    id: scopeAdministrativeUnitId,
    displayName: scopeAdministrativeUnitName
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

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    assert.strictEqual(command.name, commands.ROLE_ASSIGNMENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if roleDefinitionId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: 'foo',
      principalId: principalId,
      scopeTenant: true
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if principalId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: 'foo',
      scopeTenant: true
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scopeUserId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scopeUserId: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scopeUserName is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scopeUserName: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scopeGroupId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scopeGroupId: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scopeAdministrativeUnitId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scopeAdministrativeUnitId: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scopeUserId and scopeUserName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scopeUserId: scopeUserId,
      scopeUserName: scopeUserName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scopeUserName and scopeGroupId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scopeUserName: scopeUserName,
      scopeGroupId: scopeGroupId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scopeGroupId and scopeGroupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scopeGroupId: scopeGroupId,
      scopeGroupName: scopeGroupName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scopeGroupName and scopeAdministrativeUnitId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scopeGroupName: scopeGroupName,
      scopeAdministrativeUnitId: scopeAdministrativeUnitId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scopeAdministrativeUnitId and scopeAdministrativeUnitName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scopeAdministrativeUnitId: scopeAdministrativeUnitId,
      scopeAdministrativeUnitName: scopeAdministrativeUnitName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if scopeAdministrativeUnitName and scopeTenant is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      scopeAdministrativeUnitName: scopeAdministrativeUnitName,
      scopeTenant: true
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if roleDefinitionId and roleDefinitionName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      roleDefinitionName: roleDefinitionName,
      principalId: principalId,
      scopeTenant: true
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither roleDefinitionId nor roleDefinitionName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      principalId: principalId,
      scopeTenant: true
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if principalId and principalName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principalId: principalId,
      principalName: principalName,
      scopeTenant: true
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither principalId nor principalName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      scopeTenant: true
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
          "directoryScopeId": "/"
        })) {
        return unifiedRoleAssignmentScopeTenant;
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options:
      {
        roleDefinitionId: roleDefinitionId,
        principalId: principalId,
        scopeTenant: true,
        verbose: true
      }
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
          "directoryScopeId": "/"
        })) {
        return unifiedRoleAssignmentScopeTenant;
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options:
      {
        roleDefinitionName: roleDefinitionName,
        principalId: principalId,
        scopeTenant: true,
        verbose: true
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeTenant));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by name', async () => {
    sinon.stub(entraServicePrincipal, 'getServicePrincipalIdFromAppName').withArgs(principalName).resolves(principalId);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/"
        })) {
        return unifiedRoleAssignmentScopeTenant;
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options:
      {
        roleDefinitionId: roleDefinitionId,
        principalName: principalName,
        scopeTenant: true,
        verbose: true
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeTenant));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by id with the scope restricted to the user specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/users/7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091"
        })) {
        return unifiedRoleAssignmentScopeUser;
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options:
      {
        roleDefinitionId: roleDefinitionId,
        principalId: principalId,
        scopeUserId: scopeUserId,
        verbose: true
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeUser));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by id with the scope restricted to the user specified by name', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(scopeUserName).resolves(scopeUserId);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/users/7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091"
        })) {
        return unifiedRoleAssignmentScopeUser;
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options:
      {
        roleDefinitionId: roleDefinitionId,
        principalId: principalId,
        scopeUserName: scopeUserName,
        verbose: true
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeUser));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by id with the scope restricted to the group specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/groups/1a70e568-d286-4ad1-b036-734ff8667915"
        })) {
        return unifiedRoleAssignmentScopeGroup;
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options:
      {
        roleDefinitionId: roleDefinitionId,
        principalId: principalId,
        scopeGroupId: scopeGroupId,
        verbose: true
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeGroup));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by id with the scope restricted to the group specified by name', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs(scopeGroupName).resolves(scopeGroupId);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/groups/1a70e568-d286-4ad1-b036-734ff8667915"
        })) {
        return unifiedRoleAssignmentScopeGroup;
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options:
      {
        roleDefinitionId: roleDefinitionId,
        principalId: principalId,
        scopeGroupName: scopeGroupName,
        verbose: true
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeGroup));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by id with the scope restricted to the administrative unit specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/administrativeUnits/31537af4-6d77-4bb9-a681-d2394888ea26"
        })) {
        return unifiedRoleAssignmentScopeAdministrativeUnit;
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options:
      {
        roleDefinitionId: roleDefinitionId,
        principalId: principalId,
        scopeAdministrativeUnitId: scopeAdministrativeUnitId,
        verbose: true
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeAdministrativeUnit));
  });

  it('correctly creates the role assignment for the role specified by id and the service principal specified by id with the scope restricted to the administrative unit specified by name', async () => {
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(scopeAdministrativeUnitName).resolves(administrativeUnit);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": `/ServicePrincipals/${principalId}`,
          "directoryScopeId": "/administrativeUnits/31537af4-6d77-4bb9-a681-d2394888ea26"
        })) {
        return unifiedRoleAssignmentScopeAdministrativeUnit;
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options:
      {
        roleDefinitionId: roleDefinitionId,
        principalId: principalId,
        scopeAdministrativeUnitName: scopeAdministrativeUnitName,
        verbose: true
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeAdministrativeUnit));
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

    await assert.rejects(command.action(logger, {
      options: {
        roleDefinitionId: roleDefinitionId,
        principalId: principalId,
        scopeTenant: true
      } }), new CommandError('Invalid request'));
  });
});
