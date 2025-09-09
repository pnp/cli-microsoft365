import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { cli } from '../../../../cli/cli.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import commands from '../../commands.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { entraServicePrincipal } from '../../../../utils/entraServicePrincipal.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import command from './roleassignment-add.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { AdministrativeUnit, Application, ServicePrincipal, UnifiedRoleAssignment, UnifiedRoleDefinition } from '@microsoft/microsoft-graph-types';

describe(commands.ROLEASSIGNMENT_ADD, () => {
  const principalId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const principalUpn = 'john.doe@contoso.com';
  const principalGroupMailNickname = 'contosogroup';
  const roleDefinitionId = 'abc3aa61-cf0e-46b6-9506-f63334729876';
  const roleDefinitionName = 'Application Mail.Read';
  const userId = '7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091';
  const userName = 'john.doe@contoso.com';
  const groupId = '1a70e568-d286-4ad1-b036-734ff8667915';
  const groupName = 'Contoso Group';
  const administrativeUnitId = '31537af4-6d77-4bb9-a681-d2394888ea26';
  const administrativeUnitName = 'Contoso Administrative Unit';
  const applicationId = '77084c32-6b19-4ae8-bd9f-96a28ce9665b';
  const applicationObjectId = '8a6b3f1a-6b26-4298-b0b9-8e129fd6d197';
  const applicationName = 'Contoso Application';
  const servicePrincipalId = 'd1f2e3f4-5a6b-7c8d-9e0f-1a2b3c4d5e6f';
  const servicePrincipalName = 'Contoso Service Principal';

  const unifiedRoleDefinition: UnifiedRoleDefinition = {
    id: roleDefinitionId,
    displayName: roleDefinitionName
  };
  const administrativeUnit: AdministrativeUnit = {
    id: administrativeUnitId,
    displayName: administrativeUnitName
  };
  const application: Application = {
    id: applicationObjectId,
    appId: applicationId,
    displayName: applicationName
  };
  const servicePrincipal: ServicePrincipal = {
    id: servicePrincipalId,
    displayName: servicePrincipalName
  };

  const unifiedRoleAssignmentScopeTenant: UnifiedRoleAssignment = {
    id: 'YUb1sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHVi2I-1',
    principalId: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
    roleDefinitionId: 'abc3aa61-cf0e-46b6-9506-f63334729876',
    directoryScopeId: '/'
  };
  const unifiedRoleAssignmentScopeUser: UnifiedRoleAssignment = {
    id: 'YUb1sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHVi2I-2',
    principalId: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
    roleDefinitionId: 'abc3aa61-cf0e-46b6-9506-f63334729876',
    directoryScopeId: '/7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091'
  };
  const unifiedRoleAssignmentScopeAdministrativeUnit: UnifiedRoleAssignment = {
    id: 'YUb1sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHVi2I-4',
    principalId: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
    roleDefinitionId: 'abc3aa61-cf0e-46b6-9506-f63334729876',
    directoryScopeId: '/administrativeUnits/31537af4-6d77-4bb9-a681-d2394888ea26'
  };
  const unifiedRoleAssignmentScopeApplication: UnifiedRoleAssignment = {
    id: 'YUb1sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHVi2I-4',
    principalId: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
    roleDefinitionId: 'abc3aa61-cf0e-46b6-9506-f63334729876',
    directoryScopeId: '/8a6b3f1a-6b26-4298-b0b9-8e129fd6d197'
  };
  const unifiedRoleAssignmentScopeServicePrincipal: UnifiedRoleAssignment = {
    id: 'YUb1sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHVi2I-4',
    principalId: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
    roleDefinitionId: 'abc3aa61-cf0e-46b6-9506-f63334729876',
    directoryScopeId: '/d1f2e3f4-5a6b-7c8d-9e0f-1a2b3c4d5e6f'
  };
  const unifiedRoleAssignmentScopeGroup: UnifiedRoleAssignment = {
    id: 'YUb1sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHVi2I-4',
    principalId: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
    roleDefinitionId: 'abc3aa61-cf0e-46b6-9506-f63334729876',
    directoryScopeId: '/1a70e568-d286-4ad1-b036-734ff8667915'
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
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
      request.post,
      roleDefinition.getRoleDefinitionByDisplayName,
      entraUser.getUserIdByUpn,
      entraGroup.getGroupIdByDisplayName,
      entraGroup.getGroupIdByMailNickname,
      entraAdministrativeUnit.getAdministrativeUnitByDisplayName,
      entraServicePrincipal.getServicePrincipalByAppName,
      entraApp.getAppRegistrationByAppName,
      entraApp.getAppRegistrationByAppId
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROLEASSIGNMENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if roleDefinitionId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: 'foo',
      principal: principalId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if principal is neither a valid GUID, UPN or mail nickname', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: '@foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if principal is not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId is specified, but it is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userId: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName is specified, but it is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userName: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupId is specified, but it is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupId: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitId is specified, but it is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      administrativeUnitId: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if applicationId is specified, but it is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      applicationId: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if applicationObjectId is specified, but it is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      applicationObjectId: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if servicePrincipalId is specified, but it is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      servicePrincipalId: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId and userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userId: userId,
      userName: userName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId and groupId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userId: userId,
      groupId: groupId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId and groupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userId: userId,
      groupName: groupName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId and administrativeUnitId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userId: userId,
      administrativeUnitId: administrativeUnitId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId and administrativeUnitName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userId: userId,
      administrativeUnitName: administrativeUnitName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId and applicationId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userId: userId,
      applicationId: applicationId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId and applicationObjectId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userId: userId,
      applicationObjectId: applicationObjectId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId and applicationName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userId: userId,
      applicationName: applicationName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId and servicePrincipalId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userId: userId,
      servicePrincipalId: servicePrincipalId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId and servicePrincipalName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userId: userId,
      servicePrincipalName: servicePrincipalName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName and groupId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userName: userName,
      groupId: groupId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName and groupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userName: userName,
      groupName: groupName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName and administrativeUnitId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userName: userName,
      administrativeUnitId: administrativeUnitId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName and administrativeUnitName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userName: userName,
      administrativeUnitName: administrativeUnitName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName and applicationId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userName: userName,
      applicationId: applicationId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName and applicationObjectId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userName: userName,
      applicationObjectId: applicationObjectId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName and applicationName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userName: userName,
      applicationName: applicationName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName and servicePrincipalId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userName: userName,
      servicePrincipalId: servicePrincipalId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName and servicePrincipalName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userName: userName,
      servicePrincipalName: servicePrincipalName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupId and groupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupId: groupId,
      groupName: groupName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupId and administrativeUnitId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupId: groupId,
      administrativeUnitId: administrativeUnitId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupId and administrativeUnitName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupId: groupId,
      administrativeUnitName: administrativeUnitName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupId and applicationId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupId: groupId,
      applicationId: applicationId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupId and applicationObjectId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupId: groupId,
      applicationObjectId: applicationObjectId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupId and applicationName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupId: groupId,
      applicationName: applicationName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupId and servicePrincipalId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupId: groupId,
      servicePrincipalId: servicePrincipalId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupId and servicePrincipalName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupId: groupId,
      servicePrincipalName: servicePrincipalName
    });
    assert.notStrictEqual(actual.success, true);
  });


  it('fails validation if groupName and administrativeUnitId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupName: groupName,
      administrativeUnitId: administrativeUnitId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupName and administrativeUnitName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupName: groupName,
      administrativeUnitName: administrativeUnitName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupName and applicationId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupName: groupName,
      applicationId: applicationId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupName and applicationObjectId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupName: groupName,
      applicationObjectId: applicationObjectId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupName and applicationName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupName: groupName,
      applicationName: applicationName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupName and servicePrincipalId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupName: groupName,
      servicePrincipalId: servicePrincipalId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if groupName and servicePrincipalName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupName: groupName,
      servicePrincipalName: servicePrincipalName
    });
    assert.notStrictEqual(actual.success, true);
  });


  it('fails validation if administrativeUnitId and administrativeUnitName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      administrativeUnitId: administrativeUnitId,
      administrativeUnitName: administrativeUnitName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitId and applicationId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      administrativeUnitId: administrativeUnitId,
      applicationId: applicationId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitId and applicationObjectId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      administrativeUnitId: administrativeUnitId,
      applicationObjectId: applicationObjectId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitId and applicationName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      administrativeUnitId: administrativeUnitId,
      applicationName: applicationName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitId and servicePrincipalId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      administrativeUnitId: administrativeUnitId,
      servicePrincipalId: servicePrincipalId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitId and servicePrincipalName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      administrativeUnitId: administrativeUnitId,
      servicePrincipalName: servicePrincipalName
    });
    assert.notStrictEqual(actual.success, true);
  });


  it('fails validation if administrativeUnitName and applicationId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      administrativeUnitName: administrativeUnitName,
      applicationId: applicationId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitName and applicationObjectId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      administrativeUnitName: administrativeUnitName,
      applicationObjectId: applicationObjectId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitName and applicationName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      administrativeUnitName: administrativeUnitName,
      applicationName: applicationName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitName and servicePrincipalId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      administrativeUnitName: administrativeUnitName,
      servicePrincipalId: servicePrincipalId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if administrativeUnitName and servicePrincipalName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      administrativeUnitName: administrativeUnitName,
      servicePrincipalName: servicePrincipalName
    });
    assert.notStrictEqual(actual.success, true);
  });


  it('fails validation if applicationId and applicationObjectId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      applicationId: applicationId,
      applicationObjectId: applicationObjectId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if applicationId and applicationName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      applicationId: applicationId,
      applicationName: applicationName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if applicationId and servicePrincipalId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      applicationId: applicationId,
      servicePrincipalId: servicePrincipalId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if applicationId and servicePrincipalName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      applicationId: applicationId,
      servicePrincipalName: servicePrincipalName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if applicationObjectId and applicationName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      applicationObjectId: applicationObjectId,
      applicationName: applicationName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if applicationObjectId and servicePrincipalId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      applicationObjectId: applicationObjectId,
      servicePrincipalId: servicePrincipalId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if applicationObjectId and servicePrincipalName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      applicationObjectId: applicationObjectId,
      servicePrincipalName: servicePrincipalName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if applicationName and servicePrincipalId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      applicationName: applicationName,
      servicePrincipalId: servicePrincipalId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if applicationName and servicePrincipalName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      applicationName: applicationName,
      servicePrincipalName: servicePrincipalName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if servicePrincipalId and servicePrincipalName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      servicePrincipalId: servicePrincipalId,
      servicePrincipalName: servicePrincipalName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if roleDefinitionId and roleDefinitionName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      roleDefinitionName: roleDefinitionName,
      principal: principalId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither roleDefinitionId nor roleDefinitionName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      principal: principalId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('correctly creates the role assignment for the role specified by id and the principal specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          roleDefinitionId: roleDefinitionId,
          principalId: principalId,
          directoryScopeId: "/"
        })) {
        return unifiedRoleAssignmentScopeTenant;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      verbose: true
    });
    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeTenant));
  });

  it('correctly creates the role assignment for the role specified by name and the principal specified by id', async () => {
    sinon.stub(roleDefinition, 'getRoleDefinitionByDisplayName').withArgs(roleDefinitionName).resolves(unifiedRoleDefinition);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          roleDefinitionId: roleDefinitionId,
          principalId: principalId,
          directoryScopeId: "/"
        })) {
        return unifiedRoleAssignmentScopeTenant;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionName: roleDefinitionName,
      principal: principalId,
      verbose: true
    });
    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeTenant));
  });

  it('correctly creates the role assignment for the role specified by id and the principal specified by id with the scope restricted to the user specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          roleDefinitionId: roleDefinitionId,
          principalId: principalId,
          directoryScopeId: `/${userId}`
        })) {
        return unifiedRoleAssignmentScopeUser;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userId: userId,
      verbose: true
    });
    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeUser));
  });

  it('correctly creates the role assignment for the role specified by id and the principal specified by id with the scope restricted to the user specified by UPN', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          roleDefinitionId: roleDefinitionId,
          principalId: principalId,
          directoryScopeId: `/${userId}`
        })) {
        return unifiedRoleAssignmentScopeUser;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      userName: userName,
      verbose: true
    });
    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeUser));
  });

  it('correctly creates the role assignment for the role specified by id and the principal specified by id with the scope restricted to the administrative unit specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          roleDefinitionId: roleDefinitionId,
          principalId: principalId,
          directoryScopeId: `/administrativeUnits/${administrativeUnitId}`
        })) {
        return unifiedRoleAssignmentScopeAdministrativeUnit;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      administrativeUnitId: administrativeUnitId,
      verbose: true
    });
    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeAdministrativeUnit));
  });

  it('correctly creates the role assignment for the role specified by id and the principal specified by id with the scope restricted to the administrative unit specified by name', async () => {
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves(administrativeUnit);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          roleDefinitionId: roleDefinitionId,
          principalId: principalId,
          directoryScopeId: `/administrativeUnits/${administrativeUnitId}`
        })) {
        return unifiedRoleAssignmentScopeAdministrativeUnit;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      administrativeUnitName: administrativeUnitName,
      verbose: true
    });
    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeAdministrativeUnit));
  });

  it('correctly creates the role assignment for the role specified by id and the principal specified by id with the scope restricted to the application specified by object id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          roleDefinitionId: roleDefinitionId,
          principalId: principalId,
          directoryScopeId: `/${applicationObjectId}`
        })) {
        return unifiedRoleAssignmentScopeApplication;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      applicationObjectId: applicationObjectId,
      verbose: true
    });
    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeApplication));
  });

  it('correctly creates the role assignment for the role specified by id and the principal specified by id with the scope restricted to the application specified by name', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppName').withArgs(applicationName).resolves(application);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          roleDefinitionId: roleDefinitionId,
          principalId: principalId,
          directoryScopeId: `/${applicationObjectId}`
        })) {
        return unifiedRoleAssignmentScopeApplication;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      applicationName: applicationName,
      verbose: true
    });
    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeApplication));
  });

  it('correctly creates the role assignment for the role specified by id and the principal specified by id with the scope restricted to the application specified by app id', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').withArgs(applicationId).resolves(application);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          roleDefinitionId: roleDefinitionId,
          principalId: principalId,
          directoryScopeId: `/${applicationObjectId}`
        })) {
        return unifiedRoleAssignmentScopeApplication;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      applicationId: applicationId,
      verbose: true
    });
    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeApplication));
  });

  it('correctly creates the role assignment for the role specified by id and the principal specified by id with the scope restricted to the service principal specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          roleDefinitionId: roleDefinitionId,
          principalId: principalId,
          directoryScopeId: `/${servicePrincipalId}`
        })) {
        return unifiedRoleAssignmentScopeServicePrincipal;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      servicePrincipalId: servicePrincipalId,
      verbose: true
    });
    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeServicePrincipal));
  });

  it('correctly creates the role assignment for the role specified by id and the principal specified by id with the scope restricted to the service principal specified by name', async () => {
    sinon.stub(entraServicePrincipal, 'getServicePrincipalByAppName').withArgs(servicePrincipalName).resolves(servicePrincipal);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": principalId,
          "directoryScopeId": `/${servicePrincipalId}`
        })) {
        return unifiedRoleAssignmentScopeServicePrincipal;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      servicePrincipalName: servicePrincipalName,
      verbose: true
    });
    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeServicePrincipal));
  });

  it('correctly creates the role assignment for the role specified by id and the principal specified by id with the scope restricted to the group specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          roleDefinitionId: roleDefinitionId,
          principalId: principalId,
          directoryScopeId: `/${groupId}`
        })) {
        return unifiedRoleAssignmentScopeGroup;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupId: groupId,
      verbose: true
    });
    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeGroup));
  });

  it('correctly creates the role assignment for the role specified by id and the principal specified by id with the scope restricted to the group specified by name', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs(groupName).resolves(groupId);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          roleDefinitionId: roleDefinitionId,
          principalId: principalId,
          directoryScopeId: `/${groupId}`
        })) {
        return unifiedRoleAssignmentScopeGroup;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalId,
      groupName: groupName,
      verbose: true
    });
    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeGroup));
  });

  it('correctly creates the role assignment for the role specified by id and the principal specified by user principal name', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(principalUpn).resolves(principalId);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          roleDefinitionId: roleDefinitionId,
          principalId: principalId,
          directoryScopeId: "/"
        })) {
        return unifiedRoleAssignmentScopeTenant;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalUpn,
      verbose: true
    });
    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeTenant));
  });

  it('correctly creates the role assignment for the role specified by id and the principal specified by group mail nickname', async () => {
    sinon.stub(entraGroup, 'getGroupIdByMailNickname').withArgs(principalGroupMailNickname).resolves(principalId);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          roleDefinitionId: roleDefinitionId,
          principalId: principalId,
          directoryScopeId: "/"
        })) {
        return unifiedRoleAssignmentScopeTenant;
      }

      throw 'Invalid request';
    });
    const result = commandOptionsSchema.safeParse({
      roleDefinitionId: roleDefinitionId,
      principal: principalGroupMailNickname,
      verbose: true
    });
    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWithExactly(unifiedRoleAssignmentScopeTenant));
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
      principal: principalId
    });
    await assert.rejects(command.action(logger, {
      options: result.data
    }), new CommandError('Invalid request'));
  });
});