import assert from 'assert';
import sinon from 'sinon';
import request from '../request.js';
import { sinonUtil } from './sinonUtil.js';
import { roleAssignment } from './roleAssignment.js';

describe('utils/roleAssignment', () => {
  const roleDefinitionId = 'fe930be7-5e62-47db-91af-98c3a49a38b1';
  const userId = '2056d2f6-3257-4253-8cfc-b73393e414e5';
  const administrativeUnitId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const unifieRoleAssignmentWithAdministrativeUnitScopeResponse = {
    "id": "BH21sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHIWb7-1",
    "roleDefinitionId": "fe930be7-5e62-47db-91af-98c3a49a38b1",
    "principalId": "2056d2f6-3257-4253-8cfc-b73393e414e5",
    "directoryScopeId": "/administrativeUnits/fc33aa61-cf0e-46b6-9506-f633347202ab"
  };
  const unifieRoleAssignmentWithTenantScopeResponse = {
    "id": "BH21sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHIWb7-1",
    "roleDefinitionId": "fe930be7-5e62-47db-91af-98c3a49a38b1",
    "principalId": "2056d2f6-3257-4253-8cfc-b73393e414e5",
    "directoryScopeId": "/"
  };

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  it('correctly assigns a directory (Entra ID) role specified by id with administrative unit scope to a user specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments' && 
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": userId,
          "directoryScopeId": `/administrativeUnits/${administrativeUnitId}`
        })) {
        return unifieRoleAssignmentWithAdministrativeUnitScopeResponse;
      }

      throw 'Invalid request';
    });

    const unifiedRoleAssignment = await roleAssignment.createRoleAssignmentWithAdministrativeUnitScope(roleDefinitionId, userId, administrativeUnitId);
    assert.deepStrictEqual(unifiedRoleAssignment, {
      "id": "BH21sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHIWb7-1",
      "roleDefinitionId": "fe930be7-5e62-47db-91af-98c3a49a38b1",
      "principalId": "2056d2f6-3257-4253-8cfc-b73393e414e5",
      "directoryScopeId": "/administrativeUnits/fc33aa61-cf0e-46b6-9506-f633347202ab"
    });
  });

  it('correctly assigns a directory (Entra ID) role specified by id with tenant scope to a user specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "roleDefinitionId": roleDefinitionId,
          "principalId": userId,
          "directoryScopeId": '/'
        })) {
        return unifieRoleAssignmentWithTenantScopeResponse;
      }

      throw 'Invalid request';
    });

    const unifiedRoleAssignment = await roleAssignment.createRoleAssignmentWithTenantScope(roleDefinitionId, userId);
    assert.deepStrictEqual(unifiedRoleAssignment, {
      "id": "BH21sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHIWb7-1",
      "roleDefinitionId": "fe930be7-5e62-47db-91af-98c3a49a38b1",
      "principalId": "2056d2f6-3257-4253-8cfc-b73393e414e5",
      "directoryScopeId": "/"
    });
  });

  it('correctly handles random API error when createRoleAssignmentWithAdministrativeUnitScope is called', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'post').rejects(new Error(errorMessage));

    await assert.rejects(roleAssignment.createRoleAssignmentWithAdministrativeUnitScope(roleDefinitionId, userId, administrativeUnitId), new Error(errorMessage));
  });

  it('correctly handles random API error when createRoleAssignmentWithTenantScope is called', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'post').rejects(new Error(errorMessage));

    await assert.rejects(roleAssignment.createRoleAssignmentWithTenantScope(roleDefinitionId, userId), new Error(errorMessage));
  });
});