import assert from 'assert';
import sinon from 'sinon';
import request from "../request.js";
import { sinonUtil } from "./sinonUtil.js";
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

  it('correctly assign a directory (Entra ID) role specified by id with administrative unit scope to a user specified by id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments') {
        return unifieRoleAssignmentWithAdministrativeUnitScopeResponse;
      }

      throw 'Invalid request';
    });

    const unifiedRoleAssignment = await roleAssignment.createEntraIDRoleAssignmentWithAdministrativeUnitScope(roleDefinitionId, userId, administrativeUnitId);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      roleDefinitionId: roleDefinitionId,
      principalId: userId,
      directoryScopeId: `/administrativeUnits/${administrativeUnitId}`
    });
    assert.deepStrictEqual(unifiedRoleAssignment, {
      "id": "BH21sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHIWb7-1",
      "roleDefinitionId": "fe930be7-5e62-47db-91af-98c3a49a38b1",
      "principalId": "2056d2f6-3257-4253-8cfc-b73393e414e5",
      "directoryScopeId": "/administrativeUnits/fc33aa61-cf0e-46b6-9506-f633347202ab"
    });
  });

  it('correctly assign a directory (Entra ID) role specified by id with tenant scope to a user specified by id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments') {
        return unifieRoleAssignmentWithTenantScopeResponse;
      }

      throw 'Invalid request';
    });

    const unifiedRoleAssignment = await roleAssignment.createEntraIDRoleAssignmentWithTenantScope(roleDefinitionId, userId);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      roleDefinitionId: roleDefinitionId,
      principalId: userId,
      directoryScopeId: '/'
    });
    assert.deepStrictEqual(unifiedRoleAssignment, {
      "id": "BH21sHQtUEyvox7IA_Eu_mm3jqnUe4lEhvatluHIWb7-1",
      "roleDefinitionId": "fe930be7-5e62-47db-91af-98c3a49a38b1",
      "principalId": "2056d2f6-3257-4253-8cfc-b73393e414e5",
      "directoryScopeId": "/"
    });
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'post').rejects(new Error(errorMessage));

    await assert.rejects(roleAssignment.createEntraIDRoleAssignment('','',''), new Error(errorMessage));
  });
});