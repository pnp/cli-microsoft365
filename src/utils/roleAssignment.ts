import { UnifiedRoleAssignment } from "@microsoft/microsoft-graph-types";
import request, { CliRequestOptions } from "../request.js";

const graphResource = 'https://graph.microsoft.com';

export const roleAssignment = {
  async createRoleAssignment(roleDefinitionId: string, principalId: string, directoryScopeId: string): Promise<UnifiedRoleAssignment> {
    const requestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/roleManagement/directory/roleAssignments`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        roleDefinitionId: roleDefinitionId,
        principalId: principalId,
        directoryScopeId: directoryScopeId
      }
    };

    return await request.post<UnifiedRoleAssignment>(requestOptions);
  },

  async createRoleAssignmentWithAdministrativeUnitScope(roleDefinitionId: string, principalId: string, administrativeUnitId: string): Promise<UnifiedRoleAssignment> {
    return await this.createRoleAssignment(roleDefinitionId, principalId, `/administrativeUnits/${administrativeUnitId}`);
  },

  async createRoleAssignmentWithTenantScope(roleDefinitionId: string, principalId: string): Promise<UnifiedRoleAssignment> {
    return await this.createRoleAssignment(roleDefinitionId, principalId, '/');
  }
};