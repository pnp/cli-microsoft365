import { UnifiedRoleAssignment } from '@microsoft/microsoft-graph-types';
import request, { CliRequestOptions } from '../request.js';

const getRequestOptions = (roleDefinitionId: string, principalId: string, directoryScopeId: string): CliRequestOptions => ({
  url: `https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments`,
  headers: {
    accept: 'application/json;odata.metadata=none'
  },
  responseType: 'json',
  data: {
    roleDefinitionId: roleDefinitionId,
    principalId: principalId,
    directoryScopeId: directoryScopeId
  }
});

/**
 * Utils for RBAC.
 * Supported RBAC providers:
 *  - Directory (Entra ID)
 */
export const roleAssignment = {
  /**
   * Assigns a specific role to a principal with scope to an administrative unit
   * @param roleDefinitionId Role which lists the actions that can be performed
   * @param principalId Object that represents a user, group, service principal, or managed identity that is requesting access to resources
   * @param administrativeUnitId Administrative unit which represents a current scope for a role assignment
   * @returns Returns unified role assignment object that represents a role definition assigned to a principal with scope to an administrative unit
   */
  async createRoleAssignmentWithAdministrativeUnitScope(roleDefinitionId: string, principalId: string, administrativeUnitId: string): Promise<UnifiedRoleAssignment> {
    const requestOptions = getRequestOptions(roleDefinitionId, principalId, `/administrativeUnits/${administrativeUnitId}`);
    return await request.post<UnifiedRoleAssignment>(requestOptions);
  },

  /**
   * Assigns a specific role to a principal with scope to the whole tenant
   * @param roleDefinitionId Role which lists the actions that can be performed
   * @param principalId Object that represents a user, group, service principal, or managed identity that is requesting access to resources
   * @returns Returns unified role assignment object that represents a role definition assigned to a principal with scope to the whole tenant
   */
  async createRoleAssignmentWithTenantScope(roleDefinitionId: string, principalId: string): Promise<UnifiedRoleAssignment> {
    const requestOptions = getRequestOptions(roleDefinitionId, principalId, '/');
    return await request.post<UnifiedRoleAssignment>(requestOptions);
  }
};