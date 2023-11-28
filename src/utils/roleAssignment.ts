import { UnifiedRoleAssignment } from "@microsoft/microsoft-graph-types";
import request, { CliRequestOptions } from "../request.js";

const graphResource = 'https://graph.microsoft.com';

/**
 * Utils for RBAC.
 * Supported RBAC providers:
 *  - Directory (Entra ID)
 */
export const roleAssignment = {
  /**
   * Assigns a specific role to a principal and defines a set of resources (scope) that the access applies to.
   * Directory (Entra ID) RBAC provider
   * @param roleDefinitionId Role which lists the actions that can be performed
   * @param principalId Object that represents a user, group, service principal, or managed identity that is requesting access to resources
   * @param directoryScopeId Set of resources that the access applies to
   * @returns
   */
  async createEntraIDRoleAssignment(roleDefinitionId: string, principalId: string, directoryScopeId: string): Promise<UnifiedRoleAssignment> {
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

  async createEntraIDRoleAssignmentWithAdministrativeUnitScope(roleDefinitionId: string, principalId: string, administrativeUnitId: string): Promise<UnifiedRoleAssignment> {
    return await this.createEntraIDRoleAssignment(roleDefinitionId, principalId, `/administrativeUnits/${administrativeUnitId}`);
  },

  async createEntraIDRoleAssignmentWithTenantScope(roleDefinitionId: string, principalId: string): Promise<UnifiedRoleAssignment> {
    return await this.createEntraIDRoleAssignment(roleDefinitionId, principalId, '/');
  }
};