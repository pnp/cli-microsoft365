import { Group } from "@microsoft/microsoft-graph-types";
import request, { CliRequestOptions } from "../request";
import { formatting } from "./formatting";
import { odata } from "./odata";

const graphResource = 'https://graph.microsoft.com';

export const aadGroup = {
  /**
   * Retrieve a single group.
   * @param id Group ID.
   */
  getGroupById(id: string): Promise<Group> {
    const requestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/groups/${id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<Group>(requestOptions);
  },

  /**
   * Get a list of groups by display name.
   * @param displayName Group display name.
   * @param selectProperties Properties to select if specified.
   */
  getGroupsByDisplayName(displayName: string, selectProperties?: string[]): Promise<Group[]> {
    return odata.getAllItems<Group>(`${graphResource}/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'${(selectProperties && selectProperties.length > 0) ? `&$select=${selectProperties.join(',')}` : ''}`);
  },

  /**
   * Get a single group by its display name.
   * @param displayName Group display name.
   * @param selectProperties Properties to select if specified.
   * @throws Error when group was not found.
   * @throws Error when multiple groups with the same name were found.
   */
  async getGroupByDisplayName(displayName: string, selectProperties?: string[]): Promise<Group> {
    const groups = await this.getGroupsByDisplayName(displayName, selectProperties);

    if (!groups.length) {
      throw Error(`The specified group '${displayName}' does not exist.`);
    }

    if (groups.length > 1) {
      throw Error(`Multiple groups with name '${displayName}' found: ${groups.map(x => x.id).join(',')}.`);
    }

    return groups[0];
  }
};