import { Group } from "@microsoft/microsoft-graph-types";
import request, { CliRequestOptions } from "../request.js";
import { formatting } from "./formatting.js";
import { odata } from "./odata.js";
import { Logger } from '../cli/Logger.js';
import { Cli } from '../cli/Cli.js';

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
   */
  getGroupsByDisplayName(displayName: string): Promise<Group[]> {
    return odata.getAllItems<Group>(`${graphResource}/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`);
  },

  /**
   * Get a single group by its display name.
   * @param displayName Group display name.
   * @throws Error when group was not found.
   * @throws Error when multiple groups with the same name were found.
   */
  async getGroupByDisplayName(displayName: string): Promise<Group> {
    const groups = await this.getGroupsByDisplayName(displayName);

    if (!groups.length) {
      throw Error(`The specified group '${displayName}' does not exist.`);
    }

    if (groups.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', groups);
      groups[0] = await Cli.handleMultipleResultsFound<Group>(`Multiple groups with name '${displayName}' found.`, resultAsKeyValuePair);
    }

    return groups[0];
  },

  /**
   * Get id of a group by its display name.
   * @param displayName Group display name.
   * @throws Error when group was not found.
   * @throws Error when multiple groups with the same name were found.
   */
  async getGroupIdByDisplayName(displayName: string): Promise<string> {
    const groups = await odata.getAllItems<Group>(`${graphResource}/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`);

    if (!groups.length) {
      throw Error(`The specified group '${displayName}' does not exist.`);
    }

    if (groups.length > 1) {
      throw Error(`Multiple groups with name '${displayName}' found: ${groups.map(x => x.id).join(',')}.`);
    }

    return groups[0].id!;
  },

  async setGroup(id: string, isPrivate: boolean, logger?: Logger, verbose?: boolean): Promise<void> {
    if (verbose && logger) {
      await logger.logToStderr(`Updating Microsoft 365 Group ${id}...`);
    }

    const update: Group = {};
    if (typeof isPrivate !== 'undefined') {
      update.visibility = isPrivate ? 'Private' : 'Public';
    }

    const requestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/groups/${id}`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: update
    };

    await request.patch(requestOptions);
  }
};