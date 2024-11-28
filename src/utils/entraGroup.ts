import { Group } from "@microsoft/microsoft-graph-types";
import request, { CliRequestOptions } from "../request.js";
import { formatting } from "./formatting.js";
import { odata } from "./odata.js";
import { Logger } from '../cli/Logger.js';
import { cli } from '../cli/cli.js';

const graphResource = 'https://graph.microsoft.com';

export const entraGroup = {
  /**
   * Retrieve a single group.
   * @param id Group ID.
   * @param properties Properties to include in the response.
   */
  async getGroupById(id: string, properties?: string): Promise<Group> {
    const queryParameters: string[] = [];

    if (properties) {
      const allProperties = properties.split(',');
      const selectProperties = allProperties.filter(prop => !prop.includes('/'));

      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${selectProperties}`);
      }
    }

    const queryString = queryParameters.length > 0
      ? `?${queryParameters.join('&')}`
      : '';

    const requestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/groups/${id}${queryString}`,
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
   * @param properties Properties to include in the response.
   */
  async getGroupsByDisplayName(displayName: string, properties?: string): Promise<Group[]> {
    const queryParameters: string[] = [];

    if (properties) {
      const allProperties = properties.split(',');
      const selectProperties = allProperties.filter(prop => !prop.includes('/'));

      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${selectProperties}`);
      }
    }

    const queryString = queryParameters.length > 0
      ? `&${queryParameters.join('&')}`
      : '';

    return odata.getAllItems<Group>(`${graphResource}/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'${queryString}`);
  },

  /**
   * Get a single group by its display name.
   * @param displayName Group display name.
   * @param properties Properties to include in the response.
   * @throws Error when group was not found.
   * @throws Error when multiple groups with the same name were found.
   */
  async getGroupByDisplayName(displayName: string, properties?: string): Promise<Group> {
    const groups = await this.getGroupsByDisplayName(displayName, properties);

    if (!groups.length) {
      throw Error(`The specified group '${displayName}' does not exist.`);
    }

    if (groups.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', groups);
      return await cli.handleMultipleResultsFound<Group>(`Multiple groups with name '${displayName}' found.`, resultAsKeyValuePair);
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
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', groups);
      const result = await cli.handleMultipleResultsFound<Group>(`Multiple groups with name '${displayName}' found.`, resultAsKeyValuePair);
      return result.id!;
    }

    return groups[0].id!;
  },

  /**
   * Get id of a group by its mail nickname.
   * @param mailNickname Group mail nickname.
   * @throws Error when group was not found.
   * @throws Error when multiple groups with the same name were found.
   */
  async getGroupIdByMailNickname(mailNickname: string): Promise<string> {
    const groups = await odata.getAllItems<Group>(`${graphResource}/v1.0/groups?$filter=mailNickname eq '${formatting.encodeQueryParameter(mailNickname)}'&$select=id`);

    if (!groups.length) {
      throw Error(`The specified group '${mailNickname}' does not exist.`);
    }

    if (groups.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', groups);
      const result = await cli.handleMultipleResultsFound<Group>(`Multiple groups with mail nickname '${mailNickname}' found.`, resultAsKeyValuePair);
      return result.id!;
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
  },

  /**
   * Checks if group is a m365 group.
   * @param groupId Group id.
   * @returns whether the group is a m365 group or not
   */
  async isUnifiedGroup(groupId: string): Promise<boolean> {
    const requestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/groups/${groupId}?$select=groupTypes`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const group = await request.get<{ groupTypes: string[] }>(requestOptions);
    return group.groupTypes!.some(type => type === 'Unified');
  }
};