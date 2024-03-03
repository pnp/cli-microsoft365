import request, { CliRequestOptions } from "../request.js";
import { formatting } from "./formatting.js";

const graphResource = 'https://graph.microsoft.com';

export const entraUser = {
  /**
   * Retrieve the id of a user by its UPN.
   * @param upn User UPN.
   */
  async getUserIdByUpn(upn: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(upn)}'&$select=Id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    const res = await request.get<{ value: { id: string }[] }>(requestOptions);

    if (res.value.length === 0) {
      throw Error(`The specified user with user name ${upn} does not exist.`);
    }

    return res.value[0].id;
  },

  /**
   * Retrieve the IDs of users by their UPNs. There is no guarantee that the order of the returned IDs will match the order of the specified UPNs.
   * @param upns Array of user UPNs.
   * @returns Array of user IDs.
   */
  async getUserIdsByUpns(upns: string[]): Promise<string[]> {
    const userIds: string[] = [];

    for (let i = 0; i < upns.length; i += 20) {
      const upnsChunk = upns.slice(i, i + 20);
      const requestOptions: CliRequestOptions = {
        url: `${graphResource}/v1.0/$batch`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          requests: upnsChunk.map((upn, index) => ({
            id: index + 1,
            method: 'GET',
            url: `/users/${formatting.encodeQueryParameter(upn)}?$select=id`,
            headers: {
              accept: 'application/json;odata.metadata=none'
            }
          }))
        }
      };
      const res = await request.post<{ responses: { id: number; status: number; body: { id: string } }[] }>(requestOptions);

      for (const response of res.responses) {
        if (response.status !== 200) {
          throw Error(`The specified user with user name '${upnsChunk[response.id - 1]}' does not exist.`);
        }

        userIds.push(response.body.id);
      }
    }

    return userIds;
  },

  /**
  * Retrieve the ID of a user by its email.
  * @param mail User email.
  */
  async getUserIdByEmail(email: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/users?$filter=mail eq '${formatting.encodeQueryParameter(email)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    const res = await request.get<{ value: { id: string }[] }>(requestOptions);

    if (res.value.length === 0) {
      throw Error(`The specified user with email ${email} does not exist`);
    }

    return res.value[0].id;
  },

  /**
   * Retrieve the UPN of a user by its ID.
   * @param id User ID.
   */
  async getUpnByUserId(id: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/users/${id}?$select=userPrincipalName`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ userPrincipalName: string }>(requestOptions);

    return res.userPrincipalName;
  }
};