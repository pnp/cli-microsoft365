import request from "../request";
import { formatting } from "./formatting";

const graphResource = 'https://graph.microsoft.com';

export const aadUser = {
  /**
   * Retrieve a single group.
   * @param id Group ID.
   */
  async getUserId(userName: string): Promise<string> {
    const requestUrl: string = `${graphResource}/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(userName)}'&$select=Id`;

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ value: { id: string }[] }>(requestOptions);

    if (res.value.length === 0) {
      throw Error(`The specified user with user name ${userName} does not exist`);
    }

    return res.value[0].id;
  }
};