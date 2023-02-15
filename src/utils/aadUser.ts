import request, { CliRequestOptions } from "../request";
import { formatting } from "./formatting";

const graphResource = 'https://graph.microsoft.com';

export const aadUser = {
  /**
   * Retrieve the id of the user.
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
      throw Error(`The specified user with user name ${upn} does not exist`);
    }

    return res.value[0].id;
  }
};