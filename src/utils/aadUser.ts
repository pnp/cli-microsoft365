import request, { CliRequestOptions } from "../request";
import { formatting } from "./formatting";

const graphResource = 'https://graph.microsoft.com';

export const aadUser = {
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
  * Retrieve a single user.
  * @param mail User E - mail.
  */
  async getUserUpnByEmail(email: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/users?$filter=mail eq '${formatting.encodeQueryParameter(email)}'&$select=userPrincipalName`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    const res = await request.get<{ value: { userPrincipalName: string }[] }>(requestOptions);

    if (res.value.length === 0) {
      throw `The specified user with e-mail ${email} does not exist`;
    }

    return res.value[0].userPrincipalName;
  }
};