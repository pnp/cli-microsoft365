import { Group } from "@microsoft/microsoft-graph-types";
import { AxiosRequestConfig } from "axios";
import request from "../request";

const graphResource = 'https://graph.microsoft.com';

const getRequestOptions = (url: string, metadata: 'none' | 'minimal' | 'full'): AxiosRequestConfig => ({
  url: url,
  headers: {
    accept: `application/json;odata.metadata=${metadata}`
  },
  responseType: 'json'
});

export const aadGroup = {
  /**
   * Retrieve a unique group.
   * @param id Group ID.
   */
  async getGroupById(id: string): Promise<Group> {
    const requestOptions = getRequestOptions(`${graphResource}/v1.0/groups/${id}`, 'none');
    
    try {
      return await request.get<Group>(requestOptions);
    }
    catch(ex) {
      throw Error(`Group with ID ${id} was not found.`);
    }
  }
};