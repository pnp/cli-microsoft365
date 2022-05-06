import { PlannerPlan } from "@microsoft/microsoft-graph-types";
import { AxiosRequestConfig } from "axios";
import request from "../request";

const graphResource = 'https://graph.microsoft.com';

export const planner = {
  async getPlansByGroupId(groupId: string): Promise<PlannerPlan[]> {
    const requestOptions: AxiosRequestConfig = {
      url: `${graphResource}/v1.0/planner/plans?$filter=owner eq '${groupId}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: PlannerPlan[] }>(requestOptions);
    return response.value;
  }
};