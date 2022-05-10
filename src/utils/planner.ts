import request from "../request";
import { PlannerPlan } from "@microsoft/microsoft-graph-types";
import { AxiosRequestConfig } from "axios";

const graphResource = 'https://graph.microsoft.com';

const getRequestOptions = (url: string, metadata: 'none' | 'minimal' | 'full'): AxiosRequestConfig => ({
  url: url,
  headers: {
    accept: `application/json;odata.metadata=${metadata}`
  },
  responseType: 'json'
});

export const planner = {
  async getPlanById(id: string): Promise<PlannerPlan> {
    const requestOptions = getRequestOptions(`${graphResource}/v1.0/planner/plans/${id}`, 'none');
    
    try {
      return await request.get<PlannerPlan>(requestOptions);
    }
    catch (ex) {
      throw Error(`Planner plan with id ${id} was not found.`);
    }
  }
};