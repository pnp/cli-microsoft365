import request from "../request";
import { odata } from "./odata";
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
  /**
   * Get Planner plan by ID.
   * @param id Planner ID.
   * @param metadata OData metadata level. Default is none
   */
  async getPlanById(id: string, metadata: 'none' | 'minimal' | 'full' = 'none'): Promise<PlannerPlan> {
    const requestOptions = getRequestOptions(`${graphResource}/v1.0/planner/plans/${id}`, metadata);
    
    try {
      return await request.get<PlannerPlan>(requestOptions);
    }
    catch (ex) {
      throw Error(`Planner plan with id '${id}' was not found.`);
    }
  },

  /**
   * Get all Planner plans for a specific group.
   * @param groupId Group ID.
   */
  getPlansByGroupId(groupId: string): Promise<PlannerPlan[]> {
    return odata.getAllItems<PlannerPlan>(`${graphResource}/v1.0/groups/${groupId}/planner/plans`, 'none');
  },

  /**
   * Get Planner plan by name in a specific group. 
   * @param name Name of the Planner plan. Case insensitive.
   * @param groupId Owner group ID .
   */
  async getPlanByName(name: string, groupId: string): Promise<PlannerPlan> {
    const plans = await this.getPlansByGroupId(groupId);
    const filteredPlans = plans.filter(p => p.title && p.title.toLowerCase() === name.toLowerCase());

    if (!filteredPlans.length) {
      throw Error(`The specified plan '${name}' does not exist.`);
    }

    if (filteredPlans.length > 1) {
      throw Error(`Multiple plans with name '${name}' found: ${filteredPlans.map(x => x.id)}.`);
    }

    return filteredPlans[0];
  }
};