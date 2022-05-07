import { PlannerPlan } from "@microsoft/microsoft-graph-types";
import { odata } from "./odata";

const graphResource = 'https://graph.microsoft.com';

export const planner = {
  getPlansByGroupId(groupId: string): Promise<PlannerPlan[]> {
    return odata.getAllItems<PlannerPlan>(`${graphResource}/v1.0/groups/${groupId}/planner/plans`, undefined as any, 'none');
  }
};