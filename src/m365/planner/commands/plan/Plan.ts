export interface Plan {
  Etag: string;
  CreatedDateTime: string;
  Id: string;
  Owner: string;
  Title: string;
  CreatedBy: PlanIdentitySet;
}

interface PlanIdentitySet {
  Application?: PlanIdentity;
  Device?: PlanIdentity;
  User?: PlanIdentity;
}

interface PlanIdentity {
  DisplayName: string;
  Id: string;
}