import { ExternalUserCollection } from "./ExternalUserCollection.js";

export interface GetExternalUsersResults {
  _ObjectType_: string;
  TotalUserCount: number;
  UserCollectionPosition: number;
  ExternalUserCollection: ExternalUserCollection;
}
