import { ExternalUserCollection } from "./ExternalUserCollection";

export interface GetExternalUsersResults {
  _ObjectType_: string;
  TotalUserCount: number;
  UserCollectionPosition: number;
  ExternalUserCollection: ExternalUserCollection;
}
