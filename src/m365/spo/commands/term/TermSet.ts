export interface TermSet {
  _ObjectType_?: string;
  _ObjectIdentity_?: string;
  CreatedDate: string;
  Description: string;
  Id: string;
  LastModifiedDate: string;
  Name: string;
  CustomProperties: Hash;
}

interface Hash {
  [key: string]: string;
}