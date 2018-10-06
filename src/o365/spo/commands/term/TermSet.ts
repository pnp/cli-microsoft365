export interface TermSet {
  _ObjectType_: string;
  _ObjectIdentity_: string;
  CreatedDate: string;
  Description: string;
  Id: string;
  LastModifiedDate: string;
  Name: string;
  CustomProperties: Hash;
}

export interface Hash {
  [key: string] : string;
}