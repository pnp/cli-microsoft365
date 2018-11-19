export interface Term {
  _ObjectType_: string;
  _ObjectIdentity_: string;
  CreatedDate: string;
  CustomProperties: Hash;
  Description: string;
  Id: string;
  LastModifiedDate: string;
  LocalCustomProperties: Hash;
  Name: string;
}

export interface Hash {
  [key: string] : string;
}