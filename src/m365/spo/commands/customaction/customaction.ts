export interface CustomAction {
  "odata.null": boolean;
  ClientSideComponentId: string;
  ClientSideComponentProperties: Object;
  CommandUIExtension: Object;
  Description: string;
  Group: string;
  Id: string;
  ImageUrl: string;
  Location: string;
  Name: string;
  RegistrationId: number;
  RegistrationType: number;
  Rights: {
    High: number,
    Low: number
  };
  Scope: number;
  ScriptBlock: string;
  ScriptSrc: string;
  Sequence: number;
  Title: string;
  Url: string;
  VersionOfUserCustomAction: string;
}