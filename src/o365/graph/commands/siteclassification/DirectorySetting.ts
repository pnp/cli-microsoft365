import { DirectorySettingValue } from "./DirectorySettingValue";

export interface DirectorySetting {
  Id:                                        string;
  DeletedDateTime:                           Date;
  Description:                               string;
  DisplayName:                               string;
  values:                                    DirectorySettingValue[];
}