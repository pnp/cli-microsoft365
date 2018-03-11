import { DirectorySettingValue } from "./DirectorySettingValue";

export interface DirectorySetting {
  id:                                        string;
  displayName:                               string;
  values:                                    DirectorySettingValue[];
}