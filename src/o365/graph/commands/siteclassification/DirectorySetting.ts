import { DirectorySettingValue } from "./DirectorySettingValue";

export interface DirectorySetting {
  id: string;
  displayName: string;
  templateId: string;
  values: DirectorySettingValue[];
}