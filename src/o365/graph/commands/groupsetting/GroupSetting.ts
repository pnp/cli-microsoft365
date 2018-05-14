import { SettingValue } from "./SettingValue";

export interface GroupSetting {
  displayName: string;
  id: string;
  templateId: string;
  values: SettingValue[];
}