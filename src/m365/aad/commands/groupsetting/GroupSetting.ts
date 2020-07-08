export interface GroupSetting {
  id: string;
  displayName: string;
  templateId: string;
  values: SettingValue[];
}

export interface SettingValue {
  name: string;
  value: string;
}