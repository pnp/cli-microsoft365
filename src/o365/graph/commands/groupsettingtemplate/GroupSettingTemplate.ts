export interface GroupSettingTemplate {
  id: string;
  displayName: string;
  description: string;
  values: GroupSettingTemplateValue[];
}

export interface GroupSettingTemplateValue {
  name: string;
  type: string;
  defaultValue: string;
  description: string;
}