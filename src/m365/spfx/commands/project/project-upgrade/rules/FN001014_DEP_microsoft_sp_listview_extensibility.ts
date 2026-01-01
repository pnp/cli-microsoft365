import { DependencyRule } from "./DependencyRule.js";

export class FN001014_DEP_microsoft_sp_listview_extensibility extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: '@microsoft/sp-listview-extensibility', isOptional: true });
  }

  get id(): string {
    return 'FN001014';
  }
}