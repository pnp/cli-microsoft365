import { DependencyRule } from "./DependencyRule.js";

export class FN001029_DEP_microsoft_sp_loader extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: '@microsoft/sp-loader', isOptional: true });
  }

  get id(): string {
    return 'FN001029';
  }
}