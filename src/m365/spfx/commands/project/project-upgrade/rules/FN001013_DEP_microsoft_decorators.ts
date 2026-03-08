import { DependencyRule } from "./DependencyRule.js";

export class FN001013_DEP_microsoft_decorators extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: '@microsoft/decorators', isOptional: true });
  }

  get id(): string {
    return 'FN001013';
  }
}