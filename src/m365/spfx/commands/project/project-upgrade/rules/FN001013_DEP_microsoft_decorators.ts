import { DependencyRule } from "./DependencyRule.js";

export class FN001013_DEP_microsoft_decorators extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/decorators',
      isOptional: true,
      ...options
    });
  }

  get id(): string {
    return 'FN001013';
  }
}