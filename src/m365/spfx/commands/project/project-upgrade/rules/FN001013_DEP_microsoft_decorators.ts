import { DependencyRule } from "./DependencyRule.js";

export class FN001013_DEP_microsoft_decorators extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: '@microsoft/decorators',
      packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001013';
  }
}