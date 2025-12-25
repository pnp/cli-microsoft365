import { DependencyRule } from "./DependencyRule.js";

export class FN001013_DEP_microsoft_decorators extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      
      packageName: '@microsoft/decorators',
      packageVersion: options.packageVersion,
      
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001013';
  }
}