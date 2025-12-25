import { DependencyRule } from "./DependencyRule.js";

export class FN001026_DEP_microsoft_sp_extension_base extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: '@microsoft/sp-extension-base',
      packageVersion: options.packageVersion,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN001026';
  }
}