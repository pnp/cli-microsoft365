import { DependencyRule } from "./DependencyRule.js";

export class FN001018_DEP_microsoft_sp_client_base extends DependencyRule {
  constructor(packageVersion: string, add: boolean) {
    super({
      packageName: '@microsoft/sp-client-base',
      packageVersion,
      isOptional: true,
      add
    });
  }

  get id(): string {
    return 'FN001018';
  }
}