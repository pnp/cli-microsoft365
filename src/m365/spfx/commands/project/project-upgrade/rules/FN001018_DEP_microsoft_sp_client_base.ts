import { DependencyRule } from "./DependencyRule.js";

export class FN001018_DEP_microsoft_sp_client_base extends DependencyRule {
  constructor(options: { packageVersion: string; add: boolean }) {
    super({ ...options, packageName: '@microsoft/sp-client-base', isOptional: true });
  }

  get id(): string {
    return 'FN001018';
  }
}