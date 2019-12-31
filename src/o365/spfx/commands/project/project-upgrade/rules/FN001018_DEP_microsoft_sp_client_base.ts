import { DependencyRule } from "./DependencyRule";

export class FN001018_DEP_microsoft_sp_client_base extends DependencyRule {
  constructor(packageVersion: string, add: boolean) {
    super('@microsoft/sp-client-base', packageVersion, false, true, add);
  }

  get id(): string {
    return 'FN001018';
  }
}