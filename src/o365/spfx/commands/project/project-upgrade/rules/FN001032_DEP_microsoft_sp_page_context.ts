import { DependencyRule } from "./DependencyRule";

export class FN001032_DEP_microsoft_sp_page_context extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-page-context', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001032';
  }
}