import { DependencyRule } from "./DependencyRule";

export class FN001004_DEP_microsoft_sp_webpart_base extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-webpart-base', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001004';
  }
}