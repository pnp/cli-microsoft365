import { DependencyRule } from "./DependencyRule";

export class FN001001_DEP_microsoft_sp_core_library extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-core-library', packageVersion, false);
  }

  get id(): string {
    return 'FN001001';
  }
}