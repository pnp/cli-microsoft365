import { DependencyRule } from "./DependencyRule";

export class FN001026_DEP_microsoft_sp_extension_base extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-extension-base', packageVersion, false, true);
  }

  get id(): string {
    return 'FN001026';
  }
}