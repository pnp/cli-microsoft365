import { DependencyRule } from "./DependencyRule";

export class FN001034_DEP_microsoft_sp_adaptive_card_extension_base extends DependencyRule {
  constructor(packageVersion: string) {
    super('@microsoft/sp-adaptive-card-extension-base', packageVersion, false, false);
  }

  get id(): string {
    return 'FN001034';
  }

  get severity(): string {
    return 'Optional';
  }
}
