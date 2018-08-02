import { DependencyRule } from "./DependencyRule";

export class FN011006_DEP_microsoft_sp_client_base extends DependencyRule {
  constructor(packageVersion: string) {
    /* istanbul ignore next */
    super('@microsoft/sp-client-base', packageVersion, false, true);
  }

  get id(): string {
    return 'FN011006';
  }
}