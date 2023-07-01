import { DependencyRule } from "./DependencyRule.js";

export class FN002006_DEVDEP_types_mocha extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@types/mocha', packageVersion, true, false, add);
  }

  get id(): string {
    return 'FN002006';
  }
}