import { DependencyRule } from "./DependencyRule";

export class FN002005_DEVDEP_types_chai extends DependencyRule {
  constructor(packageVersion: string, add: boolean = true) {
    super('@types/chai', packageVersion, true, false, add);
  }

  get id(): string {
    return 'FN002005';
  }
}