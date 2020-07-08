import { DependencyRule } from "./DependencyRule";

export class FN002005_DEVDEP_types_chai extends DependencyRule {
  constructor(packageVersion: string) {
    super('@types/chai', packageVersion, true);
  }

  get id(): string {
    return 'FN002005';
  }
}