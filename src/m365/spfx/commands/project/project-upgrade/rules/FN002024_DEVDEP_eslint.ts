import { DependencyRule } from "./DependencyRule";

export class FN002024_DEVDEP_eslint extends DependencyRule {
  constructor(packageVersion: string) {
    super('eslint', packageVersion, true);
  }

  get id(): string {
    return 'FN002024';
  }
}