import { DependencyRule } from './DependencyRule.js';

export class FN002021_DEVDEP_rushstack_eslint_config extends DependencyRule {
  constructor(supportedRange: string) {
    super('@rushstack/eslint-config', supportedRange, true);
  }

  get id(): string {
    return 'FN002021';
  }
}