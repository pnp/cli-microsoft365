import { DependencyRule } from './DependencyRule.js';

export class FN002027_DEVDEP_fluentui_react extends DependencyRule {
  constructor(options: { packageVersion: string; add: boolean }) {
    super({ ...options, packageName: '@fluentui/react', isDevDep: true, isOptional: true });
  }

  get id(): string {
    return 'FN002027';
  }
}
