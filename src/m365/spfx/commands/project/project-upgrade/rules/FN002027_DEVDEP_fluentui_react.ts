import { DependencyRule } from './DependencyRule.js';

export class FN002027_DEVDEP_fluentui_react extends DependencyRule {
  constructor(packageVersion: string, add: boolean) {
    super({
      packageName: '@fluentui/react',
      packageVersion,
      isDevDep: true,
      isOptional: true,
      add
    });
  }

  get id(): string {
    return 'FN002027';
  }
}
