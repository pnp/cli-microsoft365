import { DependencyRule } from './DependencyRule';

export class FN002027_DEVDEP_fluentui_react extends DependencyRule {
  constructor(packageVersion: string, add: boolean) {
    super('@fluentui/react', packageVersion, true, true, add);
  }

  get id(): string {
    return 'FN002027';
  }
}
