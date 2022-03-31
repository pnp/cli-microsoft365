import { DependencyRule } from './DependencyRule';

export class FN002013_DEVDEP_types_webpack_env extends DependencyRule {
  constructor(supportedRange: string) {
    super('@types/webpack-env', supportedRange, true);
  }

  get id(): string {
    return 'FN002013';
  }
}