import { DependencyRule } from './DependencyRule';

export class FN002004_DEVDEP_gulp extends DependencyRule {
  constructor(supportedRange: string) {
    super('gulp', supportedRange, true);
  }

  get id(): string {
    return 'FN002004';
  }

  get supersedes(): string[] {
    return ['FN021010'];
  }
}