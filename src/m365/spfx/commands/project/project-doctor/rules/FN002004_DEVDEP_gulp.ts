import { DependencyRule } from './DependencyRule.js';

export class FN002004_DEVDEP_gulp extends DependencyRule {
  constructor(options: { supportedRange: string }) {
    super({ ...options, packageName: 'gulp', isDevDep: true });
  }

  get id(): string {
    return 'FN002004';
  }

  get supersedes(): string[] {
    return ['FN021010'];
  }
}