import { DependencyRule } from './DependencyRule.js';

export class FN002007_DEVDEP_ajv extends DependencyRule {
  constructor(options: { supportedRange: string }) {
    super({ ...options, packageName: 'ajv', isDevDep: true });
  }

  get id(): string {
    return 'FN002007';
  }

  get supersedes(): string[] {
    return ['FN021011'];
  }
}