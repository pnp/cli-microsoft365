import { DependencyRule } from './DependencyRule';

export class FN002007_DEVDEP_ajv extends DependencyRule {
  constructor(supportedRange: string) {
    super('ajv', supportedRange, true);
  }

  get id(): string {
    return 'FN002007';
  }

  get supersedes(): string[] {
    return ['FN021011'];
  }
}