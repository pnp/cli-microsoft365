import { DependencyRule } from './DependencyRule.js';

export class FN001036_DEP_microsoft_sp_copilot_component extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: '@microsoft/sp-copilot-component', isOptional: true });
  }

  get id(): string {
    return 'FN001036';
  }
}