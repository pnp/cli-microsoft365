import { DependencyRule } from "./DependencyRule.js";

export class FN002019_DEVDEP_spfx_fast_serve_helpers extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: 'spfx-fast-serve-helpers',
      isDevDep: true,
      isOptional: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002019';
  }
}