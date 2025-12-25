import { DependencyRule } from "./DependencyRule.js";

export class FN002019_DEVDEP_spfx_fast_serve_helpers extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: 'spfx-fast-serve-helpers',
      packageVersion: options.packageVersion,
      isDevDep: true,
      isOptional: true
    });
  }

  get id(): string {
    return 'FN002019';
  }
}