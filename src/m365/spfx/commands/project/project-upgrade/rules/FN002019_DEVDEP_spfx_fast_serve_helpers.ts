import { DependencyRule } from "./DependencyRule";

export class FN002019_DEVDEP_spfx_fast_serve_helpers extends DependencyRule {
  constructor(packageVersion: string) {
    super('spfx-fast-serve-helpers', packageVersion, true, true);
  }

  get id(): string {
    return 'FN002019';
  }
}