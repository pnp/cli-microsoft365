import { DependencyRule } from "./DependencyRule";

export class FN002010_DEVDEP_microsoft_rush_stack_compiler extends DependencyRule {
  constructor() {
    /* istanbul ignore next */
    // super('@microsoft/rush-stack-compiler-2.7', packageVersion, true);
    // protected packageName: string, protected packageVersion: string
    super('@microsoft/rush-stack-compiler-2.7', '0.4.0', true)
  }

  get id(): string {
    return 'FN002010';
  }

  set tscVersion(tsVersion: string) {
    this.packageName = `@microsoft/rush-stack-compiler-${tsVersion}`;
  }

  set packageVersion(packageVersion: string) {
    // BREAKS EVERYTHING
    // this.packageVersion = `${packageVersion}`;
  }

}