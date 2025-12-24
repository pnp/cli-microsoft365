import { DependencyRule } from "./DependencyRule.js";

export class FN002033_DEVDEP_css_loader extends DependencyRule {
  constructor(packageVersion: string) {
    super({
      packageName: 'css-loader',
      packageVersion,
      isDevDep: true
    });
  }

  get id(): string {
    return 'FN002033';
  }
}