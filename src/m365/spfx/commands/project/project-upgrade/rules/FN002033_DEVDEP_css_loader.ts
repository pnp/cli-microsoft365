import { DependencyRule } from "./DependencyRule.js";

export class FN002033_DEVDEP_css_loader extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({
      packageName: 'css-loader',
      isDevDep: true,
      ...options
    });
  }

  get id(): string {
    return 'FN002033';
  }
}