import { DependencyRule } from "./DependencyRule.js";

export class FN002033_DEVDEP_css_loader extends DependencyRule {
  constructor(options: { packageVersion: string }) {
    super({ ...options, packageName: 'css-loader', isDevDep: true });
  }

  get id(): string {
    return 'FN002033';
  }
}