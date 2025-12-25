import { PackageRule } from "./PackageRule.js";

export class FN021001_PKG_main extends PackageRule {
  constructor(options: { add: boolean; propertyValue?: string }) {
    super({
      propertyName: 'main',
      ...options
    });
  }

  get id(): string {
    return 'FN021001';
  }
}