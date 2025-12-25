import { PackageRule } from "./PackageRule.js";

export class FN021001_PKG_main extends PackageRule {
  constructor(options: { add: boolean; propertyValue?: string }) {
    super({
      propertyName: 'main',
      add: options.add,
      propertyValue: options.propertyValue
    });
  }

  get id(): string {
    return 'FN021001';
  }
}