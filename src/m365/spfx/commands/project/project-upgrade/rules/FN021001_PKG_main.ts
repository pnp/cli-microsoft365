import { PackageRule } from "./PackageRule.js";

export class FN021001_PKG_main extends PackageRule {
  constructor(add: boolean, propertyValue?: string) {
    super({
      propertyName: 'main',
      add,
      propertyValue
    });
  }

  get id(): string {
    return 'FN021001';
  }
}