import { PackageRule } from "./PackageRule.js";

export class FN021002_PKG_engines extends PackageRule {
  constructor(add: boolean, propertyValue?: string) {
    super({
      propertyName: 'engines',
      add,
      propertyValue
    });
  }

  get id(): string {
    return 'FN021002';
  }
}