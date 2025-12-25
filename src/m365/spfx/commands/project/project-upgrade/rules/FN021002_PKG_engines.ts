import { PackageRule } from "./PackageRule.js";

export class FN021002_PKG_engines extends PackageRule {
  constructor(options: { add: boolean; propertyValue?: string }) {
    super({
      propertyName: 'engines',
      add: options.add,
      propertyValue: options.propertyValue
    });
  }

  get id(): string {
    return 'FN021002';
  }
}