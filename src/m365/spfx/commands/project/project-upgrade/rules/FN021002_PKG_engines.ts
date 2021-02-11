import { PackageRule } from "./PackageRule";

export class FN021002_PKG_engines extends PackageRule {
  constructor(add: boolean, propertyValue?: string) {
    super('engines', add, propertyValue);
  }

  get id(): string {
    return 'FN021002';
  }
}