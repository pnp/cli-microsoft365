import { PackageRule } from "./PackageRule";

export class FN021001_PKG_main extends PackageRule {
  constructor(add: boolean, propertyValue?: string) {
    /* istanbul ignore next */
    super('main', add, propertyValue);
  }

  get id(): string {
    return 'FN021001';
  }
}