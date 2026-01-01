import { PackageRule } from "./PackageRule.js";

export class FN021002_PKG_engines extends PackageRule {
  constructor(options: { add: boolean; propertyValue?: string }) {
    super({ propertyName: 'engines', ...options });
  }

  get id(): string {
    return 'FN021002';
  }
}