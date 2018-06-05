import { Rule } from "./Rule";

export abstract class ManifestRule extends Rule {
  get title(): string {
    return '';
  }

  get description(): string {
    return '';
  };

  get resolutionType(): string {
    return 'json';
  };

  get file(): string {
    return '';
  };
}