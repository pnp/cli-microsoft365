import { createRequire } from 'module';
const require = createRequire(import.meta.url);
let packageJson: PackageJson | undefined;

export interface PackageJson {
  description: string;
  homepage: string;
  version: string;
}

export const app = {
  packageJson: (): PackageJson => {
    if (!packageJson) {
      packageJson = require('../../package.json');
    }

    return packageJson!;
  }
};