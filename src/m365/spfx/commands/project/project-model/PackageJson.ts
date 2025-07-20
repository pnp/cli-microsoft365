import { Hash } from '../../../../../utils/types.js';
import { JsonFile } from './JsonFile.js';

export interface PackageJson extends JsonFile {
  dependencies?: Hash;
  devDependencies?: Hash;
  engines?: Hash | string;
  name?: string;
  overrides?: Hash;
  resolutions?: Hash;
  scripts?: {
    build?: string;
    'build-watch'?: string;
    clean?: string;
    deploy?: string;
    'deploy-azure-storage'?: string;
    'eject-webpack'?: string;
    'package-solution'?: string;
    start?: string;
    test?: string;
  }
}