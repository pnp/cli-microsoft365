import { Hash } from '../../../../../utils/types';
import { JsonFile } from './JsonFile';

export interface PackageJson extends JsonFile {
  dependencies?: Hash;
  devDependencies?: Hash;
  engines?: Hash | string;
  name?: string;
  resolutions?: Hash;
}