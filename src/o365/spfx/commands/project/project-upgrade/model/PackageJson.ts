import { Hash } from "../";

export interface PackageJson {
  dependencies: Hash;
  devDependencies?: Hash;
  resolutions?: Hash;
}