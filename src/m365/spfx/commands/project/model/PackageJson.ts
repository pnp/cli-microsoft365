import { Hash } from "../project-upgrade/";

export interface PackageJson {
  dependencies: Hash;
  devDependencies?: Hash;
  resolutions?: Hash;
}