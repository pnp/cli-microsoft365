import { Hash } from "../project-upgrade/";
import { JsonFile } from "./JsonFile";

export interface PackageJson extends JsonFile {
  name?: string;
  dependencies: Hash;
  devDependencies?: Hash;
  resolutions?: Hash;
}