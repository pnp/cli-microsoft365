import { FileProperties } from '../file/FileProperties.js';
import { FolderProperties } from './FolderProperties.js';

// Represents folder properties along with Files and Folder within.
export interface FileFolderCollection {
  Files: FileProperties[];
  Folders: FolderProperties[];
  FolderProperties: FolderProperties
}