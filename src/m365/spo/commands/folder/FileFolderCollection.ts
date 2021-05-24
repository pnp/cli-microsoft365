import { FileProperties } from '../file/FileProperties';
import { FolderProperties } from './FolderProperties';

// Represents folder properties along with Files and Folder within.
export interface FileFolderCollection {
  Files: FileProperties[];
  Folders:FolderProperties[];
  FolderProperties:FolderProperties
}