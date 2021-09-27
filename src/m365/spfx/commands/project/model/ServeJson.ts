import { JsonFile } from ".";

export interface ServeJson extends JsonFile {
  $schema: string;
  api?: any;
  initialPage?: string;
}