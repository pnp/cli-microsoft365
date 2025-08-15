import { JsonFile } from ".";

export interface SassJson extends JsonFile {
  $schema?: string;
  extends?: string;
}
