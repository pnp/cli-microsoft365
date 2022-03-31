import { ASTNode } from "json-to-ast";

export interface JsonFile {
  ast?: ASTNode;
  source?: string;
}