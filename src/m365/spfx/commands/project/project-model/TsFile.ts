import fs from 'fs';
import path from 'path';
import ts from 'typescript';
import { tsUtil } from '../../../../../utils/tsUtil.js';

export class TsFile {
  private _sourceFile: ts.SourceFile | undefined;
  public get sourceFile(): ts.SourceFile | undefined {
    if (!this._sourceFile) {
      if (!this.source) {
        return undefined;
      }

      try {
        this._sourceFile = tsUtil.createSourceFile(path.basename(this.path), this.source, ts.ScriptTarget.Latest, true);
      }
      catch {
        // Do nothing
      }
    }

    return this._sourceFile;
  }

  private _nodes: ts.Node[] | undefined;
  public get nodes(): ts.Node[] | undefined {
    if (!this._nodes) {
      if (!this.sourceFile) {
        return undefined;
      }

      this._nodes = TsFile.getAsEnumerable(this.sourceFile, this.sourceFile);
    }

    return this._nodes;
  }

  private _source: string | undefined;
  public get source(): string | undefined {
    if (!this._source) {
      try {
        this._source = fs.readFileSync(this.path, 'utf-8');
      }
      catch {
        // Do nothing
      }
    }

    return this._source;
  }

  constructor(public path: string) {
  }

  private static getAsEnumerable(file: ts.SourceFile, node: ts.Node): ts.Node[] {
    const nodes: ts.Node[] = [node];

    node.getChildren(file).forEach(n => {
      nodes.push(...TsFile.getAsEnumerable(file, n));
    });

    return nodes;
  }
}