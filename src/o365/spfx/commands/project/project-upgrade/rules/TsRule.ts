import { Rule } from "./Rule";
import { Finding } from "../";
import * as ts from 'typescript';
import * as path from 'path';

export abstract class TsRule extends Rule {
  get resolution(): string {
    return '';
  }

  get resolutionType(): string {
    return 'ts';
  };

  get file(): string {
    return '';
  };

  protected addTsFinding(findingNumber: number, resolution: string, filePath: string, projectPath: string, node: ts.Node, findings: Finding[]): void {
    const lineChar: ts.LineAndCharacter = node.getSourceFile().getLineAndCharacterOfPosition(node.getStart());

    findings.push({
      id: `${this.id}_${findingNumber}`,
      title: this.title,
      description: this.description,
      resolution: resolution,
      resolutionType: this.resolutionType,
      file: path.relative(projectPath, filePath),
      severity: this.severity,
      position: {
        line: lineChar.line + 1,
        character: lineChar.character + 1
      }
    });
  }

  protected static getParentOfType<TParent>(node: ts.Node, typeComparer: (node: ts.Node) => boolean): TParent | undefined {
    const parent: ts.Node | undefined = node.parent;

    if (!parent) {
      return undefined;
    }

    if (typeComparer(parent)) {
      return <TParent><any>parent;
    }

    return TsRule.getParentOfType<TParent>(parent, typeComparer);
  }
}