import { Rule } from "./Rule";
import { Occurrence } from "../";
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

  protected addOccurrence(resolution: string, filePath: string, projectPath: string, node: ts.Node, occurrences: Occurrence[]): void {
    const lineChar: ts.LineAndCharacter = node.getSourceFile().getLineAndCharacterOfPosition(node.getStart());

    occurrences.push({
      file: path.relative(projectPath, filePath),
      position: {
        line: lineChar.line + 1,
        character: lineChar.character + 1
      },
      resolution: resolution
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