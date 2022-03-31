import * as parse from 'json-to-ast';
import * as path from 'path';
import { JsonRule } from '../../JsonRule';
import { Occurrence } from '../../report-model';

export abstract class ManifestRule extends JsonRule {
  get resolutionType(): string {
    return 'json';
  }

  get file(): string {
    return '';
  }

  protected addOccurrence(resolution: string, filePath: string, projectPath: string, node: parse.ASTNode | undefined, occurrences: Occurrence[]): void {
    occurrences.push({
      file: path.relative(projectPath, filePath),
      resolution: resolution,
      position: this.getPositionFromNode(node)
    });
  }
}