import * as path from 'path';
import { Occurrence } from "../";
import { Rule } from "./Rule";

export abstract class ScssRule extends Rule {
  get resolutionType(): string {
    return 'scss';
  }

  get file(): string {
    return '';
  }

  protected addOccurrence(resolution: string, filePath: string, projectPath: string, occurrences: Occurrence[]): void {
    occurrences.push({
      file: path.relative(projectPath, filePath),
      resolution: resolution
    });
  }
}