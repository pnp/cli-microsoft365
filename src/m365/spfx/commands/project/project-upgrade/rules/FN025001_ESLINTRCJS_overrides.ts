import { Finding, Occurrence } from '../../report-model';
import { Project } from '../../project-model';
import { TsRule } from './TsRule';
import ts = require('typescript');

export class FN025001_ESLINTRCJS_overrides extends TsRule {
  constructor(private contents: string) {
    super();
  }

  get id(): string {
    return 'FN025001';
  }

  get title(): string {
    return '.eslintrc.js overrides';
  }

  get description(): string {
    return `Add overrides in .eslintrc.js`;
  }

  get resolution(): string {
    return `module.exports = {
      overrides: [
        ${this.contents}
      ]
    };`;
  }

  get resolutionType(): string {
    return 'js';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './.eslintrc.js';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.esLintRcJs) {
      return;
    }

    const nodes: ts.Node[] | undefined = project.esLintRcJs.nodes;
    if (!nodes) {
      return;
    }

    const occurrences: Occurrence[] = [];
    if (nodes
      .filter(node => ts.isIdentifier(node))
      .map(node => node as ts.Identifier)
      .filter(i => i.getText() === 'overrides').length !== 0) {
      return;
    }

    const node = nodes
      .map(node => node as ts.Identifier)
      .find(i => i.text === 'module');

    if (!node) {
      return;
    }

    this.addOccurrence(this.resolution, project.esLintRcJs.path, project.path, node, occurrences);

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}
