import os from 'os';
import ts from 'typescript';
import { Project } from '../../project-model/index.js';
import { Finding, Occurrence } from '../../report-model/index.js';
import { TsRule } from './TsRule.js';

export class FN025002_ESLINTRCJS_rushstack_import_requires_chunk_name extends TsRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN025002';
  }

  get title(): string {
    return '.eslintrc.js override rule @rushstack/import-requires-chunk-name';
  }

  get description(): string {
    return `Add override rule @rushstack/import-requires-chunk-name in .eslintrc.js`;
  }

  get resolution(): string {
    return `// Require chunk names for dynamic imports in SPFx projects. https://www.npmjs.com/package/@rushstack/eslint-plugin
        '@rushstack/import-requires-chunk-name': 1,${os.EOL}`;
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
      .filter(node => ts.isPropertyName(node))
      .map(node => node as ts.PropertyNameLiteral)
      .filter(i => i.getText() === `'@rushstack/import-requires-chunk-name'`).length !== 0) {
      return;
    }

    const node = nodes
      .map(node => node as ts.Identifier)
      .find(i => i.text === 'rules');

    if (!node) {
      return;
    }

    this.addOccurrence(this.resolution, project.esLintRcJs.path, project.path, node, occurrences);

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}
