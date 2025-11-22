import ts from 'typescript';
import { Project } from '../../project-model/index.js';
import { Finding, Occurrence } from '../../report-model/index.js';
import { TsRule } from './TsRule.js';

export class FN025004_ESLINTRCJS_microsoft_spfx_import_requires_chunk_name extends TsRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN025004';
  }

  get title(): string {
    return '.eslintrc.js override rule @microsoft/spfx/import-requires-chunk-name';
  }

  get description(): string {
    return `Remove override rule @microsoft/spfx/import-requires-chunk-name in .eslintrc.js`;
  }

  get resolution(): string {
    return '';
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
    const node = nodes
      .filter(node => ts.isPropertyName(node))
      .map(node => node as ts.PropertyName)
      .find(i => i.getText() === `'@microsoft/spfx/import-requires-chunk-name'`);

    if (!node) {
      return;
    }

    this.addOccurrence(this.resolution, project.esLintRcJs.path, project.path, node, occurrences);

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}
