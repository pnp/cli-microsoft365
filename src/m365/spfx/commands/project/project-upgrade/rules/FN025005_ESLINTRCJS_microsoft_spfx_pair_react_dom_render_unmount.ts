import ts from 'typescript';
import { Project } from '../../project-model/index.js';
import { Finding, Occurrence } from '../../report-model/index.js';
import { TsRule } from './TsRule.js';

export class FN025005_ESLINTRCJS_microsoft_spfx_pair_react_dom_render_unmount extends TsRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN025005';
  }

  get title(): string {
    return '.eslintrc.js override rule @microsoft/spfx/pair-react-dom-render-unmount';
  }

  get description(): string {
    return `Remove override rule @microsoft/spfx/pair-react-dom-render-unmount in .eslintrc.js`;
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
      .find(i => i.getText() === `'@microsoft/spfx/pair-react-dom-render-unmount'`);

    if (!node) {
      return;
    }

    this.addOccurrence(this.resolution, project.esLintRcJs.path, project.path, node, occurrences);

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}
