import os from 'os';
import ts from 'typescript';
import { Project } from '../../project-model/index.js';
import { Finding, Occurrence } from '../../report-model/index.js';
import { TsRule } from './TsRule.js';

export class FN025003_ESLINTRCJS_rushstack_pair_react_dom_render_unmount extends TsRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN025003';
  }

  get title(): string {
    return '.eslintrc.js override rule @rushstack/pair-react-dom-render-unmount';
  }

  get description(): string {
    return `Add override rule @rushstack/pair-react-dom-render-unmount in .eslintrc.js`;
  }

  get resolution(): string {
    return `// Ensure that React components rendered with ReactDOM.render() are unmounted with ReactDOM.unmountComponentAtNode(). https://www.npmjs.com/package/@rushstack/eslint-plugin
        '@rushstack/pair-react-dom-render-unmount': 1,${os.EOL}`;
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
      .filter(i => i.getText() === `'@rushstack/pair-react-dom-render-unmount'`).length !== 0) {
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
