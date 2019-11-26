import { Finding, Occurrence } from "../";
import { Project } from "../../model";
import { TsRule } from "./TsRule";
import * as ts from 'typescript';

export class FN016003_TS_aadhttpclient_instance extends TsRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN016003';
  }

  get title(): string {
    return 'AadHttpClient instance';
  }

  get description(): string {
    return `Refactor the code to get AadHttpClient instance`;
  }

  get resolution(): string {
    return '';
  };

  get resolutionType(): string {
    return 'ts';
  }

  get severity(): string {
    return 'Required';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.tsFiles) {
      return;
    }

    const occurrences: Occurrence[] = [];
    project.tsFiles.forEach(file => {
      const nodes: ts.Node[] | undefined = file.nodes;
      if (!nodes) {
        return;
      }

      const newAddHttpClient = nodes
        .filter(n => ts.isNewExpression(n))
        .map(n => n as ts.NewExpression)
        .filter(n => n.expression.getText() === 'AadHttpClient');

      newAddHttpClient.forEach(n => {
        let resource: ts.Node | undefined = undefined;
        if (n.arguments && n.arguments.length === 2) {
          resource = n.arguments[1];
        }

        const varDec: ts.VariableDeclaration | undefined = TsRule.getParentOfType<ts.VariableDeclaration>(n, ts.isVariableDeclaration);
        if (varDec) {
          const resourceString = resource ? resource.getText() : '/* your resource */';
          const resolution = `this.context.aadHttpClientFactory
  .getClient(${resourceString})
  .then((client: AadHttpClient): void => {
    // use AadHttpClient here
  });`;
          this.addOccurrence(resolution, file.path, project.path, varDec, occurrences);
        }
      });
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}
