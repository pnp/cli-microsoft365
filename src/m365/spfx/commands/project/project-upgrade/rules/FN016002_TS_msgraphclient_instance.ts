import { Finding, Occurrence } from "../";
import { Project } from "../../model";
import { TsRule } from "./TsRule";
import * as ts from 'typescript';

export class FN016002_TS_msgraphclient_instance extends TsRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN016002';
  }

  get title(): string {
    return 'MSGraphClient instance';
  }

  get description(): string {
    return `Refactor the code to get MSGraphClient instance`;
  }

  get resolution(): string {
    return `this.context.msGraphClientFactory
  .getClient()
  .then((client: MSGraphClient): void => {
    // use MSGraphClient here
  });`;
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

      const propertyAccessExpressions = nodes
        .filter(n => ts.isPropertyAccessExpression(n))
        .map(n => n as ts.PropertyAccessExpression)
        .filter(p => p.name.text === 'consume' && p.expression.getText().endsWith('.serviceScope'));

      propertyAccessExpressions.forEach(p => {
        const callExpression: ts.CallExpression | undefined = TsRule.getParentOfType<ts.CallExpression>(p, ts.isCallExpression);
        if (!callExpression || callExpression.arguments.length < 1) {
          return;
        }

        if (!ts.isPropertyAccessExpression(callExpression.arguments[0])) {
          return;
        }

        const prop: ts.PropertyAccessExpression = callExpression.arguments[0] as ts.PropertyAccessExpression;
        if (prop.expression.getText() !== 'MSGraphClient' || prop.name.text !== 'serviceKey') {
          return;
        }

        const varDec: ts.VariableDeclaration | undefined = TsRule.getParentOfType<ts.VariableDeclaration>(p, ts.isVariableDeclaration);
        if (varDec) {
          this.addOccurrence(this.resolution, file.path, project.path, varDec, occurrences);
        }
      });
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}
