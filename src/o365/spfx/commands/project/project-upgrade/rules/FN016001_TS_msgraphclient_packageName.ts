import { Finding, Occurrence } from "../";
import { Project } from "../../model";
import { TsRule } from "./TsRule";
import * as ts from 'typescript';

export class FN016001_TS_msgraphclient_packageName extends TsRule {
  constructor(private packageName: string) {
    super();
  }

  get id(): string {
    return 'FN016001';
  }

  get title(): string {
    return 'MSGraphClient package name';
  }

  get description(): string {
    return `Change the name of the package to import MSGraphClient to '${this.packageName}'`;
  }

  get resolution(): string {
    return ``;
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

      const msGraphImports: ts.ImportSpecifier[] = nodes
        .filter(n => ts.isImportSpecifier(n))
        .map(n => n as ts.ImportSpecifier)
        .filter(i => i.name.text === 'MSGraphClient');
      msGraphImports.forEach(msGraphImport => {
        const msGraphImportDeclaration: ts.ImportDeclaration | undefined = TsRule.getParentOfType<ts.ImportDeclaration>(msGraphImport, ts.isImportDeclaration);
        if (!msGraphImportDeclaration) {
          return;
        }

        const moduleSpecifier: string = msGraphImportDeclaration.moduleSpecifier.getText();
        if (moduleSpecifier !== `"${this.packageName}"` &&
          moduleSpecifier !== `'${this.packageName}'`) {
          const resolution: string = msGraphImportDeclaration.getText(msGraphImportDeclaration.getSourceFile()).replace(moduleSpecifier, `"${this.packageName}"`);
          this.addOccurrence(resolution, file.path, project.path, msGraphImportDeclaration, occurrences);
        }
      });
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}
