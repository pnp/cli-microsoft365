import { Finding } from "../";
import { Project } from "../../model";
import { JsonRule } from "./JsonRule";

export class FN012019_TSC_types_es6_promise extends JsonRule {
  constructor(private add: boolean) {
    super();
  }

  get id(): string {
    return 'FN012019';
  }

  get title(): string {
    return 'tsconfig.json es6-promise types';
  }

  get description(): string {
    return `${(this.add ? 'Add' : 'Remove')} es6-promise type in tsconfig.json`;
  };

  get resolution(): string {
    return `{
  "compilerOptions": {
    "types": [
      "es6-promise"
    ]
  }
}`;
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Required';
  };

  get file(): string {
    return './tsconfig.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.tsConfigJson || !project.tsConfigJson.compilerOptions) {
      return;
    }

    if (this.add) {
      if (!project.tsConfigJson.compilerOptions.types ||
        project.tsConfigJson.compilerOptions.types.indexOf('es6-promise') < 0) {
        const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.types');
        this.addFindingWithPosition(findings, node);
      }
    }
    else {
      if (project.tsConfigJson.compilerOptions.types &&
        project.tsConfigJson.compilerOptions.types.indexOf('es6-promise') > -1) {
        const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.types[es6-promise]');
        this.addFindingWithPosition(findings, node);
      }
    }
  }
}