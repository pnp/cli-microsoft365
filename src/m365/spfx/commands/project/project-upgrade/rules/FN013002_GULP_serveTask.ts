import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN013002_GULP_serveTask extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN013002';
  }

  get title(): string {
    return 'gulpfile.js serve task';
  }

  get description(): string {
    return `Before 'build.initialize(require('gulp'));' add the serve task`;
  }

  get resolution(): string {
    return `var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};
`;
  }

  get resolutionType(): string {
    return 'js';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './gulpfile.js';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.gulpfileJs) {
      return;
    }

    if (project.gulpfileJs.source.indexOf(`result.set('serve', result.get('serve-deprecated'));`) < 0) {
      this.addFinding(findings);
    }
  }
}