import { spfx } from '../../../../../../utils';
import { Project } from "../../project-model";
import { Finding } from "../../report-model";
import { FileAddRemoveRule } from "./FileAddRemoveRule";

export class FN015008_FILE_eslintrc_js extends FileAddRemoveRule {
  constructor(add: boolean, contents: string) {
    super('./.eslintrc.js', add, contents);
  }

  get id(): string {
    return 'FN015008';
  }

  public visit(project: Project, notifications: Finding[]): void {
    if (spfx.isReactProject(project)) {
      this.contents = this.contents!.replace('@microsoft/eslint-config-spfx/lib/profiles/default', '@microsoft/eslint-config-spfx/lib/profiles/react');
    }

    super.visit(project, notifications);
  }
}
