import { spfx } from '../../../../../../utils/spfx.js';
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";
import { FileAddRemoveRule } from "./FileAddRemoveRule.js";

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
