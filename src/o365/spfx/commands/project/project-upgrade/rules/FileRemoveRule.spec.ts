import * as assert from 'assert';
import * as path from 'path';
import { Finding } from '../Finding';
import { Project } from '../model';
import { FileRemoveRule } from './FileRemoveRule';

class FileRule extends FileRemoveRule {
  constructor(filePath: string) {
    super(filePath);
  }

  get id(): string {
    return 'FN000000';
  }
}
describe('FileRemoveRule', () => {
  let findings: Finding[];
  let rule: FileRule;

  beforeEach(() => {
    findings = [];
  });

  it('doesn\'t return notification file doesn\'t exist', () => {
    rule = new FileRule('dummy.json');
    const project: Project = {
      path: path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-102-webpart-react'),
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });

  it('returns a notification if file exists', () => {
    rule = new FileRule('/typings/tsd.d.ts');
    const project: Project = {
      path: path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-102-webpart-react'),
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 1);
  });
});