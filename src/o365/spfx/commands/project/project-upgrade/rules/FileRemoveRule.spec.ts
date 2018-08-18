import * as assert from 'assert';
import * as path from 'path';
import { Finding } from '../Finding';
import { Project } from '../model';
import { FileRemoveRule } from './FileRemoveRule';

describe('FileRemoveRule', () => {
  let findings: Finding[];
  let rule: FileRemoveRule;

  beforeEach(() => {
    findings = [];
  });

  it('doesn\'t return notification file doesn\'t exist', () => {
    rule = new FileRemoveRule('dummy.json', 'FN000000');
    const project: Project = {
      path: path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-102-webpart-react'),
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });

  it('returns a notification if file exists', () => {
    rule = new FileRemoveRule('/typings/tsd.d.ts', 'FN000000');
    const project: Project = {
      path: path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-102-webpart-react'),
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 1);
  });
});