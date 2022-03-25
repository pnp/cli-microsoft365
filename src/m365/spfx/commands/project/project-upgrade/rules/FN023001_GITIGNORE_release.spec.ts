import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';
import { FN023001_GITIGNORE_release } from './FN023001_GITIGNORE_release';

describe('FN023001_GITIGNORE_release', () => {
  let findings: Finding[];
  let rule: FN023001_GITIGNORE_release;

  beforeEach(() => {
    findings = [];
    rule = new FN023001_GITIGNORE_release();
  });

  it(`doesn't return notification if the release folder is already excluded`, () => {
    const project: Project = {
      path: '/usr/tmp',
      gitignore: {
        source: rule.resolution
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't return notification if gitignore is not available`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});