import * as assert from 'assert';
import { Project } from "../../project-model";
import { Finding } from "../../report-model";
import { FN023002_GITIGNORE_heft } from './FN023002_GITIGNORE_heft';

describe('FN023002_GITIGNORE_heft', () => {
  let findings: Finding[];
  let rule: FN023002_GITIGNORE_heft;

  beforeEach(() => {
    findings = [];
    rule = new FN023002_GITIGNORE_heft();
  });

  it(`doesn't return notification if the .heft folder is already excluded`, () => {
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