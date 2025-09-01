import assert from 'assert';
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";
import { FN023003_GITIGNORE_libdts } from './FN023003_GITIGNORE_libdts.js';

describe('FN023003_GITIGNORE_libdts', () => {
  let findings: Finding[];
  let rule: FN023003_GITIGNORE_libdts;

  beforeEach(() => {
    findings = [];
    rule = new FN023003_GITIGNORE_libdts();
  });

  it(`doesn't return notification if the lib-dts folder is already excluded`, () => {
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
