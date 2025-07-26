import assert from 'assert';
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";
import { FN023004_GITIGNORE_libcommonjs } from './FN023004_GITIGNORE_libcommonjs.js';

describe('FN023004_GITIGNORE_libcommonjs', () => {
  let findings: Finding[];
  let rule: FN023004_GITIGNORE_libcommonjs;

  beforeEach(() => {
    findings = [];
    rule = new FN023004_GITIGNORE_libcommonjs();
  });

  it(`doesn't return notification if the lib-commonjs folder is already excluded`, () => {
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
