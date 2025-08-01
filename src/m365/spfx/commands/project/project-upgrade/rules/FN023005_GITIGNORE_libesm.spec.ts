import assert from 'assert';
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";
import { FN023005_GITIGNORE_libesm } from './FN023005_GITIGNORE_libesm.js';

describe('FN023005_GITIGNORE_libesm', () => {
  let findings: Finding[];
  let rule: FN023005_GITIGNORE_libesm;

  beforeEach(() => {
    findings = [];
    rule = new FN023005_GITIGNORE_libesm();
  });

  it(`doesn't return notification if the lib-esm folder is already excluded`, () => {
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
