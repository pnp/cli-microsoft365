import assert from 'assert';
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";
import { FN023006_GITIGNORE_jestoutput } from './FN023006_GITIGNORE_jestoutput.js';

describe('FN023006_GITIGNORE_jestoutput', () => {
  let findings: Finding[];
  let rule: FN023006_GITIGNORE_jestoutput;

  beforeEach(() => {
    findings = [];
    rule = new FN023006_GITIGNORE_jestoutput();
  });

  it(`doesn't return notification if the jest-output folder is already excluded`, () => {
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
