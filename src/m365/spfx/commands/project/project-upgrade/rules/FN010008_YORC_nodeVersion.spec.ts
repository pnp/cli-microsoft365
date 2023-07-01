import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN010008_YORC_nodeVersion } from './FN010008_YORC_nodeVersion.js';

describe('FN010008_YORC_nodeVersion', () => {
  let findings: Finding[];
  let rule: FN010008_YORC_nodeVersion;

  beforeEach(() => {
    findings = [];
    rule = new FN010008_YORC_nodeVersion();
  });

  it(`doesn't return notification if .yo-rc.json is not available`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns notification when @microsoft/generator-sharepoint is not set`, () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});
