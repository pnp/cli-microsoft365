import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN007003_CFG_S_api } from './FN007003_CFG_S_api.js';

describe('FN007003_CFG_S_api', () => {
  let findings: Finding[];
  let rule: FN007003_CFG_S_api;

  beforeEach(() => {
    findings = [];
    rule = new FN007003_CFG_S_api();
  });

  it('doesn\'t return notification if no serve.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});
