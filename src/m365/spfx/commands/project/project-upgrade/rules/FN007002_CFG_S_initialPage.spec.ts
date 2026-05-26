import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN007002_CFG_S_initialPage } from './FN007002_CFG_S_initialPage.js';

describe('FN007002_CFG_S_initialPage', () => {
  let findings: Finding[];
  let rule: FN007002_CFG_S_initialPage;

  beforeEach(() => {
    findings = [];
    rule = new FN007002_CFG_S_initialPage({ initialPage: 'https://enter-your-SharePoint-site/_layouts/workbench.aspx' });
  });

  it('doesn\'t return notification if no serve.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});
