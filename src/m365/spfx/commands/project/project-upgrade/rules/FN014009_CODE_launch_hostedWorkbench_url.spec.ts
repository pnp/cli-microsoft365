import * as assert from 'assert';
import * as fs from 'fs';
import { sinonUtil } from '../../../../../../utils/sinonUtil';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN014009_CODE_launch_hostedWorkbench_url } from './FN014009_CODE_launch_hostedWorkbench_url';

describe('FN014009_CODE_launch_hostedWorkbench_url', () => {
  let findings: Finding[];
  let rule: FN014009_CODE_launch_hostedWorkbench_url;
  afterEach(() => {
    sinonUtil.restore(fs.existsSync);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN014009_CODE_launch_hostedWorkbench_url('https://{tenantDomain}/_layouts/workbench.aspx');
  });

  it('doesn\'t return notifications if vscode folder doesn\'t exist', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if vscode launch file doesn\'t exist', () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if vscode launch file doesn\'t contain configurations', () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        launchJson: {
          version: '1.0'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if none of the configurations refers to hosted workbench', () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        launchJson: {
          version: '1.0',
          configurations: [{
          }]
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if the configuration already contains the specified URL', () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        launchJson: {
          version: '1.0',
          configurations: [{
            name: 'Hosted workbench',
            url: 'https://{tenantDomain}/_layouts/workbench.aspx'
          }]
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});
