import * as assert from 'assert';
import * as fs from 'fs';
import { sinonUtil } from '../../../../../../utils';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN014008_CODE_launch_hostedWorkbench_type } from './FN014008_CODE_launch_hostedWorkbench_type';

describe('FN014008_CODE_launch_hostedWorkbench_type', () => {
  let findings: Finding[];
  let rule: FN014008_CODE_launch_hostedWorkbench_type;
  afterEach(() => {
    sinonUtil.restore(fs.existsSync);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN014008_CODE_launch_hostedWorkbench_type('pwa-chrome');
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

  it('doesn\'t return notifications if the configuration already contains the specified type', () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        launchJson: {
          version: '1.0',
          configurations: [{
            name: 'Hosted workbench',
            type: 'pwa-chrome'
          }]
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});