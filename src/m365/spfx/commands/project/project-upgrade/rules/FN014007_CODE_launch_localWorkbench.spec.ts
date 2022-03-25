import * as assert from 'assert';
import * as fs from 'fs';
import { sinonUtil } from '../../../../../../utils';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN014007_CODE_launch_localWorkbench } from './FN014007_CODE_launch_localWorkbench';

describe('FN014007_CODE_launch_localWorkbench', () => {
  let findings: Finding[];
  let rule: FN014007_CODE_launch_localWorkbench;
  afterEach(() => {
    sinonUtil.restore(fs.existsSync);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN014007_CODE_launch_localWorkbench();
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

  it('doesn\'t return notifications if none of the configurations contains sourceMapPathOverrides', () => {
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

  it('doesn\'t return notifications if the configuration already contains the specified override', () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        launchJson: {
          version: '1.0',
          configurations: [{
            sourceMapPathOverrides: {
              'webpack:///.././src/*': '${webRoot}/src/*'
            }
          }]
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});