import * as assert from 'assert';
import * as fs from 'fs';
import { sinonUtil } from '../../../../../../utils';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN014006_CODE_launch_sourceMapPathOverrides } from './FN014006_CODE_launch_sourceMapPathOverrides';

describe('FN014006_CODE_launch_sourceMapPathOverrides', () => {
  let findings: Finding[];
  let rule: FN014006_CODE_launch_sourceMapPathOverrides;
  afterEach(() => {
    sinonUtil.restore(fs.existsSync);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN014006_CODE_launch_sourceMapPathOverrides('webpack:///.././src/*', '${webRoot}/src/*');
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

  it('returns notifications if the configuration does not contains the specified override', () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        launchJson: {
          version: '1.0',
          configurations: [{
            sourceMapPathOverrides: {}
          }],
          source: JSON.stringify({
            version: '1.0',
            configurations: [
              {
                sourceMapPathOverrides: {}
              }
            ]
          }, null, 2)
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 5, 'Incorrect line number');
  });
});