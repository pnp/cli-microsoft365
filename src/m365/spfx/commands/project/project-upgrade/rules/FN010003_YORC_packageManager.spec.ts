import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN010003_YORC_packageManager } from './FN010003_YORC_packageManager';

describe('FN010003_YORC_packageManager', () => {
  let findings: Finding[];
  let rule: FN010003_YORC_packageManager;

  beforeEach(() => {
    findings = [];
    rule = new FN010003_YORC_packageManager('npm');
  });

  it('doesn\'t return notification if no .yo-rc.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if packageManager is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          packageManager: 'npm'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if packageManager is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          packageManager: 'foo'
        },
        source: JSON.stringify({
          "@microsoft/generator-sharepoint": {
            packageManager: 'foo'
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });
});