import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN010006_YORC_framework } from './FN010006_YORC_framework';

describe('FN010006_YORC_framework', () => {
  let findings: Finding[];
  let rule: FN010006_YORC_framework;

  beforeEach(() => {
    findings = [];
  });

  it('doesn\'t return notification if no .yo-rc.json found', () => {
    rule = new FN010006_YORC_framework('react', true);
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if framework is already up-to-date', () => {
    rule = new FN010006_YORC_framework('react', true);
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          framework: 'react'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if framework not found and should be removed', () => {
    rule = new FN010006_YORC_framework('', false);
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if framework not found while it should be added', () => {
    rule = new FN010006_YORC_framework('react', true);
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
        },
        source: JSON.stringify({
          "@microsoft/generator-sharepoint": {
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 2, 'Incorrect line number');
  });

  it('returns notification if framework found while it should be removed', () => {
    rule = new FN010006_YORC_framework('react', false);
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          framework: 'react'
        },
        source: JSON.stringify({
          "@microsoft/generator-sharepoint": {
            framework: 'react'
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });
});