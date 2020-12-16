import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN010007_YORC_isDomainIsolated } from './FN010007_YORC_isDomainIsolated';

describe('FN010007_YORC_isDomainIsolated', () => {
  let findings: Finding[];
  let rule: FN010007_YORC_isDomainIsolated;

  beforeEach(() => {
    findings = [];
    rule = new FN010007_YORC_isDomainIsolated(false);
  });

  it('doesn\'t return notification if no .yo-rc.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if isDomainIsolated is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          isDomainIsolated: false
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if isDomainIsolated is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          isDomainIsolated: true
        },
        source: JSON.stringify({
          "@microsoft/generator-sharepoint": {
            isDomainIsolated: true
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });
});