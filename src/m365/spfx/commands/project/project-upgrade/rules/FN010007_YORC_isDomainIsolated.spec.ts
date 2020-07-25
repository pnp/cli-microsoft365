import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
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
});