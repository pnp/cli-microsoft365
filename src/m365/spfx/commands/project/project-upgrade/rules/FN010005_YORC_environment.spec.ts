import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN010005_YORC_environment } from './FN010005_YORC_environment';

describe('FN010005_YORC_environment', () => {
  let findings: Finding[];
  let rule: FN010005_YORC_environment;

  beforeEach(() => {
    findings = [];
    rule = new FN010005_YORC_environment('spo');
  });

  it('doesn\'t return notification if no .yo-rc.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if environment is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          environment: 'spo'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});