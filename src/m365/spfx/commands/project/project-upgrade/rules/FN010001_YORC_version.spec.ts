import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN010001_YORC_version } from './FN010001_YORC_version';

describe('FN010001_YORC_version', () => {
  let findings: Finding[];
  let rule: FN010001_YORC_version;

  beforeEach(() => {
    findings = [];
    rule = new FN010001_YORC_version('1.5.0');
  });

  it('doesn\'t return notification if no .yo-rc.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if version is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          version: '1.5.0'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});