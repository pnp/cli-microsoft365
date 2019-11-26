import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN010004_YORC_componentType } from './FN010004_YORC_componentType';

describe('FN010004_YORC_componentType', () => {
  let findings: Finding[];
  let rule: FN010004_YORC_componentType;

  beforeEach(() => {
    findings = [];
    rule = new FN010004_YORC_componentType('webpart');
  });

  it('doesn\'t return notification if no .yo-rc.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });

  it('doesn\'t return notification if componentType is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          componentType: 'webpart'
        }
      }
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });
});