import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN010002_YORC_isCreatingSolution } from './FN010002_YORC_isCreatingSolution';

describe('FN010002_YORC_isCreatingSolution', () => {
  let findings: Finding[];
  let rule: FN010002_YORC_isCreatingSolution;

  beforeEach(() => {
    findings = [];
    rule = new FN010002_YORC_isCreatingSolution(true);
  });

  it('doesn\'t return notification if no .yo-rc.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if isCreatingSolution is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          isCreatingSolution: true
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if isCreatingSolution is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          isCreatingSolution: false
        },
        source: JSON.stringify({
          "@microsoft/generator-sharepoint": {
            isCreatingSolution: false
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });
});