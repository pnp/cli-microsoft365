import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN010011_YORC_useGulp } from './FN010011_YORC_useGulp.js';

describe('FN010011_YORC_useGulp', () => {
  let findings: Finding[];
  let rule: FN010011_YORC_useGulp;

  beforeEach(() => {
    findings = [];
    rule = new FN010011_YORC_useGulp({ useGulp: false });
  });

  it(`doesn't return notification if .yo-rc.json is not available`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns notification when @microsoft/generator-sharepoint is not set`, () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('doesn\'t return notification if useGulp is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          useGulp: false
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if @microsoft/generator-sharepoint is not set', () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('returns notification if useGulp is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          useGulp: true
        },
        source: JSON.stringify({
          "@microsoft/generator-sharepoint": {
            useGulp: true
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });
});
