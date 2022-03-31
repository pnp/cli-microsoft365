import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';
import { FN019002_TSL_extends } from './FN019002_TSL_extends';

describe('FN019002_TSL_extends', () => {
  let findings: Finding[];
  let rule: FN019002_TSL_extends;

  beforeEach(() => {
    findings = [];
    rule = new FN019002_TSL_extends('@microsoft/sp-tslint-rules/base-tslint.json');
  });

  it('doesn\'t return notification if extends is correctly configured', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsLintJsonRoot: {
        extends: '@microsoft/sp-tslint-rules/base-tslint.json'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if extends is not correctly configured', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsLintJsonRoot: {
        extends: '@microsoft/sp-tslint-rules/old-tslint.json',
        source: JSON.stringify({
          extends: '@microsoft/sp-tslint-rules/old-tslint.json'
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 2, 'Incorrect line number');
  });
});