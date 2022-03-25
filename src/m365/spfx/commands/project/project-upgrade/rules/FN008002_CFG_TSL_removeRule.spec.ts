import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN008002_CFG_TSL_removeRule } from './FN008002_CFG_TSL_removeRule';

describe('FN008002_CFG_TSL_removeRule', () => {
  let findings: Finding[];
  let rule: FN008002_CFG_TSL_removeRule;

  beforeEach(() => {
    findings = [];
    rule = new FN008002_CFG_TSL_removeRule('no-unused-imports');
  });

  it('doesn\'t return notification if no tslint.json', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if no lintConfig in tslint.json', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsLintJson: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if no rules in tslint.json', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsLintJson: {
        lintConfig: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if rule not found', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsLintJson: {
        lintConfig: {
          rules: {}
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if rule found', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsLintJson: {
        lintConfig: {
          rules: {
            "no-unused-imports": false
          }
        },
        source: JSON.stringify({
          lintConfig: {
            rules: {
              "no-unused-imports": false
            }
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 4, 'Incorrect line number');
  });
});