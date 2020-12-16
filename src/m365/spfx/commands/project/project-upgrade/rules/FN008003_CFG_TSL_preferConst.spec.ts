import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN008003_CFG_TSL_preferConst } from './FN008003_CFG_TSL_preferConst';

describe('FN010201_CFG_TSL_preferConst', () => {
  let findings: Finding[];
  let rule: FN008003_CFG_TSL_preferConst;

  beforeEach(() => {
    findings = [];
    rule = new FN008003_CFG_TSL_preferConst();
  });

  it('doesn\'t return notification if preferConst is absent', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsLintJson: {
        $schema: "https://schema.org/dummy.json",
        lintConfig: {
          rules: {
            "class-name": false,
          }
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if preferConst is present', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsLintJson: {
        $schema: "https://schema.org/dummy.json",
        lintConfig: {
          rules: {
            "prefer-const": true,
          }
        },
        source: JSON.stringify({
          $schema: "https://schema.org/dummy.json",
          lintConfig: {
            rules: {
              "prefer-const": true,
            }
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 5, 'Incorrect line number');
  });

  it('doesn\'t return notification if tslint is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});