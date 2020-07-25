import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
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

  it('doesn\'t return notification if tslint is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});