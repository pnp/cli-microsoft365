import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../model';
import { FN008002_CFG_TSL_no_unused_imports } from './FN008002_CFG_TSL_no_unused_imports';

describe('FN008002_CFG_TSL_no_unused_imports', () => {
  let findings: Finding[];
  let rule: FN008002_CFG_TSL_no_unused_imports;

  beforeEach(() => {
    findings = [];
    rule = new FN008002_CFG_TSL_no_unused_imports(false);
  })

  it('doesn\'t return notification if no-unused-imports is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsLintJson: {
        $schema: 'test-schema',
        "no-unused-imports": false
      }
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });

  it('does return notification if no-unused-imports is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsLintJson: {
        $schema: 'test-schema',
        "no-unused-imports": true
      }
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 1);
  });

  it('exits if no ts lint', () => {
    const project: any = {
      path: '/usr/tmp',
      tsLintJson: undefined
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });
});