import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN006001_CFG_PS_schema } from './FN006001_CFG_PS_schema';

describe('FN006001_CFG_PS_schema', () => {
  let findings: Finding[];
  let rule: FN006001_CFG_PS_schema;

  beforeEach(() => {
    findings = [];
    rule = new FN006001_CFG_PS_schema('test-schema');
  });

  it('doesn\'t return notification if no package-solution.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if schema is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: 'test-schema',
        solution: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});