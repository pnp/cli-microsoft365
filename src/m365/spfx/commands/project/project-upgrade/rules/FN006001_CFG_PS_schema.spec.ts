import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
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

  it('returns notification if schema is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: 'old-schema',
        solution: {},
        source: JSON.stringify({
          $schema: 'old-schema',
          solution: {}
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 2, 'Incorrect line number');
  });
});