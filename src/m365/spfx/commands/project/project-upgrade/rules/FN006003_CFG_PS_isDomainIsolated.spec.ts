import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN006003_CFG_PS_isDomainIsolated } from './FN006003_CFG_PS_isDomainIsolated';

describe('FN006003_CFG_PS_isDomainIsolated', () => {
  let findings: Finding[];
  let rule: FN006003_CFG_PS_isDomainIsolated;

  beforeEach(() => {
    findings = [];
    rule = new FN006003_CFG_PS_isDomainIsolated(false);
  })

  it('doesn\'t return notification if isDomainIsolated is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: 'test-schema',
        solution: {
          isDomainIsolated: false
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if package-solution.json is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});