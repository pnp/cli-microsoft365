import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../model';
import { FN006003_CFG_PS_skipFeatureDeployment } from './FN006003_CFG_PS_skipFeatureDeployment';

describe('FN006002_CFG_PS_skipFeatureDeployment', () => {
  let findings: Finding[];
  let rule: FN006003_CFG_PS_skipFeatureDeployment;

  beforeEach(() => {
    findings = [];
    rule = new FN006003_CFG_PS_skipFeatureDeployment('string');
  });

  it('has empty resolution', () => {
    assert.equal(rule.resolution, '');
  });

  it('doesn\'t return notification if package-solution.json not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });

  it('doesn\'t return notification if skipFeatureDeployment is not set', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: 'test-schema',
        solution: {}
      }
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });

  it('doesn\'t return notification if skipFeatureDeployment value type is correct', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: 'test-schema',
        solution: {
          skipFeatureDeployment: "true"
        }
      }
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });

  it('returns notification if skipFeatureDeployment value is boolean while string required', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: 'test-schema',
        solution: {
          skipFeatureDeployment: true
        }
      }
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 1);
  });
});