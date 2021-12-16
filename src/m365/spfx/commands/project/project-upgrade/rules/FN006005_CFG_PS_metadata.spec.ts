import * as assert from 'assert';
import { PackageSolutionJson, Project } from '../../model';
import { Finding } from '../Finding';
import { FN006005_CFG_PS_metadata } from './FN006005_CFG_PS_metadata';

describe('FN006005_CFG_PS_metadata', () => {
  let findings: Finding[];
  let rule: FN006005_CFG_PS_metadata;

  beforeEach(() => {
    findings = [];
    rule = new FN006005_CFG_PS_metadata();
  });

  it('doesn\'t return notification if package-solution.json is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('has a default empty resolution', () => {
    assert.strictEqual(rule.resolution, '');
  });

  it(`doesn't fail if package.json is not available`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: 'test-schema',
        solution: {},
        source: JSON.stringify({
          $schema: 'test-schema',
          solution: {}
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    const resolution: PackageSolutionJson = JSON.parse(findings[0].occurrences[0].resolution);
    assert.strictEqual(resolution.solution?.metadata?.shortDescription?.default, 'undefined description', 'Unexpected shortDescription');
    assert.strictEqual(resolution.solution?.metadata.longDescription?.default, 'undefined description', 'Unexpected longDescription');
  });
});