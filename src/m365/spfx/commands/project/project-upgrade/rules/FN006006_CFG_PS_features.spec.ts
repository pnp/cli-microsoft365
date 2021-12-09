import * as assert from 'assert';
import { PackageSolutionJson, Project } from '../../model';
import { Finding } from '../Finding';
import { FN006006_CFG_PS_features } from './FN006006_CFG_PS_features';

describe('FN006006_CFG_PS_features', () => {
  let findings: Finding[];
  let rule: FN006006_CFG_PS_features;

  beforeEach(() => {
    findings = [];
    rule = new FN006006_CFG_PS_features();
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
    assert.strictEqual(resolution.solution?.features?.[0].title, 'undefined Feature', 'Unexpected title');
    assert.strictEqual(resolution.solution?.features?.[0].description, 'The feature that activates elements of the undefined solution.', 'Unexpected description');
  });
});