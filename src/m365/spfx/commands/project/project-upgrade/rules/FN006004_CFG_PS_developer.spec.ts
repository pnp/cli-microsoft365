import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN006004_CFG_PS_developer } from './FN006004_CFG_PS_developer';

describe('FN006004_CFG_PS_developer', () => {
  let findings: Finding[];
  let rule: FN006004_CFG_PS_developer;

  beforeEach(() => {
    findings = [];
    rule = new FN006004_CFG_PS_developer();
  });

  it('doesn\'t return notification if package-solution.json is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if developer section is not set', () => {
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
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });

  it('returns versioned mpnId when version specified', () => {
    rule = new FN006004_CFG_PS_developer('1.13.0-beta.20');
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
    assert(findings[0].occurrences[0].resolution.indexOf('1.13.0-beta.20') > -1);
  });
});