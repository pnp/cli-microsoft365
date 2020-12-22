import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN006002_CFG_PS_includeClientSideAssets } from './FN006002_CFG_PS_includeClientSideAssets';

describe('FN006002_CFG_PS_includeClientSideAssets', () => {
  let findings: Finding[];
  let rule: FN006002_CFG_PS_includeClientSideAssets;

  beforeEach(() => {
    findings = [];
    rule = new FN006002_CFG_PS_includeClientSideAssets(true);
  })

  it('doesn\'t return notification if includeClientSideAssets is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: 'test-schema',
        solution: {
          includeClientSideAssets: true
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if includeClientSideAssets is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: 'test-schema',
        solution: {
          includeClientSideAssets: false
        },
        source: JSON.stringify({
          $schema: 'test-schema',
          solution: {
            includeClientSideAssets: false
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 4, 'Incorrect line number');
  });

  it('doesn\'t return notification if package-solution.json is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});