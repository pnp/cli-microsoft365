import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
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

  it('doesn\'t return notification if package-solution.json is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});