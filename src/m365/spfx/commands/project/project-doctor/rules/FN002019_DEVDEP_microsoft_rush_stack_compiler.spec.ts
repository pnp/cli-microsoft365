import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN002019_DEVDEP_microsoft_rush_stack_compiler } from './FN002019_DEVDEP_microsoft_rush_stack_compiler';

describe('FN002019_DEVDEP_microsoft_rush_stack_compiler', () => {
  let findings: Finding[];
  let rule: FN002019_DEVDEP_microsoft_rush_stack_compiler;

  beforeEach(() => {
    findings = [];
    rule = new FN002019_DEVDEP_microsoft_rush_stack_compiler(['3.9']);
  });

  it('returns empty description by default', () => {
    assert.strictEqual(rule.description, '');
  });

  it('returns empty resolution by default', () => {
    assert.strictEqual(rule.resolution, '');
  });

  it(`doesn't return notifications when package.json was not collected`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns correct description when one unsupported version of rushstack found`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          "@microsoft/rush-stack-compiler-3.7": "0.6.48",
          "@microsoft/rush-stack-compiler-3.9": "0.4.48"
        }
      }
    };
    rule.visit(project, findings);
    assert(findings[0].description.includes('Uninstall unsupported version '));
  });
});