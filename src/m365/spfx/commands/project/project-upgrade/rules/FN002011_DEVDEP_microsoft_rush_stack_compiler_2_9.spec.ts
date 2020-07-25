import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN002011_DEVDEP_microsoft_rush_stack_compiler_2_9 } from './FN002011_DEVDEP_microsoft_rush_stack_compiler_2_9';

describe('FN002011_DEVDEP_microsoft_rush_stack_compiler_2_9', () => {
  let findings: Finding[];
  let rule: FN002011_DEVDEP_microsoft_rush_stack_compiler_2_9;

  beforeEach(() => {
    findings = [];
    rule = new FN002011_DEVDEP_microsoft_rush_stack_compiler_2_9('1.0.0', true);
  })

  it('should not show finding when package is not present', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@microsoft/rush-stack-compiler-3.3': '0.1.6'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});