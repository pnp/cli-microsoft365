import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN021007_PKG_only_one_rush_stack_compiler_installed } from './FN021007_PKG_only_one_rush_stack_compiler_installed';

describe('FN021007_PKG_only_one_rush_stack_compiler_installed', () => {
  let findings: Finding[];
  let rule: FN021007_PKG_only_one_rush_stack_compiler_installed;

  beforeEach(() => {
    findings = [];
    rule = new FN021007_PKG_only_one_rush_stack_compiler_installed();
  });

  it('returns empty title by default', () => {
    assert.strictEqual(rule.title, '');
  });

  it('returns empty description by default', () => {
    assert.strictEqual(rule.description, '');
  });

  it(`doesn't return notifications when project version could not be determined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't return notifications when package.json was not collected`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`uses first matched rushstack when tsconfig.json not found`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {},
        devDependencies: {
          "@microsoft/rush-stack-compiler-3.2": "0.10.48",
          "@microsoft/rush-stack-compiler-3.9": "0.4.47"
        }
      }
    };
    rule.visit(project, findings);
    assert(findings[0].occurrences[0].resolution.includes('@microsoft/rush-stack-compiler-3.9'));
  });

  it(`uses first matched rushstack when no rushstack reference found in tsconfig.json`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {},
        devDependencies: {
          "@microsoft/rush-stack-compiler-3.2": "0.10.48",
          "@microsoft/rush-stack-compiler-3.9": "0.4.47"
        }
      },
      tsConfigJson: {
        extends: 'tsconfig.json'
      }
    };
    rule.visit(project, findings);
    assert(findings[0].occurrences[0].resolution.includes('@microsoft/rush-stack-compiler-3.9'));
  });
});