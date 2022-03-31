import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN021005_PKG_types_installed_as_devdep } from './FN021005_PKG_types_installed_as_devdep';

describe('FN021005_PKG_types_installed_as_devdep', () => {
  let findings: Finding[];
  let rule: FN021005_PKG_types_installed_as_devdep;

  beforeEach(() => {
    findings = [];
    rule = new FN021005_PKG_types_installed_as_devdep();
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

  it(`doesn't return notifications when project has no dependencies`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
      } as any
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});