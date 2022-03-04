import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN021008_PKG_no_duplicate_deps } from './FN021008_PKG_no_duplicate_deps';

describe('FN021008_PKG_no_duplicate_deps', () => {
  let findings: Finding[];
  let rule: FN021008_PKG_no_duplicate_deps;

  beforeEach(() => {
    findings = [];
    rule = new FN021008_PKG_no_duplicate_deps();
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
        dependencies: {},
        devDependencies: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't return notifications when package.json was not collected`, () => {
    const project: Project = {
      path: '/usr/tmp',
      version: '1.14.0'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't return notifications when package.json has no devDependencies`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {}
      },
      version: '1.14.0'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});