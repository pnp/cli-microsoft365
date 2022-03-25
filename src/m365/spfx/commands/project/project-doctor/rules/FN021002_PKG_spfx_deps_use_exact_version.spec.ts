import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN021002_PKG_spfx_deps_use_exact_version } from './FN021002_PKG_spfx_deps_use_exact_version';

describe('FN021002_PKG_spfx_deps_use_exact_version', () => {
  let findings: Finding[];
  let rule: FN021002_PKG_spfx_deps_use_exact_version;

  beforeEach(() => {
    findings = [];
    rule = new FN021002_PKG_spfx_deps_use_exact_version();
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
});