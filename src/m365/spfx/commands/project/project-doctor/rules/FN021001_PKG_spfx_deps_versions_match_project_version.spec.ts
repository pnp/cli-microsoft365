import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN021001_PKG_spfx_deps_versions_match_project_version } from './FN021001_PKG_spfx_deps_versions_match_project_version';

describe('FN021001_PKG_spfx_deps_versions_match_project_version', () => {
  let findings: Finding[];
  let rule: FN021001_PKG_spfx_deps_versions_match_project_version;

  beforeEach(() => {
    findings = [];
    rule = new FN021001_PKG_spfx_deps_versions_match_project_version();
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