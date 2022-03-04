import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN021012_PKG_no_duplicate_pnpjs_deps } from './FN021012_PKG_no_duplicate_pnpjs_deps';

describe('FN021012_PKG_no_duplicate_pnpjs_deps', () => {
  let findings: Finding[];
  let rule: FN021012_PKG_no_duplicate_pnpjs_deps;

  beforeEach(() => {
    findings = [];
    rule = new FN021012_PKG_no_duplicate_pnpjs_deps();
  });

  it('returns empty resolution by default', () => {
    assert.strictEqual(rule.resolution, '');
  });

  it(`doesn't return notifications when project has no package.json`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't return notifications when project has no dependencies`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't return notifications when project has no PnPjs references`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns notification to uninstall sp-pnp-js from devDependencies when installed as a devDependency`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          "@pnp/sp": "^3.0.0"
        },
        devDependencies: {
          "sp-pnp-js": "^3.0.0"
        }
      }
    };
    rule.visit(project, findings);
    assert(findings[0].occurrences[0].resolution.includes('uninstallDev sp-pnp-js'));
  });
});