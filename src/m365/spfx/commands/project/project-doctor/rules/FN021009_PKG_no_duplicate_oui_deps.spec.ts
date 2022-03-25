import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN021009_PKG_no_duplicate_oui_deps } from './FN021009_PKG_no_duplicate_oui_deps';

describe('FN021009_PKG_no_duplicate_oui_deps', () => {
  let findings: Finding[];
  let rule: FN021009_PKG_no_duplicate_oui_deps;

  beforeEach(() => {
    findings = [];
    rule = new FN021009_PKG_no_duplicate_oui_deps();
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

  it(`doesn't return notifications when project has no office ui fabric references`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns notification to uninstall @fluentui/react from devDependencies when installed as a devDependency`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          "office-ui-fabric-react": "^1.0.0"
        },
        devDependencies: {
          "@fluentui/react": "^1.0.0"
        }
      }
    };
    rule.visit(project, findings);
    assert(findings[0].occurrences[0].resolution.includes('uninstallDev @fluentui/react'));
  });
});