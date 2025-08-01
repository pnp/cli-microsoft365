import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN021010_PKG_scripts_package_solution } from './FN021010_PKG_scripts_package_solution.js';

describe('FN021010_PKG_scripts_package_solution', () => {
  let findings: Finding[];
  let rule: FN021010_PKG_scripts_package_solution;

  beforeEach(() => {
    findings = [];
    rule = new FN021010_PKG_scripts_package_solution('heft package-solution');
  });

  it(`doesn't return notification if package.json is not available`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns notification if scripts property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification if scripts.package-solution property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification if scripts.package-solution property is different than expected`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {
          'package-solution': 'package-solution'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns correct node when package-solution is set to a string`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {
          'package-solution': 'package-solution'
        },
        source: JSON.stringify({
          scripts: {
            'package-solution': 'package-solution'
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3);
  });
});
