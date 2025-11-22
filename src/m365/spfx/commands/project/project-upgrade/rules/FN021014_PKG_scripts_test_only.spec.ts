import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN021014_PKG_scripts_test_only } from './FN021014_PKG_scripts_test_only.js';

describe('FN021014_PKG_scripts_test_only', () => {
  let findings: Finding[];
  let rule: FN021014_PKG_scripts_test_only;

  beforeEach(() => {
    findings = [];
    rule = new FN021014_PKG_scripts_test_only('heft run --only test --');
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

  it(`returns notification if scripts.test-only property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification if scripts.test-only property is different than expected`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {
          'test-only': 'test-only'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns correct node when scripts.test-only is set to a string`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {
          'test-only': 'test-only'
        },
        source: JSON.stringify({
          scripts: {
            'test-only': 'test-only'
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3);
  });
});
