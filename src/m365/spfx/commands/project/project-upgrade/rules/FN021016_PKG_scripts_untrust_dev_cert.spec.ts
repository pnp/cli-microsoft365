import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN021016_PKG_scripts_untrust_dev_cert } from './FN021016_PKG_scripts_untrust_dev_cert.js';

describe('FN021016_PKG_scripts_untrust_dev_cert', () => {
  let findings: Finding[];
  let rule: FN021016_PKG_scripts_untrust_dev_cert;

  beforeEach(() => {
    findings = [];
    rule = new FN021016_PKG_scripts_untrust_dev_cert('heft untrust-dev-cert');
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

  it(`returns notification if scripts.untrust-dev-cert property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification if scripts.untrust-dev-cert property is different than expected`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {
          'untrust-dev-cert': 'untrust-cert'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns correct node when scripts.untrust-dev-cert is set to a string`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {
          'untrust-dev-cert': 'untrust-cert'
        },
        source: JSON.stringify({
          scripts: {
            'untrust-dev-cert': 'untrust-cert'
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3);
  });
});
