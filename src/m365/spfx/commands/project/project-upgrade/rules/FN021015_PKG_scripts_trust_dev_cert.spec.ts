import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN021015_PKG_scripts_trust_dev_cert } from './FN021015_PKG_scripts_trust_dev_cert.js';

describe('FN021015_PKG_scripts_trust_dev_cert', () => {
  let findings: Finding[];
  let rule: FN021015_PKG_scripts_trust_dev_cert;

  beforeEach(() => {
    findings = [];
    rule = new FN021015_PKG_scripts_trust_dev_cert('heft trust-dev-cert');
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

  it(`returns notification if scripts.trust-dev-cert property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification if scripts.trust-dev-cert property is different than expected`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {
          'trust-dev-cert': 'trust-cert'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns correct node when scripts.trust-dev-cert is set to a string`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {
          'trust-dev-cert': 'trust-cert'
        },
        source: JSON.stringify({
          scripts: {
            'trust-dev-cert': 'trust-cert'
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3);
  });
});
