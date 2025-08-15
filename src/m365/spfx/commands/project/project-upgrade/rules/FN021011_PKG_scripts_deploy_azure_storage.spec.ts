import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN021011_PKG_scripts_deploy_azure_storage } from './FN021011_PKG_scripts_deploy_azure_storage.js';

describe('FN021011_PKG_scripts_deploy_azure_storage', () => {
  let findings: Finding[];
  let rule: FN021011_PKG_scripts_deploy_azure_storage;

  beforeEach(() => {
    findings = [];
    rule = new FN021011_PKG_scripts_deploy_azure_storage('heft deploy-azure-storage');
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

  it(`returns notification if scripts.deploy-azure-storage property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification if scripts.deploy-azure-storage property is different than expected`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {
          'deploy-azure-storage': 'deploy-azure-storage'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns correct node when deploy-azure-storage is set to a string`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {
          'deploy-azure-storage': 'deploy-azure-storage'
        },
        source: JSON.stringify({
          scripts: {
            'deploy-azure-storage': 'deploy-azure-storage'
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3);
  });
});
