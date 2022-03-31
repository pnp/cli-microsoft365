import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN021010_PKG_gulp_installed_as_devdep } from './FN021010_PKG_gulp_installed_as_devdep';

describe('FN021010_PKG_gulp_installed_as_devdep', () => {
  let findings: Finding[];
  let rule: FN021010_PKG_gulp_installed_as_devdep;

  beforeEach(() => {
    findings = [];
    rule = new FN021010_PKG_gulp_installed_as_devdep();
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
});