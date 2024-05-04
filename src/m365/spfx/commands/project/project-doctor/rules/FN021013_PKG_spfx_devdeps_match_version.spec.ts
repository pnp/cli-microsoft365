import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN021013_PKG_spfx_devdeps_match_version } from './FN021013_PKG_spfx_devdeps_match_version.js';

describe('FN021013_PKG_spfx_devdeps_match_version', () => {
  let findings: Finding[];
  let rule: FN021013_PKG_spfx_devdeps_match_version;

  beforeEach(() => {
    findings = [];
    rule = new FN021013_PKG_spfx_devdeps_match_version('1.0.0');
  });

  it('returns correct ID', () => {
    assert.strictEqual(rule.id, 'FN021013');
  });

  it('returns empty title by default', () => {
    assert.strictEqual(rule.title, '');
  });

  it('returns empty description by default', () => {
    assert.strictEqual(rule.description, '');
  });

  it('returns correct severity', () => {
    assert.strictEqual(rule.severity, 'Required');
  });

  it('returns correct file', () => {
    assert.strictEqual(rule.file, './package.json');
  });

  it('returns correct resolution type', () => {
    assert.strictEqual(rule.resolutionType, 'cmd');
  });

  it(`doesn't return notifications when project version could not be determined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        devDependencies: {}
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

  it(`returns a notification if one of the SPFx dev deps doesn't match the version`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        devDependencies: {
          '@microsoft/sp-build-web': '0.9.0'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});
