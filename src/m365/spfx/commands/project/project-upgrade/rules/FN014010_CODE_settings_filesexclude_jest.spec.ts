import assert from 'assert';
import fs from 'fs';
import { sinonUtil } from '../../../../../../utils/sinonUtil.js';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding.js';
import { FN014010_CODE_settings_filesexclude_jest } from './FN014010_CODE_settings_filesexclude_jest.js';

describe('FN014010_CODE_settings_filesexclude_jest', () => {
  let findings: Finding[];
  let rule: FN014010_CODE_settings_filesexclude_jest;
  afterEach(() => {
    sinonUtil.restore(fs.existsSync);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN014010_CODE_settings_filesexclude_jest();
  });

  it(`doesn't return notifications if vscode folder doesn't exist`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't return notifications if vscode settings file doesn't exist`, () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't return notifications if vscode settings file doesn't contain file exclusion rules`, () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        settingsJson: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't return notifications if vscode settings file already excludes jest output files`, () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        settingsJson: {
          "files.exclude": {
            "**/jest-output": true
          }
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});