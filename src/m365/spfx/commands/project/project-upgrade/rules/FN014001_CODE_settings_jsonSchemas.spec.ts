import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN014001_CODE_settings_jsonSchemas } from './FN014001_CODE_settings_jsonSchemas';

describe('FN014001_CODE_settings_jsonSchemas', () => {
  let findings: Finding[];
  let rule: FN014001_CODE_settings_jsonSchemas;

  beforeEach(() => {
    findings = [];
    rule = new FN014001_CODE_settings_jsonSchemas(false);
  })

  it('doesn\'t return notification if json.schemas should be removed and is not present', () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        settingsJson: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if json.schemas should be added and is already present', () => {
    rule = new FN014001_CODE_settings_jsonSchemas(true);
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        settingsJson: {
          "json.schemas": []
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if json.schemas should be added but is not present', () => {
    rule = new FN014001_CODE_settings_jsonSchemas(true);
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        settingsJson: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('doesn\'t return notification if .vscode/settings.json is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});