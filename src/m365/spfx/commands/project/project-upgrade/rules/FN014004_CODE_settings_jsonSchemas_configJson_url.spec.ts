import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN014004_CODE_settings_jsonSchemas_configJson_url } from './FN014004_CODE_settings_jsonSchemas_configJson_url';

describe('FN014004_CODE_settings_jsonSchemas_configJson_url', () => {
  let findings: Finding[];
  let rule: FN014004_CODE_settings_jsonSchemas_configJson_url;

  beforeEach(() => {
    findings = [];
    rule = new FN014004_CODE_settings_jsonSchemas_configJson_url('./node_modules/@microsoft/sp-build-core-tasks/lib/configJson/schemas/config-v1.schema.json');
  })

  it('doesn\'t return notification if config.json already has correct URL', () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        settingsJson: {
          "json.schemas": [
            {
              fileMatch: [
                '/config/config.json'
              ],
              url: './node_modules/@microsoft/sp-build-core-tasks/lib/configJson/schemas/config-v1.schema.json'
            }
          ]
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if vscode folder doesn\'t exist', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if .vscode/settings.json doesn\'t exist', () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if json.schemas in .vscode/settings.json not defined', () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        settingsJson: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if config.json JSON schema in .vscode/settings.json not defined', () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        settingsJson: {
          'json.schemas': [
            {
              fileMatch: [
                '/config/copy-assets.json'
              ],
              url: './node_modules/@microsoft/sp-build-core-tasks/lib/copyAssets/copy-assets.schema.json'
            }
          ]
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if config.json doesn\'t have correct URL', () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        settingsJson: {
          "json.schemas": [
            {
              fileMatch: [
                '/config/config.json'
              ],
              url: './node_modules/@microsoft/sp-build-web/lib/schemas/config.schema.json'
            }
          ]
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});