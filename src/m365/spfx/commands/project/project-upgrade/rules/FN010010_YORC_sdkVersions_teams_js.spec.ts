import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';
import { FN010010_YORC_sdkVersions_teams_js } from './FN010010_YORC_sdkVersions_teams_js';

describe('FN010010_YORC_sdkVersions_teams_js', () => {
  let findings: Finding[];
  let rule: FN010010_YORC_sdkVersions_teams_js;

  beforeEach(() => {
    findings = [];
    rule = new FN010010_YORC_sdkVersions_teams_js('2.4.1');
  });

  it(`doesn't return notification if .yo-rc.json is not available`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns notification when @microsoft/generator-sharepoint is not set`, () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification when sdkVersions is not set`, () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification when @microsoft/teams-js is not set`, () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          sdkVersions: {}
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification when @microsoft/teams-js version doesn't match the required version`, () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          sdkVersions: {
            "@microsoft/teams-js": "2.4.0"
          }
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});