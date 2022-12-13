import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';
import { FN010009_YORC_sdkVersions_microsoft_graph_client } from './FN010009_YORC_sdkVersions_microsoft_graph_client';

describe('FN010009_YORC_sdkVersions_microsoft_graph_client', () => {
  let findings: Finding[];
  let rule: FN010009_YORC_sdkVersions_microsoft_graph_client;

  beforeEach(() => {
    findings = [];
    rule = new FN010009_YORC_sdkVersions_microsoft_graph_client('3.0.2');
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

  it(`returns notification when @microsoft/microsoft-graph-client is not set`, () => {
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

  it(`returns notification when @microsoft/microsoft-graph-client version doesn't match the required version`, () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          sdkVersions: {
            "@microsoft/microsoft-graph-client": "3.0.1"
          }
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});
