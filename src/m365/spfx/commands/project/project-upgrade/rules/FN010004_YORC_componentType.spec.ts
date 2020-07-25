import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN010004_YORC_componentType } from './FN010004_YORC_componentType';

describe('FN010004_YORC_componentType', () => {
  let findings: Finding[];
  let rule: FN010004_YORC_componentType;

  beforeEach(() => {
    findings = [];
    rule = new FN010004_YORC_componentType();
  });

  it('has empty resolution', () => {
    assert.strictEqual(rule.resolution, '');
  });

  it('doesn\'t return notification if no .yo-rc.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if componentType is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          componentType: 'webpart'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('suggests setting componentType to webpart for a project with a web part', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        "$schema": 'https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json',
        componentType: 'WebPart',
        path: '/usr/tmp/src/webparts/helloWorld/HelloWorld.manifest.json'
      }],
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
        }
      }
    };
    rule.visit(project, findings);
    assert(findings[0].occurrences[0].resolution.indexOf('"componentType": "webpart"') > -1);
  });

  it('suggests setting componentType to extension for a project with an extension', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        "$schema": 'https://developer.microsoft.com/json-schemas/spfx/command-set-extension-manifest.schema.json',
        componentType: 'Extension',
        path: '/usr/tmp/src/extensions/helloWorld/HelloWorld.manifest.json'
      }],
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
        }
      }
    };
    rule.visit(project, findings);
    assert(findings[0].occurrences[0].resolution.indexOf('"componentType": "extension"') > -1);
  });

  it('suggests setting componentType to extension for a project with an extension and a web part', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        "$schema": 'https://developer.microsoft.com/json-schemas/spfx/command-set-extension-manifest.schema.json',
        componentType: 'Extension',
        path: '/usr/tmp/src/extensions/helloWorld/HelloWorld.manifest.json'
      },
      {
        "$schema": 'https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json',
        componentType: 'WebPart',
        path: '/usr/tmp/src/webparts/helloWorld/HelloWorld.manifest.json'
      }],
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
        }
      }
    };
    rule.visit(project, findings);
    assert(findings[0].occurrences[0].resolution.indexOf('"componentType": "extension"') > -1);
  });
});