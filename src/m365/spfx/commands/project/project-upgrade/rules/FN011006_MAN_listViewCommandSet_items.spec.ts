import * as assert from 'assert';
import { CommandSetManifest, Manifest, Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN011006_MAN_listViewCommandSet_items } from './FN011006_MAN_listViewCommandSet_items';

describe('FN011006_MAN_listViewCommandSet_items', () => {
  let findings: Finding[];
  let rule: FN011006_MAN_listViewCommandSet_items;

  beforeEach(() => {
    findings = [];
    rule = new FN011006_MAN_listViewCommandSet_items();
  });

  it('has empty resolution', () => {
    assert.strictEqual(rule.resolution, '');
  });

  it('doesn\'t return notifications if items property is in the manifest', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        path: '/usr/tmp',
        componentType: 'Extension',
        extensionType: 'ListViewCommandSet',
        items: {}
      } as CommandSetManifest]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if items property is not in the manifest', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        path: '/usr/tmp',
        componentType: 'Extension',
        extensionType: 'ListViewCommandSet',
        source: JSON.stringify({
          path: '/usr/tmp',
          componentType: 'Extension',
          extensionType: 'ListViewCommandSet'
        }, null, 2)
      } as Manifest]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 1, 'Incorrect line number');
  });

  it('returns notification if commands has to be converted to items', () => {
    const project: any = {
      path: '/usr/tmp',
      manifests: [{
        path: '/usr/tmp',
        componentType: 'Extension',
        extensionType: 'ListViewCommandSet',
        commands: {
          "COMMAND_1": {
            "title": "Command One",
            "iconImageUrl": "icons/request.png"
          },
          "COMMAND_2": {
            "title": "Command Two",
            "iconImageUrl": "icons/cancel.png"
          }
        },
        source: JSON.stringify({
          path: '/usr/tmp',
          componentType: 'Extension',
          extensionType: 'ListViewCommandSet',
          commands: {
            "COMMAND_1": {
              "title": "Command One",
              "iconImageUrl": "icons/request.png"
            },
            "COMMAND_2": {
              "title": "Command Two",
              "iconImageUrl": "icons/cancel.png"
            }
          }
        }, null, 2)
      }]
    };
    
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 5, 'Incorrect line number');
  });

  it('should correctly convert commands schema to items schema', () => {
    const project: any = {
      path: '/usr/tmp',
      manifests: [{
        path: '/usr/tmp',
        componentType: 'Extension',
        extensionType: 'ListViewCommandSet',
        commands: {
          "COMMAND_1": {
            "title": "Command One",
            "iconImageUrl": "icons/request.png"
          },
          "COMMAND_2": {
            "title": "Command Two",
            "iconImageUrl": "icons/cancel.png"
          }
        }
      }]
    };
    rule.visit(project, findings);

    const resolution: any = JSON.parse(findings[0].occurrences[0].resolution);
    const command1: any = resolution.items.COMMAND_1;
    
    assert.notStrictEqual(command1, undefined);
    assert.strictEqual(command1.title.default, 'Command One');
    assert.strictEqual(command1.iconImageUrl, 'icons/request.png');
    assert.strictEqual(command1.type, 'command');
  });

  it('exits if no manifest json', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: []
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});