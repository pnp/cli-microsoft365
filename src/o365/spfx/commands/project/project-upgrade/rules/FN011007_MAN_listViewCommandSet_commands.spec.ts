import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project, Manifest } from '../model';
import { FN011007_MAN_listViewCommandSet_commands } from './FN011007_MAN_listViewCommandSet_commands';

describe('FN011007_MAN_listViewCommandSet_commands', () => {
  let findings: Finding[];
  let rule: FN011007_MAN_listViewCommandSet_commands;

  beforeEach(() => {
    findings = [];
    rule = new FN011007_MAN_listViewCommandSet_commands();
  });

  it('has empty resolution', () => {
    assert.equal(rule.resolution, '');
  });

  it('returns notification if the "commands" property is not in the manifest', () => {
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
    assert.equal(findings.length, 1);
  });

  it('doesn\'t return notification if the "commands" property is not in the manifest', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        path: '/usr/tmp',
        componentType: 'Extension',
        extensionType: 'ListViewCommandSet'
      } as Manifest]
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });

  it('exits if no manifest json', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: []
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });
});