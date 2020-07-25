import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project, CommandSetManifest } from '../../model';
import { FN011007_MAN_listViewCommandSet_removeCommands } from './FN011007_MAN_listViewCommandSet_removeCommands';

describe('FN011007_MAN_listViewCommandSet_removeCommands', () => {
  let findings: Finding[];
  let rule: FN011007_MAN_listViewCommandSet_removeCommands;

  beforeEach(() => {
    findings = [];
    rule = new FN011007_MAN_listViewCommandSet_removeCommands();
  });

  it('has empty resolution', () => {
    assert.strictEqual(rule.resolution, '');
  });

  it('doesn\'t return notifications if no manifests found', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: []
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if no extension manifests found', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        componentType: 'WebPart',
        path: '/tmp'
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if no ListViewCommandSet manifests found', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        componentType: 'Extension',
        extensionType: 'FieldCustomizer',
        path: '/tmp'
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if commands property is not in the manifest', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        path: '/usr/tmp',
        componentType: 'Extension',
        extensionType: 'ListViewCommandSet'
      } as CommandSetManifest]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if commands property is in the manifest', () => {
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
    assert.strictEqual(findings.length, 1);
  });
});