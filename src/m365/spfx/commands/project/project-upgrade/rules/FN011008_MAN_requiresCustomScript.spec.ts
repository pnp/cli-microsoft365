import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN011008_MAN_requiresCustomScript } from './FN011008_MAN_requiresCustomScript';

describe('FN011008_MAN_requiresCustomScript', () => {
  let findings: Finding[];
  let rule: FN011008_MAN_requiresCustomScript;

  beforeEach(() => {
    findings = [];
    rule = new FN011008_MAN_requiresCustomScript();
  });

  it('has empty resolution', () => {
    assert.strictEqual(rule.resolution, '');
  });

  it('doesn\'t return notifications if no manifests collected', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if safeWithCustomScriptDisabled not defined', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        path: '/usr/tmp/manifest.json',
        $schema: 'test-schema',
        componentType: 'WebPart'
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});