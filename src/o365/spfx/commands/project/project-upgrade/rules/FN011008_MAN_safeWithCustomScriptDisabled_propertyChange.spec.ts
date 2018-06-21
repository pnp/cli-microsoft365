import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../model';
import { FN011008_MAN_safeWithCustomScriptDisabled_propertyChange } from './FN011008_MAN_safeWithCustomScriptDisabled_propertyChange';

describe('FN011005_MAN_safeWithCustomScriptDisabled_propertyChange', () => {
  let findings: Finding[];
  let rule: FN011008_MAN_safeWithCustomScriptDisabled_propertyChange;

  beforeEach(() => {
    findings = [];
    rule = new FN011008_MAN_safeWithCustomScriptDisabled_propertyChange();
  });

  it('has empty resolution', () => {
    assert.equal(rule.resolution, '');
  });

  it('doesn\'t return notifications if no manifests collected', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
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
    assert.equal(findings.length, 0);
  });
});