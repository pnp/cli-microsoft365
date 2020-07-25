import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN011009_MAN_webpart_safeWithCustomScriptDisabled } from './FN011009_MAN_webpart_safeWithCustomScriptDisabled';

describe('FN011009_MAN_webpart_safeWithCustomScriptDisabled', () => {
  let findings: Finding[];
  let rule: FN011009_MAN_webpart_safeWithCustomScriptDisabled;

  beforeEach(() => {
    findings = [];
  });

  it('doesn\'t return notifications if no manifests collected', () => {
    rule = new FN011009_MAN_webpart_safeWithCustomScriptDisabled(true);
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if no safeWithCustomScriptDisabled found while it should be removed', () => {
    rule = new FN011009_MAN_webpart_safeWithCustomScriptDisabled(false);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [
        {
          path: '/usr/tmp/manifest.json',
          $schema: 'test-schema',
          componentType: 'WebPart'
        }
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if safeWithCustomScriptDisabled found while it should be added', () => {
    rule = new FN011009_MAN_webpart_safeWithCustomScriptDisabled(true);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [
        {
          path: '/usr/tmp/manifest.json',
          $schema: 'test-schema',
          componentType: 'WebPart',
          safeWithCustomScriptDisabled: true
        }
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});