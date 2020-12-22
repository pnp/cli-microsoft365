import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
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

  it('returns notifications if safeWithCustomScriptDisabled is defined', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        path: '/usr/tmp/manifest.json',
        $schema: 'test-schema',
        componentType: 'WebPart',
        safeWithCustomScriptDisabled: true,
        source: JSON.stringify({
          path: '/usr/tmp/manifest.json',
          $schema: 'test-schema',
          componentType: 'WebPart',
          safeWithCustomScriptDisabled: true
        }, null, 2)
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 5, 'Incorrect line number');
  });
});