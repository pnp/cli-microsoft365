import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN011010_MAN_webpart_version } from './FN011010_MAN_webpart_version';

describe('FN011010_MAN_webpart_version', () => {
  let findings: Finding[];
  let rule: FN011010_MAN_webpart_version;

  beforeEach(() => {
    findings = [];
    rule = new FN011010_MAN_webpart_version();
  });

  it('doesn\'t return notifications if no manifests collected', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if version already set to *', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        path: '/usr/tmp/manifest.json',
        $schema: 'test-schema',
        componentType: 'WebPart',
        version: '*'
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if version not set to *', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        path: '/usr/tmp/manifest.json',
        $schema: 'test-schema',
        componentType: 'WebPart',
        version: '0.0.1',
        source: JSON.stringify({
          path: '/usr/tmp/manifest.json',
          $schema: 'test-schema',
          componentType: 'WebPart',
          version: '0.0.1'
        }, null, 2)
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 5, 'Incorrect line number');
  });
});