import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
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
});