import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN011005_MAN_webpart_defaultGroup } from './FN011005_MAN_webpart_defaultGroup';

describe('FN011005_MAN_webpart_defaultGroup', () => {
  let findings: Finding[];
  let rule: FN011005_MAN_webpart_defaultGroup;

  beforeEach(() => {
    findings = [];
    rule = new FN011005_MAN_webpart_defaultGroup('Under Development', 'Other');
  })

  it('doesn\'t return notifications if no manifests collected', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if web part is in a custom group', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'test-schema',
        path: '/usr/tmp/manifest.json',
        componentType: 'WebPart',
        preconfiguredEntries: [{
          group: { default: 'Custom' }
        }]
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});