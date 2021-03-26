import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN011001_MAN_webpart_schema } from './FN011001_MAN_webpart_schema';

describe('FN011001_MAN_webpart_schema', () => {
  let findings: Finding[];
  let rule: FN011001_MAN_webpart_schema;

  beforeEach(() => {
    findings = [];
    rule = new FN011001_MAN_webpart_schema('test-schema');
  });

  it('doesn\'t return notifications if no manifests collected', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notifications if webpart manifest has incorrect schema', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'old-schema',
        componentType: 'WebPart',
        path: '/usr/tmp/manifest',
        source: JSON.stringify({
          $schema: 'old-schema',
          componentType: 'WebPart',
          path: '/usr/tmp/manifest'
        }, null, 2)
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 2, 'Incorrect line number');
  });
});