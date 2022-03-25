import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN011002_MAN_applicationCustomizer_schema } from './FN011002_MAN_applicationCustomizer_schema';

describe('FN011002_MAN_applicationCustomizer_schema', () => {
  let findings: Finding[];
  let rule: FN011002_MAN_applicationCustomizer_schema;

  beforeEach(() => {
    findings = [];
    rule = new FN011002_MAN_applicationCustomizer_schema('test-schema');
  });

  it('doesn\'t return notifications if no manifests collected', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notifications if application customizer manifest has incorrect schema', () => {
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'old-schema',
        componentType: 'Extension',
        extensionType: 'ApplicationCustomizer',
        path: '/usr/tmp/manifest',
        source: JSON.stringify({
          $schema: 'old-schema',
          componentType: 'Extension',
          extensionType: 'ApplicationCustomizer',
          path: '/usr/tmp/manifest'
        }, null, 2)
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 2, 'Incorrect line number');
  });
});