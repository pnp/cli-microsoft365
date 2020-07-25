import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN011002_MAN_applicationCustomizer_schema } from './FN011002_MAN_applicationCustomizer_schema';

describe('FN011002_MAN_applicationCustomizer_schema', () => {
  let findings: Finding[];
  let rule: FN011002_MAN_applicationCustomizer_schema;

  beforeEach(() => {
    findings = [];
    rule = new FN011002_MAN_applicationCustomizer_schema('test-schema');
  })

  it('doesn\'t return notifications if no manifests collected', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});