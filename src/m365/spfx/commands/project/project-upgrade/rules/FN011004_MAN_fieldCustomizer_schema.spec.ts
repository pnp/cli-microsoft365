import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN011004_MAN_fieldCustomizer_schema } from './FN011004_MAN_fieldCustomizer_schema';

describe('FN011004_MAN_fieldCustomizer_schema', () => {
  let findings: Finding[];
  let rule: FN011004_MAN_fieldCustomizer_schema;

  beforeEach(() => {
    findings = [];
    rule = new FN011004_MAN_fieldCustomizer_schema('test-schema');
  })

  it('doesn\'t return notifications if no manifests collected', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});