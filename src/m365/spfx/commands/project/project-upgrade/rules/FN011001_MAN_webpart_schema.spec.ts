import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN011001_MAN_webpart_schema } from './FN011001_MAN_webpart_schema';

describe('FN011001_MAN_webpart_schema', () => {
  let findings: Finding[];
  let rule: FN011001_MAN_webpart_schema;

  beforeEach(() => {
    findings = [];
    rule = new FN011001_MAN_webpart_schema('test-schema');
  })

  it('doesn\'t return notifications if no manifests collected', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});