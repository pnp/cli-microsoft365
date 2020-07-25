import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN011003_MAN_listViewCommandSet_schema } from './FN011003_MAN_listViewCommandSet_schema';

describe('FN011003_MAN_listViewCommandSet_schema', () => {
  let findings: Finding[];
  let rule: FN011003_MAN_listViewCommandSet_schema;

  beforeEach(() => {
    findings = [];
    rule = new FN011003_MAN_listViewCommandSet_schema('test-schema');
  })

  it('doesn\'t return notifications if no manifests collected', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});