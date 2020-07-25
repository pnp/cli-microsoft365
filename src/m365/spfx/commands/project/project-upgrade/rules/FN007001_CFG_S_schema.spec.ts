import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN007001_CFG_S_schema } from './FN007001_CFG_S_schema';

describe('FN007001_CFG_S_schema', () => {
  let findings: Finding[];
  let rule: FN007001_CFG_S_schema;

  beforeEach(() => {
    findings = [];
    rule = new FN007001_CFG_S_schema('test-schema');
  });

  it('doesn\'t return notification if no serve.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if schema is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      serveJson: {
        $schema: 'test-schema'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});