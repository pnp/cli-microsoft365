import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN008001_CFG_TSL_schema } from './FN008001_CFG_TSL_schema';

describe('FN008001_CFG_TSL_schema', () => {
  let findings: Finding[];
  let rule: FN008001_CFG_TSL_schema;

  beforeEach(() => {
    findings = [];
    rule = new FN008001_CFG_TSL_schema('test-schema');
  });

  it('doesn\'t return notification if no tslint.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if schema is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsLintJson: {
        $schema: 'test-schema'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});