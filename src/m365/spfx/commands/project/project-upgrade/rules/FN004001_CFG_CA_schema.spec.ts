import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN004001_CFG_CA_schema } from './FN004001_CFG_CA_schema';

describe('FN004001_CFG_CA_schema', () => {
  let findings: Finding[];
  let rule: FN004001_CFG_CA_schema;

  beforeEach(() => {
    findings = [];
    rule = new FN004001_CFG_CA_schema('test-schema');
  });

  it('doesn\'t return notification if no copy-assets.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if schema is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      copyAssetsJson: {
        $schema: 'test-schema'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});