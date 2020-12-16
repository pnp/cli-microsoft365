import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN003001_CFG_schema } from './FN003001_CFG_schema';

describe('FN003001_CFG_schema', () => {
  let findings: Finding[];
  let rule: FN003001_CFG_schema;

  beforeEach(() => {
    findings = [];
    rule = new FN003001_CFG_schema('test-schema');
  });

  it('doesn\'t return notification if no config.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if schema is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      configJson: {
        $schema: 'test-schema',
        version: '2.0',
        bundles: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if schema is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      configJson: {
        $schema: 'old-schema',
        version: '2.0',
        bundles: {},
        source: JSON.stringify({
          $schema: 'old-schema',
          version: '2.0',
          bundles: {}
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of notifications');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 2, 'Incorrect line number');
  });
});