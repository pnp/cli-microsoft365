import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
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

  it('returns notification if schema is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      serveJson: {
        $schema: 'old-schema',
        source: JSON.stringify({
          $schema: 'old-schema'
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 2, 'Incorrect line number');
  });
});