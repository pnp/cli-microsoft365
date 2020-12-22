import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN005001_CFG_DAS_schema } from './FN005001_CFG_DAS_schema';

describe('FN005001_CFG_DAS_schema', () => {
  let findings: Finding[];
  let rule: FN005001_CFG_DAS_schema;

  beforeEach(() => {
    findings = [];
    rule = new FN005001_CFG_DAS_schema('test-schema');
  });

  it('doesn\'t return notification if no deploy-azure-storage.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if schema is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      deployAzureStorageJson: {
        $schema: 'test-schema'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if schema is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      deployAzureStorageJson: {
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