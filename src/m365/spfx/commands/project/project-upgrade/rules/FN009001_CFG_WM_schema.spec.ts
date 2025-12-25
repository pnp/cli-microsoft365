import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN009001_CFG_WM_schema } from './FN009001_CFG_WM_schema.js';

describe('FN009001_CFG_WM_schema', () => {
  let findings: Finding[];
  let rule: FN009001_CFG_WM_schema;

  beforeEach(() => {
    findings = [];
    rule = new FN009001_CFG_WM_schema({ version: 'test-schema' });
  });

  it('doesn\'t return notification if write-manifests.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if schema is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      writeManifestsJson: {
        $schema: 'test-schema'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if schema is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      writeManifestsJson: {
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
