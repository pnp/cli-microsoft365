import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
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
        $schema: 'test-schema',
        deployCdnPath: './release/assets/'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if schema is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      copyAssetsJson: {
        $schema: 'old-schema',
        deployCdnPath: './release/assets/',
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