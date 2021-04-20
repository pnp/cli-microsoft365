import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN004002_CFG_CA_deployCdnPath } from './FN004002_CFG_CA_deployCdnPath';

describe('FN004002_CFG_CA_deployCdnPath', () => {
  let findings: Finding[];
  let rule: FN004002_CFG_CA_deployCdnPath;

  beforeEach(() => {
    findings = [];
    rule = new FN004002_CFG_CA_deployCdnPath('./release/assets/');
  });

  it(`doesn't return notification if no copy-assets.json found`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't return notification if deployCdnPath is already up-to-date`, () => {
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

  it('returns notification if deployCdnPath is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      copyAssetsJson: {
        $schema: 'test-schema',
        deployCdnPath: './tmp/deploy/',
        source: JSON.stringify({
          $schema: 'test-schema',
          deployCdnPath: './tmp/deploy/'
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });
});