import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN005002_CFG_DAS_workingDir } from './FN005002_CFG_DAS_workingDir';

describe('FN005002_CFG_DAS_workingDir', () => {
  let findings: Finding[];
  let rule: FN005002_CFG_DAS_workingDir;

  beforeEach(() => {
    findings = [];
    rule = new FN005002_CFG_DAS_workingDir('./release/assets/');
  });

  it('doesn\'t return notification if no deploy-azure-storage.json found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if workingDir is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      deployAzureStorageJson: {
        $schema: 'test-schema',
        workingDir: './release/assets/'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if workingDir is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      deployAzureStorageJson: {
        $schema: 'test-schema',
        workingDir: './temp/deploy/',
        source: JSON.stringify({
          $schema: 'test-schema',
          workingDir: './temp/deploy/'
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });
});