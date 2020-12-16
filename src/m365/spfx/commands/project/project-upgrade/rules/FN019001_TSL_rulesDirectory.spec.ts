import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN019001_TSL_rulesDirectory } from './FN019001_TSL_rulesDirectory';

describe('FN019001_TSL_rulesDirectory', () => {
  let findings: Finding[];
  let rule: FN019001_TSL_rulesDirectory;

  beforeEach(() => {
    findings = [];
    rule = new FN019001_TSL_rulesDirectory();
  });

  it('doesn\'t return notification if rulesDirectory is undefined', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsLintJsonRoot: {
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if rulesDirectory is defined', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsLintJsonRoot: {
        rulesDirectory: [],
        source: JSON.stringify({
          rulesDirectory: []
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 2, 'Incorrect line number');
  });
});