import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
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
});