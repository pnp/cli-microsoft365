import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN013002_GULP_serveTask } from './FN013002_GULP_serveTask';

describe('FN013002_GULP_serveTask', () => {
  let findings: Finding[];
  let rule: FN013002_GULP_serveTask;

  beforeEach(() => {
    findings = [];
    rule = new FN013002_GULP_serveTask();
  })

  it('doesn\'t return notification if serve task is already present', () => {
    const project: Project = {
      path: '/usr/tmp',
      gulpfileJs: {
        source: rule.resolution
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if gulpfile.js is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});