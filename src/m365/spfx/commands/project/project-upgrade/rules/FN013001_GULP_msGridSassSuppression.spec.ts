import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN013001_GULP_msGridSassSuppression } from './FN013001_GULP_msGridSassSuppression';

describe('FN013001_GULP_msGridSassSuppression', () => {
  let findings: Finding[];
  let rule: FN013001_GULP_msGridSassSuppression;

  beforeEach(() => {
    findings = [];
    rule = new FN013001_GULP_msGridSassSuppression();
  })

  it('doesn\'t return notification if ms-grid sass suppression is already present', () => {
    const project: Project = {
      path: '/usr/tmp',
      gulpfileJs: {
        src: rule.resolution
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