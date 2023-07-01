import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN012020_TSC_noImplicitAny } from './FN012020_TSC_noImplicitAny.js';

describe('FN012020_TSC_noImplicitAny', () => {
  let findings: Finding[];
  let rule: FN012020_TSC_noImplicitAny;

  beforeEach(() => {
    findings = [];
    rule = new FN012020_TSC_noImplicitAny(true);
  });

  it('doesn\'t return notification if noImplicitAny is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          noImplicitAny: true
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if noImplicitAny is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          noImplicitAny: false
        },
        source: JSON.stringify({
          compilerOptions: {
            noImplicitAny: false
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });

  it('doesn\'t return notification if tsconfig is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});
