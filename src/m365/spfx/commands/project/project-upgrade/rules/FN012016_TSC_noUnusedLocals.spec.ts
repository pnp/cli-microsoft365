import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN012016_TSC_noUnusedLocals } from './FN012016_TSC_noUnusedLocals';

describe('FN012016_TSC_noUnusedLocals', () => {
  let findings: Finding[];
  let rule: FN012016_TSC_noUnusedLocals;

  beforeEach(() => {
    findings = [];
    rule = new FN012016_TSC_noUnusedLocals(false);
  });

  it('doesn\'t return notification if noUnusedLocals is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          noUnusedLocals: false
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if object is missing', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: undefined
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if noUnusedLocals has the wrong value', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          noUnusedLocals: true
        },
        source: JSON.stringify({
          compilerOptions: {
            noUnusedLocals: true
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });

  it('returns notification if noUnusedLocals is missing', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});