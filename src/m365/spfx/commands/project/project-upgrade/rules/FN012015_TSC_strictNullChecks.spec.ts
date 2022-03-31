import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN012015_TSC_strictNullChecks } from './FN012015_TSC_strictNullChecks';

describe('FN012015_TSC_strictNullChecks', () => {
  let findings: Finding[];
  let rule: FN012015_TSC_strictNullChecks;

  beforeEach(() => {
    findings = [];
    rule = new FN012015_TSC_strictNullChecks(false);
  });

  it('doesn\'t return notification if strictNullChecks is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          strictNullChecks: false
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

  it('returns notification if strictNullChecks has the wrong value', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          strictNullChecks: true
        },
        source: JSON.stringify({
          compilerOptions: {
            strictNullChecks: true
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });

  it('returns notification if strictNullChecks is missing', () => {
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