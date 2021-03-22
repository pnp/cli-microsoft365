import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN012019_TSC_types_es6_promise } from './FN012019_TSC_types_es6_promise';

describe('FN012019_TSC_types_es6_promise', () => {
  let findings: Finding[];
  let rule: FN012019_TSC_types_es6_promise;

  beforeEach(() => {
    findings = [];
    rule = new FN012019_TSC_types_es6_promise(false);
  })

  it('doesn\'t return notification if es6-promise should be removed and is not present', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          types: []
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if es6-promise should be removed and is present', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          types: [
            'es6-promise'
          ]
        },
        source: JSON.stringify({
          compilerOptions: {
            types: [
              'es6-promise'
            ]
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 4, 'Incorrect line number');
  });

  it('doesn\'t return notification if es6-promise should be added and is already present', () => {
    rule = new FN012019_TSC_types_es6_promise(true);
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          types: ['es6-promise']
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if es6-promise should be added but is not present', () => {
    rule = new FN012019_TSC_types_es6_promise(true);
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          types: []
        },
        source: JSON.stringify({
          compilerOptions: {
            types: []
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