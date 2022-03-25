import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN012018_TSC_lib_es2015_promise } from './FN012018_TSC_lib_es2015_promise';

describe('FN012018_TSC_lib_es2015_promise', () => {
  let findings: Finding[];
  let rule: FN012018_TSC_lib_es2015_promise;

  beforeEach(() => {
    findings = [];
    rule = new FN012018_TSC_lib_es2015_promise();
  });

  it('doesn\'t return notification if es2015.promise is already present', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          lib: ['es2015.promise']
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if es2015.promise is not present', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          lib: []
        },
        source: JSON.stringify({
          compilerOptions: {
            lib: []
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