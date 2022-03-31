import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN012009_TSC_lib_es2015_collection } from './FN012009_TSC_lib_es2015_collection';

describe('FN012009_TSC_lib_es2015_collection', () => {
  let findings: Finding[];
  let rule: FN012009_TSC_lib_es2015_collection;

  beforeEach(() => {
    findings = [];
    rule = new FN012009_TSC_lib_es2015_collection();
  });

  it('doesn\'t return notification if es2015.collection is already present', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          lib: ['es2015.collection']
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if es2015.collection is not present', () => {
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