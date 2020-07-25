import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN012006_TSC_types_es6_collections } from './FN012006_TSC_types_es6_collections';

describe('FN012006_TSC_types_es6_collections', () => {
  let findings: Finding[];
  let rule: FN012006_TSC_types_es6_collections;

  beforeEach(() => {
    findings = [];
    rule = new FN012006_TSC_types_es6_collections(false);
  })

  it('doesn\'t return notification if es6-collection should be removed and is not present', () => {
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

  it('doesn\'t return notification if es6-collection should be added and is already present', () => {
    rule = new FN012006_TSC_types_es6_collections(true);
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          types: ['es6-collections']
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if es6-collection should be added but is not present', () => {
    rule = new FN012006_TSC_types_es6_collections(true);
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          types: []
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('doesn\'t return notification if tsconfig is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});