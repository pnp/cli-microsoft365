import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN012004_TSC_typeRoots_types } from './FN012004_TSC_typeRoots_types';

describe('FN012004_TSC_typeRoots_types', () => {
  let findings: Finding[];
  let rule: FN012004_TSC_typeRoots_types;

  beforeEach(() => {
    findings = [];
    rule = new FN012004_TSC_typeRoots_types();
  })

  it('doesn\'t return notification if ./node_modules/@types is already present', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          typeRoots: ['./node_modules/@types']
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if tsconfig is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});