import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN012007_TSC_lib_es5 } from './FN012007_TSC_lib_es5';

describe('FN012007_TSC_lib_es5', () => {
  let findings: Finding[];
  let rule: FN012007_TSC_lib_es5;

  beforeEach(() => {
    findings = [];
    rule = new FN012007_TSC_lib_es5();
  })

  it('doesn\'t return notification if es5 is already present', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          lib: ['es5']
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