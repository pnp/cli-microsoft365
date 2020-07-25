import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN012008_TSC_lib_dom } from './FN012008_TSC_lib_dom';

describe('FN012008_TSC_lib_dom', () => {
  let findings: Finding[];
  let rule: FN012008_TSC_lib_dom;

  beforeEach(() => {
    findings = [];
    rule = new FN012008_TSC_lib_dom();
  })

  it('doesn\'t return notification if dom is already present', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          lib: ['dom']
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