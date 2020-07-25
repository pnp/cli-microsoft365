import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN012005_TSC_typeRoots_microsoft } from './FN012005_TSC_typeRoots_microsoft';

describe('FN012005_TSC_typeRoots_microsoft', () => {
  let findings: Finding[];
  let rule: FN012005_TSC_typeRoots_microsoft;

  beforeEach(() => {
    findings = [];
    rule = new FN012005_TSC_typeRoots_microsoft();
  })

  it('doesn\'t return notification if ./node_modules/@microsoft is already present', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          typeRoots: ['./node_modules/@microsoft']
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