import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN012001_TSC_module } from './FN012001_TSC_module';

describe('FN012001_TSC_module', () => {
  let findings: Finding[];
  let rule: FN012001_TSC_module;

  beforeEach(() => {
    findings = [];
    rule = new FN012001_TSC_module('esnext');
  })

  it('doesn\'t return notification if module type is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          module: 'esnext'
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