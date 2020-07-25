import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN012002_TSC_moduleResolution } from './FN012002_TSC_moduleResolution';

describe('FN012002_TSC_moduleResolution', () => {
  let findings: Finding[];
  let rule: FN012002_TSC_moduleResolution;

  beforeEach(() => {
    findings = [];
    rule = new FN012002_TSC_moduleResolution('node');
  })

  it('doesn\'t return notification if moduleResolution is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          moduleResolution: 'node'
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