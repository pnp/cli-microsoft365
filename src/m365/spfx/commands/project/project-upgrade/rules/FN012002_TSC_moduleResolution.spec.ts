import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN012002_TSC_moduleResolution } from './FN012002_TSC_moduleResolution';

describe('FN012002_TSC_moduleResolution', () => {
  let findings: Finding[];
  let rule: FN012002_TSC_moduleResolution;

  beforeEach(() => {
    findings = [];
    rule = new FN012002_TSC_moduleResolution('node');
  });

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

  it('returns notification if moduleResolution is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          moduleResolution: 'browser'
        },
        source: JSON.stringify({
          compilerOptions: {
            moduleResolution: 'browser'
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