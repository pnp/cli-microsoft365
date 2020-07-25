import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN012016_TSC_noUnusedLocals } from './FN012016_TSC_noUnusedLocals';

describe('FN012016_TSC_noUnusedLocals', () => {
  let findings: Finding[];
  let rule: FN012016_TSC_noUnusedLocals;

  beforeEach(() => {
    findings = [];
    rule = new FN012016_TSC_noUnusedLocals(false);
  });

  it('doesn\'t return notification if noUnusedLocals is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          noUnusedLocals: false
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if object is missing', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: undefined
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if noUnusedLocals has the wrong value', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          noUnusedLocals: true
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('returns notification if noUnusedLocals is missing', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});