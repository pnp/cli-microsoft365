import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN012014_TSC_inlineSources } from './FN012014_TSC_inlineSources';

describe('FN012014_TSC_inlineSources', () => {
  let findings: Finding[];
  let rule: FN012014_TSC_inlineSources;

  beforeEach(() => {
    findings = [];
    rule = new FN012014_TSC_inlineSources(false);
  });

  it('doesn\'t return notification if inlineSources is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          inlineSources: false
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

  it('returns notification if inlineSources has the wrong value', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          inlineSources: true
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('returns notification if inlineSources is missing', () => {
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