import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN012013_TSC_exclude } from './FN012013_TSC_exclude';

describe('FN012013_TSC_exclude', () => {
  let findings: Finding[];
  let rule: FN012013_TSC_exclude;

  beforeEach(() => {
    findings = [];
    rule = new FN012013_TSC_exclude(['node_modules', 'lib']);
  });

  it('doesn\'t return notification if exclude has the exact same elements', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        exclude: [
          'node_modules',
          'lib'
        ]
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if exclude has the exact same elements in different order', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        exclude: [
          'lib',
          'node_modules'
        ]
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if exclude has all required elements', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        exclude: [
          'node_modules',
          'tmp',
          'lib'
        ]
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
});