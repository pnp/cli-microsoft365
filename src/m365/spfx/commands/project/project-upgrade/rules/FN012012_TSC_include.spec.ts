import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN012012_TSC_include } from './FN012012_TSC_include';

describe('FN012012_TSC_include', () => {
  let findings: Finding[];
  let rule: FN012012_TSC_include;

  beforeEach(() => {
    findings = [];
    rule = new FN012012_TSC_include(['src/**/*.ts']);
  });

  it('doesn\'t return notification if include has the exact same elements', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        include: [
          'src/**/*.ts'
        ]
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if include has the required elements', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        include: [
          'foo',
          'src/**/*.ts'
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