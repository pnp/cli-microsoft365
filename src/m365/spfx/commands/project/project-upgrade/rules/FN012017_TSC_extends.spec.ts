import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN012017_TSC_extends } from './FN012017_TSC_extends';

describe('FN012017_TSC_extends', () => {
  let findings: Finding[];
  let rule: FN012017_TSC_extends;

  beforeEach(() => {
    findings = [];
    rule = new FN012017_TSC_extends('./node_modules/@microsoft/rush-stack-compiler-2.7/includes/tsconfig-web.json');
  });

  it('doesn\'t return notification if extends has the exact same elements', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        extends: './node_modules/@microsoft/rush-stack-compiler-2.7/includes/tsconfig-web.json'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if extends has the exact same elements in different order', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        extends: './node_modules/@microsoft/rush-stack-compiler-2.7/includes/tsconfig-web.json'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if extends has all required elements', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        extends: './node_modules/@microsoft/rush-stack-compiler-2.7/includes/tsconfig-web.json'
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

  it('returns notification if extends value has to be changed', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        extends: 'abc'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});