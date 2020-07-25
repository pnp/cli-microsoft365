import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN019002_TSL_extends } from './FN019002_TSL_extends';

describe('FN019002_TSL_extends', () => {
  let findings: Finding[];
  let rule: FN019002_TSL_extends;

  beforeEach(() => {
    findings = [];
    rule = new FN019002_TSL_extends('@microsoft/sp-tslint-rules/base-tslint.json');
  });

  it('doesn\'t return notification if extends is correctly configured', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsLintJsonRoot: {
        extends: '@microsoft/sp-tslint-rules/base-tslint.json'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});