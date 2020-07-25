import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN012010_TSC_experimentalDecorators } from './FN012010_TSC_experimentalDecorators';

describe('FN012010_TSC_experimentalDecorators', () => {
  let findings: Finding[];
  let rule: FN012010_TSC_experimentalDecorators;

  beforeEach(() => {
    findings = [];
    rule = new FN012010_TSC_experimentalDecorators();
  });

  it('doesn\'t return notification if schema is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          experimentalDecorators: true,
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
});