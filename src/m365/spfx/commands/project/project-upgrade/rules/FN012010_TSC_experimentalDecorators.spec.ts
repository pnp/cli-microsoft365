import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN012010_TSC_experimentalDecorators } from './FN012010_TSC_experimentalDecorators';

describe('FN012010_TSC_experimentalDecorators', () => {
  let findings: Finding[];
  let rule: FN012010_TSC_experimentalDecorators;

  beforeEach(() => {
    findings = [];
    rule = new FN012010_TSC_experimentalDecorators();
  });

  it('doesn\'t return notification if experimentalDecorators is already up-to-date', () => {
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

  it('returns notification if experimentalDecorators is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          experimentalDecorators: false,
        },
        source: JSON.stringify({
          compilerOptions: {
            experimentalDecorators: false,
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });

  it('doesn\'t return notification if tsConfigJson is missing', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: undefined
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});