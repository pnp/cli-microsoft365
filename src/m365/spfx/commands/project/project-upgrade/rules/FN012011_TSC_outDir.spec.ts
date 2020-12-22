import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN012011_TSC_outDir } from './FN012011_TSC_outDir';

describe('FN012011_TSC_outDir', () => {
  let findings: Finding[];
  let rule: FN012011_TSC_outDir;

  beforeEach(() => {
    findings = [];
    rule = new FN012011_TSC_outDir('lib');
  });

  it('doesn\'t return notification if outDir is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          outDir: 'lib'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if outDir is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          outDir: 'tmp'
        },
        source: JSON.stringify({
          compilerOptions: {
            outDir: 'tmp'
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
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