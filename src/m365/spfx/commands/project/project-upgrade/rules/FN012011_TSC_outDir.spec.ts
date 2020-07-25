import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
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

  it('doesn\'t return notification if object is missing', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: undefined
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});