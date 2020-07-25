import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN012003_TSC_skipLibCheck } from './FN012003_TSC_skipLibCheck';

describe('FN012003_TSC_skipLibCheck', () => {
  let findings: Finding[];
  let rule: FN012003_TSC_skipLibCheck;

  beforeEach(() => {
    findings = [];
    rule = new FN012003_TSC_skipLibCheck(true);
  })

  it('doesn\'t return notification if skipLibCheck is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          skipLibCheck: true
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if tsconfig is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});