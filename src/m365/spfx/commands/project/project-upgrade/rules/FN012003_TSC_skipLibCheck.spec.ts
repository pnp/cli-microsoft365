import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
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

  it('returns notification if skipLibCheck is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          skipLibCheck: false
        },
        source: JSON.stringify({
          compilerOptions: {
            skipLibCheck: false
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });

  it('doesn\'t return notification if tsconfig is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});