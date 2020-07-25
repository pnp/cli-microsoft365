import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN014003_CODE_launch } from './FN014003_CODE_launch';

describe('FN014003_CODE_launch', () => {
  let findings: Finding[];
  let rule: FN014003_CODE_launch;

  beforeEach(() => {
    findings = [];
    rule = new FN014003_CODE_launch();
  })

  it('doesn\'t return notification if launch.json already exists', () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        launchJson: {
          version: '2.0'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});