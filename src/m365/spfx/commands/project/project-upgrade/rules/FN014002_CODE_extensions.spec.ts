import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN014002_CODE_extensions } from './FN014002_CODE_extensions';

describe('FN014002_CODE_extensions', () => {
  let findings: Finding[];
  let rule: FN014002_CODE_extensions;

  beforeEach(() => {
    findings = [];
    rule = new FN014002_CODE_extensions();
  });

  it('doesn\'t return notification if extensions.json already exists', () => {
    const project: Project = {
      path: '/usr/tmp',
      vsCode: {
        extensionsJson: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});