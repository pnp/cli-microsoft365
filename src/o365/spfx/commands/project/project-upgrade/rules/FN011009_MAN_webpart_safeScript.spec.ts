import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../model';
import { FN011009_MAN_webpart_safeScript } from './FN011009_MAN_webpart_safeScript';

describe('FN011009_MAN_webpart_safeScript', () => {
  let findings: Finding[];
  let rule: FN011009_MAN_webpart_safeScript;

  beforeEach(() => {
    findings = [];
    rule = new FN011009_MAN_webpart_safeScript();
  });

  it('doesn\'t return notifications if no manifests collected', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });
});