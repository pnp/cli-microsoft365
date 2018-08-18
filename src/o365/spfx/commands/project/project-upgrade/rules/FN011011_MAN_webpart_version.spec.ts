import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../model';
import { FN011011_MAN_webpart_version } from './FN011011_MAN_webpart_version';

describe('FN011011_MAN_webpart_version', () => {
  let findings: Finding[];
  let rule: FN011011_MAN_webpart_version;

  beforeEach(() => {
    findings = [];
    rule = new FN011011_MAN_webpart_version();
  });

  it('doesn\'t return notifications if no manifests collected', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });
});