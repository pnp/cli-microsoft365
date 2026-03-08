import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { sinonUtil } from '../../../../../../utils/sinonUtil.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FileAddRemoveRule } from './FileAddRemoveRule.js';

class FileAddRule extends FileAddRemoveRule {
  constructor(options: { add: boolean }) {
    super({ filePath: 'test-file.ext', ...options });
  }

  get id(): string {
    return 'FN000000';
  }
}

describe('FileAddRemoveRule', () => {
  let findings: Finding[];
  let rule: FileAddRemoveRule;

  afterEach(() => {
    sinonUtil.restore(fs.existsSync);
  });

  beforeEach(() => {
    findings = [];
  });

  it('doesn\'t return notification if the file doesn\'t exist and should be deleted', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    rule = new FileAddRule({ add: false });
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if the file exists and should be added', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    rule = new FileAddRule({ add: true });
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('adjusts description when the file should be created', () => {
    rule = new FileAddRule({ add: true });
    assert(rule.description.indexOf('Add') > -1);
  });

  it('adjusts description when the file should be removed', () => {
    rule = new FileAddRule({ add: false });
    assert(rule.description.indexOf('Remove') > -1);
  });

  it('adjusts resolution when the file should be created', () => {
    rule = new FileAddRule({ add: true });
    assert(rule.resolution.indexOf('add_cmd') > -1);
  });

  it('adjusts resolution when the file should be removed', () => {
    rule = new FileAddRule({ add: false });
    assert(rule.resolution.indexOf('remove_cmd') > -1);
  });
});
