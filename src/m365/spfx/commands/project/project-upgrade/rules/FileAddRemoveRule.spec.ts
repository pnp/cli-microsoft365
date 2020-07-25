import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FileAddRemoveRule } from './FileAddRemoveRule';
import Utils from '../../../../../../Utils';

class FileAddRule extends FileAddRemoveRule {
  constructor(add: boolean) {
    super('test-file.ext', add);
  }

  get id(): string {
    return 'FN000000';
  }
}

describe('FileAddRemoveRule', () => {
  let findings: Finding[];
  let rule: FileAddRemoveRule;

  afterEach(() => {
    Utils.restore(fs.existsSync);
  });

  beforeEach(() => {
    findings = [];
  });

  it('doesn\'t return notification if the file doesn\'t exist and should be deleted', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    rule = new FileAddRule(false);
    const project: Project = {
      path: '/usr/tmp',
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if the file exists and should be added', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    rule = new FileAddRule(true);
    const project: Project = {
      path: '/usr/tmp',
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('adjusts description when the file should be created', () => {
    rule = new FileAddRule(true);
    assert(rule.description.indexOf('Add') > -1);
  });

  it('adjusts description when the file should be removed', () => {
    rule = new FileAddRule(false);
    assert(rule.description.indexOf('Remove') > -1);
  });

  it('adjusts resolution when the file should be created', () => {
    rule = new FileAddRule(true);
    assert(rule.resolution.indexOf('add_cmd') > -1);
  });

  it('adjusts resolution when the file should be removed', () => {
    rule = new FileAddRule(false);
    assert(rule.resolution.indexOf('remove_cmd') > -1);
  });
});