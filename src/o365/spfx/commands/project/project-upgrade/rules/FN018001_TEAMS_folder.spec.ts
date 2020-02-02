import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN018001_TEAMS_folder } from './FN018001_TEAMS_folder';
import Utils from '../../../../../../Utils';

describe('FN018001_TEAMS_folder', () => {
  let findings: Finding[];
  let rule: FN018001_TEAMS_folder;
  afterEach(() => {
    Utils.restore(fs.existsSync);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN018001_TEAMS_folder();
  });

  it('returns empty resolution by default', () => {
    assert.equal(rule.resolution, '');
  });

  it('returns empty file name by default', () => {
    assert.equal(rule.file, '');
  });

  it('doesn\'t return notifications if no manifests are present', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });

  it('doesn\'t return notifications if teams folder exists', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        componentType: 'WebPart',
        path: '/usr/tmp/webpart'
      }]
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });
});