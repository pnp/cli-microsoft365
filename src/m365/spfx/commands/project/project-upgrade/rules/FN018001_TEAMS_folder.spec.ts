import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
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
    assert.strictEqual(rule.resolution, '');
  });

  it('returns empty file name by default', () => {
    assert.strictEqual(rule.file, '');
  });

  it('doesn\'t return notifications if no manifests are present', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if teams folder exists', () => {
    sinon.stub(fs, 'existsSync').callsFake((filePath) => filePath.toString().endsWith(`${path.sep}teams`));
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        componentType: 'WebPart',
        path: '/usr/tmp/webpart'
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns 1 finding with 1 occurrence for one web part if the teams folder does not exists', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        componentType: 'WebPart',
        path: '/usr/tmp/webpart'
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences.length, 1, 'Incorrect number of occurrences');
  });

  it('returns 1 finding with 1 occurrence for two web parts if the teams folder does not exists', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [
        {
          $schema: 'schema',
          componentType: 'WebPart',
          path: '/usr/tmp/webpart1'
        },
        {
          $schema: 'schema',
          componentType: 'WebPart',
          path: '/usr/tmp/webpart2'
        }
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences.length, 1, 'Incorrect number of occurrences');
  });
});
