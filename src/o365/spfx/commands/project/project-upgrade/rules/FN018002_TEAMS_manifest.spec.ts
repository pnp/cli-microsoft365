import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { Finding } from '../Finding';
import { Project } from '../model';
import { FN018002_TEAMS_manifest } from './FN018002_TEAMS_manifest';
import Utils from '../../../../../../Utils';

describe('FN018002_TEAMS_manifest', () => {
  let findings: Finding[];
  let rule: FN018002_TEAMS_manifest;
  afterEach(() => {
    Utils.restore(fs.existsSync);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN018002_TEAMS_manifest();
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

  it('returns undefined packageName if no preconfigured entries specified in the web part', () => {
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
    assert.equal(findings.length, 1, 'No findings reported while expected');
    assert.equal(findings[0].occurrences.length, 1, 'No occurrences reported while expected');
    assert(findings[0].occurrences[0].resolution.indexOf('"packageName": "undefined",') > -1);
  });

  it('returns undefined packageName if no title specified in the web part', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        componentType: 'WebPart',
        path: '/usr/tmp/webpart',
        preconfiguredEntries: [{}]
      }]
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 1, 'No findings reported while expected');
    assert.equal(findings[0].occurrences.length, 1, 'No occurrences reported while expected');
    assert(findings[0].occurrences[0].resolution.indexOf('"packageName": "undefined",') > -1);
  });

  it('returns undefined short description if no description specified in the web part', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        componentType: 'WebPart',
        path: '/usr/tmp/webpart',
        preconfiguredEntries: [{}]
      }]
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 1, 'No findings reported while expected');
    assert.equal(findings[0].occurrences.length, 1, 'No occurrences reported while expected');
    assert(findings[0].occurrences[0].resolution.indexOf('"short": "undefined",') > -1);
  });

  it('returns undefined full description if no description specified in the web part', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        componentType: 'WebPart',
        path: '/usr/tmp/webpart',
        preconfiguredEntries: [{}]
      }]
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 1, 'No findings reported while expected');
    assert.equal(findings[0].occurrences.length, 1, 'No occurrences reported while expected');
    assert(findings[0].occurrences[0].resolution.indexOf('"full": "undefined"') > -1);
  });

  it('returns undefined id if no id specified in the web part', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        componentType: 'WebPart',
        path: '/usr/tmp/webpart',
        preconfiguredEntries: [{}]
      }]
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 1, 'No findings reported while expected');
    assert.equal(findings[0].occurrences.length, 1, 'No occurrences reported while expected');
    assert(findings[0].occurrences[0].resolution.indexOf('"id": "undefined",') > -1);
  });
});