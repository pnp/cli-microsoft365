import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import * as sinon from 'sinon';
import { Finding } from '../Finding';
import { Project } from '../../model';
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
    assert.strictEqual(rule.file, '');
  });

  it(`doesn't return notifications if no web part manifests are present`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`looks for Teams manifest for a web part using the correct path`, () => {
    const existsSyncFake: sinon.SinonStub = sinon.stub(fs, 'existsSync').callsFake(() => true);
    const project: Project = {
      path: os.platform() === 'win32' ? 'c:\\tmp' : '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        componentType: 'WebPart',
        path: os.platform() === 'win32' ? 'c:\\tmp\\webpart\\webpart.manifest.json' : '/usr/tmp/webpart/webpart.manifest.json'
      }]
    };
    rule.visit(project, findings);
    if (os.platform() === 'win32') {
      assert(existsSyncFake.calledWith('c:\\tmp\\teams\\manifest_webpart.json'));
    }
    else {
      assert(existsSyncFake.calledWith('/usr/tmp/teams/manifest_webpart.json'));
    }
  });

  it(`doesn't return notifications if the Teams manifest for the given web part already exists`, () => {
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
    assert.strictEqual(findings.length, 0);
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
    assert.strictEqual(findings.length, 1, 'No findings reported while expected');
    assert.strictEqual(findings[0].occurrences.length, 1, 'No occurrences reported while expected');
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
    assert.strictEqual(findings.length, 1, 'No findings reported while expected');
    assert.strictEqual(findings[0].occurrences.length, 1, 'No occurrences reported while expected');
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
    assert.strictEqual(findings.length, 1, 'No findings reported while expected');
    assert.strictEqual(findings[0].occurrences.length, 1, 'No occurrences reported while expected');
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
    assert.strictEqual(findings.length, 1, 'No findings reported while expected');
    assert.strictEqual(findings[0].occurrences.length, 1, 'No occurrences reported while expected');
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
    assert.strictEqual(findings.length, 1, 'No findings reported while expected');
    assert.strictEqual(findings[0].occurrences.length, 1, 'No occurrences reported while expected');
    assert(findings[0].occurrences[0].resolution.indexOf('"id": "undefined",') > -1);
  });

  it('creates manifest with a unique name following the web part name (single web part)', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        componentType: 'WebPart',
        path: '/usr/tmp/webpart/webpart.manifest.json',
        preconfiguredEntries: [{}]
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences[0].file, path.join('teams', 'manifest_webpart.json'));
  });

  it('creates manifest with a unique name following the web part name (multiple web parts)', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [
        {
          $schema: 'schema',
          componentType: 'WebPart',
          path: '/usr/tmp/webpart1/webpart1.manifest.json',
          preconfiguredEntries: [{}]
        },
        {
          $schema: 'schema',
          componentType: 'WebPart',
          path: '/usr/tmp/webpart2/webpart2.manifest.json',
          preconfiguredEntries: [{}]
        }
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences[0].file, path.join('teams', 'manifest_webpart1.json'), 'Incorrect manifest path for web part 1');
    assert.strictEqual(findings[0].occurrences[1].file, path.join('teams', 'manifest_webpart2.json'), 'Incorrect manifest path for web part 2');
  });
});