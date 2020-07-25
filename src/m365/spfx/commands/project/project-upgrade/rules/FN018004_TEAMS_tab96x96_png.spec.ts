import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import * as sinon from 'sinon';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN018004_TEAMS_tab96x96_png } from './FN018004_TEAMS_tab96x96_png';
import Utils from '../../../../../../Utils';

describe('FN018004_TEAMS_tab96x96_png', () => {
  let findings: Finding[];
  let rule: FN018004_TEAMS_tab96x96_png;
  afterEach(() => {
    Utils.restore(fs.existsSync);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN018004_TEAMS_tab96x96_png();
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

  it('doesn\'t return notifications if the icon exists', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        id: 'c93e90e5-6222-45c6-b241-995df0029e3c',
        componentType: 'WebPart',
        path: '/usr/tmp/webpart'
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns path to icon with the specified name when fixed name used', () => {
    rule = new FN018004_TEAMS_tab96x96_png('tab96x96.png');
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        id: 'c93e90e5-6222-45c6-b241-995df0029e3c',
        componentType: 'WebPart',
        path: '/usr/tmp/webpart'
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences[0].file, path.join('teams', 'tab96x96.png'));
  });

  it('returns path to icon with name following web part ID when no fixed name specified', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        id: 'c93e90e5-6222-45c6-b241-995df0029e3c',
        componentType: 'WebPart',
        path: '/usr/tmp/webpart'
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences[0].file, path.join('teams', 'c93e90e5-6222-45c6-b241-995df0029e3c_color.png'));
  });

  it(`doesn't return notification when web part ID not specified`, () => {
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
    assert.strictEqual(findings.length, 0);
  });
});