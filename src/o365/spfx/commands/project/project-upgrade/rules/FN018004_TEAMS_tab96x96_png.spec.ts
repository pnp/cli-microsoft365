import * as assert from 'assert';
import * as fs from 'fs';
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

  it('doesn\'t return notifications if the icon exists', () => {
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