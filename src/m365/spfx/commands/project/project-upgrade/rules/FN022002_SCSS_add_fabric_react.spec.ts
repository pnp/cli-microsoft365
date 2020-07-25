import * as assert from 'assert';
import * as sinon from 'sinon';
import { Finding } from '../Finding';
import { Project, ScssFile } from '../../model';
import { FN022002_SCSS_add_fabric_react } from './FN022002_SCSS_add_fabric_react';
import * as fs from 'fs';
import { Utils }  from '../';

describe('FN022002_SCSS_add_fabric_react', () => {
  let findings: Finding[];
  let rule: FN022002_SCSS_add_fabric_react;
  let fileStub: sinon.SinonStub;
  let utilsStub: sinon.SinonStub;

  beforeEach(() => {
    findings = [];
    utilsStub = sinon.stub(Utils, 'isReactProject').returns(true);
  });

  afterEach(() => {
    fileStub.restore();
    utilsStub.restore();
  });

  it('doesn\'t return notifications if import is already there', () => {
    rule = new FN022002_SCSS_add_fabric_react('~fabric-ui/react');
    
    fileStub = sinon.stub(fs, 'readFileSync').returns('~fabric-ui/react');
    const project: Project = {
      path: '/usr/tmp',
      scssFiles: [
        new ScssFile('some/path')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notifications if import is missing and no condition', () => {
    rule = new FN022002_SCSS_add_fabric_react('~fabric-ui/react');
    fileStub = sinon.stub(fs, 'readFileSync').returns('');
    const project: Project = {
      path: '/usr/tmp',
      scssFiles: [
        new ScssFile('some/path')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('doesn\'t return notifications if import is missing but condition is not met', () => {
    rule = new FN022002_SCSS_add_fabric_react('~fabric-ui/react', '~old-fabric-ui/react');
    
    fileStub = sinon.stub(fs, 'readFileSync').returns('');
    const project: Project = {
      path: '/usr/tmp',
      scssFiles: [
        new ScssFile('some/path')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notifications if import is missing and condition is met', () => {
    rule = new FN022002_SCSS_add_fabric_react('~fabric-ui/react', '~old-fabric-ui/react');
    fileStub = sinon.stub(fs, 'readFileSync').returns('~old-fabric-ui/react');
    const project: Project = {
      path: '/usr/tmp',
      scssFiles: [
        new ScssFile('some/path')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('doesn\'t return notifications if scss is not in react web part', () => {
    rule = new FN022002_SCSS_add_fabric_react('~fabric-ui/react');
    utilsStub.restore();
    utilsStub = sinon.stub(Utils, 'isReactProject').returns(false);

    fileStub = sinon.stub(fs, 'readFileSync').returns('');

    const project: Project = {
      path: '/usr/tmp',
      scssFiles: [
        new ScssFile('some/path')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if no scss files', () => {
    rule = new FN022002_SCSS_add_fabric_react('~fabric-ui/react');
    fileStub = sinon.stub(fs, 'readFileSync').returns('');

    const project: Project = {
      path: '/usr/tmp',
      scssFiles: []
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('rule file name is empty', () => {
    rule = new FN022002_SCSS_add_fabric_react('~fabric-ui/react');
    fileStub = sinon.stub(fs, 'readFileSync').returns('');

    const project: Project = {
      path: '/usr/tmp',
      scssFiles: []
    };
    rule.visit(project, findings);
    assert.strictEqual(rule.file, '');
  });
});