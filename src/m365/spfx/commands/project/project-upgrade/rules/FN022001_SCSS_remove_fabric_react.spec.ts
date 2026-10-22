import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { spfx } from '../../../../../../utils/spfx.js';
import { Project, ScssFile } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN022001_SCSS_remove_fabric_react } from './FN022001_SCSS_remove_fabric_react.js';

describe('FN022001_SCSS_remove_fabric_react', () => {
  let findings: Finding[];
  let rule: FN022001_SCSS_remove_fabric_react;
  let fileStub: sinon.SinonStub;
  let utilsStub: sinon.SinonStub;

  beforeEach(() => {
    findings = [];
    utilsStub = sinon.stub(spfx, 'isReactProject').returns(true);
  });

  afterEach(() => {
    fileStub.restore();
    utilsStub.restore();
  });

  it('doesn\'t return notifications if import is already removed', () => {
    rule = new FN022001_SCSS_remove_fabric_react({ importValue: '~fabric-ui/react' });

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

  it('returns notifications if import is not removed', () => {
    rule = new FN022001_SCSS_remove_fabric_react({ importValue: '~fabric-ui/react' });
    fileStub = sinon.stub(fs, 'readFileSync').returns('~fabric-ui/react');
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
    rule = new FN022001_SCSS_remove_fabric_react({ importValue: '~fabric-ui/react' });
    utilsStub.restore();
    utilsStub = sinon.stub(spfx, 'isReactProject').returns(false);

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
    rule = new FN022001_SCSS_remove_fabric_react({ importValue: '~fabric-ui/react' });
    fileStub = sinon.stub(fs, 'readFileSync').returns('');

    const project: Project = {
      path: '/usr/tmp',
      scssFiles: []
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('rule file name is empy', () => {
    rule = new FN022001_SCSS_remove_fabric_react({ importValue: '~fabric-ui/react' });
    fileStub = sinon.stub(fs, 'readFileSync').returns('');

    const project: Project = {
      path: '/usr/tmp',
      scssFiles: []
    };
    rule.visit(project, findings);
    assert.strictEqual(rule.file, '');
  });
});
