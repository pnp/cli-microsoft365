import * as sinon from 'sinon';
import * as assert from 'assert';
import { Project } from '../../model';
import { DynamicRule } from './DynamicRule';
import Utils from '../../../../../../Utils';
import request from '../../../../../../request';
import * as fs from 'fs';

describe('DynamicRule', () => {
  let rule: DynamicRule;

  beforeEach(() => {
    rule = new DynamicRule();
  })

  afterEach(() => {
    Utils.restore([
      fs.readFileSync,
      request.head,
      request.post,
    ]);
  });

  it('doesn\'t return anything if project json is missing', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: undefined
    };
    const findings = await rule.visit(project);
    assert.equal(findings.entries.length, 0);
  });

  it('returns something is package.json is here', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/pnpjs': '1.3.5'
        }
      }
    };
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('@pnp/pnpjs/package.json')) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake(() => Promise.reject());
    const findings = await rule.visit(project);
    assert.equal(findings.entries.length, 1);
  });

  it('doesnt return anything is package is unsupported', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/sp-taxonomy': '1.3.5',
          '@pnp/sp-clientsvc': '1.3.5',
        }
      }
    };
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('@pnp/sp-taxonomy/package.json') || path.toString().endsWith('@pnp/sp-clientsvc/package.json')) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake(() => Promise.reject());
    const findings = await rule.visit(project);
    assert.equal(findings.entries.length, 0);
  });

  it('returns from main if module is missing', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/pnpjs': '1.3.5'
        }
      }
    };
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('@pnp/pnpjs/package.json')) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
        });
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake(() => Promise.resolve({ scriptType: 'script' }));
    const findings = await rule.visit(project);
    assert.equal(findings.entries.length, 1);
  });

  it('doesn\'t return anything if both module and main are missing', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/pnpjs': '1.3.5'
        }
      }
    };
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('@pnp/pnpjs/package.json')) {
        return JSON.stringify({
        });
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());
    const findings = await rule.visit(project);
    assert.equal(findings.entries.length, 0);
  });

  it('doesn\'t return anything if file is not present on CDN', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/pnpjs': '1.3.5'
        }
      }
    };
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path) => {
      if (path.toString().endsWith('@pnp/pnpjs/package.json')) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      }
      else {
        return originalReadFileSync(path);
      }
    });
    sinon.stub(request, 'head').callsFake(() => Promise.reject());
    sinon.stub(request, 'post').callsFake(() => Promise.resolve({ scriptType: 'module' }));
    const findings = await rule.visit(project);
    assert.equal(findings.entries.length, 0);
  });
});