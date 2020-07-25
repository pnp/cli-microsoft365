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
    assert.strictEqual(findings.entries.length, 0);
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
    assert.strictEqual(findings.entries.length, 1);
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
    assert.strictEqual(findings.entries.length, 0);
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
    sinon.stub(fs, 'readFileSync').callsFake((path) => {
      if (path.toString().endsWith('@pnp/pnpjs/package.json')) {
        return JSON.stringify({
        });
      }
      else {
        return originalReadFileSync(path);
      }
    });
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 0);
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
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
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
    sinon.stub(request, 'post').callsFake(() => Promise.resolve({ scriptType: 'UMD' }));
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 0);
  });
  it('doesn\'t return anything if module type is not supported', async () => {
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
    sinon.stub(request, 'post').callsFake(() => Promise.resolve({ scriptType: 'CommonJs' }));
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 0);
  });
  it('adds missing file extension', async () => {
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
          main: "./dist/pnpjs.es5.umd.bundle",
          module: "./dist/pnpjs.es5.umd.bundle.min"
        });
      }
      else {
        return originalReadFileSync(path);
      }
    });
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake(() => Promise.resolve({ scriptType: 'UMD' }));
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 1);
  });
  it('uses exports from API', async () => {
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
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake(() => Promise.resolve({ scriptType: 'UMD', exports: ['pnpjs'] }));
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 1);
    assert.strictEqual(findings.entries[0].globalName, 'pnpjs');
  });
  it('considers all package entries', async () => {
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
          module: "./dist/pnpjs.es5.umd.bundle.min.js",
          es2015: "./dist/pnpjs.es5.umd.bundle.min.js",
          jspm: {
            main: "./dist/pnpjs.es5.umd.bundle.min.js",
            files: ["./dist/pnpjs.es5.umd.bundle.min.js"],
          },
          spm: {
            main: "./dist/pnpjs.es5.umd.bundle.min.js",
          }
        });
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake(() => Promise.resolve({ scriptType: 'UMD' }));
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 1);
  });
  it('doesnt return anything if package json is missing', async () => {
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
        throw new Error('file doesnt exist');
      }
      else {
        return originalReadFileSync(path);
      }
    });
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 0);
  });
  it('returns something for es2015 modules', async () => {
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
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake(() => Promise.resolve({ scriptType: 'ES2015' }));
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 1);
  });
  it('returns something for AMD modules', async () => {
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
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake(() => Promise.resolve({ scriptType: 'AMD' }));
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 1);
  });
});