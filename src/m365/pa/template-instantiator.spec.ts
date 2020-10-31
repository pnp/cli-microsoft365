import * as assert from 'assert';
import * as fs from "fs";
import * as path from 'path';
import * as sinon from 'sinon';
import { v4 } from 'uuid';
import { Logger } from '../../cli';
import Utils from '../../Utils';
import { PcfInitVariables } from './commands/pcf/pcf-init/pcf-init-variables';
import TemplateInstantiator from './template-instantiator';

describe('TemplateInstantiator', () => {
  let log: string[];
  let logger: Logger;
  let fsMkdirSync: sinon.SinonStub;
  let fsCopyFileSync: sinon.SinonStub;
  let fsWriteFileSync: sinon.SinonStub;
  const assetsRoot = path.join(__dirname, 'commands', 'pcf', 'pcf-init', 'assets');
  const componentAssetsRoot = path.join(assetsRoot, 'control', 'field-template');
  const projectDirectory = process.cwd();
  const componentDirectory = path.join(projectDirectory, 'Example1Name');
  const variables: PcfInitVariables = {
    "$namespaceplaceholder$": "Example1.Namespace",
    "$controlnameplaceholder$": "Example1Name",
    "$pcfProjectName$": "ExampleComponentProject",
    "pcfprojecttype": "ExampleComponentProject",
    "$pcfProjectGuid$": v4()
  };

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    fsMkdirSync = sinon.stub(fs, 'mkdirSync').callsFake(() => {});
    fsCopyFileSync = sinon.stub(fs, 'copyFileSync').callsFake(() => {});
    fsWriteFileSync = sinon.stub(fs, 'writeFileSync').callsFake(() => {});
  });

  afterEach(() => {
    Utils.restore([
      fs.existsSync,
      fs.mkdirSync,
      fs.copyFileSync,
      fs.writeFileSync,
      TemplateInstantiator.mkdirSyncIfNotExists
    ]);
  });

  it('doesn\'t try to create destinationPath if it already exists', () => {
    const fsExistsSync = sinon.stub(fs, 'existsSync').callsFake(() => true);
    
    TemplateInstantiator.mkdirSyncIfNotExists(logger, componentDirectory, false);

    assert(fsExistsSync.withArgs(componentDirectory).calledOnce);
    assert(fsMkdirSync.withArgs(componentDirectory).notCalled);
  });

  it('doesn\'t try to create destinationPath if it already exists (verbose)', () => {
    const fsExistsSync = sinon.stub(fs, 'existsSync').callsFake(() => true);

    TemplateInstantiator.mkdirSyncIfNotExists(logger, componentDirectory, true);

    assert(fsExistsSync.withArgs(componentDirectory).calledOnce);
    assert(fsMkdirSync.withArgs(componentDirectory).notCalled);
  });

  it('creates destinationPath when it doesn\'t exist yet', () => {
    const fsExistsSync = sinon.stub(fs, 'existsSync').callsFake(() => false);

    TemplateInstantiator.mkdirSyncIfNotExists(logger, componentDirectory, false);

    assert(fsExistsSync.withArgs(componentDirectory).calledOnce);
    assert(fsMkdirSync.withArgs(componentDirectory).calledOnce);
  });

  it('creates destinationPath when it doesn\'t exist yet (verbose)', () => {
    const fsExistsSync = sinon.stub(fs, 'existsSync').callsFake(() => false);

    TemplateInstantiator.mkdirSyncIfNotExists(logger, componentDirectory, true);

    assert(fsExistsSync.withArgs(componentDirectory).calledOnce);
    assert(fsMkdirSync.withArgs(componentDirectory).calledOnce);
  });

  it('creates structure for shared files', () => {
    const mkdirSyncIfNotExists = sinon.stub(TemplateInstantiator, 'mkdirSyncIfNotExists').callsFake(() => {});

    TemplateInstantiator.instantiate(logger, assetsRoot, projectDirectory, false, variables, false);

    assert.strictEqual(mkdirSyncIfNotExists.callCount, 6);
    assert.strictEqual(fsCopyFileSync.callCount, 3);
    assert.strictEqual(fsWriteFileSync.callCount, 2);
  });

  it('creates structure for shared files (verbose)', () => {
    const mkdirSyncIfNotExists = sinon.stub(TemplateInstantiator, 'mkdirSyncIfNotExists').callsFake(() => {});

    TemplateInstantiator.instantiate(logger, assetsRoot, projectDirectory, false, variables, true);

    assert.strictEqual(mkdirSyncIfNotExists.callCount, 6);
    assert.strictEqual(fsCopyFileSync.callCount, 3);
    assert.strictEqual(fsWriteFileSync.callCount, 2);
  });

  it('creates structure for field component', () => {
    const mkdirSyncIfNotExists = sinon.stub(TemplateInstantiator, 'mkdirSyncIfNotExists').callsFake(() => {});

    TemplateInstantiator.instantiate(logger, componentAssetsRoot, componentDirectory, true, variables, false);

    assert.strictEqual(mkdirSyncIfNotExists.callCount, 4);
    assert.strictEqual(fsCopyFileSync.callCount, 1);
    assert.strictEqual(fsWriteFileSync.callCount, 2);
  });

  it('creates structure for field component (verbose)', () => {
    const mkdirSyncIfNotExists = sinon.stub(TemplateInstantiator, 'mkdirSyncIfNotExists').callsFake(() => {});

    TemplateInstantiator.instantiate(logger, componentAssetsRoot, componentDirectory, true, variables, true);

    assert.strictEqual(mkdirSyncIfNotExists.callCount, 4);
    assert.strictEqual(fsCopyFileSync.callCount, 1);
    assert.strictEqual(fsWriteFileSync.callCount, 2);
  });

  it('creates structure for dataset component', () => {
    const mkdirSyncIfNotExists = sinon.stub(TemplateInstantiator, 'mkdirSyncIfNotExists').callsFake(() => {});

    TemplateInstantiator.instantiate(logger, componentAssetsRoot, componentDirectory, true, variables, false);

    assert.strictEqual(mkdirSyncIfNotExists.callCount, 4);
    assert.strictEqual(fsCopyFileSync.callCount, 1);
    assert.strictEqual(fsWriteFileSync.callCount, 2);
  });

  it('creates structure for dataset component (verbose)', () => {
    const mkdirSyncIfNotExists = sinon.stub(TemplateInstantiator, 'mkdirSyncIfNotExists').callsFake(() => {});

    TemplateInstantiator.instantiate(logger, componentAssetsRoot, componentDirectory, true, variables, true);

    assert.strictEqual(mkdirSyncIfNotExists.callCount, 4);
    assert.strictEqual(fsCopyFileSync.callCount, 1);
    assert.strictEqual(fsWriteFileSync.callCount, 2);
  });
});