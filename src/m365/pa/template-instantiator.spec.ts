import assert from 'assert';
import fs from "fs";
import path from 'path';
import sinon from 'sinon';
import url from 'url';
import { v4 } from 'uuid';
import { Logger } from '../../cli/Logger.js';
import { sinonUtil } from '../../utils/sinonUtil.js';
import { PcfInitVariables } from './commands/pcf/pcf-init/pcf-init-variables.js';
import TemplateInstantiator from './template-instantiator.js';

const __dirname = url.fileURLToPath(new URL('.', import.meta.url));

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
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    fsMkdirSync = sinon.stub(fs, 'mkdirSync').returns('');
    fsCopyFileSync = sinon.stub(fs, 'copyFileSync').returns();
    fsWriteFileSync = sinon.stub(fs, 'writeFileSync').returns();
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.mkdirSync,
      fs.copyFileSync,
      fs.writeFileSync,
      TemplateInstantiator.mkdirSyncIfNotExists
    ]);
  });

  it('doesn\'t try to create destinationPath if it already exists', () => {
    const fsExistsSync = sinon.stub(fs, 'existsSync').returns(true);

    TemplateInstantiator.mkdirSyncIfNotExists(logger, componentDirectory, false);

    assert(fsExistsSync.withArgs(componentDirectory).calledOnce);
    assert(fsMkdirSync.withArgs(componentDirectory).notCalled);
  });

  it('doesn\'t try to create destinationPath if it already exists (verbose)', () => {
    const fsExistsSync = sinon.stub(fs, 'existsSync').returns(true);

    TemplateInstantiator.mkdirSyncIfNotExists(logger, componentDirectory, true);

    assert(fsExistsSync.withArgs(componentDirectory).calledOnce);
    assert(fsMkdirSync.withArgs(componentDirectory).notCalled);
  });

  it('creates destinationPath when it doesn\'t exist yet', () => {
    const fsExistsSync = sinon.stub(fs, 'existsSync').returns(false);

    TemplateInstantiator.mkdirSyncIfNotExists(logger, componentDirectory, false);

    assert(fsExistsSync.withArgs(componentDirectory).calledOnce);
    assert(fsMkdirSync.withArgs(componentDirectory).calledOnce);
  });

  it('creates destinationPath when it doesn\'t exist yet (verbose)', async () => {
    const fsExistsSync = sinon.stub(fs, 'existsSync').returns(false);

    await TemplateInstantiator.mkdirSyncIfNotExists(logger, componentDirectory, true);

    assert(fsExistsSync.withArgs(componentDirectory).calledOnce);
    assert(fsMkdirSync.withArgs(componentDirectory).calledOnce);
  });

  it('creates structure for shared files', async () => {
    const mkdirSyncIfNotExists = sinon.stub(TemplateInstantiator, 'mkdirSyncIfNotExists').returns(Promise.resolve());

    await TemplateInstantiator.instantiate(logger, assetsRoot, projectDirectory, false, variables, false);

    assert.strictEqual(mkdirSyncIfNotExists.callCount, 6);
    assert.strictEqual(fsCopyFileSync.callCount, 3);
    assert.strictEqual(fsWriteFileSync.callCount, 2);
  });

  it('creates structure for shared files (verbose)', async () => {
    const mkdirSyncIfNotExists = sinon.stub(TemplateInstantiator, 'mkdirSyncIfNotExists').returns(Promise.resolve());

    await TemplateInstantiator.instantiate(logger, assetsRoot, projectDirectory, false, variables, true);

    assert.strictEqual(mkdirSyncIfNotExists.callCount, 6);
    assert.strictEqual(fsCopyFileSync.callCount, 3);
    assert.strictEqual(fsWriteFileSync.callCount, 2);
  });

  it('creates structure for field component', async () => {
    const mkdirSyncIfNotExists = sinon.stub(TemplateInstantiator, 'mkdirSyncIfNotExists').returns(Promise.resolve());

    await TemplateInstantiator.instantiate(logger, componentAssetsRoot, componentDirectory, true, variables, false);

    assert.strictEqual(mkdirSyncIfNotExists.callCount, 4);
    assert.strictEqual(fsCopyFileSync.callCount, 1);
    assert.strictEqual(fsWriteFileSync.callCount, 2);
  });

  it('creates structure for field component (verbose)', async () => {
    const mkdirSyncIfNotExists = sinon.stub(TemplateInstantiator, 'mkdirSyncIfNotExists').returns(Promise.resolve());

    await TemplateInstantiator.instantiate(logger, componentAssetsRoot, componentDirectory, true, variables, true);

    assert.strictEqual(mkdirSyncIfNotExists.callCount, 4);
    assert.strictEqual(fsCopyFileSync.callCount, 1);
    assert.strictEqual(fsWriteFileSync.callCount, 2);
  });

  it('creates structure for dataset component', async () => {
    const mkdirSyncIfNotExists = sinon.stub(TemplateInstantiator, 'mkdirSyncIfNotExists').returns(Promise.resolve());

    await TemplateInstantiator.instantiate(logger, componentAssetsRoot, componentDirectory, true, variables, false);

    assert.strictEqual(mkdirSyncIfNotExists.callCount, 4);
    assert.strictEqual(fsCopyFileSync.callCount, 1);
    assert.strictEqual(fsWriteFileSync.callCount, 2);
  });

  it('creates structure for dataset component (verbose)', async () => {
    const mkdirSyncIfNotExists = sinon.stub(TemplateInstantiator, 'mkdirSyncIfNotExists').returns(Promise.resolve());

    await TemplateInstantiator.instantiate(logger, componentAssetsRoot, componentDirectory, true, variables, true);

    assert.strictEqual(mkdirSyncIfNotExists.callCount, 4);
    assert.strictEqual(fsCopyFileSync.callCount, 1);
    assert.strictEqual(fsWriteFileSync.callCount, 2);
  });
});
