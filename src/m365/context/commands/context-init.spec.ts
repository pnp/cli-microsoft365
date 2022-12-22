import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import { Logger } from '../../../cli/Logger';
import Command, { CommandError } from '../../../Command';
import { sinonUtil } from '../../../utils/sinonUtil';
import commands from '../commands';
const command: Command = require('./context-init');

describe(commands.INIT, () => {
  let log: any[];
  let logger: Logger;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync
    ]);
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.INIT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('logs an error if an error occurred while reading the .m365rc.json', async () => {
    const originalFsExistsSync = fs.existsSync;
    const originalFsReadFileSync = fs.readFileSync;

    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.m365rc.json') > -1) {
        return true;
      }
      else {
        return originalFsExistsSync(path);
      }
    });
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().indexOf('.m365rc.json') > -1) {
        throw new Error('An error has occurred');
      }
      else {
        return originalFsReadFileSync(path, options);
      }
    });

    await assert.rejects(command.action(logger, { options: { verbose: true } }), new CommandError('Error reading .m365rc.json: Error: An error has occurred. Please add context info to .m365rc.json manually.'));
  });

  it(`logs an error if the .m365rc.json file contents couldn't be parsed`, async () => {
    const originalFsExistsSync = fs.existsSync;
    const originalFsReadFileSync = fs.readFileSync;

    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.m365rc.json') > -1) {
        return true;
      }
      else {
        return originalFsExistsSync(path);
      }
    });
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().indexOf('.m365rc.json') > -1) {
        return '{';
      }
      else {
        return originalFsReadFileSync(path, options);
      }
    });

    await assert.rejects(command.action(logger, { options: { verbose: true } }), new CommandError('Error reading .m365rc.json: SyntaxError: Unexpected end of JSON input. Please add context info to .m365rc.json manually.'));
  });

  it(`logs an error if the content can't be written to the .m365rc.json file`, async () => {
    const originalFsExistsSync = fs.existsSync;
    const originalFsReadFileSync = fs.readFileSync;

    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.m365rc.json') > -1) {
        return false;
      }
      else {
        return originalFsExistsSync(path);
      }
    });
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().indexOf('.m365rc.json') > -1) {
        return JSON.stringify({
          "context": {}
        });
      }
      else {
        return originalFsReadFileSync(path, options);
      }
    });
    sinon.stub(fs, 'writeFileSync').callsFake(_ => { throw new Error('An error has occurred'); });

    await assert.rejects(() => command.action(logger, { options: { verbose: true } }), new CommandError('Error writing .m365rc.json: Error: An error has occurred. Please add context info to .m365rc.json manually.'));
  });

  it(`creates the .m365rc.json file if it doesn't exist and saves context info`, async () => {
    const originalFsExistsSync = fs.existsSync;

    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.m365rc.json') > -1) {
        return false;
      }
      else {
        return originalFsExistsSync(path);
      }
    });
    const fsWriteFileSyncStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });

    await command.action(logger, { options: { verbose: true } });

    assert(fsWriteFileSyncStub.calledWith('.m365rc.json', JSON.stringify({
      context: {}
    }, null, 2)));
  });

  it(`adds the context info to the existing .m365rc.json file`, async () => {
    const originalFsExistsSync = fs.existsSync;
    const originalFsReadFileSync = fs.readFileSync;

    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.m365rc.json') > -1) {
        return false;
      }
      else {
        return originalFsExistsSync(path);
      }
    });
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().indexOf('.m365rc.json') > -1) {
        return JSON.stringify({});
      }
      else {
        return originalFsReadFileSync(path, options);
      }
    });
    const fsWriteFileSyncStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });

    await command.action(logger, { options: { verbose: true } });

    assert(fsWriteFileSyncStub.calledWith('.m365rc.json', JSON.stringify({
      context: {}
    }, null, 2)));
  });

  it(`reads context info from the .m365rc.json file`, async () => {
    const originalFsExistsSync = fs.existsSync;
    const originalFsReadFileSync = fs.readFileSync;

    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.m365rc.json') > -1) {
        return true;
      }
      else {
        return originalFsExistsSync(path);
      }
    });
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().indexOf('.m365rc.json') > -1) {
        return JSON.stringify({
          "context": {}
        });
      }
      else {
        return originalFsReadFileSync(path, options);
      }
    });
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    await command.action(logger, { options: { verbose: true } });

    assert(fsWriteFileSyncSpy.notCalled);
  });


  it(`doesn't save context info in the .m365rc.json file when there was an error reading file contents`, async () => {
    const originalFsExistsSync = fs.existsSync;
    const originalFsReadFileSync = fs.readFileSync;

    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.m365rc.json') > -1) {
        return true;
      }
      else {
        return originalFsExistsSync(path);
      }
    });
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().indexOf('.m365rc.json') > -1) {
        throw new Error('An error has occurred');
      }
      else {
        return originalFsReadFileSync(path, options);
      }
    });
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    await assert.rejects(command.action(logger, { options: { verbose: true } }), new CommandError('Error reading .m365rc.json: Error: An error has occurred. Please add context info to .m365rc.json manually.'));
    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`doesn't save context info in the .m365rc.json file when there was error writing file contents`, async () => {
    const originalFsExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.m365rc.json') > -1) {
        return false;
      }
      else {
        return originalFsExistsSync(path);
      }
    });
    sinon.stub(fs, 'writeFileSync').callsFake(_ => { throw new Error('An error has occurred'); });

    await assert.rejects(command.action(logger, { options: { verbose: true } }), new CommandError('Error writing .m365rc.json: Error: An error has occurred. Please add context info to .m365rc.json manually.'));
  });
});