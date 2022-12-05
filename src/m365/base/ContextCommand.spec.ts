import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../appInsights';
import { Logger } from '../../cli/Logger';
import { CommandError } from '../../Command';
import ContextCommand from './ContextCommand';
import { sinonUtil } from '../../utils/sinonUtil';
import * as fs from 'fs';
import { Hash } from '../../utils/types';

class MockCommand extends ContextCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public mockSaveContextInfo(contextInfo: Hash, logger: Logger) {
    this.saveContextInfo(contextInfo, logger);
  }

  public async commandAction(): Promise<void> {
  }

  public commandHelp(): void {
  }
}

describe('ContextCommand', () => {
  let cmd: MockCommand;
  let log: any[];
  let logger: Logger;
  const contextInfo: Hash = {};
  let loggerLogSpy: sinon.SinonSpy;


  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
  });

  beforeEach(() => {
    cmd = new MockCommand();
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
    loggerLogSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      appInsights.trackEvent,
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync
    ]);
  });

  it('logs an error if an error occured while reading the .m365rc.json', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => { throw new Error('An error has occurred'); });
    cmd.mockSaveContextInfo(contextInfo, logger);
    assert.strictEqual(loggerLogSpy.lastCall.args[0], 'Error reading .m365rc.json: Error: An error has occurred. Please add context info to .m365rc.json manually.');
  });

  it(`logs an error if the .m365rc.json file contents couldn't be parsed`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '{');
    cmd.mockSaveContextInfo(contextInfo, logger);
    assert.strictEqual(loggerLogSpy.lastCall.args[0], 'Error reading .m365rc.json: SyntaxError: Unexpected end of JSON input. Please add context info to .m365rc.json manually.');
  });

  it(`logs an error if the content can't be written to the .m365rc.json file`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => false);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "context": {}
    }));
    sinon.stub(fs, 'writeFileSync').callsFake(_ => { throw new Error('An error has occurred'); });
    cmd.mockSaveContextInfo(contextInfo, logger);
    assert.strictEqual(loggerLogSpy.lastCall.args[0], 'Error writing .m365rc.json: Error: An error has occurred. Please add context info to .m365rc.json manually.');
  });


  it(`creates the .m365rc.json file if it doesn't exist and saves context info`, async () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;

    sinon.stub(fs, 'existsSync').callsFake(_ => false);
    sinon.stub(fs, 'writeFileSync').callsFake((_, contents) => {
      filePath = _.toString();
      fileContents = contents as string;
    });

    cmd.mockSaveContextInfo(contextInfo, logger);

    assert.strictEqual(filePath, '.m365rc.json');
    assert.strictEqual(fileContents, JSON.stringify({
      context: {}
    }, null, 2));
  });

  it(`adds the context info to the existing .m365rc.json file`, async () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;

    sinon.stub(fs, 'existsSync').callsFake(_ => false);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "context": {}
    }));
    sinon.stub(fs, 'writeFileSync').callsFake((_, contents) => {
      filePath = _.toString();
      fileContents = contents as string;
    });

    cmd.mockSaveContextInfo(contextInfo, logger);

    assert.strictEqual(filePath, '.m365rc.json');
    assert.strictEqual(fileContents, JSON.stringify({
      context: {}
    }, null, 2));
  });

  it(`reads context info from the .m365rc.json file`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "context": {}
    }));
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    cmd.mockSaveContextInfo(contextInfo, logger);

    assert(fsWriteFileSyncSpy.notCalled);
  });


  it(`doesn't save context info in the .m365rc.json file when there was an error reading file contents`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => { throw new Error('An error has occurred'); });
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    await cmd.mockSaveContextInfo(contextInfo, logger);
    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`doesn't save context info in the .m365rc.json file when there was error writing file contents`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => false);
    sinon.stub(fs, 'writeFileSync').callsFake(_ => { throw new Error('An error has occurred'); });

    await cmd.mockSaveContextInfo(contextInfo, logger), new CommandError('An error has occurred');
  });
});