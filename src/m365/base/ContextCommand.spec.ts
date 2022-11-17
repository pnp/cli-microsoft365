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
  });

  afterEach(() => {
    sinonUtil.restore([
      appInsights.trackEvent,
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync
    ]);
  });

  it(`creates the .m365rc.json file if it doesn't exist and saves context info`, async () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;

    sinon.stub(fs, 'existsSync').callsFake(_ => false);
    sinon.stub(fs, 'writeFileSync').callsFake((_, contents) => {
      filePath = _.toString();
      fileContents = contents as string;
    });

    await cmd.saveContextInfo(contextInfo, logger);

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

    await cmd.saveContextInfo(contextInfo, logger);

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

    await cmd.saveContextInfo(contextInfo, logger);

    assert(fsWriteFileSyncSpy.notCalled);
  });


  it(`doesn't save context info in the .m365rc.json file when there was an error reading file contents`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => { throw new Error('An error has occurred'); });
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    await cmd.saveContextInfo(contextInfo, logger);
    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`doesn't save context info in the .m365rc.json file when there was error writing file contents`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => false);
    sinon.stub(fs, 'writeFileSync').callsFake(_ => { throw new Error('An error has occurred'); });

    await cmd.saveContextInfo(contextInfo, logger), new CommandError('An error has occurred');
  });
});