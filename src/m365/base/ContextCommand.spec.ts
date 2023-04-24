import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { CommandError } from '../../Command';
import { telemetry } from '../../telemetry';
import { sinonUtil } from '../../utils/sinonUtil';
import { Hash } from '../../utils/types';
import ContextCommand from './ContextCommand';

class MockCommand extends ContextCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public mockSaveContextInfo(contextInfo: Hash) {
    this.saveContextInfo(contextInfo);
  }

  public async commandAction(): Promise<void> {
  }

  public commandHelp(): void {
  }
}

describe('ContextCommand', () => {
  let cmd: MockCommand;
  const contextInfo: Hash = {};

  before(() => {
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
  });

  beforeEach(() => {
    cmd = new MockCommand();
  });

  afterEach(() => {
    sinonUtil.restore([
      telemetry.trackEvent,
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync
    ]);
  });

  it('logs an error if an error occurred while reading the .m365rc.json', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => { throw new Error('An error has occurred'); });

    assert.throws(() => cmd.mockSaveContextInfo(contextInfo), new CommandError('Error reading .m365rc.json: Error: An error has occurred. Please add context info to .m365rc.json manually.'));
  });

  it(`logs an error if the .m365rc.json file contents couldn't be parsed`, () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '{');

    let errorMessage;
    try {
      JSON.parse('{');
    }
    catch (err: any) {
      errorMessage = err;
    }

    assert.throws(() => cmd.mockSaveContextInfo(contextInfo), new CommandError(`Error reading .m365rc.json: ${errorMessage}. Please add context info to .m365rc.json manually.`));
  });

  it(`logs an error if the content can't be written to the .m365rc.json file`, () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => false);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "context": {}
    }));
    sinon.stub(fs, 'writeFileSync').callsFake(_ => { throw new Error('An error has occurred'); });

    assert.throws(() => cmd.mockSaveContextInfo(contextInfo), new CommandError('Error writing .m365rc.json: Error: An error has occurred. Please add context info to .m365rc.json manually.'));
  });

  it(`creates the .m365rc.json file if it doesn't exist and saves context info`, () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;

    sinon.stub(fs, 'existsSync').callsFake(_ => false);
    sinon.stub(fs, 'writeFileSync').callsFake((_, contents) => {
      filePath = _.toString();
      fileContents = contents as string;
    });

    cmd.mockSaveContextInfo(contextInfo);

    assert.strictEqual(filePath, '.m365rc.json');
    assert.strictEqual(fileContents, JSON.stringify({
      context: {}
    }, null, 2));
  });

  it(`adds the context info to the existing .m365rc.json file`, () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;

    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({}));
    sinon.stub(fs, 'writeFileSync').callsFake((_, contents) => {
      filePath = _.toString();
      fileContents = contents as string;
    });

    cmd.mockSaveContextInfo(contextInfo);

    assert.strictEqual(filePath, '.m365rc.json');
    assert.strictEqual(fileContents, JSON.stringify({
      context: {}
    }, null, 2));
  });

  it(`doesn't initiate context when it's already present`, () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "context": {}
    }));
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    cmd.mockSaveContextInfo(contextInfo);

    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`doesn't save context info in the .m365rc.json file when there was an error reading file contents`, () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => { throw new Error(); });
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    assert.throws(() => cmd.mockSaveContextInfo(contextInfo), new CommandError('Error reading .m365rc.json: Error. Please add context info to .m365rc.json manually.'));
    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`doesn't save context info in the .m365rc.json file when there was error writing file contents`, () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => false);
    sinon.stub(fs, 'writeFileSync').callsFake(_ => { throw new Error('An error has occurred'); });

    assert.throws(() => cmd.mockSaveContextInfo(contextInfo), new CommandError('Error writing .m365rc.json: Error: An error has occurred. Please add context info to .m365rc.json manually.'));
  });
});