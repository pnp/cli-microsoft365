import * as assert from 'assert';
import * as fs from 'fs';
import { Cli } from '../../cli/Cli';
import { CommandInfo } from '../../cli/CommandInfo';
import { Logger } from '../../cli/Logger';
import { sinonUtil } from '../../utils/sinonUtil';
import AppCommand from './AppCommand';
import sinon = require('sinon');
import Command, { CommandError } from '../../Command';

class MockCommand extends AppCommand {
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

describe('AppCommand', () => {
  let cmd: MockCommand;
  let logger: Logger;
  let log: string[];
  let commandInfo: CommandInfo;

  before(() => {
    commandInfo = Cli.getCommandInfo(new MockCommand());
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
      fs.existsSync,
      fs.readFileSync,
      Cli.prompt
    ]);
  });

  it('defines correct resource', () => {
    assert.strictEqual((cmd as any).resource, 'https://graph.microsoft.com');
  });

  it('returns error if .m365rc.json file not found in the current directory', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    await assert.rejects(cmd.action(logger, { options: {} }), new CommandError('Could not find file: .m365rc.json'));
  });

  it('returns error if the .m365rc.json file is empty', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '');
    await assert.rejects(cmd.action(logger, { options: {} }), new CommandError('File .m365rc.json is empty'));
  });

  it(`returns error if the .m365rc.json file contents couldn't be parsed`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '{');
    await assert.rejects(cmd.action(logger, { options: {} }), new CommandError('Could not parse file: .m365rc.json'));
  });

  it(`returns error if the .m365rc.json file is empty doesn't contain any apps`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      apps: []
    }));
    await assert.rejects(cmd.action(logger, { options: {} }), new CommandError('No Azure AD apps found in .m365rc.json'));
  });

  it(`returns error if the specified appId not found in the .m365rc.json file`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      apps: [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        }
      ]
    }));
    await assert.rejects(cmd.action(logger, { options: { appId: 'e23d235c-fcdf-45d1-ac5f-24ab2ee06951' } }),
      new CommandError('App e23d235c-fcdf-45d1-ac5f-24ab2ee06951 not found in .m365rc.json'));
  });

  it(`prompts to choose an app when multiple apps found in .m365rc.json and no appId specified`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      apps: [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        },
        {
          "appId": "9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d",
          "name": "CLI app1"
        }
      ]
    }));
    const cliPromptStub = sinon.stub(Cli, 'prompt').callsFake(async () => (
      { appIdIndex: 0 }
    ));
    await assert.rejects(cmd.action(logger, { options: {} }));
    assert(cliPromptStub.called);
  });

  it(`uses app selected by the user in the prompt`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      apps: [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        },
        {
          "appId": "9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d",
          "name": "CLI app1"
        }
      ]
    }));
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { appIdIndex: 1 }
    ));
    sinon.stub(Command.prototype, 'action').callsFake(async () => Promise.resolve());

    try {
      await cmd.action(logger, { options: {} });
      assert.strictEqual((cmd as any).appId, '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d');
    }
    finally {
      sinonUtil.restore(Command.prototype.action);
    }
  });

  it(`uses app specified in the appId command option`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      apps: [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        },
        {
          "appId": "9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d",
          "name": "CLI app1"
        }
      ]
    }));
    await assert.rejects(cmd.action(logger, { options: { appId: '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d' } }));
    assert.strictEqual((cmd as any).appId, '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d');
  });

  it(`uses app from the .m365rc.json if only one app defined`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      apps: [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        }
      ]
    }));
    await assert.rejects(cmd.action(logger, { options: {} }));
    assert.strictEqual((cmd as any).appId, 'e23d235c-fcdf-45d1-ac5f-24ab2ee0695d');
  });

  it('fails validation if the specified appId is not a valid GUID', async () => {
    const actual = await cmd.validate({ options: { appId: 'e23d235c-fcdf-45d1-ac5f-24ab2ee0695' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the specified appId is a valid GUID', async () => {
    const actual = await cmd.validate({ options: { appId: 'e23d235c-fcdf-45d1-ac5f-24ab2ee0695d' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});