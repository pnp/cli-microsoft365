import * as assert from 'assert';
import * as sinon from 'sinon';
import commands from '../commands';
import { Logger } from '../../../cli/Logger';
import { sinonUtil } from '../../../utils/sinonUtil';
import appInsights from '../../../appInsights';
import Command from '../../../Command';
import * as fs from 'fs';
import { Cli } from '../../../cli/Cli';
const command: Command = require('./context-remove');

describe(commands.REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let promptOptions: any;

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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      appInsights.trackEvent,
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync,
      fs.unlinkSync,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing the context from the .m365rc.json file when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        debug: false
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it(`removes the .m365rc.json file if it exist and only consists of the context parameter`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "context": {}
    }));
    const unlinkSyncStub = sinon.stub(fs, 'unlinkSync').callsFake(_ => { });
    await command.action(logger, { options: { debug: true, confirm: true } });

    assert(unlinkSyncStub.called);
  });

  it(`removes the context info from the existing .m365rc.json file`, async () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "apps": [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        }
      ],
      "context": {}
    }));
    sinon.stub(fs, 'writeFileSync').callsFake((_, contents) => {
      filePath = _.toString();
      fileContents = contents as string;
    });

    await command.action(logger, { options: { debug: true } });

    assert.strictEqual(filePath, '.m365rc.json');
    assert.strictEqual(fileContents, JSON.stringify({
      apps: [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        }
      ]
    }, null, 2));
  });

  it(`doesn't save context info in the .m365rc.json file when there was an error reading file contents`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => { throw new Error('An error has occurred'); });
    const fsWriteFileSyncSpy = sinon.spy(fs, 'unlinkSync');

    await command.action(logger, { options: { debug: true, confirm: true } });
    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`doesn't save context info in the .m365rc.json file when there was error writing file contents`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "apps": [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        }
      ],
      "context": {}
    }));
    sinon.stub(fs, 'writeFileSync').callsFake(_ => { throw new Error('An error has occurred'); });
    command.action(logger, { options: { debug: true, confirm: true } });
  });

  it(`doesn't remove the .m365rc.json file when there was error removing the file`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "context": {}
    }));
    sinon.stub(fs, 'unlinkSync').callsFake(_ => { throw new Error('An error has occurred'); });
    command.action(logger, { options: { debug: true, confirm: true } });
  });

  it(`doesn't remove context info from the .m365rc.json file when there was no context`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      apps: [{
        appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
        name: 'My AAD app'
      }]
    }));
    const fsWriteFileSyncSpy = sinon.spy(fs, 'unlinkSync');

    await command.action(logger, { options: { debug: true, confirm: true } });
    assert(fsWriteFileSyncSpy.notCalled);
  });

});