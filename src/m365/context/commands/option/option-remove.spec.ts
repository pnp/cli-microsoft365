import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './option-remove.js';

describe(commands.OPTION_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let promptOptions: any;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
  });

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
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.OPTION_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing the context option from the .m365rc.json file when confirm option not passed', async () => {
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

  it('handles an error when reading file contents fails', async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => { throw new Error('An error has occurred'); });

    await assert.rejects(command.action(logger, { options: { debug: true, name: 'listName', force: true } }), new CommandError(`Error reading .m365rc.json: Error: An error has occurred. Please remove context option listName from .m365rc.json manually.`));
  });

  it('handles an error when writing file contents fails', async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "apps": [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        }
      ],
      "context": {
        "listName": "listNameValue"
      }
    }));
    sinon.stub(fs, 'writeFileSync').callsFake(_ => { throw new Error('An error has occurred'); });

    await assert.rejects(command.action(logger, { options: { debug: true, name: 'listName', force: true } }), new CommandError(`Error writing .m365rc.json: Error: An error has occurred. Please remove context option listName from .m365rc.json manually.`));
  });

  it(`removes a context info option from the existing .m365rc.json file`, async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "apps": [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        }
      ],
      "context": {
        "listName": "listNameValue"
      }
    }));
    sinon.stub(fs, 'writeFileSync').callsFake(_ => { });

    await assert.doesNotReject(command.action(logger, { options: { debug: true, name: 'listName' } }));
  });

  it(`removes a context info option from the existing .m365rc.json file without prompt`, async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "apps": [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        }
      ],
      "context": {
        "listName": "listNameValue"
      }
    }));
    sinon.stub(fs, 'writeFileSync').callsFake(_ => { });

    await assert.doesNotReject(command.action(logger, { options: { debug: true, name: 'listName', force: true } }));
  });

  it('handles an error when option is not present in the context', async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "apps": [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        }
      ],
      "context": {
        "listId": "5"
      }
    }));

    await assert.rejects(command.action(logger, { options: { debug: true, name: 'listName', force: true } }), new CommandError(`There is no option listName in the context info`));
  });

  it('handles an error when context is not present in the .m365rc.json file', async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "apps": [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        }
      ]
    }));

    await assert.rejects(command.action(logger, { options: { debug: true, name: 'listName', force: true } }), new CommandError(`There is no option listName in the context info`));
  });

});