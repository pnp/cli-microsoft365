import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './option-set.js';

describe(commands.OPTION_SET, () => {
  let log: any[];
  let logger: Logger;

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
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.OPTION_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('handles an error when reading file contents fails', async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => { throw new Error('An error has occurred'); });

    await assert.rejects(command.action(logger, { options: { debug: true, name: 'listName', value: 'testList' } }), new CommandError(`Error reading .m365rc.json: Error: An error has occurred. Please add listName to .m365rc.json manually.`));
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
      "context": {}
    }));
    sinon.stub(fs, 'writeFileSync').callsFake(_ => { throw new Error('An error has occurred'); });

    await assert.rejects(command.action(logger, { options: { debug: true, name: 'listName', value: 'testList' } }), new CommandError(`Error writing .m365rc.json: Error: An error has occurred. Please add listName to .m365rc.json manually.`));
  });

  it('adds a new key with value when context is present', async () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;

    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({}));
    sinon.stub(fs, 'writeFileSync').callsFake((_, contents) => {
      filePath = _.toString();
      fileContents = contents as string;
    });

    await command.action(logger, { options: { verbose: true, name: 'listName', value: 'testList' } });
    assert.strictEqual(filePath, '.m365rc.json');
    assert.strictEqual(fileContents, JSON.stringify({
      context: { listName: 'testList' }
    }, null, 2));
  });

  it('adds a new key with value when no context is present', async () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;

    sinon.stub(fs, 'existsSync').callsFake(_ => false);
    sinon.stub(fs, 'writeFileSync').callsFake((_, contents) => {
      filePath = _.toString();
      fileContents = contents as string;
    });
    await assert.doesNotReject(command.action(logger, { options: { debug: true, name: 'listName', value: 'testList' } }));
    assert.strictEqual(filePath, '.m365rc.json');
    assert.strictEqual(fileContents, JSON.stringify({
      context: { listName: 'testList' }
    }, null, 2));
  });

  it('Updates an existing key with a new value', async () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;

    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "apps": [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        }
      ],
      "context": {
        listName: "oldListName"
      }
    }));
    sinon.stub(fs, 'writeFileSync').callsFake((_, contents) => {
      filePath = _.toString();
      fileContents = contents as string;
    });

    await command.action(logger, { options: { verbose: true, name: 'listName', value: 'testList' } });
    assert.strictEqual(filePath, '.m365rc.json');
    assert.strictEqual(fileContents, JSON.stringify({
      "apps": [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        }
      ],
      context: { listName: 'testList' }
    }, null, 2));
  });
});