import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { telemetry } from '../../../../telemetry';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./externalconnection-schema-add');

describe(commands.EXTERNALCONNECTION_SCHEMA_ADD, () => {
  const externalConnectionId = 'TestConnectionForCLI';
  const schema = '{"baseType": "microsoft.graph.externalItem","properties": [{"name": "ticketTitle","type": "String"}]}';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.EXTERNALCONNECTION_SCHEMA_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds an external connection schema', async () => {
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/${externalConnectionId}/schema`) {
        return;
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: { schema: schema, externalConnectionId: externalConnectionId, verbose: true } } as any);
  });

  it('correctly handles error when request is malformed or schema already exists', async () => {
    const errorMessage = 'Error: The request is malformed or incorrect.';
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/${externalConnectionId}/schema`) {
        throw errorMessage;
      }
      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: { schema: schema, externalConnectionId: externalConnectionId } } as any),
      new CommandError(errorMessage));
  });

  it('fails validation if id is less than 3 characters', async () => {
    const actual = await command.validate({
      options: {
        externalConnectionId: 'T',
        schema: schema
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('fails validation if id is more than 32 characters', async () => {
    const actual = await command.validate({
      options: {
        externalConnectionId: externalConnectionId + 'zzzzzzzzzzzzzzzzzz',
        schema: schema
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('fails validation if id is not alphanumeric', async () => {
    const actual = await command.validate({
      options: {
        externalConnectionId: externalConnectionId + '!',
        schema: schema
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('fails validation if id starts with Microsoft', async () => {
    const actual = await command.validate({
      options: {
        externalConnectionId: 'Microsoft' + externalConnectionId,
        schema: schema
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('fails validation if schema does not contain baseType', async () => {
    const actual = await command.validate({
      options: {
        externalConnectionId: externalConnectionId,
        schema: '{"properties": [{"name": "ticketTitle","type": "String"}]}'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('fails validation if schema does not contain properties', async () => {
    const actual = await command.validate({
      options: {
        externalConnectionId: externalConnectionId,
        schema: '{"baseType": "microsoft.graph.externalItem"}'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('fails validation if schema does contain more than 128 properties', async () => {
    const schemaObject = JSON.parse(schema);
    for (let i = 0; i < 128; i++) {
      schemaObject.properties.push({
        name: `Test${i}`,
        type: 'String'
      });
    }
    const actual = await command.validate({
      options: {
        externalConnectionId: externalConnectionId,
        schema: JSON.stringify(schemaObject)
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('passes validation with a correct schema and external connection id', async () => {
    const actual = await command.validate({
      options: {
        externalConnectionId: externalConnectionId,
        schema: schema
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});