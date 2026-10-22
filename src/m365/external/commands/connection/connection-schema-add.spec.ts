import { ExternalConnectors } from '@microsoft/microsoft-graph-types';
import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './connection-schema-add.js';

describe(commands.CONNECTION_SCHEMA_ADD, () => {
  const externalConnectionId = 'TestConnectionForCLI';
  const schema = '{"baseType": "microsoft.graph.externalItem","properties": [{"name": "ticketTitle","type": "String"}]}';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
      request.get,
      request.patch,
      global.setTimeout
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONNECTION_SCHEMA_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('creates an external connection schema without waiting for provisioning to complete', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/${externalConnectionId}/schema`) {
        return {
          headers: {
            location: 'https://graph.microsoft.com/v1.0/external/connections/fromcli/operations/1.weu-b.0251D560C889F660594C3F098392B322.98B2D9481B87CDDEBDF3DABCB52A9D22'
          }
        };
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ schema: schema, externalConnectionId: externalConnectionId, verbose: true }) });
  });

  it('creates an external connection schema and waits for provisioning to complete', async () => {
    let waitsForCompletion = false;
    let i = 0;
    sinon.stub(request, 'patch').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/${externalConnectionId}/schema`) {
        return {
          headers: {
            location: 'https://graph.microsoft.com/v1.0/external/connections/fromcli/operations/1.weu-b.0251D560C889F660594C3F098392B322.98B2D9481B87CDDEBDF3DABCB52A9D22'
          }
        };
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/external/connections/fromcli/operations/1.weu-b.0251D560C889F660594C3F098392B322.98B2D9481B87CDDEBDF3DABCB52A9D22') {
        if (i++ < 2) {
          return {
            status: 'inprogress'
          };
        }

        waitsForCompletion = true;
        return {
          status: 'succeeded'
        };
      }
      throw 'Invalid request';
    });
    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ schema: schema, externalConnectionId: externalConnectionId, verbose: true, wait: true }) });
    assert.strictEqual(waitsForCompletion, true);
  });

  it('correctly handles error when waiting for provisioning to complete and provisioning failed', async () => {
    let i = 0;
    sinon.stub(request, 'patch').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/${externalConnectionId}/schema`) {
        return {
          headers: {
            location: 'https://graph.microsoft.com/v1.0/external/connections/fromcli/operations/1.weu-b.0251D560C889F660594C3F098392B322.98B2D9481B87CDDEBDF3DABCB52A9D22'
          }
        };
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/external/connections/fromcli/operations/1.weu-b.0251D560C889F660594C3F098392B322.98B2D9481B87CDDEBDF3DABCB52A9D22') {
        if (i++ < 2) {
          return {
            status: 'inprogress'
          } as ExternalConnectors.ConnectionOperation;
        }

        return {
          status: 'failed',
          error: {
            message: 'An error has occurred'
          }
        } as ExternalConnectors.ConnectionOperation;
      }
      throw 'Invalid request';
    });
    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });
    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ schema: schema, externalConnectionId: externalConnectionId, debug: true, wait: true }) }),
      new CommandError('Provisioning schema failed: An error has occurred'));
  });

  it('correctly handles error when request is malformed or schema already exists', async () => {
    const errorMessage = 'Error: The request is malformed or incorrect.';
    sinon.stub(request, 'patch').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/${externalConnectionId}/schema`) {
        throw errorMessage;
      }
      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ schema: schema, externalConnectionId: externalConnectionId }) }),
      new CommandError(errorMessage));
  });

  it('fails validation if id is less than 3 characters', () => {
    const actual = commandOptionsSchema.safeParse({
      externalConnectionId: 'T',
      schema: schema
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if id is more than 32 characters', () => {
    const actual = commandOptionsSchema.safeParse({
      externalConnectionId: externalConnectionId + 'zzzzzzzzzzzzzzzzzz',
      schema: schema
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if id is not alphanumeric', () => {
    const actual = commandOptionsSchema.safeParse({
      externalConnectionId: externalConnectionId + '!',
      schema: schema
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if id starts with Microsoft', () => {
    const actual = commandOptionsSchema.safeParse({
      externalConnectionId: 'Microsoft' + externalConnectionId,
      schema: schema
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if schema does not contain baseType', () => {
    const actual = commandOptionsSchema.safeParse({
      externalConnectionId: externalConnectionId,
      schema: '{"properties": [{"name": "ticketTitle","type": "String"}]}'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if schema does not contain properties', () => {
    const actual = commandOptionsSchema.safeParse({
      externalConnectionId: externalConnectionId,
      schema: '{"baseType": "microsoft.graph.externalItem"}'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if schema does contain more than 128 properties', () => {
    const schemaObject = JSON.parse(schema);
    for (let i = 0; i < 128; i++) {
      schemaObject.properties.push({
        name: `Test${i}`,
        type: 'String'
      });
    }
    const actual = commandOptionsSchema.safeParse({
      externalConnectionId: externalConnectionId,
      schema: JSON.stringify(schemaObject)
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation with a correct schema and external connection id', () => {
    const actual = commandOptionsSchema.safeParse({
      externalConnectionId: externalConnectionId,
      schema: schema
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if schema is not valid JSON', () => {
    const actual = commandOptionsSchema.safeParse({
      externalConnectionId: externalConnectionId,
      schema: 'not valid json'
    });
    assert.strictEqual(actual.success, false);
  });
});