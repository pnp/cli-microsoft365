import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './schemaextension-set.js';

describe(commands.SCHEMAEXTENSION_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandOptionsSchema = command.getSchemaToParse() as typeof options;
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SCHEMAEXTENSION_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates schema extension', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/schemaExtensions/ext6kguklm2_TestSchemaExtension`) {
        return {
          "id": "ext6kguklm2_TestSchemaExtension",
          "description": "Test Description",
          "targetTypes": [
            "Group"
          ],
          "status": "InDevelopment",
          "owner": "b07a45b3-f7b7-489b-9269-da6f3f93dff0",
          "properties": [
            {
              "name": "MyInt",
              "type": "Integer"
            },
            {
              "name": "MyString",
              "type": "String"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: 'ext6kguklm2_TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        status: 'Available',
        properties: '[{"name":"MyInt","type":"Integer"},{"name":"MyString","type":"String"}]'
      })
    });
    assert.strictEqual(log.length, 0);
  });

  it('updates schema extension (debug)', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/schemaExtensions/ext6kguklm2_TestSchemaExtension`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        id: 'ext6kguklm2_TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        status: 'Available',
        properties: '[{"name":"MyInt","type":"Integer"},{"name":"MyString","type":"String"}]'
      })
    });
    assert(loggerLogToStderrSpy.calledWith("Schema extension successfully updated."));
  });

  it('updates schema extension (verbose)', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/schemaExtensions/ext6kguklm2_TestSchemaExtension`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: true,
        id: 'ext6kguklm2_TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        status: 'Available'
      })
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'patch').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"Integer"},{"name":"MyString","type":"String"}]'
      })
    }), new CommandError('An error has occurred'));
  });

  it('fails validation if the owner is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'TestSchemaExtension',
      description: 'Test Description',
      owner: 'invalid',
      targetTypes: 'Group',
      properties: '[{"name":"MyInt","type":"Integer"},{"name":"MyString","type":"String"}]'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if no update information is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'TestSchemaExtension',
      owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if properties is not valid JSON string', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'TestSchemaExtension',
      description: 'Test Description',
      owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
      targetTypes: 'Group',
      properties: 'foobar'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if properties have no valid type', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'TestSchemaExtension',
      description: 'Test Description',
      owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
      targetTypes: 'Group',
      properties: '[{"name":"MyInt","type":"Foo"},{"name":"MyString","type":"String"}]'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if a specified property has missing type', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'TestSchemaExtension',
      description: 'Test Description',
      owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
      targetTypes: 'Group',
      properties: '[{"name":"MyInt"},{"name":"MyString","type":"String"}]'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if a specified property has missing name', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'TestSchemaExtension',
      description: 'Test Description',
      owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
      targetTypes: 'Group',
      properties: '[{"type":"Integer"},{"name":"MyString","type":"String"}]'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if properties JSON string is not an array', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'TestSchemaExtension',
      description: 'Test Description',
      owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
      targetTypes: 'Group',
      properties: '{}'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if status is not valid', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'TestSchemaExtension',
      description: 'Test Description',
      owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
      status: 'invalid'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if required parameters are set and at least one property to update (description) is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'TestSchemaExtension',
      owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
      description: 'test'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if the property type is Binary', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'TestSchemaExtension',

      owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
      targetTypes: 'Group',
      properties: '[{"name":"MyInt","type":"Binary"}]'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if the property type is Boolean', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'TestSchemaExtension',

      owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
      targetTypes: 'Group',
      properties: '[{"name":"MyInt","type":"Boolean"}]'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if the property type is DateTime', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'TestSchemaExtension',

      owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
      targetTypes: 'Group',
      properties: '[{"name":"MyInt","type":"DateTime"}]'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if the property type is Integer', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'TestSchemaExtension',

      owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
      targetTypes: 'Group',
      properties: '[{"name":"MyInt","type":"Integer"}]'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if the property type is String', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'TestSchemaExtension',

      owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
      targetTypes: 'Group',
      properties: '[{"name":"MyInt","type":"String"}]'
    });
    assert.strictEqual(actual.success, true);
  });
});
