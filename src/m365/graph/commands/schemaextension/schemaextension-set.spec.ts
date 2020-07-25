import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./schemaextension-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.SCHEMAEXTENSION_SET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.patch
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SCHEMAEXTENSION_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates schema extension', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/schemaExtensions/ext6kguklm2_TestSchemaExtension`) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        id: 'ext6kguklm2_TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        status: 'Available',
        properties: '[{"name":"MyInt","type":"Integer"},{"name":"MyString","type":"String"}]'
      }
    }, () => {
      try {
        assert.strictEqual(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates schema extension (debug)', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/schemaExtensions/ext6kguklm2_TestSchemaExtension`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        id: 'ext6kguklm2_TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        status: 'Available',
        properties: '[{"name":"MyInt","type":"Integer"},{"name":"MyString","type":"String"}]'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith("Schema extension successfully updated."));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates schema extension (verbose)', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/schemaExtensions/ext6kguklm2_TestSchemaExtension`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        debug: false,
        id: 'ext6kguklm2_TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        status: 'Available'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error correctly', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"Integer"},{"name":"MyString","type":"String"}]'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the owner is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'invalid',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"Integer"},{"name":"MyString","type":"String"}]'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if no update information is specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if properties is not valid JSON string', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: 'foobar'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if properties have no valid type', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"Foo"},{"name":"MyString","type":"String"}]'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if a specified property has missing type', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt"},{"name":"MyString","type":"String"}]'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if a specified property has missing name', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"type":"Integer"},{"name":"MyString","type":"String"}]'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if properties JSON string is not an array', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '{}'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if status is not valid', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        status: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required parameters are set and at least one property to update (description) is specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        description: 'test',
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if the property type is Binary', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        description: null,
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"Binary"}]'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if the property type is Boolean', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        description: null,
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"Boolean"}]'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if the property type is DateTime', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        description: null,
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"DateTime"}]'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if the property type is Integer', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        description: null,
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"Integer"}]'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if the property type is String', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        description: null,
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"String"}]'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});