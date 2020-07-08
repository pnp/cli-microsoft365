import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./schemaextension-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.SCHEMAEXTENSION_SET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
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
      vorpal.find,
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
    assert.equal(command.name.startsWith(commands.SCHEMAEXTENSION_SET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
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
        assert.equal(log.length, 0);
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
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
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
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the id is not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: null
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if owner is not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension'
      }
    });
    assert.notEqual(actual, true);
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
    assert.notEqual(actual, true);
  });

  it('fails validation if no update information is specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'TestSchemaExtension',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0'
      }
    });
    assert.notEqual(actual, true);
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
    assert.notEqual(actual, true);
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
    assert.notEqual(actual, true);
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
    assert.notEqual(actual, true);
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
    assert.notEqual(actual, true);
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
    assert.notEqual(actual, true);
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
    assert.notEqual(actual, true);
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
    assert.equal(actual, true);
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
    assert.equal(actual, true);
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
    assert.equal(actual, true);
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
    assert.equal(actual, true);
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
    assert.equal(actual, true);
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
    assert.equal(actual, true);
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.SCHEMAEXTENSION_SET));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});